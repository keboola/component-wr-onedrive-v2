"""Microsoft Graph HTTP client for keboola.wr-onedrive-v2.

Design spec: ``docs/superpowers/specs/2026-08-17-wr-onedrive-v2-design.md`` §4 "Mechanics of the
in-scope surface" (rate limits) and §6 "Code architecture" (retries, error mapping).

Everything talks to Graph through a single ``requests.Session``-based client with **one** retry
strategy (not the three overlapping layers the sibling extractor had): 429/503 honor
``Retry-After`` (falling back to exponential backoff when the header is absent), 500/502/504
always use exponential backoff, and 405/409 ("workbook transient" — Excel's
``EditModeCannotAcquireLockTooManyRequests``-style codes) are retried only when the caller opts
in per call. A total-wait cap bounds how long any single logical request will keep retrying;
once it would be exceeded, :class:`~client.exceptions.GraphRateLimitCapExceededError` is raised
rather than retrying forever. All durations here are in seconds throughout — the extractor's
seconds-vs-milliseconds bug is not repeated.
"""

import logging
import time
from collections.abc import Callable, Iterator
from importlib import metadata as importlib_metadata

import requests

from client.auth import TokenProvider
from client.exceptions import (
    GraphBadRequestError,
    GraphClientError,
    GraphNotFoundError,
    GraphPermissionError,
    GraphQuotaExceededError,
    GraphRateLimitCapExceededError,
)

logger = logging.getLogger(__name__)

BASE_URL = "https://graph.microsoft.com/v1.0"

# Statuses that honor a `Retry-After` header (falling back to exponential backoff when absent).
_RATE_LIMIT_STATUSES = frozenset({429, 503})
# Statuses that always use exponential backoff (Graph does not send `Retry-After` for these).
_SERVER_ERROR_STATUSES = frozenset({500, 502, 504})
# Statuses retried only when the caller opts in per call (Excel workbook-lock transients).
_WORKBOOK_TRANSIENT_STATUSES = frozenset({405, 409})

# Exponential backoff when no `Retry-After` header is present: `min(BACKOFF_BASE * 2**attempt,
# BACKOFF_MAX)` seconds, attempt starting at 0 for the first retry.
_BACKOFF_BASE_SECONDS = 1.0
_BACKOFF_MAX_SECONDS = 60.0

# Total time (seconds) a single logical request is allowed to spend sleeping between retries
# before giving up with a user-facing error. Overridable per `GraphClient` instance.
DEFAULT_TOTAL_WAIT_CAP_SECONDS = 300.0

DEFAULT_TIMEOUT_SECONDS = 60.0


def _resolve_user_agent() -> str:
    try:
        version = importlib_metadata.version("wr-onedrive-v2")
    except importlib_metadata.PackageNotFoundError:
        version = "0.0.0"
    return f"NONISV|Keboola|wr-onedrive-v2/{version}"


USER_AGENT = _resolve_user_agent()


def _parse_graph_error(response: requests.Response) -> tuple[str | None, str]:
    """Extract ``error.code``/``error.message`` from a Graph error body, if present.

    Falls back to the raw response text when the body isn't the documented Graph error JSON
    shape (``{"error": {"code": ..., "message": ...}}``).
    """
    try:
        body = response.json()
    except ValueError:
        return None, response.text or f"HTTP {response.status_code}"

    error = body.get("error") if isinstance(body, dict) else None
    if not isinstance(error, dict):
        return None, response.text or f"HTTP {response.status_code}"

    code = error.get("code")
    message = error.get("message") or response.text or f"HTTP {response.status_code}"
    return code, message


def _raise_for_error_response(response: requests.Response) -> None:
    """Map a non-retried (or retry-exhausted) error response to a typed exception and raise it."""
    status = response.status_code
    error_code, message = _parse_graph_error(response)
    detail = f"{message} (Graph error code: {error_code})" if error_code else message

    if status in (401, 403):
        raise GraphPermissionError(detail, status_code=status, error_code=error_code)
    if status == 404:
        raise GraphNotFoundError(detail, status_code=status, error_code=error_code)
    if status == 507:
        raise GraphQuotaExceededError(detail, status_code=status, error_code=error_code)
    if status == 400:
        raise GraphBadRequestError(detail, status_code=status, error_code=error_code)
    raise GraphClientError(detail, status_code=status, error_code=error_code)


class GraphClient:
    """Thin ``requests.Session`` wrapper implementing Graph's auth, retry, and paging rules.

    A fresh access token is fetched from the injected :class:`TokenProvider` on *every* request
    (not cached client-side) so the provider's own proactive-refresh logic is always in the
    driver's seat — the client never needs to know whether a refresh actually happened.
    """

    def __init__(
        self,
        token_provider: TokenProvider,
        session: requests.Session | None = None,
        base_url: str = BASE_URL,
        total_wait_cap_seconds: float = DEFAULT_TOTAL_WAIT_CAP_SECONDS,
        sleep_fn: Callable[[float], None] | None = None,
    ):
        """
        Args:
            token_provider: supplies the bearer token for every request (and handles refresh).
            session: injectable ``requests.Session`` (tests only); a fresh one is created
                otherwise.
            base_url: Graph API root; overridable for tests.
            total_wait_cap_seconds: maximum cumulative sleep time across retries for a single
                logical request before giving up with :class:`GraphRateLimitCapExceededError`.
            sleep_fn: injectable sleep function (tests only). When ``None``, ``time.sleep`` is
                resolved dynamically on each call, so ``unittest.mock.patch("time.sleep")`` (or
                patching ``client.graph_client.time.sleep``) works without needing this argument.
        """
        self._token_provider = token_provider
        self._session = session if session is not None else requests.Session()
        self._base_url = base_url.rstrip("/")
        self._total_wait_cap_seconds = total_wait_cap_seconds
        self._sleep_fn = sleep_fn

    def request(
        self,
        method: str,
        url: str,
        *,
        json: dict | list | None = None,
        params: dict | None = None,
        headers: dict | None = None,
        data=None,
        absolute: bool = False,
        auth: bool = True,
        retry: bool = True,
        retry_transient_workbook: bool = False,
        stream: bool = False,
        timeout: float | None = DEFAULT_TIMEOUT_SECONDS,
    ) -> requests.Response:
        """Send a request to Graph, applying auth, retries, and error mapping.

        Args:
            method: HTTP verb (``"GET"``, ``"POST"``, ...).
            url: a path relative to ``base_url`` (e.g. ``"/me"``), or a full URL when
                ``absolute=True`` (needed for ``@odata.nextLink`` pages and pre-signed
                ``uploadUrl``s, which must not be prefixed with the Graph base URL).
            auth: attach an ``Authorization: Bearer`` header fetched from the
                ``TokenProvider``. Set ``False`` for pre-signed URLs (e.g. upload-session chunk
                PUTs) which reject/ignore the header entirely.
            retry: when ``False``, send the request exactly once and raise immediately on any
                error response — no 401 re-fetch, no backoff, no `Retry-After`. Required for
                non-idempotent upload-session chunk PUTs, whose resume logic (a later task) is
                itself the retry strategy.
            retry_transient_workbook: opt-in retry of 405/409 ("workbook transient" — Excel
                session/lock contention), used by the Excel writer (a later task).
            stream: passed through to ``requests`` for large response bodies.
            timeout: passed through to ``requests``; ``None`` disables the timeout.

        Returns:
            The successful (``2xx``) ``requests.Response``.

        Raises:
            client.exceptions.GraphClientError (or a typed subclass): on any error response
                that isn't retried away.
            client.exceptions.GraphRateLimitCapExceededError: when honoring `Retry-After` or the
                next backoff step would exceed ``total_wait_cap_seconds``.
        """
        full_url = url if absolute else self._join_url(url)
        elapsed_wait = 0.0
        attempt = 0
        retried_401 = False

        while True:
            request_headers = self._build_headers(headers, auth=auth)
            response = self._session.request(
                method,
                full_url,
                json=json,
                params=params,
                headers=request_headers,
                data=data,
                stream=stream,
                timeout=timeout,
            )
            if response.ok:
                return response

            status = response.status_code

            if not retry:
                _raise_for_error_response(response)

            if status == 401 and auth and not retried_401:
                logger.info("Graph request returned 401; re-fetching the access token and retrying once.")
                retried_401 = True
                continue

            if self._is_retryable(status, retry_transient_workbook):
                wait_seconds = self._compute_wait_seconds(response, attempt)
                remaining = self._total_wait_cap_seconds - elapsed_wait
                if wait_seconds > remaining:
                    error_code, message = _parse_graph_error(response)
                    raise GraphRateLimitCapExceededError(
                        f"Microsoft Graph rate-limited the request (HTTP {status}: {message}) and "
                        f"the required wait ({wait_seconds:.0f}s) exceeds the remaining retry "
                        f"budget ({remaining:.0f}s of {self._total_wait_cap_seconds:.0f}s total). "
                        f"Please retry the job later.",
                        status_code=status,
                        error_code=error_code,
                    )
                self._sleep(wait_seconds)
                elapsed_wait += wait_seconds
                attempt += 1
                continue

            _raise_for_error_response(response)

    def get(self, url: str, **kwargs) -> requests.Response:
        return self.request("GET", url, **kwargs)

    def post(self, url: str, **kwargs) -> requests.Response:
        return self.request("POST", url, **kwargs)

    def patch(self, url: str, **kwargs) -> requests.Response:
        return self.request("PATCH", url, **kwargs)

    def put(self, url: str, **kwargs) -> requests.Response:
        return self.request("PUT", url, **kwargs)

    def delete(self, url: str, **kwargs) -> requests.Response:
        return self.request("DELETE", url, **kwargs)

    def get_paged(
        self,
        url: str,
        *,
        params: dict | None = None,
        headers: dict | None = None,
        absolute: bool = False,
    ) -> Iterator[dict]:
        """Yield every item from the ``value`` array across all ``@odata.nextLink`` pages."""
        next_url: str | None = url
        next_params = params
        next_absolute = absolute

        while next_url:
            response = self.get(next_url, params=next_params, headers=headers, absolute=next_absolute)
            body = response.json()
            yield from body.get("value", [])
            next_url = body.get("@odata.nextLink")
            # `@odata.nextLink` is always a complete, self-contained absolute URL.
            next_params = None
            next_absolute = True

    def _join_url(self, url: str) -> str:
        return f"{self._base_url}/{url.lstrip('/')}"

    def _build_headers(self, caller_headers: dict | None, *, auth: bool) -> dict:
        headers = {"User-Agent": USER_AGENT, "Accept": "application/json"}
        if auth:
            headers["Authorization"] = f"Bearer {self._token_provider.get_access_token()}"
        if caller_headers:
            headers.update(caller_headers)
        return headers

    @staticmethod
    def _is_retryable(status: int, retry_transient_workbook: bool) -> bool:
        if status in _RATE_LIMIT_STATUSES or status in _SERVER_ERROR_STATUSES:
            return True
        return retry_transient_workbook and status in _WORKBOOK_TRANSIENT_STATUSES

    @staticmethod
    def _compute_wait_seconds(response: requests.Response, attempt: int) -> float:
        if response.status_code in _RATE_LIMIT_STATUSES:
            retry_after = response.headers.get("Retry-After")
            if retry_after is not None:
                try:
                    return float(retry_after)
                except ValueError:
                    logger.warning("Ignoring non-numeric Retry-After header: %r", retry_after)
        return min(_BACKOFF_BASE_SECONDS * (2**attempt), _BACKOFF_MAX_SECONDS)

    def _sleep(self, seconds: float) -> None:
        (self._sleep_fn or time.sleep)(seconds)
