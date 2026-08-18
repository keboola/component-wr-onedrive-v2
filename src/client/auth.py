"""OAuth token acquisition and rotation for keboola.wr-onedrive-v2.

Design spec: ``docs/superpowers/specs/2026-08-17-wr-onedrive-v2-design.md`` §3.

Microsoft Graph's Excel API is delegated-only (no application-permission mode), so this
component authenticates via the OAuth 2.0 refresh-token grant against the Microsoft identity
platform — plain ``requests``, no MSAL, no ``requests-oauthlib`` (same deliberate choice as the
sibling extractor, ``kds-team.ex-onedrive``).

``TokenProvider`` is a narrow interface so a future service-principal / client-credentials
provider (CFTL-702, deferred) can be added without changing any caller (``GraphClient``, the
uploader, the Excel writer all depend on the interface, never on ``RefreshTokenProvider``
directly).
"""

import logging
import time
from abc import ABC, abstractmethod
from collections.abc import Callable, Sequence
from dataclasses import dataclass

import requests

logger = logging.getLogger(__name__)

TOKEN_ENDPOINT_TEMPLATE = "https://login.microsoftonline.com/{authority}/oauth2/v2.0/token"

# Must match the scopes already granted to v1's Azure app registration (design spec §3).
SCOPES = "offline_access User.Read Files.ReadWrite.All Sites.ReadWrite.All"

# Access tokens are refreshed this many seconds before their reported `expires_in` elapses,
# so a long-running row never hits a 401 mid-request (fixes the extractor's 60-minute-runtime
# limitation, where the token was only refreshed reactively).
REFRESH_SAFETY_MARGIN_SECONDS = 300

# Fallback lifetime assumed when the token response omits `expires_in` (Microsoft's documented
# default is ~3599s).
DEFAULT_EXPIRES_IN_SECONDS = 3599

_REAUTHORIZE_HINT = "Please reauthorize the configuration in the Keboola UI."


class AuthenticationError(Exception):
    """Raised when a valid Graph access token cannot be obtained.

    Deliberately local to the ``client`` package (no dependency on ``keboola.component`` here) —
    ``component.py`` catches this at the call site and re-raises it as a ``UserException`` so it
    surfaces as an exit-1, user-facing error.
    """


class _InvalidGrantError(Exception):
    """Internal: a single refresh-token candidate was rejected by the token endpoint."""


class TokenProvider(ABC):
    """Minimal interface for supplying a valid Graph access token.

    Callers (``GraphClient`` and anything built on top of it) depend only on this interface,
    never on a concrete provider — this is what lets a future service-principal provider
    (CFTL-702) slot in without reworking upload/Excel code.
    """

    @abstractmethod
    def get_access_token(self) -> str:
        """Return a currently-valid access token, transparently refreshing it first if stale."""

    @property
    @abstractmethod
    def rotated_refresh_token(self) -> str | None:
        """The most recently rotated refresh token, or ``None`` if no refresh has happened yet.

        The caller persists this to row state after each run (see ``RefreshTokenProvider``'s
        docstring for the exact state payload shape) so the next run doesn't need to reauthorize.
        """

    def invalidate(self) -> None:
        """Discard any cached access token so the next :meth:`get_access_token` call refreshes.

        Called by ``GraphClient`` right before its single 401-triggered retry, so a stale
        cached token can't make that retry fail identically. Default is a no-op — a provider
        with nothing cached (e.g. a future stateless service-principal provider, CFTL-702) has
        nothing to discard.
        """


@dataclass
class _TokenState:
    access_token: str
    refresh_token: str
    stale_at: float  # value from `clock()` at/after which the access token must be refreshed


class RefreshTokenProvider(TokenProvider):
    """Delegated OAuth refresh-token grant against the Microsoft identity platform.

    On each refresh, ``POST {token_endpoint}`` is called with a ``grant_type=refresh_token``
    form body (``client_id``, ``client_secret``, ``scope``, ``grant_type``, ``refresh_token``).
    ``token_endpoint`` is ``https://login.microsoftonline.com/{authority}/oauth2/v2.0/token``,
    where ``authority`` is ``"common"`` for private OneDrive accounts and the tenant id for
    business/SharePoint accounts (the caller computes and passes this — this class has no
    knowledge of account types).

    Refresh-token fallback: ``refresh_token_candidates`` is tried in order (state token first,
    config token second, per design spec §3). A candidate rejected with ``invalid_grant`` is
    logged as a warning and the next candidate is tried; once every candidate has been
    exhausted, ``AuthenticationError`` is raised telling the user to reauthorize. Any other
    error response (network error, non-``invalid_grant`` 4xx/5xx) aborts immediately without
    trying further candidates, since it isn't necessarily specific to that refresh token.

    Rotation: every successful refresh returns a *new* refresh token (Microsoft always rotates
    it). It is captured and exposed via :attr:`rotated_refresh_token` so the caller can persist
    it to row state under the v1-compatible key ``#refreshed_auth_data``. The persisted value
    should be the JSON encoding of a payload shaped like::

        {"refresh_token": "<new refresh token>", "access_token": "<new access token>",
         "expires_in": 3599}

    (v1's state payload shape — only ``refresh_token`` is actually required on read; the other
    keys are kept for parity/debuggability.)

    No network call happens in ``__init__`` — the first refresh happens lazily on the first
    :meth:`get_access_token` call.
    """

    def __init__(
        self,
        client_id: str,
        client_secret: str,
        authority: str,
        refresh_token_candidates: Sequence[str],
        session: requests.Session | None = None,
        clock: Callable[[], float] = time.monotonic,
    ):
        """
        Args:
            client_id: OAuth app (client) id — v1's Azure app registration id.
            client_secret: OAuth app client secret.
            authority: ``"common"`` for private OneDrive, tenant id for business/SharePoint.
            refresh_token_candidates: refresh tokens to try, in order (state token first,
                config token second). Must contain at least one non-empty token.
            session: injectable ``requests.Session`` (tests only); a fresh one is created
                otherwise. No request is made until the first :meth:`get_access_token` call.
            clock: injectable monotonic-style clock (tests only); defaults to
                ``time.monotonic``.
        """
        candidates = [token for token in refresh_token_candidates if token]
        if not candidates:
            raise AuthenticationError(
                f"No OAuth refresh token is available for this configuration. {_REAUTHORIZE_HINT}"
            )
        self._client_id = client_id
        self._client_secret = client_secret
        self._token_url = TOKEN_ENDPOINT_TEMPLATE.format(authority=authority)
        self._candidates = candidates
        self._session = session if session is not None else requests.Session()
        self._clock = clock
        self._state: _TokenState | None = None

    def get_access_token(self) -> str:
        if self._state is None or self._clock() >= self._state.stale_at:
            self._refresh()
        return self._state.access_token

    @property
    def rotated_refresh_token(self) -> str | None:
        return self._state.refresh_token if self._state is not None else None

    def invalidate(self) -> None:
        """Drop the cached access token so the next :meth:`get_access_token` call refreshes.

        Used by ``GraphClient``'s 401-retry path: a 401 means the cached access token Graph just
        rejected is no longer trustworthy, even if it isn't "stale" by our own expiry bookkeeping
        (e.g. it was revoked out-of-band). Clearing ``_state`` forces ``get_access_token`` to call
        ``_refresh`` again instead of handing back the same rejected token.
        """
        self._state = None

    def _refresh(self) -> None:
        last_error: Exception | None = None
        for index, candidate in enumerate(self._candidates):
            try:
                payload = self._request_token(candidate)
            except _InvalidGrantError as exc:
                logger.warning(
                    "Refresh token candidate %d/%d was rejected (invalid_grant); trying the next one.",
                    index + 1,
                    len(self._candidates),
                )
                last_error = exc
                continue
            expires_in = int(payload.get("expires_in", DEFAULT_EXPIRES_IN_SECONDS))
            self._state = _TokenState(
                access_token=payload["access_token"],
                refresh_token=payload["refresh_token"],
                stale_at=self._clock() + expires_in - REFRESH_SAFETY_MARGIN_SECONDS,
            )
            return
        raise AuthenticationError(
            f"Unable to refresh the OneDrive/SharePoint access token: the refresh token was "
            f"rejected by Microsoft (invalid_grant). {_REAUTHORIZE_HINT}"
        ) from last_error

    def _request_token(self, refresh_token: str) -> dict:
        try:
            response = self._session.post(
                self._token_url,
                data={
                    "client_id": self._client_id,
                    "client_secret": self._client_secret,
                    "scope": SCOPES,
                    "grant_type": "refresh_token",
                    "refresh_token": refresh_token,
                },
                timeout=30,
            )
        except requests.RequestException as exc:
            raise AuthenticationError(
                f"Could not reach the Microsoft login endpoint to refresh the access token: {exc}. "
                "This is likely a transient network issue; please retry the job."
            ) from exc
        if response.status_code == 200:
            return response.json()

        error_body = _safe_json(response)
        error_code = error_body.get("error")
        if error_code == "invalid_grant":
            raise _InvalidGrantError(error_body.get("error_description", "invalid_grant"))

        description = error_body.get("error_description") or response.text
        raise AuthenticationError(
            f"Token refresh request failed with HTTP {response.status_code}: {description}"
        )


def _safe_json(response: requests.Response) -> dict:
    try:
        return response.json()
    except ValueError:
        return {}
