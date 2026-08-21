"""Unit tests for keboola.wr-onedrive-v2's client/configuration modules (mocked HTTP, no network).

Merged from the retired per-module ``tests/unit/test_{auth,graph_client,drives,uploader,headers,
configuration}.py`` files into one flat module, organized below by source module with a section
banner + its own test classes per module. ``client.exceptions.sanitize_exception_text`` and every
secret-leak-prevention assertion (the query-string/identity redaction tests that used to live in
this module's auth/uploader sections, plus the old ``test_exceptions.py``) moved to
``tests/test_security.py`` instead — see that module's docstring. ``client.excel_writer`` stays in
its own file, ``tests/test_excel_writer.py`` (1,000+ lines; a distinct rendering engine).
"""

from datetime import UTC, datetime
from unittest.mock import MagicMock, call, patch

import pytest
import requests
from pydantic import ValidationError

from client.auth import REFRESH_SAFETY_MARGIN_SECONDS, AuthenticationError, RefreshTokenProvider
from client.drives import get_site_id, list_drives, resolve_drive_id
from client.exceptions import (
    FileAlreadyExistsError,
    GraphBadRequestError,
    GraphClientError,
    GraphConnectionError,
    GraphNotFoundError,
    GraphPermissionError,
    GraphQuotaExceededError,
    GraphRateLimitCapExceededError,
    InvalidPathError,
    UploadSessionError,
)
from client.graph_client import BASE_URL, USER_AGENT, GraphClient
from client.headers import normalize_header_row, to_ascii
from client.uploader import (
    CHUNK_SIZE,
    MAX_RESUME_ATTEMPTS,
    SIMPLE_UPLOAD_THRESHOLD,
    encode_path_segments,
    ensure_folder,
    resolve_placeholders,
    upload_file,
    validate_path,
)
from configuration import (
    Account,
    AccountType,
    ConflictBehavior,
    CsvOptions,
    Destination,
    Mode,
    RowConfig,
    Workbook,
    WorkbookTargeting,
    Worksheet,
    WorksheetSelection,
    WriteMode,
)

# ----------------------------------------------------------------------------------------------
# client.auth — RefreshTokenProvider (token refresh, rotation, fallback order)
# ----------------------------------------------------------------------------------------------


def _make_response(status_code: int, json_body: dict | None = None, text: str = ""):
    response = MagicMock()
    response.status_code = status_code
    response.text = text
    if json_body is None:
        response.json.side_effect = ValueError("no JSON body")
    else:
        response.json.return_value = json_body
    return response


def _token_payload(access_token="access-1", refresh_token="refresh-rotated-1", expires_in=3599):
    return {"access_token": access_token, "refresh_token": refresh_token, "expires_in": expires_in}


class FakeClock:
    """Injectable monotonic-style clock, advanced explicitly by tests."""

    def __init__(self, start: float = 0.0):
        self._now = start

    def __call__(self) -> float:
        return self._now

    def advance(self, seconds: float) -> None:
        self._now += seconds


class TestRefreshSuccess:
    def test_get_access_token_returns_token_from_response(self):
        session = MagicMock()
        session.post.return_value = _make_response(200, _token_payload(access_token="access-abc"))
        provider = RefreshTokenProvider(
            client_id="client-1",
            client_secret="secret-1",
            authority="common",
            refresh_token_candidates=["refresh-config"],
            session=session,
        )

        token = provider.get_access_token()

        assert token == "access-abc"

    def test_posts_expected_token_endpoint_and_form_fields(self):
        session = MagicMock()
        session.post.return_value = _make_response(200, _token_payload())
        provider = RefreshTokenProvider(
            client_id="client-1",
            client_secret="secret-1",
            authority="common",
            refresh_token_candidates=["refresh-config"],
            session=session,
        )

        provider.get_access_token()

        session.post.assert_called_once()
        args, kwargs = session.post.call_args
        assert args[0] == "https://login.microsoftonline.com/common/oauth2/v2.0/token"
        assert kwargs["data"] == {
            "client_id": "client-1",
            "client_secret": "secret-1",
            "scope": "offline_access User.Read Files.ReadWrite.All Sites.ReadWrite.All",
            "grant_type": "refresh_token",
            "refresh_token": "refresh-config",
        }

    def test_uses_tenant_id_as_authority_when_given(self):
        session = MagicMock()
        session.post.return_value = _make_response(200, _token_payload())
        provider = RefreshTokenProvider(
            client_id="client-1",
            client_secret="secret-1",
            authority="tenant-xyz",
            refresh_token_candidates=["refresh-config"],
            session=session,
        )

        provider.get_access_token()

        args, _ = session.post.call_args
        assert args[0] == "https://login.microsoftonline.com/tenant-xyz/oauth2/v2.0/token"

    def test_second_call_before_expiry_does_not_refresh_again(self):
        session = MagicMock()
        session.post.return_value = _make_response(200, _token_payload())
        provider = RefreshTokenProvider(
            client_id="client-1",
            client_secret="secret-1",
            authority="common",
            refresh_token_candidates=["refresh-config"],
            session=session,
        )

        provider.get_access_token()
        provider.get_access_token()

        assert session.post.call_count == 1

    def test_no_network_call_at_construction_time(self):
        session = MagicMock()
        RefreshTokenProvider(
            client_id="client-1",
            client_secret="secret-1",
            authority="common",
            refresh_token_candidates=["refresh-config"],
            session=session,
        )

        session.post.assert_not_called()


class TestRotation:
    def test_rotated_refresh_token_is_none_before_first_refresh(self):
        provider = RefreshTokenProvider(
            client_id="client-1",
            client_secret="secret-1",
            authority="common",
            refresh_token_candidates=["refresh-config"],
            session=MagicMock(),
        )

        assert provider.rotated_refresh_token is None

    def test_rotated_refresh_token_captured_after_refresh(self):
        session = MagicMock()
        session.post.return_value = _make_response(200, _token_payload(refresh_token="brand-new-refresh-token"))
        provider = RefreshTokenProvider(
            client_id="client-1",
            client_secret="secret-1",
            authority="common",
            refresh_token_candidates=["refresh-config"],
            session=session,
        )

        provider.get_access_token()

        assert provider.rotated_refresh_token == "brand-new-refresh-token"


class TestFallbackOrder:
    def test_state_token_tried_before_config_token(self):
        session = MagicMock()
        session.post.return_value = _make_response(200, _token_payload())
        RefreshTokenProvider(
            client_id="client-1",
            client_secret="secret-1",
            authority="common",
            refresh_token_candidates=["refresh-state", "refresh-config"],
            session=session,
        ).get_access_token()

        _, kwargs = session.post.call_args
        assert kwargs["data"]["refresh_token"] == "refresh-state"

    def test_falls_back_to_config_token_when_state_token_invalid_grant(self):
        session = MagicMock()
        invalid_grant_response = _make_response(400, {"error": "invalid_grant", "error_description": "token expired"})
        success_response = _make_response(200, _token_payload(access_token="access-from-config"))
        session.post.side_effect = [invalid_grant_response, success_response]

        provider = RefreshTokenProvider(
            client_id="client-1",
            client_secret="secret-1",
            authority="common",
            refresh_token_candidates=["refresh-state", "refresh-config"],
            session=session,
        )

        token = provider.get_access_token()

        assert token == "access-from-config"
        assert session.post.call_count == 2
        second_call_data = session.post.call_args_list[1].kwargs["data"]
        assert second_call_data["refresh_token"] == "refresh-config"

    def test_logs_warning_when_a_candidate_is_rejected(self, caplog):
        session = MagicMock()
        invalid_grant_response = _make_response(400, {"error": "invalid_grant", "error_description": "token expired"})
        success_response = _make_response(200, _token_payload())
        session.post.side_effect = [invalid_grant_response, success_response]

        provider = RefreshTokenProvider(
            client_id="client-1",
            client_secret="secret-1",
            authority="common",
            refresh_token_candidates=["refresh-state", "refresh-config"],
            session=session,
        )

        with caplog.at_level("WARNING"):
            provider.get_access_token()

        assert any("invalid_grant" in record.message for record in caplog.records)


class TestAllCandidatesFail:
    def test_raises_authentication_error_with_reauthorize_hint(self):
        session = MagicMock()
        session.post.return_value = _make_response(
            400, {"error": "invalid_grant", "error_description": "token expired"}
        )
        provider = RefreshTokenProvider(
            client_id="client-1",
            client_secret="secret-1",
            authority="common",
            refresh_token_candidates=["refresh-state", "refresh-config"],
            session=session,
        )

        with pytest.raises(AuthenticationError, match="reauthorize"):
            provider.get_access_token()

        assert session.post.call_count == 2

    def test_empty_candidate_list_raises_at_construction(self):
        with pytest.raises(AuthenticationError, match="reauthorize"):
            RefreshTokenProvider(
                client_id="client-1",
                client_secret="secret-1",
                authority="common",
                refresh_token_candidates=[],
                session=MagicMock(),
            )

    def test_blank_candidates_are_filtered_out(self):
        with pytest.raises(AuthenticationError, match="reauthorize"):
            RefreshTokenProvider(
                client_id="client-1",
                client_secret="secret-1",
                authority="common",
                refresh_token_candidates=["", None],
                session=MagicMock(),
            )

    def test_non_invalid_grant_error_aborts_without_trying_other_candidates(self):
        session = MagicMock()
        session.post.return_value = _make_response(500, {"error": "server_error"}, text="boom")
        provider = RefreshTokenProvider(
            client_id="client-1",
            client_secret="secret-1",
            authority="common",
            refresh_token_candidates=["refresh-state", "refresh-config"],
            session=session,
        )

        with pytest.raises(AuthenticationError):
            provider.get_access_token()

        assert session.post.call_count == 1

    def test_network_failure_raises_authentication_error(self):
        session = MagicMock()
        session.post.side_effect = requests.exceptions.ConnectionError("connection refused")
        provider = RefreshTokenProvider(
            client_id="client-1",
            client_secret="secret-1",
            authority="common",
            refresh_token_candidates=["refresh-state"],
            session=session,
        )

        with pytest.raises(AuthenticationError, match="Microsoft login endpoint"):
            provider.get_access_token()

    def test_timeout_raises_authentication_error_without_trying_other_candidates(self):
        session = MagicMock()
        session.post.side_effect = requests.exceptions.Timeout("timed out")
        provider = RefreshTokenProvider(
            client_id="client-1",
            client_secret="secret-1",
            authority="common",
            refresh_token_candidates=["refresh-state", "refresh-config"],
            session=session,
        )

        with pytest.raises(AuthenticationError):
            provider.get_access_token()

        assert session.post.call_count == 1


class TestInvalidate:
    def test_invalidate_forces_refresh_on_next_call_even_before_expiry(self):
        clock = FakeClock(start=0.0)
        session = MagicMock()
        session.post.side_effect = [
            _make_response(200, _token_payload(access_token="access-1")),
            _make_response(200, _token_payload(access_token="access-2")),
        ]
        provider = RefreshTokenProvider(
            client_id="client-1",
            client_secret="secret-1",
            authority="common",
            refresh_token_candidates=["refresh-config"],
            session=session,
            clock=clock,
        )

        first_token = provider.get_access_token()
        provider.invalidate()
        # Clock hasn't advanced at all -> without invalidate() this would be a cache hit.
        second_token = provider.get_access_token()

        assert first_token == "access-1"
        assert second_token == "access-2"
        assert session.post.call_count == 2

    def test_invalidate_before_any_refresh_is_a_noop(self):
        provider = RefreshTokenProvider(
            client_id="client-1",
            client_secret="secret-1",
            authority="common",
            refresh_token_candidates=["refresh-config"],
            session=MagicMock(),
        )

        provider.invalidate()

        assert provider.rotated_refresh_token is None


class TestProactiveRefresh:
    def test_refreshes_again_after_expiry_elapses(self):
        clock = FakeClock(start=0.0)
        session = MagicMock()
        session.post.side_effect = [
            _make_response(200, _token_payload(access_token="access-1", expires_in=3599)),
            _make_response(200, _token_payload(access_token="access-2", expires_in=3599)),
        ]
        provider = RefreshTokenProvider(
            client_id="client-1",
            client_secret="secret-1",
            authority="common",
            refresh_token_candidates=["refresh-config"],
            session=session,
            clock=clock,
        )

        first_token = provider.get_access_token()
        # Not yet within the safety margin of expiry -> no re-refresh.
        clock.advance(3599 - REFRESH_SAFETY_MARGIN_SECONDS - 60)
        second_token = provider.get_access_token()
        # Past the safety-margin threshold -> proactive re-refresh happens transparently.
        clock.advance(120)
        third_token = provider.get_access_token()

        assert first_token == "access-1"
        assert second_token == "access-1"
        assert third_token == "access-2"
        assert session.post.call_count == 2

    def test_uses_default_lifetime_when_expires_in_missing(self):
        clock = FakeClock(start=0.0)
        session = MagicMock()
        payload = _token_payload()
        del payload["expires_in"]
        session.post.return_value = _make_response(200, payload)
        provider = RefreshTokenProvider(
            client_id="client-1",
            client_secret="secret-1",
            authority="common",
            refresh_token_candidates=["refresh-config"],
            session=session,
            clock=clock,
        )

        provider.get_access_token()
        clock.advance(1)
        provider.get_access_token()

        assert session.post.call_count == 1


# ----------------------------------------------------------------------------------------------
# client.graph_client — GraphClient HTTP behavior (headers, retries, error mapping, paging)
# ----------------------------------------------------------------------------------------------


def _gc_response(status_code: int, json_body: dict | None = None, headers: dict | None = None, text: str = ""):
    response = MagicMock()
    response.status_code = status_code
    response.ok = 200 <= status_code < 400
    response.headers = headers or {}
    response.text = text
    if json_body is None:
        response.json.side_effect = ValueError("no JSON body")
    else:
        response.json.return_value = json_body
    return response


def _gc_error_body(code: str, message: str) -> dict:
    return {"error": {"code": code, "message": message}}


def _gc_token_provider(token: str = "token-1"):
    provider = MagicMock()
    provider.get_access_token.return_value = token
    return provider


def _gc_client(session=None, token_provider=None, **kwargs) -> GraphClient:
    return GraphClient(
        token_provider=token_provider or _gc_token_provider(),
        session=session or MagicMock(),
        **kwargs,
    )


class TestRequestHeaders:
    def test_sends_authorization_user_agent_and_base_url(self):
        session = MagicMock()
        session.request.return_value = _gc_response(200, {"ok": True})
        provider = _gc_token_provider("abc-token")
        client = _gc_client(session=session, token_provider=provider)

        client.get("/me")

        args, kwargs = session.request.call_args
        assert args[0] == "GET"
        assert args[1] == f"{BASE_URL}/me"
        assert kwargs["headers"]["Authorization"] == "Bearer abc-token"
        assert kwargs["headers"]["User-Agent"] == USER_AGENT
        assert USER_AGENT.startswith("NONISV|Keboola|wr-onedrive-v2/")

    def test_auth_false_omits_authorization_header(self):
        session = MagicMock()
        session.request.return_value = _gc_response(200, {"ok": True})
        provider = _gc_token_provider()
        client = _gc_client(session=session, token_provider=provider)

        client.put("https://upload.example/session-url", absolute=True, auth=False, data=b"chunk")

        provider.get_access_token.assert_not_called()
        _, kwargs = session.request.call_args
        assert "Authorization" not in kwargs["headers"]

    def test_absolute_url_is_not_prefixed_with_base_url(self):
        session = MagicMock()
        session.request.return_value = _gc_response(200, {"ok": True})
        client = _gc_client(session=session)

        client.get("https://graph.microsoft.com/v1.0/absolute/path", absolute=True)

        args, _ = session.request.call_args
        assert args[1] == "https://graph.microsoft.com/v1.0/absolute/path"


class TestRetryAfter:
    @patch("client.graph_client.time.sleep")
    def test_honors_retry_after_header_in_seconds(self, mock_sleep):
        session = MagicMock()
        session.request.side_effect = [
            _gc_response(429, _gc_error_body("TooManyRequests", "slow down"), headers={"Retry-After": "7"}),
            _gc_response(200, {"ok": True}),
        ]
        client = _gc_client(session=session)

        response = client.get("/me")

        assert response.ok
        mock_sleep.assert_called_once_with(7.0)

    @patch("client.graph_client.time.sleep")
    def test_503_without_retry_after_falls_back_to_backoff(self, mock_sleep):
        session = MagicMock()
        session.request.side_effect = [
            _gc_response(503, _gc_error_body("ServiceUnavailable", "busy")),
            _gc_response(200, {"ok": True}),
        ]
        client = _gc_client(session=session)

        client.get("/me")

        mock_sleep.assert_called_once_with(1.0)


class TestServerErrorBackoff:
    @patch("client.graph_client.time.sleep")
    def test_5xx_backs_off_then_succeeds(self, mock_sleep):
        session = MagicMock()
        session.request.side_effect = [
            _gc_response(500, _gc_error_body("InternalServerError", "oops")),
            _gc_response(502, _gc_error_body("BadGateway", "oops")),
            _gc_response(200, {"ok": True}),
        ]
        client = _gc_client(session=session)

        response = client.get("/me")

        assert response.ok
        assert mock_sleep.call_args_list == [((1.0,),), ((2.0,),)]

    @patch("client.graph_client.time.sleep")
    def test_5xx_exhausting_cap_raises_rate_limit_cap_error(self, mock_sleep):
        session = MagicMock()
        # Backoff sequence 1, 2, 4, 8, ... will exceed a tiny cap quickly.
        session.request.return_value = _gc_response(500, _gc_error_body("InternalServerError", "oops"))
        client = _gc_client(session=session, total_wait_cap_seconds=1.5)

        with pytest.raises(GraphRateLimitCapExceededError):
            client.get("/me")


class TestRateLimitCap:
    def test_single_retry_after_exceeding_cap_raises_immediately(self):
        session = MagicMock()
        session.request.return_value = _gc_response(
            429, _gc_error_body("TooManyRequests", "slow down"), headers={"Retry-After": "600"}
        )
        client = _gc_client(session=session, total_wait_cap_seconds=300.0)

        with pytest.raises(GraphRateLimitCapExceededError) as exc_info:
            client.get("/me")

        assert "600" in str(exc_info.value) or "600s" in str(exc_info.value)
        session.request.assert_called_once()

    @patch("client.graph_client.time.sleep")
    def test_retry_after_zero_loop_terminates_via_attempt_cap(self, mock_sleep):
        """A server that always sends `Retry-After: 0` never accumulates cumulative wait, so only
        the max-attempts guard (MINOR-3) can stop this from looping forever."""
        session = MagicMock()
        session.request.return_value = _gc_response(
            429, _gc_error_body("TooManyRequests", "slow down"), headers={"Retry-After": "0"}
        )
        client = _gc_client(session=session, total_wait_cap_seconds=10_000.0, max_retry_attempts=15)

        with pytest.raises(GraphRateLimitCapExceededError):
            client.get("/me")

        assert session.request.call_count == 16  # initial attempt + 15 retries
        assert mock_sleep.call_count == 15

    @patch("client.graph_client.time.sleep")
    def test_cumulative_retries_exhausting_cap_raises(self, mock_sleep):
        session = MagicMock()
        session.request.side_effect = [
            _gc_response(429, _gc_error_body("TooManyRequests", "slow"), headers={"Retry-After": "5"}),
            _gc_response(429, _gc_error_body("TooManyRequests", "slow"), headers={"Retry-After": "5"}),
            _gc_response(429, _gc_error_body("TooManyRequests", "slow"), headers={"Retry-After": "5"}),
        ]
        client = _gc_client(session=session, total_wait_cap_seconds=8.0)

        with pytest.raises(GraphRateLimitCapExceededError):
            client.get("/me")

        # First 5s wait succeeds (elapsed=5 <= 8), second 5s wait would bring elapsed to 10 > 8.
        mock_sleep.assert_called_once_with(5.0)


class TestWorkbookTransientOptIn:
    @patch("client.graph_client.time.sleep")
    def test_409_not_retried_without_opt_in(self, mock_sleep):
        session = MagicMock()
        session.request.return_value = _gc_response(409, _gc_error_body("nameAlreadyExists", "exists"))
        client = _gc_client(session=session)

        with pytest.raises(GraphClientError):
            client.get("/me")

        session.request.assert_called_once()
        mock_sleep.assert_not_called()

    @patch("client.graph_client.time.sleep")
    def test_409_retried_with_opt_in(self, mock_sleep):
        session = MagicMock()
        session.request.side_effect = [
            _gc_response(409, _gc_error_body("EditModeCannotAcquireLockTooManyRequests", "locked")),
            _gc_response(200, {"ok": True}),
        ]
        client = _gc_client(session=session)

        response = client.get("/me", retry_transient_workbook=True)

        assert response.ok
        mock_sleep.assert_called_once()

    @patch("client.graph_client.time.sleep")
    def test_405_retried_with_opt_in(self, mock_sleep):
        session = MagicMock()
        session.request.side_effect = [
            _gc_response(405, _gc_error_body("MethodNotAllowed", "not allowed")),
            _gc_response(200, {"ok": True}),
        ]
        client = _gc_client(session=session)

        response = client.patch("/me", retry_transient_workbook=True)

        assert response.ok


class TestUnauthorizedRetry:
    def test_401_refetches_token_and_retries_once(self):
        session = MagicMock()
        session.request.side_effect = [
            _gc_response(401, _gc_error_body("InvalidAuthenticationToken", "expired")),
            _gc_response(200, {"ok": True}),
        ]
        provider = MagicMock()
        provider.get_access_token.side_effect = ["token-old", "token-new"]
        client = _gc_client(session=session, token_provider=provider)

        response = client.get("/me")

        assert response.ok
        assert provider.get_access_token.call_count == 2
        second_call_headers = session.request.call_args_list[1].kwargs["headers"]
        assert second_call_headers["Authorization"] == "Bearer token-new"

    def test_401_invalidates_token_provider_before_retrying(self):
        """The 401 path must call `invalidate()` so a provider's own cache can't hand back the
        same rejected token on retry — this is what actually makes `get_access_token`'s second
        call return something different, not just a mock configured to do so."""
        session = MagicMock()
        session.request.side_effect = [
            _gc_response(401, _gc_error_body("InvalidAuthenticationToken", "expired")),
            _gc_response(200, {"ok": True}),
        ]
        provider = MagicMock()
        provider.get_access_token.side_effect = ["token-old", "token-new"]
        client = _gc_client(session=session, token_provider=provider)

        client.get("/me")

        provider.invalidate.assert_called_once()
        # invalidate() must happen strictly between the two get_access_token() calls.
        call_order = [call[0] for call in provider.method_calls]
        assert call_order == ["get_access_token", "invalidate", "get_access_token"]

    def test_401_does_not_invalidate_when_auth_is_false(self):
        session = MagicMock()
        session.request.return_value = _gc_response(401, _gc_error_body("InvalidAuthenticationToken", "expired"))
        provider = MagicMock()
        client = _gc_client(session=session, token_provider=provider)

        with pytest.raises(GraphPermissionError):
            client.put("https://upload.example/session-url", absolute=True, auth=False)

        provider.invalidate.assert_not_called()

    def test_401_persisting_after_retry_raises_permission_error(self):
        session = MagicMock()
        session.request.return_value = _gc_response(401, _gc_error_body("InvalidAuthenticationToken", "still bad"))
        client = _gc_client(session=session)

        with pytest.raises(GraphPermissionError):
            client.get("/me")

        assert session.request.call_count == 2


class TestNoRetryFlag:
    @patch("client.graph_client.time.sleep")
    def test_retry_false_raises_immediately_on_first_error(self, mock_sleep):
        session = MagicMock()
        session.request.return_value = _gc_response(503, _gc_error_body("ServiceUnavailable", "busy"))
        client = _gc_client(session=session)

        with pytest.raises(GraphClientError):
            client.put("/drives/1/items/2/content", retry=False)

        session.request.assert_called_once()
        mock_sleep.assert_not_called()

    def test_retry_false_does_not_retry_401(self):
        session = MagicMock()
        session.request.return_value = _gc_response(401, _gc_error_body("InvalidAuthenticationToken", "expired"))
        provider = _gc_token_provider()
        client = _gc_client(session=session, token_provider=provider)

        with pytest.raises(GraphPermissionError):
            client.get("/me", retry=False)

        session.request.assert_called_once()
        provider.get_access_token.assert_called_once()


class TestErrorMapping:
    @pytest.mark.parametrize(
        "status_code,expected_exception",
        [
            (401, GraphPermissionError),
            (403, GraphPermissionError),
            (404, GraphNotFoundError),
            (507, GraphQuotaExceededError),
            (400, GraphBadRequestError),
            (418, GraphClientError),
        ],
    )
    def test_status_maps_to_typed_exception(self, status_code, expected_exception):
        session = MagicMock()
        session.request.return_value = _gc_response(status_code, _gc_error_body("SomeCode", "Some message"))
        client = _gc_client(session=session)

        with pytest.raises(expected_exception) as exc_info:
            client.get("/me", retry=False)

        assert exc_info.value.status_code == status_code
        assert exc_info.value.error_code == "SomeCode"
        assert "Some message" in str(exc_info.value)

    def test_non_json_error_body_falls_back_to_response_text(self):
        session = MagicMock()
        session.request.return_value = _gc_response(404, json_body=None, text="Not Found")
        client = _gc_client(session=session)

        with pytest.raises(GraphNotFoundError) as exc_info:
            client.get("/me", retry=False)

        assert "Not Found" in str(exc_info.value)
        assert exc_info.value.error_code is None


class TestConnectionErrors:
    @patch("client.graph_client.time.sleep")
    def test_connection_error_retried_then_succeeds(self, mock_sleep):
        session = MagicMock()
        session.request.side_effect = [
            requests.exceptions.ConnectionError("connection refused"),
            _gc_response(200, {"ok": True}),
        ]
        client = _gc_client(session=session)

        response = client.get("/me")

        assert response.ok
        mock_sleep.assert_called_once_with(1.0)

    @patch("client.graph_client.time.sleep")
    def test_timeout_retried_then_succeeds(self, mock_sleep):
        session = MagicMock()
        session.request.side_effect = [
            requests.exceptions.Timeout("timed out"),
            _gc_response(200, {"ok": True}),
        ]
        client = _gc_client(session=session)

        response = client.get("/me")

        assert response.ok
        mock_sleep.assert_called_once_with(1.0)

    @patch("client.graph_client.time.sleep")
    def test_retries_exhausted_raises_graph_connection_error(self, mock_sleep):
        session = MagicMock()
        session.request.side_effect = requests.exceptions.ConnectionError("connection refused")
        client = _gc_client(session=session, total_wait_cap_seconds=1.5)

        with pytest.raises(GraphConnectionError):
            client.get("/me")

    def test_retry_false_raises_graph_connection_error_immediately(self):
        session = MagicMock()
        session.request.side_effect = requests.exceptions.ConnectionError("connection refused")
        client = _gc_client(session=session)

        with pytest.raises(GraphConnectionError):
            client.put("/drives/1/items/2/content", retry=False)

        session.request.assert_called_once()

    @patch("client.graph_client.time.sleep")
    def test_attempt_cap_raises_graph_connection_error(self, mock_sleep):
        # A connection error every time never accumulates cumulative wait beyond a *huge* cap,
        # so only the attempt cap can stop this loop.
        session = MagicMock()
        session.request.side_effect = requests.exceptions.ConnectionError("connection refused")
        client = _gc_client(session=session, total_wait_cap_seconds=10_000.0, max_retry_attempts=3)

        with pytest.raises(GraphConnectionError):
            client.get("/me")

        assert session.request.call_count == 4  # initial attempt + 3 retries
        assert mock_sleep.call_count == 3


class TestPaging:
    def test_get_paged_yields_items_across_two_pages(self):
        session = MagicMock()
        session.request.side_effect = [
            _gc_response(
                200,
                {"value": [{"id": 1}, {"id": 2}], "@odata.nextLink": "https://graph.microsoft.com/v1.0/next-page"},
            ),
            _gc_response(200, {"value": [{"id": 3}]}),
        ]
        client = _gc_client(session=session)

        items = list(client.get_paged("/sites/site-1/drives"))

        assert items == [{"id": 1}, {"id": 2}, {"id": 3}]
        assert session.request.call_count == 2
        second_call_args, _ = session.request.call_args_list[1]
        assert second_call_args[1] == "https://graph.microsoft.com/v1.0/next-page"

    def test_get_paged_stops_when_no_next_link(self):
        session = MagicMock()
        session.request.return_value = _gc_response(200, {"value": [{"id": 1}]})
        client = _gc_client(session=session)

        items = list(client.get_paged("/sites/site-1/drives"))

        assert items == [{"id": 1}]
        session.request.assert_called_once()

    def test_get_paged_defaults_to_empty_list_when_value_missing(self):
        session = MagicMock()
        session.request.return_value = _gc_response(200, {})
        client = _gc_client(session=session)

        assert list(client.get_paged("/sites/site-1/drives")) == []


# ----------------------------------------------------------------------------------------------
# client.drives — site/drive resolution
# ----------------------------------------------------------------------------------------------


def _client_with_get(return_value=None, side_effect=None) -> MagicMock:
    client = MagicMock()
    if side_effect is not None:
        client.get.side_effect = side_effect
    else:
        response = MagicMock()
        response.json.return_value = return_value
        client.get.return_value = response
    return client


class TestGetSiteId:
    def test_builds_hostname_and_server_relative_path_lookup_url(self):
        client = _client_with_get(return_value={"id": "contoso.sharepoint.com,GUID1,GUID2"})

        site_id = get_site_id(client, "https://contoso.sharepoint.com/sites/marketing")

        client.get.assert_called_once_with("/sites/contoso.sharepoint.com:/sites/marketing")
        assert site_id == "contoso.sharepoint.com,GUID1,GUID2"

    def test_root_site_url_omits_the_colon_path_suffix(self):
        client = _client_with_get(return_value={"id": "contoso.sharepoint.com,GUID1,GUID2"})

        get_site_id(client, "https://contoso.sharepoint.com")

        client.get.assert_called_once_with("/sites/contoso.sharepoint.com")

    def test_trailing_slash_in_path_is_normalized(self):
        client = _client_with_get(return_value={"id": "site-id"})

        get_site_id(client, "https://contoso.sharepoint.com/sites/marketing/")

        client.get.assert_called_once_with("/sites/contoso.sharepoint.com:/sites/marketing")

    def test_missing_hostname_raises_without_any_network_call(self):
        client = MagicMock()

        with pytest.raises(GraphNotFoundError):
            get_site_id(client, "not-a-url")

        client.get.assert_not_called()

    def test_404_is_wrapped_with_a_helpful_message(self):
        client = _client_with_get(
            side_effect=GraphNotFoundError("Item not found", status_code=404, error_code="itemNotFound")
        )

        with pytest.raises(GraphNotFoundError) as exc_info:
            get_site_id(client, "https://contoso.sharepoint.com/sites/marketing")

        assert "https://contoso.sharepoint.com/sites/marketing" in str(exc_info.value)
        assert exc_info.value.status_code == 404
        assert exc_info.value.error_code == "itemNotFound"


class TestListDrives:
    def test_delegates_to_the_paging_helper_and_returns_a_list(self):
        client = MagicMock()
        client.get_paged.return_value = iter(
            [
                {"id": "drive-1", "name": "Documents", "webUrl": "https://x/Documents"},
                {"id": "drive-2", "name": "Marketing Assets", "webUrl": "https://x/Marketing"},
            ]
        )

        drives = list_drives(client, "site-1")

        client.get_paged.assert_called_once_with("/sites/site-1/drives")
        assert drives == [
            {"id": "drive-1", "name": "Documents", "webUrl": "https://x/Documents"},
            {"id": "drive-2", "name": "Marketing Assets", "webUrl": "https://x/Marketing"},
        ]


class TestResolveDriveId:
    @pytest.mark.parametrize("account_type", [AccountType.PRIVATE_ONEDRIVE, AccountType.ONEDRIVE_FOR_BUSINESS])
    def test_personal_and_business_accounts_without_configured_drive_id_resolve_via_me_drive(self, account_type):
        client = _client_with_get(return_value={"id": "my-drive-id"})
        account = Account(account_type=account_type, tenant_id="tenant-1")

        drive_id = resolve_drive_id(client, account, destination_drive_id=None)

        client.get.assert_called_once_with("/me/drive")
        assert drive_id == "my-drive-id"

    @pytest.mark.parametrize("account_type", [AccountType.PRIVATE_ONEDRIVE, AccountType.ONEDRIVE_FOR_BUSINESS])
    def test_personal_and_business_accounts_with_configured_drive_id_use_it_verbatim(self, account_type):
        # Drives are globally addressable in Graph — a configured `destination.drive_id` is now
        # honored for every account type, not just `sharepoint`, and never triggers a `/me/drive`
        # lookup.
        client = MagicMock()
        account = Account(account_type=account_type, tenant_id="tenant-1")

        drive_id = resolve_drive_id(client, account, destination_drive_id="drive-configured")

        assert drive_id == "drive-configured"
        client.get.assert_not_called()

    def test_sharepoint_with_configured_drive_id_makes_no_network_call(self):
        client = MagicMock()
        account = Account(
            account_type=AccountType.SHAREPOINT,
            tenant_id="tenant-1",
            site_url="https://contoso.sharepoint.com/sites/marketing",
        )

        drive_id = resolve_drive_id(client, account, destination_drive_id="drive-configured")

        assert drive_id == "drive-configured"
        client.get.assert_not_called()
        client.get_paged.assert_not_called()

    def test_sharepoint_without_configured_drive_id_resolves_site_default_drive(self):
        client = MagicMock()
        site_response = MagicMock()
        site_response.json.return_value = {"id": "site-1"}
        drive_response = MagicMock()
        drive_response.json.return_value = {"id": "default-drive-id"}
        client.get.side_effect = [site_response, drive_response]
        account = Account(
            account_type=AccountType.SHAREPOINT,
            tenant_id="tenant-1",
            site_url="https://contoso.sharepoint.com/sites/marketing",
        )

        drive_id = resolve_drive_id(client, account, destination_drive_id=None)

        assert drive_id == "default-drive-id"
        assert client.get.call_args_list[0].args[0] == "/sites/contoso.sharepoint.com:/sites/marketing"
        assert client.get.call_args_list[1].args[0] == "/sites/site-1/drive"


# ----------------------------------------------------------------------------------------------
# client.uploader — path placeholders/validation, folder creation, upload dispatch
# ----------------------------------------------------------------------------------------------


def _uploader_response(status_code: int, json_body=None, headers: dict | None = None) -> MagicMock:
    response = MagicMock()
    response.status_code = status_code
    response.ok = 200 <= status_code < 400
    response.headers = headers or {}
    if json_body is not None:
        response.json.return_value = json_body
    return response


def _uploader_graph_error(status_code: int, error_code: str | None = None) -> GraphClientError:
    return GraphClientError(f"HTTP {status_code}", status_code=status_code, error_code=error_code)


def _make_sparse_file(path, size: int) -> str:
    """Create a file of exactly ``size`` bytes without allocating real content (sparse file)."""
    with open(path, "wb") as fh:
        if size:
            fh.truncate(size)
    return str(path)


class TestResolvePlaceholders:
    def test_resolves_a_single_date_token(self):
        now = datetime(2026, 8, 17, 12, 30, tzinfo=UTC)

        result = resolve_placeholders("reports/{date:%Y-%m-%d}", now)

        assert result == "reports/2026-08-17"

    def test_resolves_multiple_tokens_with_arbitrary_strftime_formats(self):
        now = datetime(2026, 8, 17, 9, 5, tzinfo=UTC)

        result = resolve_placeholders("{date:%Y}/{date:%m}/{date:%d}-report", now)

        assert result == "2026/08/17-report"

    def test_path_without_placeholders_is_unchanged(self):
        now = datetime(2026, 8, 17, tzinfo=UTC)

        assert resolve_placeholders("static/folder", now) == "static/folder"

    def test_unknown_token_raises_invalid_path_error(self):
        now = datetime(2026, 8, 17, tzinfo=UTC)

        with pytest.raises(InvalidPathError, match="Unknown placeholder"):
            resolve_placeholders("reports/{table_name}", now)

    def test_empty_date_format_raises(self):
        now = datetime(2026, 8, 17, tzinfo=UTC)

        with pytest.raises(InvalidPathError, match="Empty date format"):
            resolve_placeholders("reports/{date:}", now)


class TestDoubleBraceDatePlaceholder:
    """Change 1: `{{date}}` is the current, documented placeholder — always `now` formatted
    `%Y-%m-%d`, no arguments. The old `{date:<strftime-format>}` form (`TestResolvePlaceholders`
    above) keeps resolving silently for already-recorded platform rows, but is never the
    recommended way to write a new one."""

    def test_resolves_in_a_folder_path(self):
        now = datetime(2026, 8, 17, 12, 30, tzinfo=UTC)

        assert resolve_placeholders("acme/reports/{{date}}/", now) == "acme/reports/2026-08-17/"

    def test_resolves_in_a_file_name(self):
        now = datetime(2026, 1, 5, tzinfo=UTC)

        assert resolve_placeholders("orders-{{date}}.csv", now) == "orders-2026-01-05.csv"

    def test_always_formats_yyyy_mm_dd_regardless_of_time_of_day(self):
        now = datetime(2026, 8, 17, 23, 59, 59, tzinfo=UTC)

        assert resolve_placeholders("{{date}}", now) == "2026-08-17"

    def test_unknown_double_brace_token_raises_invalid_path_error(self):
        now = datetime(2026, 8, 17, tzinfo=UTC)

        with pytest.raises(InvalidPathError, match="Unknown placeholder '\\{\\{table_name\\}\\}'"):
            resolve_placeholders("reports/{{table_name}}", now)

    def test_legacy_strftime_placeholder_still_resolves_alongside_double_brace(self):
        # Silent legacy handling: an existing row could plausibly mix both forms across separate
        # fields (folder_path still on the old syntax, csv.file_name migrated) — both must resolve
        # from the same `now` in one call.
        now = datetime(2026, 8, 17, tzinfo=UTC)

        result = resolve_placeholders("reports/{date:%Y}/{{date}}", now)

        assert result == "reports/2026/2026-08-17"

    def test_path_without_any_placeholder_is_unchanged(self):
        now = datetime(2026, 8, 17, tzinfo=UTC)

        assert resolve_placeholders("static/folder", now) == "static/folder"


class TestValidatePath:
    @pytest.mark.parametrize("char", list('"*:<>?\\|'))
    def test_reserved_character_rejected(self, char):
        with pytest.raises(InvalidPathError, match="reserved character"):
            validate_path(f"folder/na{char}me", business=False)

    def test_hash_and_percent_allowed_on_personal_onedrive(self):
        validate_path("folder/na#me%2", business=False)  # does not raise

    @pytest.mark.parametrize("char", ["#", "%"])
    def test_hash_and_percent_rejected_on_business(self, char):
        with pytest.raises(InvalidPathError, match="reserved character"):
            validate_path(f"folder/na{char}me", business=True)

    @pytest.mark.parametrize("name", ["CON", "con", "PRN", "AUX", "NUL", "COM1", "com9", "LPT0", "LPT5"])
    def test_reserved_device_names_rejected(self, name):
        with pytest.raises(InvalidPathError, match="reserved device name"):
            validate_path(f"folder/{name}", business=False)

    @pytest.mark.parametrize("name", [".lock", "desktop.ini", "DESKTOP.INI"])
    def test_reserved_exact_names_rejected(self, name):
        with pytest.raises(InvalidPathError, match="reserved name"):
            validate_path(f"folder/{name}", business=False)

    def test_name_starting_with_tilde_dollar_rejected(self):
        with pytest.raises(InvalidPathError, match="starts with '~'"):
            validate_path("folder/~$temp.docx", business=False)

    def test_vti_substring_anywhere_rejected(self):
        with pytest.raises(InvalidPathError, match="_vti_"):
            validate_path("folder/my_vti_stuff", business=False)

    def test_forms_at_root_rejected(self):
        with pytest.raises(InvalidPathError, match="forms"):
            validate_path("forms", business=False)

    def test_forms_not_at_root_is_allowed(self):
        validate_path("folder/forms", business=False)  # does not raise

    def test_segment_ending_with_dot_rejected(self):
        with pytest.raises(InvalidPathError, match="ends with '.'"):
            validate_path("folder/name.", business=False)

    def test_segment_with_leading_trailing_spaces_rejected(self):
        with pytest.raises(InvalidPathError, match="leading/trailing spaces"):
            validate_path("folder/ name", business=False)

    def test_segment_too_long_rejected(self):
        with pytest.raises(InvalidPathError, match="256 characters"):
            validate_path("a" * 256, business=False)

    def test_segment_at_max_length_is_allowed(self):
        validate_path("a" * 255, business=False)  # does not raise

    def test_total_path_too_long_rejected(self):
        long_path = "/".join(["seg"] * 150)
        assert len(long_path) > 400
        with pytest.raises(InvalidPathError, match="maximum is 400"):
            validate_path(long_path, business=False)

    def test_valid_path_does_not_raise(self):
        validate_path("reports/2026-08-17/monthly export", business=True)


class TestEncodePathSegments:
    def test_space_is_percent_encoded(self):
        assert encode_path_segments("my folder/report") == "my%20folder/report"

    def test_hash_is_percent_encoded(self):
        assert encode_path_segments("folder#1") == "folder%231"

    def test_diacritics_are_percent_encoded(self):
        assert encode_path_segments("faktury/účetní") == "faktury/%C3%BA%C4%8Detn%C3%AD"

    def test_slash_within_a_conceptual_segment_never_occurs_but_separators_are_preserved(self):
        assert encode_path_segments("a/b/c") == "a/b/c"

    def test_empty_path_encodes_to_empty_string(self):
        assert encode_path_segments("") == ""


class TestEnsureFolder:
    def test_empty_path_returns_root_id(self):
        client = MagicMock()
        client.get.return_value = _uploader_response(200, {"id": "root-id"})

        folder_id = ensure_folder(client, "drive-1", "")

        client.get.assert_called_once_with("/drives/drive-1/root")
        assert folder_id == "root-id"

    def test_existing_folders_are_found_via_get(self):
        client = MagicMock()
        client.get.side_effect = [
            _uploader_response(200, {"id": "a-id"}),
            _uploader_response(200, {"id": "b-id"}),
        ]

        folder_id = ensure_folder(client, "drive-1", "a/b")

        assert client.get.call_args_list == [
            call("/drives/drive-1/root:/a"),
            call("/drives/drive-1/root:/a/b"),
        ]
        client.post.assert_not_called()
        assert folder_id == "b-id"

    def test_missing_segment_is_created(self):
        client = MagicMock()
        client.get.side_effect = GraphNotFoundError("not found", status_code=404)
        client.post.return_value = _uploader_response(201, {"id": "new-id"})

        folder_id = ensure_folder(client, "drive-1", "new")

        client.post.assert_called_once_with(
            "/drives/drive-1/root/children",
            json={"name": "new", "folder": {}, "@microsoft.graph.conflictBehavior": "fail"},
        )
        assert folder_id == "new-id"

    def test_second_level_folder_is_created_under_the_first_levels_id(self):
        client = MagicMock()
        client.get.side_effect = [
            _uploader_response(200, {"id": "a-id"}),
            GraphNotFoundError("not found", status_code=404),
        ]
        client.post.return_value = _uploader_response(201, {"id": "b-id"})

        folder_id = ensure_folder(client, "drive-1", "a/b")

        client.post.assert_called_once_with(
            "/drives/drive-1/items/a-id/children",
            json={"name": "b", "folder": {}, "@microsoft.graph.conflictBehavior": "fail"},
        )
        assert folder_id == "b-id"

    def test_409_creation_race_reregets_the_folder(self):
        client = MagicMock()
        client.get.side_effect = [
            GraphNotFoundError("not found", status_code=404),
            _uploader_response(200, {"id": "existing-id"}),
        ]
        client.post.side_effect = _uploader_graph_error(409, "nameAlreadyExists")

        folder_id = ensure_folder(client, "drive-1", "raced")

        assert folder_id == "existing-id"
        assert client.get.call_args_list == [
            call("/drives/drive-1/root:/raced"),
            call("/drives/drive-1/root:/raced"),
        ]

    def test_non_conflict_creation_error_propagates(self):
        client = MagicMock()
        client.get.side_effect = GraphNotFoundError("not found", status_code=404)
        client.post.side_effect = _uploader_graph_error(403)

        with pytest.raises(GraphClientError):
            ensure_folder(client, "drive-1", "denied")


class TestUploadDispatch:
    def test_small_file_uses_simple_put_with_conflict_behavior_query_param(self, tmp_path):
        local_path = _make_sparse_file(tmp_path / "small.csv", 1024)
        client = MagicMock()
        client.put.return_value = _uploader_response(201, {"id": "item-1", "name": "small.csv"})

        result = upload_file(client, "drive-1", "parent-1", local_path, "small.csv", "fail")

        client.put.assert_called_once()
        args, kwargs = client.put.call_args
        assert args[0] == "/drives/drive-1/items/parent-1:/small.csv:/content"
        assert kwargs["params"] == {"@microsoft.graph.conflictBehavior": "fail"}
        assert kwargs["headers"]["Content-Type"] == "application/octet-stream"
        assert kwargs["data"].name == local_path
        client.post.assert_not_called()
        assert result == {"id": "item-1", "name": "small.csv"}

    def test_simple_put_conflict_behavior_is_always_present_even_for_replace(self, tmp_path):
        local_path = _make_sparse_file(tmp_path / "small.csv", 10)
        client = MagicMock()
        client.put.return_value = _uploader_response(200, {"id": "item-1"})

        upload_file(client, "drive-1", "parent-1", local_path, "small.csv", "replace")

        _, kwargs = client.put.call_args
        assert kwargs["params"] == {"@microsoft.graph.conflictBehavior": "replace"}

    def test_file_exactly_at_threshold_uses_simple_put(self, tmp_path):
        local_path = _make_sparse_file(tmp_path / "exact.bin", SIMPLE_UPLOAD_THRESHOLD)
        client = MagicMock()
        client.put.return_value = _uploader_response(201, {"id": "item-1"})

        upload_file(client, "drive-1", "parent-1", local_path, "exact.bin", "fail")

        client.put.assert_called_once()
        client.post.assert_not_called()

    def test_simple_put_name_conflict_raises_file_already_exists(self, tmp_path):
        local_path = _make_sparse_file(tmp_path / "small.csv", 10)
        client = MagicMock()
        client.put.side_effect = _uploader_graph_error(409, "nameAlreadyExists")

        with pytest.raises(FileAlreadyExistsError, match="small.csv"):
            upload_file(client, "drive-1", "parent-1", local_path, "small.csv", "fail")

    def test_simple_put_other_error_propagates_unwrapped(self, tmp_path):
        local_path = _make_sparse_file(tmp_path / "small.csv", 10)
        client = MagicMock()
        client.put.side_effect = _uploader_graph_error(403)

        with pytest.raises(GraphClientError):
            upload_file(client, "drive-1", "parent-1", local_path, "small.csv", "fail")

    def test_file_over_threshold_uses_upload_session(self, tmp_path):
        size = SIMPLE_UPLOAD_THRESHOLD + 1
        local_path = _make_sparse_file(tmp_path / "big.bin", size)
        client = MagicMock()
        client.post.return_value = _uploader_response(200, {"uploadUrl": "https://upload.example/session"})
        client.put.side_effect = [
            _uploader_response(202),
            _uploader_response(201, {"id": "item-big"}),
        ]

        result = upload_file(client, "drive-1", "parent-1", local_path, "big.bin", "fail")

        client.post.assert_called_once_with(
            "/drives/drive-1/items/parent-1:/big.bin:/createUploadSession",
            json={"item": {"@microsoft.graph.conflictBehavior": "fail", "name": "big.bin"}},
        )
        assert result == {"id": "item-big"}
        assert client.put.call_count == 2
        first_headers = client.put.call_args_list[0].kwargs["headers"]
        second_headers = client.put.call_args_list[1].kwargs["headers"]
        assert first_headers["Content-Range"] == f"bytes 0-{SIMPLE_UPLOAD_THRESHOLD - 1}/{size}"
        assert second_headers["Content-Range"] == f"bytes {SIMPLE_UPLOAD_THRESHOLD}-{size - 1}/{size}"
        for kwargs in (client.put.call_args_list[0].kwargs, client.put.call_args_list[1].kwargs):
            assert kwargs["absolute"] is True
            assert kwargs["auth"] is False
            assert kwargs["retry"] is False


class TestChunkMath:
    def test_25_mib_file_uploads_in_three_chunks_with_a_partial_final_chunk(self, tmp_path):
        size = 25 * 1024 * 1024
        local_path = _make_sparse_file(tmp_path / "big.bin", size)
        client = MagicMock()
        client.post.return_value = _uploader_response(200, {"uploadUrl": "https://upload.example/session"})
        client.put.side_effect = [
            _uploader_response(202),
            _uploader_response(202),
            _uploader_response(201, {"id": "item-big"}),
        ]

        result = upload_file(client, "drive-1", "parent-1", local_path, "big.bin", "fail")

        assert result == {"id": "item-big"}
        assert client.put.call_count == 3
        ranges = [c.kwargs["headers"]["Content-Range"] for c in client.put.call_args_list]
        lengths = [c.kwargs["headers"]["Content-Length"] for c in client.put.call_args_list]
        assert ranges == [
            f"bytes 0-{CHUNK_SIZE - 1}/{size}",
            f"bytes {CHUNK_SIZE}-{2 * CHUNK_SIZE - 1}/{size}",
            f"bytes {2 * CHUNK_SIZE}-{size - 1}/{size}",
        ]
        assert lengths == [str(CHUNK_SIZE), str(CHUNK_SIZE), str(size - 2 * CHUNK_SIZE)]
        # Final chunk is smaller than a full CHUNK_SIZE.
        assert int(lengths[-1]) < CHUNK_SIZE


class TestUploadSessionResume:
    def test_transient_failure_resumes_from_next_expected_ranges(self, tmp_path):
        size = 25 * 1024 * 1024
        local_path = _make_sparse_file(tmp_path / "big.bin", size)
        client = MagicMock()
        client.post.return_value = _uploader_response(200, {"uploadUrl": "https://upload.example/session"})
        resumed_offset = CHUNK_SIZE + 2048
        client.put.side_effect = [
            _uploader_response(202),  # chunk 1 ok
            _uploader_graph_error(503),  # chunk 2 fails transiently
            _uploader_response(202),  # resumed partial chunk accepted
            _uploader_response(201, {"id": "item-big"}),  # final chunk
        ]
        client.get.return_value = _uploader_response(200, {"nextExpectedRanges": [f"{resumed_offset}-"]})

        result = upload_file(client, "drive-1", "parent-1", local_path, "big.bin", "fail")

        assert result == {"id": "item-big"}
        client.get.assert_called_once_with("https://upload.example/session", absolute=True, auth=False, retry=False)
        resumed_range = client.put.call_args_list[2].kwargs["headers"]["Content-Range"]
        expected_chunk_end = min(resumed_offset + CHUNK_SIZE, size) - 1
        assert resumed_range == f"bytes {resumed_offset}-{expected_chunk_end}/{size}"

    def test_graph_connection_error_resumes_from_next_expected_ranges(self, tmp_path):
        """A connection error/timeout at the `GraphClient` transport boundary now surfaces as
        `GraphConnectionError` (a `GraphClientError` subclass with no `status_code`) instead of a
        raw `requests.RequestException` — the resume dance must treat it as transient, same as a
        503, rather than aborting immediately because `None not in _TRANSIENT_CHUNK_STATUSES`."""
        size = 25 * 1024 * 1024
        local_path = _make_sparse_file(tmp_path / "big.bin", size)
        client = MagicMock()
        client.post.return_value = _uploader_response(200, {"uploadUrl": "https://upload.example/session"})
        resumed_offset = CHUNK_SIZE + 2048
        client.put.side_effect = [
            _uploader_response(202),  # chunk 1 ok
            GraphConnectionError("connection refused"),  # chunk 2: connection error
            _uploader_response(202),  # resumed partial chunk accepted
            _uploader_response(201, {"id": "item-big"}),  # final chunk
        ]
        client.get.return_value = _uploader_response(200, {"nextExpectedRanges": [f"{resumed_offset}-"]})

        result = upload_file(client, "drive-1", "parent-1", local_path, "big.bin", "fail")

        assert result == {"id": "item-big"}
        client.get.assert_called_once_with("https://upload.example/session", absolute=True, auth=False, retry=False)
        resumed_range = client.put.call_args_list[2].kwargs["headers"]["Content-Range"]
        expected_chunk_end = min(resumed_offset + CHUNK_SIZE, size) - 1
        assert resumed_range == f"bytes {resumed_offset}-{expected_chunk_end}/{size}"

    def test_session_404_restarts_once(self, tmp_path):
        size = 5
        local_path = _make_sparse_file(tmp_path / "small.bin", SIMPLE_UPLOAD_THRESHOLD + size)
        client = MagicMock()
        client.post.side_effect = [
            _uploader_response(200, {"uploadUrl": "https://upload.example/session-1"}),
            _uploader_response(200, {"uploadUrl": "https://upload.example/session-2"}),
        ]
        client.put.side_effect = [
            _uploader_graph_error(404),  # session-1's first chunk: session is gone
            _uploader_response(202),  # session-2's first chunk
            _uploader_response(201, {"id": "item-1"}),  # session-2's final chunk
        ]

        result = upload_file(client, "drive-1", "parent-1", local_path, "small.bin", "fail")

        assert result == {"id": "item-1"}
        assert client.post.call_count == 2
        urls_used = [c.args[0] for c in client.put.call_args_list]
        assert urls_used == [
            "https://upload.example/session-1",
            "https://upload.example/session-2",
            "https://upload.example/session-2",
        ]

    def test_session_404_twice_gives_up(self, tmp_path):
        size = 5
        local_path = _make_sparse_file(tmp_path / "small.bin", SIMPLE_UPLOAD_THRESHOLD + size)
        client = MagicMock()
        client.post.return_value = _uploader_response(200, {"uploadUrl": "https://upload.example/session"})
        client.put.side_effect = _uploader_graph_error(404)

        with pytest.raises(UploadSessionError):
            upload_file(client, "drive-1", "parent-1", local_path, "small.bin", "fail")

        assert client.post.call_count == 2

    def test_unrecoverable_failure_deletes_session_and_raises(self, tmp_path):
        size = 5
        local_path = _make_sparse_file(tmp_path / "small.bin", SIMPLE_UPLOAD_THRESHOLD + size)
        client = MagicMock()
        client.post.return_value = _uploader_response(200, {"uploadUrl": "https://upload.example/session"})
        # Every chunk PUT attempt fails transiently, and every resume status check succeeds but
        # never advances beyond the resume-attempt budget.
        client.put.side_effect = _uploader_graph_error(503)
        client.get.return_value = _uploader_response(200, {"nextExpectedRanges": ["0-"]})

        with pytest.raises(UploadSessionError):
            upload_file(client, "drive-1", "parent-1", local_path, "small.bin", "fail")

        assert client.put.call_count == MAX_RESUME_ATTEMPTS + 1
        client.delete.assert_called_once_with("https://upload.example/session", absolute=True, auth=False, retry=False)

    def test_non_retryable_status_aborts_immediately_without_resuming(self, tmp_path):
        size = 5
        local_path = _make_sparse_file(tmp_path / "small.bin", SIMPLE_UPLOAD_THRESHOLD + size)
        client = MagicMock()
        client.post.return_value = _uploader_response(200, {"uploadUrl": "https://upload.example/session"})
        client.put.side_effect = _uploader_graph_error(400)

        with pytest.raises(UploadSessionError):
            upload_file(client, "drive-1", "parent-1", local_path, "small.bin", "fail")

        client.get.assert_not_called()
        client.delete.assert_called_once()

    def test_final_chunk_conflict_fail_raises_file_already_exists(self, tmp_path):
        size = 5
        local_path = _make_sparse_file(tmp_path / "small.bin", SIMPLE_UPLOAD_THRESHOLD + size)
        client = MagicMock()
        client.post.return_value = _uploader_response(200, {"uploadUrl": "https://upload.example/session"})
        client.put.side_effect = [
            _uploader_response(202),
            _uploader_graph_error(409, "nameAlreadyExists"),
        ]

        with pytest.raises(FileAlreadyExistsError, match="small.bin"):
            upload_file(client, "drive-1", "parent-1", local_path, "small.bin", "fail")

        client.delete.assert_not_called()

    @pytest.mark.parametrize("conflict_behavior", ["replace", "rename"])
    def test_final_chunk_conflict_replace_or_rename_raises_clear_retryable_error(self, tmp_path, conflict_behavior):
        size = 5
        local_path = _make_sparse_file(tmp_path / "small.bin", SIMPLE_UPLOAD_THRESHOLD + size)
        client = MagicMock()
        client.post.return_value = _uploader_response(200, {"uploadUrl": "https://upload.example/session"})
        client.put.side_effect = [
            _uploader_response(202),
            _uploader_graph_error(409, "nameAlreadyExists"),
        ]

        with pytest.raises(UploadSessionError, match="late name conflict"):
            upload_file(client, "drive-1", "parent-1", local_path, "small.bin", conflict_behavior)


# ----------------------------------------------------------------------------------------------
# client.headers — v1-parity table-header normalization
# ----------------------------------------------------------------------------------------------


class TestToAscii:
    def test_space_becomes_underscore(self):
        assert to_ascii("Col 1") == "Col_1"

    def test_diacritics_are_stripped_to_base_letter(self):
        # v1 fixture: worksheet name "Zošit 2" (untouched — only header *cells* go through
        # `toAscii`, not sheet names); this exercises the same NFD + combining-mark-strip
        # machinery on a header cell instead.
        assert to_ascii("Zošit") == "Zosit"

    def test_run_of_disallowed_characters_collapses_to_one_underscore(self):
        assert to_ascii("a   b") == "a_b"
        assert to_ascii("a!!!b") == "a_b"

    def test_leading_and_trailing_underscores_are_trimmed(self):
        assert to_ascii("  leading and trailing  ") == "leading_and_trailing"

    def test_dot_and_hyphen_are_preserved(self):
        assert to_ascii("col-1.2") == "col-1.2"

    def test_digits_are_preserved(self):
        assert to_ascii("2024") == "2024"

    def test_empty_string_stays_empty(self):
        assert to_ascii("") == ""

    def test_equals_sign_is_not_a_reserved_character(self):
        # v1 fixture: a worksheet literally named "sheet=4" — `=` isn't ASCII-folded away (it's
        # not alnum/`-`/`.`, so it *would* normally become `_`, but this asserts the header
        # normalizer's behavior on a similarly "special" character for completeness).
        assert to_ascii("sheet=4") == "sheet_4"


class TestNormalizeHeaderRow:
    def test_v1_fixture_col_1_2_3(self):
        assert normalize_header_row(["Col 1", "Col 2", "Col 3"]) == ["Col_1", "Col_2", "Col_3"]

    def test_empty_sheet_single_empty_cell_normalizes_to_no_header(self):
        # Graph still returns a one-cell row (`text: [[""]]`) for a genuinely empty worksheet;
        # v1 special-cases this as "no header at all" rather than a one-column sheet.
        assert normalize_header_row([""]) == []

    def test_zero_cells_normalizes_to_no_header(self):
        assert normalize_header_row([]) == []

    def test_blank_cell_becomes_positional_column_name(self):
        assert normalize_header_row(["a", "", "c"]) == ["a", "column-2", "c"]

    def test_duplicate_names_get_dash_suffixes_in_column_order(self):
        assert normalize_header_row(["id", "id", "id"]) == ["id", "id-1", "id-2"]

    def test_duplicate_after_normalization_still_gets_suffixed(self):
        # "Col!" and "Col?" both normalize to "Col_" before trimming to "Col" — still a collision.
        assert normalize_header_row(["Col", "Col!"]) == ["Col", "Col-1"]

    def test_row_with_multiple_blank_cells_is_not_the_empty_shortcut(self):
        # `len(cells) <= 1` is the empty-sheet shortcut; three blank cells go through normal
        # positional naming instead (v1: `TableHeader::parseColumns`).
        assert normalize_header_row(["", "", ""]) == ["column-1", "column-2", "column-3"]


# ----------------------------------------------------------------------------------------------
# configuration — Pydantic models (Account/Destination/Workbook/Worksheet/RowConfig/...)
# ----------------------------------------------------------------------------------------------


class TestAccount:
    def test_private_onedrive_requires_nothing(self):
        account = Account(account_type="private_onedrive")
        assert account.account_type == AccountType.PRIVATE_ONEDRIVE
        assert account.tenant_id is None
        assert account.site_url is None

    def test_business_requires_tenant_id(self):
        with pytest.raises(ValidationError, match="account.tenant_id is required"):
            Account(account_type="onedrive_for_business")

    def test_business_with_tenant_id_is_valid(self):
        account = Account(account_type="onedrive_for_business", tenant_id="tenant-1")
        assert account.tenant_id == "tenant-1"

    def test_sharepoint_requires_tenant_id_and_site_url(self):
        with pytest.raises(ValidationError, match="account.tenant_id is required"):
            Account(account_type="sharepoint", site_url="https://contoso.sharepoint.com/sites/x")

    def test_sharepoint_requires_site_url_even_with_tenant_id(self):
        with pytest.raises(ValidationError, match="account.site_url is required"):
            Account(account_type="sharepoint", tenant_id="tenant-1")

    def test_sharepoint_with_both_is_valid(self):
        account = Account(
            account_type="sharepoint",
            tenant_id="tenant-1",
            site_url="https://contoso.sharepoint.com/sites/x",
        )
        assert account.tenant_id == "tenant-1"
        assert account.site_url == "https://contoso.sharepoint.com/sites/x"

    def test_invalid_account_type_raises(self):
        with pytest.raises(ValidationError):
            Account(account_type="not_a_real_type")

    def test_unknown_fields_are_ignored(self):
        account = Account(account_type="private_onedrive", some_future_field="x")
        assert not hasattr(account, "some_future_field")


class TestDestination:
    def test_defaults(self):
        destination = Destination()
        assert destination.drive_id is None
        assert destination.folder_path is None
        assert destination.date is None
        assert destination.conflict_behavior == ConflictBehavior.FAIL

    def test_invalid_conflict_behavior_raises(self):
        with pytest.raises(ValidationError):
            Destination(conflict_behavior="overwrite")

    def test_date_accepts_a_free_form_string(self):
        destination = Destination(date="yesterday")
        assert destination.date == "yesterday"

    def test_ui_shaped_payload_with_blank_untouched_fields_normalizes_to_defaults(self):
        # The Keboola UI submits an untouched text/select field as "" — including
        # `conflict_behavior`, whose declared default ("fail") must still apply rather than
        # failing enum validation on the empty string.
        destination = Destination.model_validate(
            {"drive_id": "", "folder_path": "", "date": "", "conflict_behavior": ""}
        )
        assert destination.drive_id is None
        assert destination.folder_path is None
        assert destination.date is None
        assert destination.conflict_behavior == ConflictBehavior.FAIL

    def test_row_schema_helper_account_type_field_is_ignored(self):
        # `configRowSchema.json`'s hidden `destination.helper_account_type` (root-watch UX
        # addition — gates the Document Library dropdown to SharePoint accounts in the UI) is
        # submitted alongside the real fields; `Destination`'s `extra="ignore"` must tolerate it.
        destination = Destination.model_validate(
            {"helper_account_type": "sharepoint", "drive_id": "drive-1", "folder_path": "reports"}
        )
        assert destination.drive_id == "drive-1"
        assert destination.folder_path == "reports"
        assert not hasattr(destination, "helper_account_type")


class TestCsvOptions:
    def test_defaults(self):
        csv_options = CsvOptions()
        assert csv_options.file_name is None
        assert csv_options.delimiter == ","
        assert csv_options.enclosure == '"'
        assert csv_options.include_header is True


class TestWorkbook:
    def test_path_only_is_valid(self):
        workbook = Workbook(path="/reports/book.xlsx")
        assert workbook.path == "/reports/book.xlsx"

    def test_ids_together_are_valid(self):
        workbook = Workbook(drive_id="drive-1", file_id="file-1")
        assert workbook.drive_id == "drive-1"
        assert workbook.file_id == "file-1"

    def test_drive_id_alone_is_invalid(self):
        with pytest.raises(ValidationError, match="must be provided together"):
            Workbook(drive_id="drive-1")

    def test_file_id_alone_is_invalid(self):
        with pytest.raises(ValidationError, match="must be provided together"):
            Workbook(file_id="file-1")

    def test_path_and_ids_together_is_invalid(self):
        # Change B: humanized legacy (no `targeting`) mutual-exclusivity message.
        with pytest.raises(ValidationError, match="Choose one way to target the workbook"):
            Workbook(path="/book.xlsx", drive_id="drive-1", file_id="file-1")

    def test_neither_path_nor_ids_is_invalid(self):
        with pytest.raises(ValidationError, match="requires either workbook.path"):
            Workbook()

    def test_metadata_is_accepted_and_passed_through(self):
        workbook = Workbook(path="/book.xlsx", metadata={"pickerId": "abc", "nested": [1, 2]})
        assert workbook.metadata == {"pickerId": "abc", "nested": [1, 2]}

    def test_metadata_defaults_to_none(self):
        workbook = Workbook(path="/book.xlsx")
        assert workbook.metadata is None

    def test_ui_shaped_payload_with_blank_untouched_path_resolves_to_ids_mode(self):
        # The Keboola UI submits an untouched text field as "" (not omitted, not null) — a row
        # edited to pick Library + Workbook from the dropdowns still carries the Path field's
        # placeholder value verbatim.
        workbook = Workbook.model_validate({"path": "", "drive_id": "b!drive-id", "file_id": "01file-id"})
        assert workbook.path is None
        assert workbook.drive_id == "b!drive-id"
        assert workbook.file_id == "01file-id"

    def test_blank_path_alone_is_still_invalid(self):
        with pytest.raises(ValidationError, match="requires either workbook.path"):
            Workbook.model_validate({"path": "  "})


class TestWorksheet:
    def test_id_only_is_valid(self):
        worksheet = Worksheet(id="sheet-1")
        assert worksheet.id == "sheet-1"

    def test_position_only_is_valid(self):
        worksheet = Worksheet(position=2)
        assert worksheet.position == 2

    def test_name_only_is_valid(self):
        worksheet = Worksheet(name="Sheet1")
        assert worksheet.name == "Sheet1"

    def test_position_string_is_coerced_to_int(self):
        worksheet = Worksheet(position="0")
        assert worksheet.position == 0
        assert isinstance(worksheet.position, int)

    def test_non_numeric_position_string_raises(self):
        with pytest.raises(ValidationError, match="must be an integer"):
            Worksheet(position="first")

    def test_id_and_position_together_is_invalid(self):
        # Change C: humanized legacy (no `selection`) mutual-exclusivity message.
        with pytest.raises(ValidationError, match="Choose one way to target the worksheet"):
            Worksheet(id="sheet-1", position=0)

    def test_id_and_name_together_is_valid(self):
        worksheet = Worksheet(id="sheet-1", name="Renamed")
        assert worksheet.id == "sheet-1"
        assert worksheet.name == "Renamed"

    def test_position_and_name_together_is_valid(self):
        worksheet = Worksheet(position=1, name="Renamed")
        assert worksheet.position == 1
        assert worksheet.name == "Renamed"

    def test_none_of_id_name_position_is_invalid(self):
        with pytest.raises(ValidationError, match="at least one of id, name, or position"):
            Worksheet()

    def test_metadata_is_accepted_and_passed_through(self):
        worksheet = Worksheet(name="Sheet1", metadata={"pickerId": "abc"})
        assert worksheet.metadata == {"pickerId": "abc"}

    def test_ui_shaped_payload_with_blank_untouched_id_and_position_resolves_to_name_mode(self):
        # A row edited via the UI to target a worksheet by name still carries the ID/Position
        # fields' untouched "" placeholder values.
        worksheet = Worksheet.model_validate({"name": "X", "id": "", "position": ""})
        assert worksheet.name == "X"
        assert worksheet.id is None
        assert worksheet.position is None

    def test_blank_id_and_position_alone_is_still_invalid(self):
        with pytest.raises(ValidationError, match="at least one of id, name, or position"):
            Worksheet.model_validate({"id": "", "position": "", "name": ""})


class TestRowConfig:
    def _account_params(self, **overrides):
        params = {"account_type": "private_onedrive"}
        params.update(overrides)
        return params

    def test_file_mode_minimal_config(self):
        config = RowConfig(mode="file", account=self._account_params())
        assert config.mode == Mode.FILE
        assert config.write_mode == WriteMode.OVERWRITE
        assert config.key_columns == []
        assert config.batch_size == 5000
        assert isinstance(config.destination, Destination)
        assert isinstance(config.csv, CsvOptions)
        assert config.workbook is None
        assert config.worksheet is None

    def test_worksheet_mode_requires_workbook_and_worksheet(self):
        with pytest.raises(ValidationError, match="workbook configuration is required"):
            RowConfig(mode="worksheet", account=self._account_params())

    def test_worksheet_mode_requires_worksheet_even_with_workbook(self):
        with pytest.raises(ValidationError, match="worksheet configuration is required"):
            RowConfig(
                mode="worksheet",
                account=self._account_params(),
                workbook={"path": "/book.xlsx"},
            )

    def test_worksheet_mode_with_workbook_and_worksheet_is_valid(self):
        config = RowConfig(
            mode="worksheet",
            account=self._account_params(),
            workbook={"path": "/book.xlsx"},
            worksheet={"name": "Sheet1"},
        )
        assert config.workbook.path == "/book.xlsx"
        assert config.worksheet.name == "Sheet1"

    def test_invalid_mode_raises(self):
        with pytest.raises(ValidationError):
            RowConfig(mode="not_a_mode", account=self._account_params())

    def test_zero_batch_size_raises(self):
        # IMPORTANT-4 (phase 8 audit): 0 would silently write nothing (batched in chunks of zero
        # rows) — reject it as a normal configuration error instead.
        with pytest.raises(ValidationError, match="batch_size"):
            RowConfig(mode="file", account=self._account_params(), batch_size=0)

    def test_negative_batch_size_raises(self):
        # A negative value used to reach the Excel writer's batching helper and blow up there
        # with an unmapped `ValueError` (exit 2) instead of a clean configuration error.
        with pytest.raises(ValidationError, match="batch_size"):
            RowConfig(mode="file", account=self._account_params(), batch_size=-1)

    def test_positive_batch_size_is_valid(self):
        config = RowConfig(mode="file", account=self._account_params(), batch_size=1)
        assert config.batch_size == 1

    def test_unknown_top_level_keys_are_ignored(self):
        config = RowConfig(
            mode="file",
            account=self._account_params(),
            some_future_top_level_field="x",
        )
        assert not hasattr(config, "some_future_top_level_field")

    def test_full_merged_parameters_shape_from_spec(self):
        merged_parameters = {
            "mode": "table_excel",  # legacy alias — also exercises Change A's normalization end-to-end
            "account": {
                "account_type": "sharepoint",
                "tenant_id": "00000000-0000-0000-0000-000000000000",
                "site_url": "https://contoso.sharepoint.com/sites/DataTeam",
            },
            "destination": {
                "drive_id": "b!driveid",
                "folder_path": "reports/{date:%Y-%m-%d}",
                "conflict_behavior": "replace",
            },
            "csv": {
                "file_name": "output.csv",
                "delimiter": ";",
                "enclosure": "'",
                "include_header": False,
            },
            "workbook": {
                "drive_id": "b!driveid",
                "file_id": "01ABCDEF",
                "metadata": {"pickerOrigin": "oneDrivePicker"},
            },
            "worksheet": {
                "position": "0",
                "name": "Sheet1",
                "metadata": {"pickerOrigin": "oneDrivePicker"},
            },
            "append": True,
            "batch_size": 2500,
        }

        config = RowConfig.model_validate(merged_parameters)

        assert config.mode == Mode.WORKSHEET
        assert config.account.account_type == AccountType.SHAREPOINT
        assert config.destination.conflict_behavior == ConflictBehavior.REPLACE
        assert config.csv.delimiter == ";"
        assert config.workbook.file_id == "01ABCDEF"
        assert config.worksheet.position == 0
        assert config.write_mode == WriteMode.APPEND
        assert config.batch_size == 2500

    def test_account_error_propagates_through_row_config(self):
        with pytest.raises(ValidationError, match="account.tenant_id is required"):
            RowConfig(mode="file", account={"account_type": "sharepoint"})


class TestWriteModeAndKeyColumns:
    """Change 3: `write_mode` (enum) replaces `append: bool`; `key_columns` is required
    (non-empty) when `write_mode` is 'upsert'. A pre-existing `append: true`/`false` (every
    already-recorded VCR cassette and every platform row created before this change) is silently
    normalized to its `write_mode` equivalent."""

    def _account_params(self, **overrides):
        params = {"account_type": "private_onedrive"}
        params.update(overrides)
        return params

    def _worksheet_row(self, **overrides):
        params = {
            "mode": "worksheet",
            "account": self._account_params(),
            "workbook": {"path": "/book.xlsx"},
            "worksheet": {"name": "Sheet1"},
        }
        params.update(overrides)
        return params

    def test_default_write_mode_is_overwrite(self):
        config = RowConfig(**self._worksheet_row())
        assert config.write_mode == WriteMode.OVERWRITE
        assert config.key_columns == []

    def test_write_mode_upsert_accepted_with_key_columns(self):
        config = RowConfig(**self._worksheet_row(write_mode="upsert", key_columns=["id"]))
        assert config.write_mode == WriteMode.UPSERT
        assert config.key_columns == ["id"]

    def test_write_mode_upsert_with_composite_key_columns(self):
        config = RowConfig(**self._worksheet_row(write_mode="upsert", key_columns=["region", "id"]))
        assert config.key_columns == ["region", "id"]

    def test_write_mode_upsert_without_key_columns_raises(self):
        with pytest.raises(ValidationError, match="key_columns is required"):
            RowConfig(**self._worksheet_row(write_mode="upsert"))

    def test_write_mode_upsert_with_empty_key_columns_list_raises(self):
        with pytest.raises(ValidationError, match="key_columns is required"):
            RowConfig(**self._worksheet_row(write_mode="upsert", key_columns=[]))

    def test_key_columns_without_upsert_is_not_required(self):
        config = RowConfig(**self._worksheet_row(write_mode="append"))
        assert config.key_columns == []

    def test_invalid_write_mode_raises(self):
        with pytest.raises(ValidationError):
            RowConfig(**self._worksheet_row(write_mode="not_a_mode"))

    def test_legacy_append_true_normalizes_to_write_mode_append(self):
        config = RowConfig(**self._worksheet_row(append=True))
        assert config.write_mode == WriteMode.APPEND

    def test_legacy_append_false_normalizes_to_write_mode_overwrite(self):
        config = RowConfig(**self._worksheet_row(append=False))
        assert config.write_mode == WriteMode.OVERWRITE

    def test_legacy_append_absent_defaults_to_overwrite(self):
        config = RowConfig(**self._worksheet_row())
        assert config.write_mode == WriteMode.OVERWRITE

    def test_explicit_write_mode_wins_over_legacy_append_if_both_present(self):
        # Not a real config shape (no such row exists today), but the before-validator must not
        # let a stray legacy `append` clobber an explicit `write_mode`.
        config = RowConfig(**self._worksheet_row(append=True, write_mode="overwrite"))
        assert config.write_mode == WriteMode.OVERWRITE


class TestModeAliasNormalization:
    """Change A: the pre-merge mode names are silently accepted and normalized — every
    already-recorded VCR cassette config and every platform row created before this change uses
    them, and none of them should ever need to be re-saved or re-recorded."""

    def _account_params(self):
        return {"account_type": "private_onedrive"}

    def test_table_csv_alias_normalizes_to_file(self):
        config = RowConfig(mode="table_csv", account=self._account_params())
        assert config.mode == Mode.FILE

    def test_table_excel_alias_normalizes_to_worksheet(self):
        config = RowConfig(
            mode="table_excel",
            account=self._account_params(),
            workbook={"path": "/book.xlsx"},
            worksheet={"name": "Sheet1"},
        )
        assert config.mode == Mode.WORKSHEET

    def test_canonical_names_pass_through_unchanged(self):
        assert RowConfig(mode="file", account=self._account_params()).mode == Mode.FILE

    def test_unrecognized_mode_value_still_raises(self):
        # A typo'd/unknown mode must not be silently swallowed by the alias lookup.
        with pytest.raises(ValidationError):
            RowConfig(mode="not_a_real_mode", account=self._account_params())


class TestWorkbookTargeting:
    """Change B: `workbook.targeting` makes the "pick via dropdowns" vs. "by path" choice
    explicit, ignoring whatever the *other* form's field(s) happen to hold — a stale hidden value
    left over from switching `targeting` back and forth in the UI can never break validation or
    get used by mistake."""

    def test_pick_requires_both_ids(self):
        with pytest.raises(ValidationError, match="both workbook.drive_id and workbook.file_id are required"):
            Workbook(targeting="pick")

    def test_pick_with_only_drive_id_is_invalid(self):
        with pytest.raises(ValidationError, match="both workbook.drive_id and workbook.file_id are required"):
            Workbook(targeting="pick", drive_id="drive-1")

    def test_pick_with_both_ids_is_valid(self):
        workbook = Workbook(targeting=WorkbookTargeting.PICK, drive_id="drive-1", file_id="file-1")
        assert workbook.drive_id == "drive-1"
        assert workbook.file_id == "file-1"
        assert workbook.path is None

    def test_pick_ignores_a_stale_hidden_path(self):
        workbook = Workbook(targeting="pick", drive_id="drive-1", file_id="file-1", path="/stale.xlsx")
        assert workbook.path is None
        assert workbook.drive_id == "drive-1"
        assert workbook.file_id == "file-1"

    def test_path_requires_path(self):
        with pytest.raises(ValidationError, match='"By path": workbook.path is required'):
            Workbook(targeting="path")

    def test_path_with_path_is_valid(self):
        workbook = Workbook(targeting=WorkbookTargeting.PATH, path="/book.xlsx")
        assert workbook.path == "/book.xlsx"
        assert workbook.drive_id is None
        assert workbook.file_id is None

    def test_path_ignores_stale_hidden_ids(self):
        workbook = Workbook(targeting="path", path="/book.xlsx", drive_id="stale-drive", file_id="stale-file")
        assert workbook.path == "/book.xlsx"
        assert workbook.drive_id is None
        assert workbook.file_id is None

    def test_blank_targeting_string_is_treated_as_legacy(self):
        # UI untouched-field convention: an unset select still submits "".
        workbook = Workbook.model_validate({"targeting": "", "path": "/book.xlsx"})
        assert workbook.targeting is None
        assert workbook.path == "/book.xlsx"

    def test_legacy_ids_only_is_still_valid_without_targeting(self):
        workbook = Workbook(drive_id="drive-1", file_id="file-1")
        assert workbook.targeting is None
        assert workbook.drive_id == "drive-1"


class TestWorksheetSelection:
    """Change C: `worksheet.selection` makes the "pick existing" vs. "by name" choice explicit,
    ignoring whatever the *other* form's field(s) happen to hold (including a legacy `position`
    value) — no rename ever happens under either explicit branch."""

    def test_pick_requires_id(self):
        with pytest.raises(ValidationError, match='"Pick existing": worksheet.id is required'):
            Worksheet(selection="pick")

    def test_pick_with_id_is_valid(self):
        worksheet = Worksheet(selection=WorksheetSelection.PICK, id="sheet-1")
        assert worksheet.id == "sheet-1"
        assert worksheet.name is None

    def test_pick_ignores_stale_hidden_name_and_position(self):
        worksheet = Worksheet(selection="pick", id="sheet-1", name="StaleRenameTarget", position=3)
        assert worksheet.id == "sheet-1"
        assert worksheet.name is None  # no rename under "pick"
        assert worksheet.position is None

    def test_name_requires_name(self):
        with pytest.raises(ValidationError, match='"By name \\(creates if missing\\)": worksheet.name is required'):
            Worksheet(selection="name")

    def test_name_with_name_is_valid(self):
        worksheet = Worksheet(selection=WorksheetSelection.NAME, name="Sheet1")
        assert worksheet.name == "Sheet1"
        assert worksheet.id is None

    def test_name_ignores_stale_hidden_id_and_position(self):
        worksheet = Worksheet(selection="name", name="Sheet1", id="stale-id", position=2)
        assert worksheet.name == "Sheet1"
        assert worksheet.id is None
        assert worksheet.position is None

    def test_blank_selection_string_is_treated_as_legacy(self):
        worksheet = Worksheet.model_validate({"selection": "", "name": "Sheet1"})
        assert worksheet.selection is None
        assert worksheet.name == "Sheet1"

    def test_legacy_id_and_name_together_still_renames_without_selection(self):
        # v1-parity behavior preserved exactly when `selection` is absent (row API payloads,
        # already-recorded VCR cassette configs).
        worksheet = Worksheet(id="sheet-1", name="Renamed")
        assert worksheet.selection is None
        assert worksheet.id == "sheet-1"
        assert worksheet.name == "Renamed"


class TestOffModeSectionsDropped:
    """File-mode rows must ignore stray hidden-section defaults the UI saves (seen live)."""

    def test_file_mode_ignores_stray_workbook_and_worksheet_defaults(self):
        config = RowConfig.model_validate(
            {
                "mode": "file",
                "account": {"account_type": "private_onedrive"},
                "workbook": {"targeting": "pick"},
                "worksheet": {"selection": "pick"},
            }
        )
        assert config.workbook is None
        assert config.worksheet is None

    def test_worksheet_mode_still_validates_workbook_strictly(self):
        import pytest

        with pytest.raises(Exception, match="drive_id"):
            RowConfig.model_validate(
                {
                    "mode": "worksheet",
                    "account": {"account_type": "sharepoint", "tenant_id": "t", "site_url": "https://x"},
                    "workbook": {"targeting": "pick"},
                    "worksheet": {"selection": "name", "name": "S"},
                }
            )
