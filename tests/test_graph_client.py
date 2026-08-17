from unittest.mock import MagicMock, patch

import pytest

from client.exceptions import (
    GraphBadRequestError,
    GraphClientError,
    GraphNotFoundError,
    GraphPermissionError,
    GraphQuotaExceededError,
    GraphRateLimitCapExceededError,
)
from client.graph_client import BASE_URL, USER_AGENT, GraphClient


def _response(status_code: int, json_body: dict | None = None, headers: dict | None = None, text: str = ""):
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


def _error_body(code: str, message: str) -> dict:
    return {"error": {"code": code, "message": message}}


def _token_provider(token: str = "token-1"):
    provider = MagicMock()
    provider.get_access_token.return_value = token
    return provider


def _client(session=None, token_provider=None, **kwargs) -> GraphClient:
    return GraphClient(
        token_provider=token_provider or _token_provider(),
        session=session or MagicMock(),
        **kwargs,
    )


class TestRequestHeaders:
    def test_sends_authorization_user_agent_and_base_url(self):
        session = MagicMock()
        session.request.return_value = _response(200, {"ok": True})
        provider = _token_provider("abc-token")
        client = _client(session=session, token_provider=provider)

        client.get("/me")

        args, kwargs = session.request.call_args
        assert args[0] == "GET"
        assert args[1] == f"{BASE_URL}/me"
        assert kwargs["headers"]["Authorization"] == "Bearer abc-token"
        assert kwargs["headers"]["User-Agent"] == USER_AGENT
        assert USER_AGENT.startswith("NONISV|Keboola|wr-onedrive-v2/")

    def test_auth_false_omits_authorization_header(self):
        session = MagicMock()
        session.request.return_value = _response(200, {"ok": True})
        provider = _token_provider()
        client = _client(session=session, token_provider=provider)

        client.put("https://upload.example/session-url", absolute=True, auth=False, data=b"chunk")

        provider.get_access_token.assert_not_called()
        _, kwargs = session.request.call_args
        assert "Authorization" not in kwargs["headers"]

    def test_absolute_url_is_not_prefixed_with_base_url(self):
        session = MagicMock()
        session.request.return_value = _response(200, {"ok": True})
        client = _client(session=session)

        client.get("https://graph.microsoft.com/v1.0/absolute/path", absolute=True)

        args, _ = session.request.call_args
        assert args[1] == "https://graph.microsoft.com/v1.0/absolute/path"


class TestRetryAfter:
    @patch("client.graph_client.time.sleep")
    def test_honors_retry_after_header_in_seconds(self, mock_sleep):
        session = MagicMock()
        session.request.side_effect = [
            _response(429, _error_body("TooManyRequests", "slow down"), headers={"Retry-After": "7"}),
            _response(200, {"ok": True}),
        ]
        client = _client(session=session)

        response = client.get("/me")

        assert response.ok
        mock_sleep.assert_called_once_with(7.0)

    @patch("client.graph_client.time.sleep")
    def test_503_without_retry_after_falls_back_to_backoff(self, mock_sleep):
        session = MagicMock()
        session.request.side_effect = [
            _response(503, _error_body("ServiceUnavailable", "busy")),
            _response(200, {"ok": True}),
        ]
        client = _client(session=session)

        client.get("/me")

        mock_sleep.assert_called_once_with(1.0)


class TestServerErrorBackoff:
    @patch("client.graph_client.time.sleep")
    def test_5xx_backs_off_then_succeeds(self, mock_sleep):
        session = MagicMock()
        session.request.side_effect = [
            _response(500, _error_body("InternalServerError", "oops")),
            _response(502, _error_body("BadGateway", "oops")),
            _response(200, {"ok": True}),
        ]
        client = _client(session=session)

        response = client.get("/me")

        assert response.ok
        assert mock_sleep.call_args_list == [((1.0,),), ((2.0,),)]

    @patch("client.graph_client.time.sleep")
    def test_5xx_exhausting_cap_raises_rate_limit_cap_error(self, mock_sleep):
        session = MagicMock()
        # Backoff sequence 1, 2, 4, 8, ... will exceed a tiny cap quickly.
        session.request.return_value = _response(500, _error_body("InternalServerError", "oops"))
        client = _client(session=session, total_wait_cap_seconds=1.5)

        with pytest.raises(GraphRateLimitCapExceededError):
            client.get("/me")


class TestRateLimitCap:
    def test_single_retry_after_exceeding_cap_raises_immediately(self):
        session = MagicMock()
        session.request.return_value = _response(
            429, _error_body("TooManyRequests", "slow down"), headers={"Retry-After": "600"}
        )
        client = _client(session=session, total_wait_cap_seconds=300.0)

        with pytest.raises(GraphRateLimitCapExceededError) as exc_info:
            client.get("/me")

        assert "600" in str(exc_info.value) or "600s" in str(exc_info.value)
        session.request.assert_called_once()

    @patch("client.graph_client.time.sleep")
    def test_cumulative_retries_exhausting_cap_raises(self, mock_sleep):
        session = MagicMock()
        session.request.side_effect = [
            _response(429, _error_body("TooManyRequests", "slow"), headers={"Retry-After": "5"}),
            _response(429, _error_body("TooManyRequests", "slow"), headers={"Retry-After": "5"}),
            _response(429, _error_body("TooManyRequests", "slow"), headers={"Retry-After": "5"}),
        ]
        client = _client(session=session, total_wait_cap_seconds=8.0)

        with pytest.raises(GraphRateLimitCapExceededError):
            client.get("/me")

        # First 5s wait succeeds (elapsed=5 <= 8), second 5s wait would bring elapsed to 10 > 8.
        mock_sleep.assert_called_once_with(5.0)


class TestWorkbookTransientOptIn:
    @patch("client.graph_client.time.sleep")
    def test_409_not_retried_without_opt_in(self, mock_sleep):
        session = MagicMock()
        session.request.return_value = _response(409, _error_body("nameAlreadyExists", "exists"))
        client = _client(session=session)

        with pytest.raises(GraphClientError):
            client.get("/me")

        session.request.assert_called_once()
        mock_sleep.assert_not_called()

    @patch("client.graph_client.time.sleep")
    def test_409_retried_with_opt_in(self, mock_sleep):
        session = MagicMock()
        session.request.side_effect = [
            _response(409, _error_body("EditModeCannotAcquireLockTooManyRequests", "locked")),
            _response(200, {"ok": True}),
        ]
        client = _client(session=session)

        response = client.get("/me", retry_transient_workbook=True)

        assert response.ok
        mock_sleep.assert_called_once()

    @patch("client.graph_client.time.sleep")
    def test_405_retried_with_opt_in(self, mock_sleep):
        session = MagicMock()
        session.request.side_effect = [
            _response(405, _error_body("MethodNotAllowed", "not allowed")),
            _response(200, {"ok": True}),
        ]
        client = _client(session=session)

        response = client.patch("/me", retry_transient_workbook=True)

        assert response.ok


class TestUnauthorizedRetry:
    def test_401_refetches_token_and_retries_once(self):
        session = MagicMock()
        session.request.side_effect = [
            _response(401, _error_body("InvalidAuthenticationToken", "expired")),
            _response(200, {"ok": True}),
        ]
        provider = MagicMock()
        provider.get_access_token.side_effect = ["token-old", "token-new"]
        client = _client(session=session, token_provider=provider)

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
            _response(401, _error_body("InvalidAuthenticationToken", "expired")),
            _response(200, {"ok": True}),
        ]
        provider = MagicMock()
        provider.get_access_token.side_effect = ["token-old", "token-new"]
        client = _client(session=session, token_provider=provider)

        client.get("/me")

        provider.invalidate.assert_called_once()
        # invalidate() must happen strictly between the two get_access_token() calls.
        call_order = [call[0] for call in provider.method_calls]
        assert call_order == ["get_access_token", "invalidate", "get_access_token"]

    def test_401_does_not_invalidate_when_auth_is_false(self):
        session = MagicMock()
        session.request.return_value = _response(401, _error_body("InvalidAuthenticationToken", "expired"))
        provider = MagicMock()
        client = _client(session=session, token_provider=provider)

        with pytest.raises(GraphPermissionError):
            client.put("https://upload.example/session-url", absolute=True, auth=False)

        provider.invalidate.assert_not_called()

    def test_401_persisting_after_retry_raises_permission_error(self):
        session = MagicMock()
        session.request.return_value = _response(401, _error_body("InvalidAuthenticationToken", "still bad"))
        client = _client(session=session)

        with pytest.raises(GraphPermissionError):
            client.get("/me")

        assert session.request.call_count == 2


class TestNoRetryFlag:
    @patch("client.graph_client.time.sleep")
    def test_retry_false_raises_immediately_on_first_error(self, mock_sleep):
        session = MagicMock()
        session.request.return_value = _response(503, _error_body("ServiceUnavailable", "busy"))
        client = _client(session=session)

        with pytest.raises(GraphClientError):
            client.put("/drives/1/items/2/content", retry=False)

        session.request.assert_called_once()
        mock_sleep.assert_not_called()

    def test_retry_false_does_not_retry_401(self):
        session = MagicMock()
        session.request.return_value = _response(401, _error_body("InvalidAuthenticationToken", "expired"))
        provider = _token_provider()
        client = _client(session=session, token_provider=provider)

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
        session.request.return_value = _response(status_code, _error_body("SomeCode", "Some message"))
        client = _client(session=session)

        with pytest.raises(expected_exception) as exc_info:
            client.get("/me", retry=False)

        assert exc_info.value.status_code == status_code
        assert exc_info.value.error_code == "SomeCode"
        assert "Some message" in str(exc_info.value)

    def test_non_json_error_body_falls_back_to_response_text(self):
        session = MagicMock()
        session.request.return_value = _response(404, json_body=None, text="Not Found")
        client = _client(session=session)

        with pytest.raises(GraphNotFoundError) as exc_info:
            client.get("/me", retry=False)

        assert "Not Found" in str(exc_info.value)
        assert exc_info.value.error_code is None


class TestPaging:
    def test_get_paged_yields_items_across_two_pages(self):
        session = MagicMock()
        session.request.side_effect = [
            _response(
                200,
                {"value": [{"id": 1}, {"id": 2}], "@odata.nextLink": "https://graph.microsoft.com/v1.0/next-page"},
            ),
            _response(200, {"value": [{"id": 3}]}),
        ]
        client = _client(session=session)

        items = list(client.get_paged("/sites/site-1/drives"))

        assert items == [{"id": 1}, {"id": 2}, {"id": 3}]
        assert session.request.call_count == 2
        second_call_args, _ = session.request.call_args_list[1]
        assert second_call_args[1] == "https://graph.microsoft.com/v1.0/next-page"

    def test_get_paged_stops_when_no_next_link(self):
        session = MagicMock()
        session.request.return_value = _response(200, {"value": [{"id": 1}]})
        client = _client(session=session)

        items = list(client.get_paged("/sites/site-1/drives"))

        assert items == [{"id": 1}]
        session.request.assert_called_once()

    def test_get_paged_defaults_to_empty_list_when_value_missing(self):
        session = MagicMock()
        session.request.return_value = _response(200, {})
        client = _client(session=session)

        assert list(client.get_paged("/sites/site-1/drives")) == []
