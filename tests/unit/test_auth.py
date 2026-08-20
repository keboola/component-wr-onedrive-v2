from unittest.mock import MagicMock

import pytest
import requests

from client.auth import (
    REFRESH_SAFETY_MARGIN_SECONDS,
    AuthenticationError,
    RefreshTokenProvider,
    _redact_identities,
)


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

    def test_network_failure_message_has_query_strings_redacted(self):
        # IMPORTANT-6 (phase 8 audit): a `requests.RequestException`'s message can embed the
        # request URL verbatim — must go through `sanitize_exception_text`, not a raw f-string.
        session = MagicMock()
        session.post.side_effect = requests.exceptions.ConnectionError(
            "failed to connect to https://login.microsoftonline.com/common/oauth2/v2.0/token?client_secret=leaked-secret"
        )
        provider = RefreshTokenProvider(
            client_id="client-1",
            client_secret="secret-1",
            authority="common",
            refresh_token_candidates=["refresh-state"],
            session=session,
        )

        with pytest.raises(AuthenticationError) as exc_info:
            provider.get_access_token()

        assert "leaked-secret" not in str(exc_info.value)
        assert "?<redacted>" in str(exc_info.value)

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


def test_invalid_grant_detail_redacts_account_identifiers():
    """AADSTS descriptions can embed the account UPN — it must not reach the job log."""
    text = "AADSTS50034: The user account john.doe@contoso.com does not exist in tenant."
    redacted = _redact_identities(text)
    assert "john.doe@contoso.com" not in redacted
    assert "<redacted-account>" in redacted
    assert "AADSTS50034" in redacted
