"""Secret-leak-prevention tests for keboola.wr-onedrive-v2 (mocked HTTP, no network).

Every assertion here proves that a secret or PII-shaped value used by the component can never
reach a log line, an exception message, or a job's stdout/stderr — as distinct from the
functional-correctness tests for the same client modules in ``tests/test_unit.py``. Gathered from
across the retired per-module unit test files into one place:

- ``client.exceptions.sanitize_exception_text`` itself (was ``tests/unit/test_exceptions.py`` in
  full) — the shared query-string-redaction helper every other test below relies on.
- ``client.auth._redact_identities`` (was ``tests/unit/test_auth.py``) — strips an AADSTS error
  description's account UPN before it can reach the job log.
- Two "IMPORTANT-6 (phase 8 audit)" regressions (were ``tests/unit/test_auth.py`` /
  ``tests/unit/test_uploader.py``): a `requests.RequestException`'s message, and an upload
  session's resume-status-check failure message, can each embed a live URL's query string
  verbatim unless routed through ``sanitize_exception_text``.
- One "IMPORTANT-2 (phase 8 audit)" regression (was ``tests/unit/test_uploader.py``): a failed
  best-effort abort-session cleanup must never log the pre-signed, credential-bearing
  ``uploadUrl`` itself.
"""

import logging
from unittest.mock import MagicMock

import pytest
import requests

from client.auth import AuthenticationError, RefreshTokenProvider, _redact_identities
from client.exceptions import GraphClientError, UploadSessionError, sanitize_exception_text
from client.uploader import SIMPLE_UPLOAD_THRESHOLD, upload_file


def _make_sparse_file(path, size: int) -> str:
    """Create a file of exactly ``size`` bytes without allocating real content (sparse file)."""
    with open(path, "wb") as fh:
        if size:
            fh.truncate(size)
    return str(path)


def _response(status_code: int, json_body=None, headers: dict | None = None) -> MagicMock:
    response = MagicMock()
    response.status_code = status_code
    response.ok = 200 <= status_code < 400
    response.headers = headers or {}
    if json_body is not None:
        response.json.return_value = json_body
    return response


def _graph_error(status_code: int, error_code: str | None = None) -> GraphClientError:
    return GraphClientError(f"HTTP {status_code}", status_code=status_code, error_code=error_code)


# -- client.exceptions.sanitize_exception_text ---------------------------------------------------


def test_sanitizes_query_string_from_exception_text():
    exc = RuntimeError("Max retries exceeded with url: https://sn.example.com/up/abc?tempauth=SECRET&x=1 (oops)")
    text = sanitize_exception_text(exc)
    assert "SECRET" not in text
    assert "tempauth" not in text
    assert "?<redacted>" in text


def test_leaves_plain_text_untouched():
    assert sanitize_exception_text(ValueError("connection reset")) == "connection reset"


# -- client.auth._redact_identities ---------------------------------------------------------------


def test_invalid_grant_detail_redacts_account_identifiers():
    """AADSTS descriptions can embed the account UPN — it must not reach the job log."""
    text = "AADSTS50034: The user account john.doe@contoso.com does not exist in tenant."
    redacted = _redact_identities(text)
    assert "john.doe@contoso.com" not in redacted
    assert "<redacted-account>" in redacted
    assert "AADSTS50034" in redacted


# -- client.auth.RefreshTokenProvider: network-failure message redaction -------------------------


def test_network_failure_message_has_query_strings_redacted():
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


# -- client.uploader.upload_file: upload-session error message / log redaction -------------------


def test_resume_status_check_failure_message_has_query_strings_redacted(tmp_path):
    """IMPORTANT-6 (phase 8 audit): when the resume status check itself fails (not a 404,
    so not a session restart), the raised `UploadSessionError` must sanitize that error's
    message — a `GraphClientError` can otherwise reproduce a query string verbatim."""
    size = 5
    local_path = _make_sparse_file(tmp_path / "small.bin", SIMPLE_UPLOAD_THRESHOLD + size)
    client = MagicMock()
    client.post.return_value = _response(200, {"uploadUrl": "https://upload.example/session"})
    client.put.side_effect = _graph_error(503)
    client.get.side_effect = GraphClientError(
        "status check failed: https://upload.example/session?tempauth=super-secret-token",
        status_code=500,
    )

    with pytest.raises(UploadSessionError) as exc_info:
        upload_file(client, "drive-1", "parent-1", local_path, "small.bin", "fail")

    assert "super-secret-token" not in str(exc_info.value)
    assert "?<redacted>" in str(exc_info.value)
    client.delete.assert_called_once_with("https://upload.example/session", absolute=True, auth=False, retry=False)


def test_abort_session_failure_does_not_log_the_upload_url(tmp_path, caplog):
    """`uploadUrl` is a pre-signed, credential-bearing URL — a failed best-effort cleanup
    DELETE must never log it (IMPORTANT-2)."""
    size = 5
    local_path = _make_sparse_file(tmp_path / "small.bin", SIMPLE_UPLOAD_THRESHOLD + size)
    secret_url = "https://upload.example/session?token=super-secret-signed-token"
    client = MagicMock()
    client.post.return_value = _response(200, {"uploadUrl": secret_url})
    client.put.side_effect = _graph_error(400)  # non-retryable -> triggers _abort_session
    client.delete.side_effect = _graph_error(500)  # cleanup DELETE itself fails

    with caplog.at_level(logging.WARNING, logger="client.uploader"), pytest.raises(UploadSessionError):
        upload_file(client, "drive-1", "parent-1", local_path, "small.bin", "fail")

    assert secret_url not in caplog.text
    assert "abandoned upload session" in caplog.text
