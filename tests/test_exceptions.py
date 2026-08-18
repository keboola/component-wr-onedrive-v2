"""Tests for the exception-text sanitizer (pre-signed URL redaction)."""

from client.exceptions import sanitize_exception_text


def test_sanitizes_query_string_from_exception_text():
    exc = RuntimeError(
        "Max retries exceeded with url: https://sn.example.com/up/abc?tempauth=SECRET&x=1 (oops)"
    )
    text = sanitize_exception_text(exc)
    assert "SECRET" not in text
    assert "tempauth" not in text
    assert "?<redacted>" in text


def test_leaves_plain_text_untouched():
    assert sanitize_exception_text(ValueError("connection reset")) == "connection reset"
