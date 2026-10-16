"""Unit tests for the VCR cassette sanitizers in ``src/vcr_sanitizers.py`` (no network).

Mirrors ``component-ex-medallia``'s ``tests/test_recording_support.py`` philosophy: prove each
sanitizer's guarantee that no real value can survive into a committed cassette, by planting
real-looking sentinels (UPN/email, display name, GUIDs, a bearer-token-shaped secret, a tempauth
token) and asserting NONE of them appear anywhere in the sanitized output (Medallia's
``REAL_TOKENS``/allowlist pattern) — plus a deterministic-placeholder guarantee for
``_GuidRedactor`` specifically (a replayed request built from an already-redacted response must
sanitize to the exact same fixed value, or cassette matching breaks).

Doubles here are minimal, faithful-enough stand-ins for vcrpy's real ``Request``/response shapes:
``_FakeRequest`` mirrors the ``.body`` property being backed by a ``_body`` attribute (exactly the
split ``_SafeBodyRedactor``'s own docstring describes), and ``_resp``/plain dicts mirror the
``{"body": {"string": ...}}`` response-dict shape ``VCRRecorder`` produces. Every sanitizer under
test here also runs, unmodified, against every recorded cassette replayed in
``tests/test_functional.py`` — this module is the fast, network-free, allowlist-style proof of the
underlying guarantee those cassettes rely on.

``keboola.vcr`` is a dev-only dependency (see ``pyproject.toml``'s ``[dependency-groups] dev`` —
the production Docker stage runs ``uv sync --no-dev``); this module is import-guarded with
``pytest.importorskip`` so a hypothetical `--no-dev` test invocation still collects cleanly, even
though the dependency is in fact installed in every environment these tests actually run in today.
``vcr_sanitizers`` itself is imported directly — never through ``component.py``, which would pull
in the component's entire import graph as a side effect just to reach this one module.
"""

import json

import pytest

pytest.importorskip("keboola.vcr", reason="keboola.vcr is a dev-only dependency (uv sync --no-dev omits it)")

from keboola.vcr import BaseSanitizer, DefaultSanitizer, QueryParamSanitizer

import vcr_sanitizers

# Real-looking sentinels planted in the fake request/response bodies below. After sanitization
# NONE of these may survive anywhere in the cassette-bound text — the allowlist guarantee.
REAL_TOKENS = [
    "john.doe@contoso.com",  # a real UPN/email
    "Jane Q. Doe",  # a real display name
    "7f8e9d6c-5b4a-4321-9fed-0123456789ab",  # a real-looking GUID (e.g. tenant id)
    "AbCdEf01-2345-6789-AbCd-Ef0123456789",  # a real-looking GUID, mixed case (composite site id)
    "eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9.super-secret-bearer-token",  # a real-looking secret/token
    "eyJ0ZW1wYXV0aC10b2tlbi12YWx1ZS1kby1ub3QtbGVhayI",  # a real-looking tempauth token
]


def _assert_no_real_values(text) -> None:
    if isinstance(text, bytes):
        text = text.decode("utf-8", errors="ignore")
    for token in REAL_TOKENS:
        assert token not in text, f"real value leaked into cassette: {token!r}"


class _FakeRequest:
    """Faithful-enough stand-in for vcrpy's ``Request``: ``.body`` is a property backed by
    ``_body`` (matching the split ``_SafeBodyRedactor``'s own docstring describes for the real
    class), plus the ``uri``/``_was_file``/``_was_iter`` attributes the sanitizers read/write.
    """

    def __init__(self, uri: str = "https://graph.microsoft.com/v1.0/me", body=None, was_file=False, was_iter=False):
        self.uri = uri
        self._body = body
        self._was_file = was_file
        self._was_iter = was_iter

    @property
    def body(self):
        return self._body

    @body.setter
    def body(self, value):
        self._body = value


def _resp(text) -> dict:
    """Wrap response text in the vcrpy response-dict shape the sanitizers expect."""
    return {"status": {"code": 200}, "body": {"string": text}}


# -- _StreamedBodySerializationFix -----------------------------------------------------------


class TestStreamedBodySerializationFix:
    def test_was_file_flag_cleared(self):
        request = _FakeRequest(was_file=True, was_iter=False)

        result = vcr_sanitizers._StreamedBodySerializationFix().before_record_request(request)

        assert result is request
        assert request._was_file is False
        assert request._was_iter is False

    def test_was_iter_flag_cleared(self):
        request = _FakeRequest(was_file=False, was_iter=True)

        vcr_sanitizers._StreamedBodySerializationFix().before_record_request(request)

        assert request._was_file is False
        assert request._was_iter is False

    def test_both_flags_cleared_when_both_set(self):
        # A real file object is both readable *and* an iterator — both flags can end up set.
        request = _FakeRequest(was_file=True, was_iter=True)

        vcr_sanitizers._StreamedBodySerializationFix().before_record_request(request)

        assert request._was_file is False
        assert request._was_iter is False

    def test_neither_flag_set_is_a_noop(self):
        request = _FakeRequest(was_file=False, was_iter=False)

        vcr_sanitizers._StreamedBodySerializationFix().before_record_request(request)

        assert request._was_file is False
        assert request._was_iter is False

    def test_missing_flags_entirely_does_not_raise(self):
        class _Bare:
            pass

        request = _Bare()

        result = vcr_sanitizers._StreamedBodySerializationFix().before_record_request(request)

        assert result is request
        assert not hasattr(request, "_was_file")


# -- _SafeBodyRedactor (exercised via _GuidRedactor, since the base class has no _sanitize_text) --


class TestSafeBodyRedactorGuards:
    def test_was_file_request_body_is_never_touched(self):
        """Uploaded file *content* (CSV/text/xlsx fixtures) must never be text-redacted — a
        lossy utf-8 round-trip would silently mangle binary content."""
        binary_ish = b"upload-content-with-a-real-guid-7f8e9d6c-5b4a-4321-9fed-0123456789ab"
        request = _FakeRequest(body=binary_ish, was_file=True)

        vcr_sanitizers._GuidRedactor().before_record_request(request)

        assert request.body == binary_ish  # untouched, even though it "looks" redactable

    def test_bytes_body_is_decoded_redacted_and_reencoded(self):
        request = _FakeRequest(body=b'{"tenantId": "7f8e9d6c-5b4a-4321-9fed-0123456789ab"}')

        vcr_sanitizers._GuidRedactor().before_record_request(request)

        assert isinstance(request.body, bytes)
        assert b"7f8e9d6c-5b4a-4321-9fed-0123456789ab" not in request.body
        assert vcr_sanitizers._GuidRedactor.PLACEHOLDER.encode() in request.body

    def test_string_body_is_redacted_and_encoded_to_bytes(self):
        # `_redact_request_body` always `.encode()`s its result, even for a `str` input body —
        # matches vcrpy's own `Request.body` setter, which expects bytes.
        request = _FakeRequest(body='{"tenantId": "7f8e9d6c-5b4a-4321-9fed-0123456789ab"}')

        vcr_sanitizers._GuidRedactor().before_record_request(request)

        assert isinstance(request.body, bytes)
        assert b"7f8e9d6c-5b4a-4321-9fed-0123456789ab" not in request.body

    def test_non_str_bytes_body_is_left_untouched(self):
        sentinel = object()
        request = _FakeRequest(body=sentinel)

        vcr_sanitizers._GuidRedactor().before_record_request(request)

        assert request.body is sentinel  # not a redactable text type -> skipped entirely

    def test_response_body_missing_string_key_is_left_untouched(self):
        response = {"body": {}}

        vcr_sanitizers._GuidRedactor().before_record_response(response)

        assert response == {"body": {}}

    def test_response_with_no_body_key_is_left_untouched(self):
        response = {"status": {"code": 200}}

        vcr_sanitizers._GuidRedactor().before_record_response(response)

        assert response == {"status": {"code": 200}}


# -- _GuidRedactor ----------------------------------------------------------------------------


class TestGuidRedactor:
    def test_guid_in_uri_collapsed_to_placeholder(self):
        request = _FakeRequest(
            uri="https://login.microsoftonline.com/7f8e9d6c-5b4a-4321-9fed-0123456789ab/oauth2/v2.0/token"
        )

        vcr_sanitizers._GuidRedactor().before_record_request(request)

        assert "7f8e9d6c-5b4a-4321-9fed-0123456789ab" not in request.uri
        assert vcr_sanitizers._GuidRedactor.PLACEHOLDER in request.uri

    def test_uppercase_and_mixed_case_guid_also_matches(self):
        request = _FakeRequest(uri="https://x/AbCdEf01-2345-6789-AbCd-Ef0123456789/y")

        vcr_sanitizers._GuidRedactor().before_record_request(request)

        assert "AbCdEf01-2345-6789-AbCd-Ef0123456789" not in request.uri
        assert vcr_sanitizers._GuidRedactor.PLACEHOLDER in request.uri

    def test_composite_site_id_both_guids_collapse_to_the_same_fixed_placeholder(self):
        # SharePoint composite site id shape: "{hostname},{guid},{guid}".
        text = "contoso.sharepoint.com,7f8e9d6c-5b4a-4321-9fed-0123456789ab,AbCdEf01-2345-6789-AbCd-Ef0123456789"

        sanitized = vcr_sanitizers._GuidRedactor()._sanitize_text(text)

        placeholder = vcr_sanitizers._GuidRedactor.PLACEHOLDER
        assert sanitized == f"contoso.sharepoint.com,{placeholder},{placeholder}"

    def test_deterministic_across_instances_and_calls(self):
        text = "7f8e9d6c-5b4a-4321-9fed-0123456789ab"

        first = vcr_sanitizers._GuidRedactor()._sanitize_text(text)
        second = vcr_sanitizers._GuidRedactor()._sanitize_text(text)

        assert first == second == vcr_sanitizers._GuidRedactor.PLACEHOLDER

    def test_already_redacted_placeholder_is_a_fixed_point(self):
        """Critical for replay: a request built from an already-redacted response (e.g. reusing
        a site id read from a prior response) must sanitize to the identical value, not drift."""
        placeholder = vcr_sanitizers._GuidRedactor.PLACEHOLDER

        assert vcr_sanitizers._GuidRedactor()._sanitize_text(placeholder) == placeholder

    def test_request_without_uri_attribute_is_safe(self):
        class _NoUri:
            _body = None

        request = _NoUri()

        result = vcr_sanitizers._GuidRedactor().before_record_request(request)

        assert result is request  # no AttributeError

    def test_non_guid_resource_ids_are_left_untouched(self):
        # Opaque (non-GUID) Graph drive/item ids are deliberately NOT redacted (module docstring:
        # treated as non-secret resource identifiers, same as a non-secret base_url host).
        text = "b!SYhH9exampleOpaqueDriveOrItemId_not-a-guid"

        assert vcr_sanitizers._GuidRedactor()._sanitize_text(text) == text

    def test_response_body_guid_collapsed(self):
        response = _resp('{"id": "7f8e9d6c-5b4a-4321-9fed-0123456789ab"}')

        vcr_sanitizers._GuidRedactor().before_record_response(response)

        assert "7f8e9d6c-5b4a-4321-9fed-0123456789ab" not in response["body"]["string"]
        assert vcr_sanitizers._GuidRedactor.PLACEHOLDER in response["body"]["string"]

    def test_bytes_response_body_roundtrips(self):
        payload = json.dumps({"id": "7f8e9d6c-5b4a-4321-9fed-0123456789ab"}).encode("utf-8")
        response = {"body": {"string": payload}}

        vcr_sanitizers._GuidRedactor().before_record_response(response)

        text = response["body"]["string"]
        assert isinstance(text, bytes)
        assert b"7f8e9d6c-5b4a-4321-9fed-0123456789ab" not in text


# -- _IdentityFieldRedactor --------------------------------------------------------------------


class TestIdentityFieldRedactor:
    def test_user_principal_name_redacted_in_response(self):
        response = _resp(json.dumps({"userPrincipalName": "john.doe@contoso.com"}))

        vcr_sanitizers._IdentityFieldRedactor().before_record_response(response)

        body = json.loads(response["body"]["string"])
        assert body["userPrincipalName"] == "REDACTED"
        assert "john.doe@contoso.com" not in response["body"]["string"]

    def test_mail_field_redacted(self):
        response = _resp(json.dumps({"mail": "john.doe@contoso.com"}))

        vcr_sanitizers._IdentityFieldRedactor().before_record_response(response)

        assert json.loads(response["body"]["string"])["mail"] == "REDACTED"

    def test_display_name_given_name_surname_redacted(self):
        payload = {"displayName": "Jane Q. Doe", "givenName": "Jane", "surname": "Doe"}
        response = _resp(json.dumps(payload))

        vcr_sanitizers._IdentityFieldRedactor().before_record_response(response)

        body = json.loads(response["body"]["string"])
        assert body == {"displayName": "REDACTED", "givenName": "REDACTED", "surname": "REDACTED"}

    def test_nested_created_by_last_modified_by_blocks_redacted(self):
        payload = {
            "id": "file-1",
            "createdBy": {"user": {"email": "john.doe@contoso.com", "displayName": "Jane Q. Doe"}},
            "lastModifiedBy": {"user": {"email": "john.doe@contoso.com"}},
        }
        response = _resp(json.dumps(payload))

        vcr_sanitizers._IdentityFieldRedactor().before_record_response(response)

        body = json.loads(response["body"]["string"])
        assert body["createdBy"]["user"]["email"] == "REDACTED"
        assert body["createdBy"]["user"]["displayName"] == "REDACTED"
        assert body["lastModifiedBy"]["user"]["email"] == "REDACTED"
        assert body["id"] == "file-1"  # non-identity field untouched

    def test_identity_field_inside_a_list_is_redacted(self):
        payload = {"members": [{"email": "john.doe@contoso.com"}, {"email": "second@contoso.com"}]}
        response = _resp(json.dumps(payload))

        vcr_sanitizers._IdentityFieldRedactor().before_record_response(response)

        body = json.loads(response["body"]["string"])
        assert [m["email"] for m in body["members"]] == ["REDACTED", "REDACTED"]

    def test_non_identity_fields_untouched(self):
        payload = {"id": "file-1", "name": "report.xlsx"}
        response = _resp(json.dumps(payload))

        vcr_sanitizers._IdentityFieldRedactor().before_record_response(response)

        assert json.loads(response["body"]["string"]) == payload

    def test_non_json_body_passed_through_unchanged(self):
        response = _resp("not-json-at-all")

        vcr_sanitizers._IdentityFieldRedactor().before_record_response(response)

        assert response["body"]["string"] == "not-json-at-all"

    def test_bytes_response_body_roundtrips(self):
        payload = {"mail": "john.doe@contoso.com"}
        response = {"body": {"string": json.dumps(payload).encode("utf-8")}}

        vcr_sanitizers._IdentityFieldRedactor().before_record_response(response)

        text = response["body"]["string"]
        assert isinstance(text, bytes)
        assert b"john.doe@contoso.com" not in text

    def test_request_bodies_are_never_touched_identity_fields_are_response_only(self):
        """``_IdentityFieldRedactor`` deliberately has no ``before_record_request`` override —
        identity fields appear in Graph *responses* only (module docstring); calling it on a
        request falls through to ``BaseSanitizer``'s pass-through default."""
        original_body = json.dumps({"mail": "john.doe@contoso.com"})
        request = _FakeRequest(body=original_body)

        result = vcr_sanitizers._IdentityFieldRedactor().before_record_request(request)

        assert result is request
        assert request.body == original_body  # unchanged


# -- tempauth QueryParamSanitizer wiring (VCR_SANITIZERS[1]) -----------------------------------


class TestTempauthQueryParamSanitizerWiring:
    def _tempauth_sanitizer(self) -> QueryParamSanitizer:
        return next(s for s in vcr_sanitizers.VCR_SANITIZERS if isinstance(s, QueryParamSanitizer))

    def test_configured_for_tempauth_with_redacted_replacement(self):
        sanitizer = self._tempauth_sanitizer()

        assert sanitizer.parameters == ["tempauth"]
        assert sanitizer.replacement == "REDACTED"

    def test_tempauth_redacted_in_upload_session_response_body(self):
        sanitizer = self._tempauth_sanitizer()
        payload = {
            "uploadUrl": (
                "https://contoso-my.sharepoint.com/_api/v2.0/UploadSession"
                "?tempauth=eyJ0ZW1wYXV0aC10b2tlbi12YWx1ZS1kby1ub3QtbGVhayI&foo=bar"
            )
        }
        response = _resp(json.dumps(payload))

        sanitizer.before_record_response(response)

        assert "eyJ0ZW1wYXV0aC10b2tlbi12YWx1ZS1kby1ub3QtbGVhayI" not in response["body"]["string"]
        assert "tempauth=REDACTED" in response["body"]["string"]
        assert "foo=bar" in response["body"]["string"]  # other params untouched

    def test_tempauth_redacted_in_chunk_put_request_uri(self):
        sanitizer = self._tempauth_sanitizer()
        request = _FakeRequest(
            uri=(
                "https://contoso-my.sharepoint.com/_api/v2.0/UploadSession"
                "?tempauth=eyJ0ZW1wYXV0aC10b2tlbi12YWx1ZS1kby1ub3QtbGVhayI"
            )
        )

        sanitizer.before_record_request(request)

        assert "eyJ0ZW1wYXV0aC10b2tlbi12YWx1ZS1kby1ub3QtbGVhayI" not in request.uri
        assert "tempauth=REDACTED" in request.uri


# -- DefaultSanitizer's DEFAULT_SENSITIVE_FIELDS patch (module import side effect) --------------


class TestDefaultSensitiveFieldsPatch:
    def test_code_removed_from_default_sensitive_fields(self):
        # Graph's error taxonomy uses a top-level error.code (itemNotFound, nameAlreadyExists,
        # ...) unrelated to an OAuth authorization *code* — redacting it would desync logs.json
        # between record and replay for every deliberate-failure cassette.
        assert "code" not in DefaultSanitizer.DEFAULT_SENSITIVE_FIELDS

    def test_other_default_sensitive_fields_are_untouched(self):
        for field in ("access_token", "refresh_token", "client_id", "client_secret", "password"):
            assert field in DefaultSanitizer.DEFAULT_SENSITIVE_FIELDS


# -- VCR_SANITIZERS assembly --------------------------------------------------------------------


class TestVcrSanitizersAssembly:
    def test_is_a_list_of_five_sanitizers_in_documented_order(self):
        types = [type(s) for s in vcr_sanitizers.VCR_SANITIZERS]

        assert types == [
            DefaultSanitizer,
            QueryParamSanitizer,
            vcr_sanitizers._IdentityFieldRedactor,
            vcr_sanitizers._GuidRedactor,
            vcr_sanitizers._StreamedBodySerializationFix,
        ]

    def test_streamed_body_fix_runs_last(self):
        # Its own docstring: must run last so upstream sanitizers still see the original
        # _was_file=True flag and correctly skip redacting uploaded file *content*.
        assert isinstance(vcr_sanitizers.VCR_SANITIZERS[-1], vcr_sanitizers._StreamedBodySerializationFix)

    def test_every_entry_is_a_base_sanitizer(self):
        assert all(isinstance(s, BaseSanitizer) for s in vcr_sanitizers.VCR_SANITIZERS)

    def test_importable_without_keboola_vcr_would_yield_an_empty_list(self):
        """``component.py`` guards the ``from vcr_sanitizers import VCR_SANITIZERS`` import in a
        ``try/except ImportError`` and falls back to ``[]`` (see ``vcr_sanitizers.py``'s own
        module docstring) — this only ever happens when ``keboola.vcr`` itself is missing, which
        makes the *module-level* ``from keboola.vcr import ...`` at the top of ``vcr_sanitizers.py``
        raise before ``VCR_SANITIZERS`` is even defined. Simulate that by re-executing the
        module's source with the ``keboola.vcr`` import blocked, mirroring exactly what
        ``component.py``'s guard observes."""
        import builtins
        import importlib
        import sys

        real_import = builtins.__import__

        def _blocking_import(name, *args, **kwargs):
            if name == "keboola.vcr" or name.startswith("keboola.vcr."):
                raise ImportError("simulated: keboola.vcr not installed")
            return real_import(name, *args, **kwargs)

        module_name = vcr_sanitizers.__name__
        spec = importlib.util.find_spec(module_name)
        assert spec is not None and spec.loader is not None

        builtins.__import__ = _blocking_import
        try:
            with pytest.raises(ImportError):
                fresh_module = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(fresh_module)
        finally:
            builtins.__import__ = real_import
            # Re-importing here is unnecessary (the real module in sys.modules was never
            # touched by the blocked exec above), but guards against any accidental pollution.
            assert sys.modules[module_name] is vcr_sanitizers


# -- Full-pipeline leak-proof (every sanitizer, in VCR_SANITIZERS order) ------------------------


class TestFullPipelineNoRealValueSurvives:
    """Runs every sanitizer in ``VCR_SANITIZERS``, in order, over realistic-looking request and
    response bodies planted with ``REAL_TOKENS`` — proves the end-to-end guarantee the individual
    class-level tests above only prove in isolation."""

    @staticmethod
    def _run_request(request):
        for sanitizer in vcr_sanitizers.VCR_SANITIZERS:
            request = sanitizer.before_record_request(request)
        return request

    @staticmethod
    def _run_response(response):
        for sanitizer in vcr_sanitizers.VCR_SANITIZERS:
            response = sanitizer.before_record_response(response)
        return response

    def test_token_refresh_request_is_fully_sanitized(self):
        request = _FakeRequest(
            uri="https://login.microsoftonline.com/7f8e9d6c-5b4a-4321-9fed-0123456789ab/oauth2/v2.0/token",
            body=(
                "client_id=abc"
                "&client_secret=eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9.super-secret-bearer-token"
                "&refresh_token=some-refresh-token"
                "&grant_type=refresh_token"
            ),
        )

        result = self._run_request(request)

        _assert_no_real_values(result.uri)
        _assert_no_real_values(result.body)

    def test_upload_session_response_is_fully_sanitized(self):
        payload = {
            "uploadUrl": (
                "https://contoso-my.sharepoint.com/_api/v2.0/UploadSession"
                "?tempauth=eyJ0ZW1wYXV0aC10b2tlbi12YWx1ZS1kby1ub3QtbGVhayI"
            ),
            "createdBy": {"user": {"email": "john.doe@contoso.com", "displayName": "Jane Q. Doe"}},
            "parentReference": {
                "siteId": (
                    "contoso.sharepoint.com,7f8e9d6c-5b4a-4321-9fed-0123456789ab,"
                    "AbCdEf01-2345-6789-AbCd-Ef0123456789"
                )
            },
        }
        response = _resp(json.dumps(payload))

        result = self._run_response(response)

        _assert_no_real_values(result["body"]["string"])

    def test_streamed_file_upload_content_survives_untouched_but_flags_cleared(self):
        """The one deliberate non-redaction: an uploaded file's own bytes (the writer
        component's own content, never a secret) must reach the cassette unmodified — but the
        was_file/was_iter flags must still end up cleared for JSON serialization to succeed."""
        file_bytes = b"col1,col2\n1,2\n"
        request = _FakeRequest(
            uri="https://graph.microsoft.com/v1.0/drives/d/items/i:/f.csv:/content",
            body=file_bytes,
            was_file=True,
        )

        result = self._run_request(request)

        assert result.body == file_bytes
        assert result._was_file is False
        assert result._was_iter is False
