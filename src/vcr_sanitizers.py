"""VCR cassette sanitizers for keboola.wr-onedrive-v2 (plan Task 11; extracted phase 8 audit).

``keboola.vcr``/``vcrpy`` are dev-only dependencies (see ``pyproject.toml``'s
``[dependency-groups] dev`` — the production Docker stage runs ``uv sync --no-dev``), so this
module must never be imported unconditionally: doing so would break every real job. It is only
ever imported from inside ``component.py``'s own ``try/except ImportError`` guard around
``VCR_SANITIZERS`` — if ``keboola.vcr`` isn't installed, the ``from keboola.vcr import ...`` line
below raises ``ImportError``, which propagates straight through that guarded import and is caught
there exactly as if the classes were still defined inline. The ``keboola.datadirtest`` VCR tester
(``tests/test_functional_vcr.py``) picks up ``VCR_SANITIZERS`` automatically via
``keboola.datadirtest.vcr.tester._load_vcr_sanitizers_from_script``, which itself tolerates a
missing/empty list.
"""

import json
import re

from keboola.vcr import BaseSanitizer, DefaultSanitizer, QueryParamSanitizer

# Graph's error taxonomy uses a top-level ``error.code`` field (e.g. "itemNotFound",
# "nameAlreadyExists", "notAllowed") that has nothing to do with an OAuth authorization
# *code* (this component only ever uses the refresh-token grant — no auth-code flow exists
# to protect here). ``DefaultSanitizer.DEFAULT_SENSITIVE_FIELDS`` redacting bare ``"code"``
# would replace that value in every recorded error body, desyncing ``logs.json`` between
# record and replay for every deliberate-failure cassette (the component logs the real code
# verbatim). Passing ``sensitive_fields=`` to *our own* ``DefaultSanitizer`` instance below
# isn't enough: ``VCRRecorder.__init__`` always additionally builds its *own* internal
# ``DefaultSanitizer`` from ``create_default_sanitizer(secrets)`` using the unmodified class
# defaults (see ``keboola/vcr/recorder.py``), so the class attribute itself has to be patched
# — this runs once per test (``VCR_SANITIZERS`` is re-probed fresh per scaffolded test
# directory), is idempotent, and only affects sanitizer construction, never live requests.
if "code" in DefaultSanitizer.DEFAULT_SENSITIVE_FIELDS:
    DefaultSanitizer.DEFAULT_SENSITIVE_FIELDS.remove("code")


class _StreamedBodySerializationFix(BaseSanitizer):
    """Work around vcrpy ``Request``/keboola.vcr not supporting streamed (file-object) bodies
    for cassette *serialization*, hit by every simple-PUT file/CSV/workbook-fixture upload in
    this component (``requests.put(..., data=open(path, "rb"))``).

    vcrpy's ``Request.__init__`` sets ``_was_file=True`` whenever the original body was a
    file-like object (``hasattr(body, "read")``) — reading it into ``_body`` as plain bytes —
    and independently also computes ``_was_iter`` (``_is_nonsequence_iterator(body)``); both
    flags can end up set on the same ``Request`` for a streamed-file body (a real file object
    is both readable *and* an iterator), even though only the ``_was_file`` branch actually
    ran in ``__init__``. Its ``.body`` *property getter* checks them in order and, whichever
    is set, always re-wraps ``_body`` on **every** access: ``BytesIO(self._body)`` if
    ``_was_file``, or ``iter(self._body)`` if ``_was_iter`` — including inside
    ``Request._to_dict()`` (``"body": self.body``), which is exactly what
    ``VCRRecorder._append_interaction`` serializes to the cassette JSON. keboola.vcr's custom
    ``_BytesEncoder`` only knows how to encode ``bytes`` — not a ``BytesIO`` (``TypeError:
    Object of type BytesIO is not JSON serializable``) and not a ``bytes_iterator`` (the same
    error, different type name) either. This is independent of any sanitizer's own body
    handling (reproduces with an empty ``VCR_SANITIZERS`` list too) — clearing *both* flags
    is required; clearing only ``_was_file`` still leaves the ``_was_iter`` branch active.

    Clearing both flags here (after the real, live upload has already happened — this only
    affects what gets *serialized*) makes every later ``.body`` access, including inside
    ``_to_dict()``, fall through to ``return self._body`` (plain bytes) instead. Placed
    **last** in ``VCR_SANITIZERS`` (after ``_GuidRedactor``/``_IdentityFieldRedactor``, which
    rely on seeing the original ``_was_file=True`` to correctly skip redacting uploaded file
    *content* — see ``_SafeBodyRedactor``) so nothing downstream needs to special-case it.
    """

    def before_record_request(self, request):
        if getattr(request, "_was_file", False) or getattr(request, "_was_iter", False):
            request._was_file = False
            request._was_iter = False
        return request


class _SafeBodyRedactor(BaseSanitizer):
    """Base class for sanitizers that rewrite request/response body *text*.

    Provides ``_body_text``/``_apply_text`` helpers that read/write a request's body via its
    **raw** ``_body`` attribute rather than the ``.body`` property, and skip entirely when
    it isn't plain ``str``/``bytes``.

    This matters for every simple-PUT file/CSV/workbook-fixture upload in this component
    (``requests.put(..., data=open(path, "rb"))``): vcrpy's ``Request`` marks
    ``_was_file=True`` for those and its ``.body`` *getter* unconditionally re-wraps whatever
    ``_body`` currently holds in a fresh ``BytesIO(...)`` on every access —
    ``BytesIO(an_already_BytesIO_instance)`` raises ``TypeError``. keboola.vcr's own
    ``BodyFieldSanitizer`` hits exactly this: its body-type check falls through unchanged for
    a non-str/bytes value and writes that (a ``BytesIO``) straight back into ``_body`` via
    the setter, permanently corrupting it for every sanitizer (and the cassette-serialization
    step) that touches ``.body`` afterward — which is why it's not used here at all. Reading
    ``_body`` directly sidesteps the rewrap entirely; skipping non-str/bytes bodies is exactly
    the desired behavior anyway (streamed file *content* needs no text redaction).

    Also skips entirely whenever ``_was_file`` is set (regardless of what ``_body`` currently
    looks like): those are our own uploaded file fixtures (CSV/text/xlsx), never containing
    secrets, and a lossy utf-8 decode/re-encode round-trip would otherwise silently mangle
    binary (e.g. ``.xlsx``) content in the cassette. Harmless either way for *matching* —
    ``VCRRecorder``'s default ``match_on`` never includes body content — but there is no
    reason to touch it.
    """

    def _body_text(self, request) -> str | None:
        if getattr(request, "_was_file", False):
            return None
        raw_body = getattr(request, "_body", None)
        if isinstance(raw_body, bytes):
            return raw_body.decode("utf-8", errors="ignore")
        if isinstance(raw_body, str):
            return raw_body
        return None

    def _redact_request_body(self, request, transform) -> None:
        text = self._body_text(request)
        if text is not None:
            request.body = transform(text).encode("utf-8")

    @staticmethod
    def _redact_response_body(response, transform) -> None:
        body = response.get("body")
        if not (isinstance(body, dict) and "string" in body):
            return
        value = body["string"]
        if isinstance(value, bytes):
            body["string"] = transform(value.decode("utf-8", errors="ignore")).encode("utf-8")
        elif isinstance(value, str):
            body["string"] = transform(value)


class _GuidRedactor(_SafeBodyRedactor):
    """Collapse every GUID-shaped substring (request URIs/bodies, response bodies) to one
    fixed placeholder — the OAuth ``tenant_id`` (embedded directly in every token-refresh
    URL, built from ``account.tenant_id`` in config) and the two GUIDs inside a SharePoint
    composite site id (``"{hostname},{guid},{guid}"``, a value the component reads from one
    response and reuses verbatim in later request URLs within the same run).

    Uses a single **fixed** placeholder (not a stable-but-distinct-per-value mapping)
    deliberately: ``before_record_request`` runs on *every* outgoing request, both when
    recording and — critically — when replaying (vcrpy applies it before matching, so the
    request the live component just built can line up against the sanitized cassette). A
    replayed request built from an *already-redacted* response (e.g. the site-id case above)
    would otherwise get **re-redacted** with a fresh replay-time sanitizer instance that has
    no memory of "I already assigned this value a placeholder", assigning it a *different*
    one and breaking the match. A single fixed placeholder sidesteps this: it's a fixed point
    of the substitution (redacting an already-redacted placeholder is a no-op), so the same
    text survives being sanitized twice without drifting.

    ``tests/setup/configs.json``'s dummy ``account.tenant_id`` is deliberately set to this
    exact placeholder value, so after the scaffolder restores ``config.json`` to its dummy
    values post-recording, replay's config-constructed token URL already matches what's
    stored in the cassette — no secrets-file-aware sanitizer needed for tenant_id at all.

    Deliberately **not** ``scrub_before_read``: that flavor scrubs the response the *live*
    component reads mid-recording, which would send the placeholder back to the real Graph
    API on the very next call (a 404), breaking the live recording run itself.
    """

    PLACEHOLDER = "00000000-0000-4000-8000-000000000000"
    _GUID_RE = re.compile(r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}")

    def _sanitize_text(self, text: str) -> str:
        return self._GUID_RE.sub(self.PLACEHOLDER, text)

    def before_record_request(self, request):
        if hasattr(request, "uri"):
            request.uri = self._sanitize_text(request.uri)
        self._redact_request_body(request, self._sanitize_text)
        return request

    def before_record_response(self, response):
        self._redact_response_body(response, self._sanitize_text)
        return response


class _IdentityFieldRedactor(_SafeBodyRedactor):
    """Redact user-identity JSON fields (UPN/mail/display name/...) in response bodies.

    A hand-rolled, minimal replacement for ``keboola.vcr``'s ``BodyFieldSanitizer`` — see
    ``_SafeBodyRedactor`` for why that class isn't safe to use here. Response-only (these
    identity fields appear in Graph *responses* — ``/me``, ``createdBy``/``lastModifiedBy``
    blocks — never in this component's own outgoing request bodies).
    """

    FIELDS = frozenset({"userPrincipalName", "mail", "email", "displayName", "givenName", "surname"})
    REPLACEMENT = "REDACTED"

    def _redact_value(self, value):
        if isinstance(value, dict):
            return {k: (self.REPLACEMENT if k in self.FIELDS else self._redact_value(v)) for k, v in value.items()}
        if isinstance(value, list):
            return [self._redact_value(item) for item in value]
        return value

    def _sanitize_text(self, text: str) -> str:
        try:
            data = json.loads(text)
        except (json.JSONDecodeError, TypeError, ValueError):
            return text
        return json.dumps(self._redact_value(data))

    def before_record_response(self, response):
        self._redact_response_body(response, self._sanitize_text)
        return response


VCR_SANITIZERS = [
    # DefaultSanitizer's own defaults (minus "code", patched above) already cover
    # access_token/refresh_token/client_secret/client_id/password/id_token/client_assertion
    # in the form-encoded token-refresh body, the JSON token response, and query strings; and
    # **all headers except content-type/content-length/accept** — which is what strips
    # Authorization/Set-Cookie/Cookie/WWW-Authenticate.
    DefaultSanitizer(),
    # Upload-session ``uploadUrl`` values are pre-signed with a ``tempauth`` query-string
    # token (SharePoint/OneDrive for Business) that grants unauthenticated write access to the
    # session — must never leak into a committed cassette, in either the createUploadSession
    # response body (where the URL first appears) or the subsequent chunk PUT/GET/DELETE
    # request URIs that use it (design spec §4 "Uploads"; ``client/uploader.py`` never logs
    # ``upload_url`` for the same reason).
    QueryParamSanitizer(parameters=["tempauth"], replacement="REDACTED"),
    # User principal name / mail / display name appear in `/me` (testConnection) and any
    # Graph response embedding `createdBy`/`lastModifiedBy` identity blocks.
    _IdentityFieldRedactor(),
    _GuidRedactor(),
    # Must run last — see its docstring for why.
    _StreamedBodySerializationFix(),
]

# Deliberately NOT sanitized: ``account.site_url`` (the SharePoint site host + path, e.g.
# "keboolaconnection.sharepoint.com"/"allcompany") and Graph's opaque (non-GUID) drive/item ids
# (e.g. "b!SYhH9ex...") are treated as non-secret resource identifiers recorded as-is — matching
# the same reasoning `vcr-sanitizers.md` gives for a non-secret ``base_url`` host, and
# `vcr-configs-format.md`'s coverage guidance ("Resource IDs ... are not secrets — use real
# ones"). This is Keboola's own internal M365 test tenant, not customer data.
