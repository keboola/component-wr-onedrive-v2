"""
Component main class for keboola.wr-onedrive-v2.

Design spec: ``docs/superpowers/specs/2026-08-17-wr-onedrive-v2-design.md`` §3 (auth), §5
(configuration & sync actions), and §6 (run orchestration, error mapping).
"""

import csv
import json
import logging
import os
import re
import sys
import tempfile
from datetime import UTC, datetime

from keboola.component.base import ComponentBase, sync_action
from keboola.component.dao import FileDefinition, TableDefinition
from keboola.component.exceptions import UserException
from keboola.component.sync_actions import SelectElement
from pydantic import ValidationError

from client.auth import AuthenticationError, RefreshTokenProvider, TokenProvider
from client.drives import get_site_id, list_drives, resolve_drive_id
from client.excel_writer import (
    list_worksheets_with_headers,
    resolve_workbook,
    resolve_worksheet,
    search_workbook,
    workbook_session,
    write_table,
)
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
    InvalidWorkbookFormatError,
    InvalidWorkbookPathError,
    MultipleSitesFoundError,
    UploadSessionError,
    WorkbookNotFoundError,
    WorksheetNotFoundError,
)
from client.graph_client import GraphClient
from client.uploader import ensure_folder, resolve_placeholders, upload_file, validate_path
from configuration import Account, AccountType, CsvOptions, Mode, RowConfig, Workbook, Worksheet

logger = logging.getLogger(__name__)

# --- VCR cassette sanitizers (plan Task 11) -------------------------------------------------
#
# ``keboola.vcr``/``vcrpy`` are dev-only dependencies (see ``pyproject.toml``'s
# ``[dependency-groups] dev`` — the production Docker stage runs ``uv sync --no-dev``), so this
# import must never be unconditional at module level: it would break every real job. The
# ``keboola.datadirtest`` VCR tester (``tests/test_functional_vcr.py``) picks up ``VCR_SANITIZERS``
# automatically via ``keboola.datadirtest.vcr.tester._load_vcr_sanitizers_from_script``, which
# itself tolerates a missing/empty list — so ``[]`` here is a safe, inert fallback outside tests.
try:
    from keboola.vcr import BaseSanitizer, DefaultSanitizer, QueryParamSanitizer
except ImportError:  # pragma: no cover - exercised only when keboola.vcr isn't installed (prod)
    VCR_SANITIZERS: list = []
else:
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
# --- end VCR cassette sanitizers -----------------------------------------------------------
#
# Deliberately NOT sanitized: ``account.site_url`` (the SharePoint site host + path, e.g.
# "keboolaconnection.sharepoint.com"/"allcompany") and Graph's opaque (non-GUID) drive/item ids
# (e.g. "b!SYhH9ex...") are treated as non-secret resource identifiers recorded as-is — matching
# the same reasoning `vcr-sanitizers.md` gives for a non-secret ``base_url`` host, and
# `vcr-configs-format.md`'s coverage guidance ("Resource IDs ... are not secrets — use real
# ones"). This is Keboola's own internal M365 test tenant, not customer data.

# v1-compatible state key (design spec §2/§3) — holds the JSON-encoded payload produced by
# `RefreshTokenProvider.rotated_refresh_token` (only `refresh_token` is actually read back).
STATE_KEY_REFRESHED_AUTH_DATA = "#refreshed_auth_data"

# Storage's own CSV defaults — an input table downloaded via the input mapping is always a
# comma-delimited, double-quote-enclosed CSV with a header row (design spec §6 "Streaming").
# When the row's `csv` options match all three, the input file is uploaded byte-for-byte; any
# difference triggers a streaming rewrite (see `_rewrite_csv`).
_STORAGE_DEFAULT_DELIMITER = ","
_STORAGE_DEFAULT_ENCLOSURE = '"'

# The run boundary (`Component.run`) maps exactly these exceptions to `UserException` (exit 1);
# everything else propagates to exit 2 (design spec §6 "Error mapping"). Note this deliberately
# excludes a bare `GraphClientError` (raised for unmapped/unexpected statuses after retries).
_USER_FACING_ERRORS = (
    AuthenticationError,
    GraphPermissionError,
    GraphNotFoundError,
    GraphQuotaExceededError,
    GraphBadRequestError,
    GraphRateLimitCapExceededError,
    GraphConnectionError,
    InvalidPathError,
    FileAlreadyExistsError,
    UploadSessionError,
    InvalidWorkbookPathError,
    InvalidWorkbookFormatError,
    MultipleSitesFoundError,
    WorkbookNotFoundError,
    WorksheetNotFoundError,
)


class Component(ComponentBase):
    """
    Extends base class for general Python components. Initializes the CommonInterface
    and performs configuration validation.

    For easier debugging the data folder is picked up by default from `../data` path,
    relative to working directory.
    """

    def __init__(self):
        super().__init__()

    def run(self) -> None:
        """Thin orchestrator (design spec §6): load config, dispatch on mode, persist the token.

        Token persistence happens in ``finally`` so a rotated refresh token is never lost even
        when the run itself fails (design spec §2/§3 — rotation must survive a failed job).
        """
        config = self._load_configuration()
        token_provider = self._build_token_provider(config.account)
        client = GraphClient(token_provider=token_provider)
        now = datetime.now(UTC)
        try:
            drive_id = resolve_drive_id(client, config.account, config.destination.drive_id)
            self._dispatch_mode(config, client, drive_id, now)
        except _USER_FACING_ERRORS as exc:
            raise UserException(str(exc)) from exc
        finally:
            self._persist_token_state(token_provider)

    @sync_action("testConnection")
    def test_connection(self) -> None:
        """Verify the OAuth credentials can obtain a Graph access token and call the API.

        Built from the root configuration only (no row parameters exist yet when this runs).
        """
        account = self._load_account()
        client = self._build_client(account)
        try:
            client.get("/me", params={"$select": "userPrincipalName"})
        except (AuthenticationError, GraphClientError) as exc:
            raise UserException(str(exc)) from exc

    @sync_action("listLibraries")
    def list_libraries(self) -> list[SelectElement]:
        """List the document libraries (drives) of the configured SharePoint site.

        Built from the **root** configuration only — reading row parameters here would tie the
        dropdown to whatever row happens to be open in the UI (the extractor's regression this
        component must not repeat, design spec §5).
        """
        account = self._load_account()
        self._require_sharepoint_site(account)
        client = self._build_client(account)
        try:
            site_id = get_site_id(client, account.site_url)
            drives = list_drives(client, site_id)
        except (AuthenticationError, GraphClientError) as exc:
            raise UserException(str(exc)) from exc
        return [SelectElement(label=drive["name"], value=drive["id"]) for drive in drives]

    @sync_action("search")
    def search(self) -> dict:
        """v1-parity ``search`` sync action (design spec §5): find a workbook by ``workbook.path``.

        Unlike the other three v1-parity actions, ``search`` only ever accepts a ``path`` (v1's
        ``WorkbooksFinder::search`` takes a single search string, never ``driveId``/``fileId``).
        A well-formed path that resolves to nothing returns ``{"file": null}``; an unrecognized
        path format or an inaccessible sharing link is always a ``UserException`` — v1 never
        treats those as "not found".
        """
        account = self._load_account()
        client = self._build_client(account)
        path = self._require_workbook_path('To search for a workbook please configure "parameters.workbook.path".')
        try:
            result = search_workbook(client, account, path)
        except (AuthenticationError, GraphClientError) as exc:
            raise UserException(str(exc)) from exc
        if result is None:
            return {"file": None}
        return {
            "file": {
                "driveId": result.drive_id,
                "fileId": result.file_id,
                "name": result.name,
                "path": result.path,
            }
        }

    @sync_action("getWorksheets")
    def get_worksheets(self) -> dict:
        """v1-parity ``getWorksheets`` sync action (design spec §5): list a workbook's sheets.

        ``workbook`` accepts either targeting form (ids or path, design spec §5's ``Workbook``
        model). A path-mode target that doesn't exist is a ``UserException`` here — unlike
        ``search``'s ``{"file": null}`` and unlike row-run Excel mode's create-on-missing,
        listing a workbook's worksheets should never have the side effect of creating it (a
        deliberate deviation from v1, which silently creates the workbook in this case; see the
        Task 8 report for detail).
        """
        account = self._load_account()
        client = self._build_client(account)
        workbook = self._load_workbook_param()
        try:
            drive_id, file_id, _created = resolve_workbook(client, account, workbook, create_if_missing=False)
            worksheets = list_worksheets_with_headers(client, drive_id, file_id)
        except (AuthenticationError, GraphClientError) as exc:
            raise UserException(str(exc)) from exc
        return {"worksheets": worksheets}

    @sync_action("createWorkbook")
    def create_workbook(self) -> dict:
        """v1-parity ``createWorkbook`` sync action (design spec §5): create an empty workbook.

        Only ``workbook.path`` is accepted (v1 parity — there is no "create by ids" concept).
        Reuses :func:`~client.excel_writer.resolve_workbook`'s create-on-missing path-mode
        resolution; its ``created`` flag distinguishes "just created" from "already existed",
        which is exactly the check v1's ``SheetProvider::createFile`` performs too.
        """
        account = self._load_account()
        client = self._build_client(account)
        path = self._require_workbook_path('To create workbook please configure "parameters.workbook.path".')
        try:
            drive_id, file_id, created = resolve_workbook(client, account, Workbook(path=path))
        except (AuthenticationError, GraphClientError) as exc:
            raise UserException(str(exc)) from exc
        if not created:
            raise UserException(f'Workbook "{path}" already exists.')
        return {"file": {"driveId": drive_id, "fileId": file_id}}

    @sync_action("createWorksheet")
    def create_worksheet(self) -> dict:
        """v1-parity ``createWorksheet`` sync action (design spec §5): add a named worksheet.

        ``workbook`` accepts either targeting form (ids or path); the workbook itself is never
        created here (``create_if_missing=False`` — same reasoning as ``getWorksheets``).
        ``worksheet.name`` is required; :func:`~client.excel_writer.resolve_worksheet` (the same
        name-mode targeting row-run Excel mode uses) creates the sheet when it's missing, and its
        ``is_new`` return value is the "already exists" check, mirroring v1's
        ``SheetProvider::createSheet``.
        """
        account = self._load_account()
        client = self._build_client(account)
        workbook = self._load_workbook_param()
        worksheet = self._require_worksheet_name('To create worksheet please configure "parameters.worksheet.name".')
        try:
            drive_id, file_id, _workbook_created = resolve_workbook(client, account, workbook, create_if_missing=False)
            worksheet_id, created, _actual_name = resolve_worksheet(client, drive_id, file_id, worksheet, session=None)
        except (AuthenticationError, GraphClientError) as exc:
            raise UserException(str(exc)) from exc
        if not created:
            raise UserException(f'Worksheet "{worksheet.name}" already exists.')
        return {"worksheet": {"driveId": drive_id, "fileId": file_id, "worksheetId": worksheet_id}}

    def _require_workbook_path(self, missing_message: str) -> str:
        """Validate ``parameters.workbook.path`` through the ``Workbook`` partial model.

        Used by ``search``/``createWorkbook``, which — unlike ``getWorksheets``/
        ``createWorksheet`` — only ever accept a ``path`` (no ids targeting, design spec §5). The
        raw-dict presence check happens *before* full-model validation so a ``parameters.workbook``
        section that's missing/empty entirely (the common "not configured yet" case) gets the
        exact v1-parity ``missing_message`` below, rather than the generic "Invalid workbook
        configuration" one ``Workbook``'s own "requires either path or drive_id+file_id" validator
        would otherwise raise for an empty dict.
        """
        workbook_params = self.configuration.parameters.get("workbook")
        if not isinstance(workbook_params, dict) or not workbook_params.get("path"):
            raise UserException(missing_message)
        try:
            workbook = Workbook.model_validate(workbook_params)
        except ValidationError as e:
            raise UserException(f"Invalid workbook configuration: {_format_validation_error(e)}") from e
        # The presence check above guarantees `workbook_params["path"]` was truthy, and `Workbook`
        # never clears a supplied `path` during validation.
        assert workbook.path is not None
        return workbook.path

    def _require_worksheet_name(self, missing_message: str) -> Worksheet:
        """Validate ``parameters.worksheet.name`` through the ``Worksheet`` partial model.

        Used by ``createWorksheet``, which — like v1's ``SheetProvider::createSheet`` — only ever
        targets a worksheet by name (never id/position); see :meth:`_require_workbook_path` for
        why the raw-dict presence check happens before full-model validation.
        """
        worksheet_params = self.configuration.parameters.get("worksheet")
        if not isinstance(worksheet_params, dict) or not worksheet_params.get("name"):
            raise UserException(missing_message)
        try:
            return Worksheet.model_validate({"name": worksheet_params["name"]})
        except ValidationError as e:
            raise UserException(f"Invalid worksheet configuration: {_format_validation_error(e)}") from e

    def _load_workbook_param(self) -> Workbook:
        """Validate just ``parameters.workbook`` (ids or path) for a sync action.

        Used by ``getWorksheets``/``createWorksheet``, which — unlike ``search``/
        ``createWorkbook`` — accept either targeting form (design spec §5).
        """
        workbook_params = self.configuration.parameters.get("workbook")
        try:
            return Workbook.model_validate(workbook_params or {})
        except ValidationError as e:
            raise UserException(f"Invalid workbook configuration: {_format_validation_error(e)}") from e

    def _load_configuration(self) -> RowConfig:
        try:
            return RowConfig.model_validate(self.configuration.parameters)
        except ValidationError as e:
            raise UserException(f"Invalid configuration: {_format_validation_error(e)}") from e

    def _dispatch_mode(self, config: RowConfig, client: GraphClient, drive_id: str, now: datetime) -> None:
        if config.mode == Mode.FILE:
            self._run_file_mode(config, client, drive_id, now)
        elif config.mode == Mode.TABLE_CSV:
            self._run_csv_mode(config, client, drive_id, now)
        else:
            self._run_excel_mode(config, client, drive_id)

    def _run_file_mode(self, config: RowConfig, client: GraphClient, drive_id: str, now: datetime) -> None:
        """Upload every file from the row's file input mapping (design spec §2/§6)."""
        files: list[FileDefinition] = self.get_input_files_definitions()
        if not files:
            raise UserException(
                "No files found in the input mapping. Add at least one file to this row's file "
                "input mapping before running mode 'file'."
            )
        parent_id, folder_path = self._resolve_destination_folder(config, client, drive_id, now)
        conflict_behavior = config.destination.conflict_behavior.value
        for file_def in files:
            upload_file(client, drive_id, parent_id, file_def.full_path, file_def.name, conflict_behavior)
            target_path = f"{folder_path}/{file_def.name}" if folder_path else file_def.name
            logger.info("Uploaded file '%s' to '%s'.", file_def.name, target_path)

    def _run_csv_mode(self, config: RowConfig, client: GraphClient, drive_id: str, now: datetime) -> None:
        """Upload the row's single input table as a CSV file (design spec §2/§6)."""
        table = self._require_single_input_table()
        parent_id, folder_path = self._resolve_destination_folder(config, client, drive_id, now)
        table_base_name = table.name.removesuffix(".csv")
        file_name = config.csv.file_name or f"{table_base_name}.csv"
        upload_path, is_temp_file = self._prepare_csv_upload_source(table, config.csv)
        try:
            conflict_behavior = config.destination.conflict_behavior.value
            upload_file(client, drive_id, parent_id, upload_path, file_name, conflict_behavior)
        finally:
            if is_temp_file:
                os.remove(upload_path)
        target_path = f"{folder_path}/{file_name}" if folder_path else file_name
        logger.info("Uploaded CSV file '%s' to '%s'.", file_name, target_path)

    def _run_excel_mode(self, config: RowConfig, client: GraphClient, drive_id: str) -> None:
        """Write the row's single input table into an Excel worksheet (design spec §5/§6).

        Unlike file/CSV mode, Excel mode never uses the ``drive_id`` the caller resolved from
        ``destination.drive_id`` (that field doesn't even apply here) — the target drive comes
        entirely from ``workbook.{path,drive_id,file_id}``, resolved below via
        :func:`~client.excel_writer.resolve_workbook`.
        """
        if config.account.account_type == AccountType.PRIVATE_ONEDRIVE:
            raise UserException(
                "Mode 'table_excel' is not supported for account_type 'private_onedrive': the "
                "Microsoft Graph Excel API is only available for OneDrive for Business and "
                "SharePoint accounts."
            )
        table = self._require_single_input_table()
        # `RowConfig._validate_mode_requirements` guarantees both are set for mode 'table_excel'.
        assert config.workbook is not None and config.worksheet is not None

        workbook_drive_id, workbook_file_id, workbook_created = resolve_workbook(
            client, config.account, config.workbook
        )
        with workbook_session(client, workbook_drive_id, workbook_file_id) as session:
            worksheet_id, worksheet_created, _actual_name = resolve_worksheet(
                client, workbook_drive_id, workbook_file_id, config.worksheet, session
            )
            wrote = write_table(
                client,
                workbook_drive_id,
                workbook_file_id,
                worksheet_id,
                table.full_path,
                append=config.append,
                batch_size=config.batch_size,
                is_new_sheet=workbook_created or worksheet_created,
                session=session,
            )
        if not wrote:
            logger.warning('Ignored empty CSV file "%s".', table.name)
            return
        logger.info("Wrote table '%s' to the Excel worksheet.", table.name)

    def _resolve_destination_folder(
        self, config: RowConfig, client: GraphClient, drive_id: str, now: datetime
    ) -> tuple[str, str]:
        """Resolve `destination.folder_path` (placeholders, validation) and ensure it exists.

        Shared by file and CSV mode (design spec §5: `destination.folder_path` applies to both).
        """
        business = config.account.account_type != AccountType.PRIVATE_ONEDRIVE
        folder_path = resolve_placeholders(config.destination.folder_path or "", now)
        validate_path(folder_path, business)
        parent_id = ensure_folder(client, drive_id, folder_path)
        return parent_id, folder_path

    def _require_single_input_table(self) -> TableDefinition:
        """Exactly one input table is required for CSV/Excel modes (v1-parity messages)."""
        tables: list[TableDefinition] = self.get_input_tables_definitions()
        if not tables:
            raise UserException('No CSV file found in "/data/in/tables".')
        if len(tables) > 1:
            names = ", ".join(f'"{name}"' for name in sorted(table.name for table in tables))
            raise UserException(f"Expected one CSV file, found multiple: {names}.")
        return tables[0]

    def _prepare_csv_upload_source(self, table: TableDefinition, csv_options: CsvOptions) -> tuple[str, bool]:
        """Return ``(path_to_upload, is_temp_file)``.

        When the row's CSV options match Storage's own defaults, the input file is uploaded
        as-is (streamed, no rewrite). Otherwise it's rewritten to a `/tmp` file — never under
        `data/out/` (design spec §2 "Scratch files") — and the caller is responsible for
        removing it once the upload finishes (see `_run_csv_mode`'s `finally`).
        """
        if (
            csv_options.delimiter == _STORAGE_DEFAULT_DELIMITER
            and csv_options.enclosure == _STORAGE_DEFAULT_ENCLOSURE
            and csv_options.include_header
        ):
            return table.full_path, False
        return _rewrite_csv(table.full_path, csv_options), True

    def _persist_token_state(self, token_provider: TokenProvider) -> None:
        """Persist a rotated refresh token to row state, if the provider rotated one.

        Called from `run()`'s `finally` block so rotation survives a failed run too (design
        spec §2/§3) — a token fetched before a mid-run failure would otherwise be lost, forcing
        the next run to fall back to a (possibly already-consumed) older refresh token.
        """
        rotated_token = token_provider.rotated_refresh_token
        if not rotated_token:
            return
        state = {STATE_KEY_REFRESHED_AUTH_DATA: json.dumps({"refresh_token": rotated_token})}
        self.write_state_file(state)

    def _load_account(self) -> Account:
        """Validate only the `account` section of the merged parameters.

        Sync actions run against the root configuration before any row exists, so validating the
        full `RowConfig` (which requires `mode`) would fail. Partial instantiation of just
        `Account` is the model the configuration checklist expects for sync actions that need
        fewer fields than a full row.
        """
        account_params = self.configuration.parameters.get("account", {})
        try:
            return Account.model_validate(account_params)
        except ValidationError as e:
            raise UserException(f"Invalid account configuration: {_format_validation_error(e)}") from e

    @staticmethod
    def _require_sharepoint_site(account: Account) -> None:
        if not account.tenant_id or not account.site_url:
            raise UserException(
                "listLibraries requires a SharePoint account with both account.tenant_id and "
                "account.site_url configured."
            )

    def _build_client(self, account: Account) -> GraphClient:
        return GraphClient(token_provider=self._build_token_provider(account))

    def _build_token_provider(self, account: Account) -> TokenProvider:
        """Build the `RefreshTokenProvider` for `account`, from OAuth credentials + state.

        Fallback order (design spec §3): the rotated refresh token from row state
        (`#refreshed_auth_data`) is tried first, the config's `#data` refresh token second.
        """
        oauth_credentials = self.configuration.oauth_credentials
        if oauth_credentials is None:
            raise UserException(
                "The component is not authorized. Please authorize the configuration in the "
                "Keboola UI and run the job again."
            )
        config_refresh_token = oauth_credentials.data.get("refresh_token")
        authority = "common" if account.account_type == AccountType.PRIVATE_ONEDRIVE else account.tenant_id
        try:
            return RefreshTokenProvider(
                client_id=oauth_credentials.appKey,
                client_secret=oauth_credentials.appSecret,
                authority=authority,
                refresh_token_candidates=[self._state_refresh_token(), config_refresh_token],
            )
        except AuthenticationError as exc:
            raise UserException(str(exc)) from exc

    def _state_refresh_token(self) -> str | None:
        raw = self.get_state_file().get(STATE_KEY_REFRESHED_AUTH_DATA)
        if not raw:
            return None
        try:
            return json.loads(raw).get("refresh_token")
        except (TypeError, ValueError, AttributeError):
            logger.warning("Ignoring malformed %s value found in state.", STATE_KEY_REFRESHED_AUTH_DATA)
            return None


def _format_validation_error(error: ValidationError) -> str:
    messages = [f"{'.'.join(str(loc) for loc in err['loc'])}: {err['msg']}" for err in error.errors()]
    return "; ".join(messages)


def _rewrite_csv(source_path: str, csv_options: CsvOptions) -> str:
    """Stream-rewrite a Storage-default CSV (`source_path`) into a new temp file per `csv_options`.

    Row-by-row via the stdlib `csv` module — the whole file is never buffered in memory (design
    spec §6 "Streaming"). `csv_options.include_header=False` drops the first row. The temp file
    lives under the OS temp directory (`tempfile` module — design spec §2 "Scratch files": `/tmp`
    only, never `data/out/`); the caller owns removing it once the upload finishes.
    """
    descriptor, tmp_path = tempfile.mkstemp(prefix="wr-onedrive-v2-csv-", suffix=".csv")
    os.close(descriptor)
    with (
        open(source_path, newline="", encoding="utf-8") as source_file,
        open(tmp_path, "w", newline="", encoding="utf-8") as dest_file,
    ):
        reader = csv.reader(source_file, delimiter=_STORAGE_DEFAULT_DELIMITER, quotechar=_STORAGE_DEFAULT_ENCLOSURE)
        writer = csv.writer(dest_file, delimiter=csv_options.delimiter, quotechar=csv_options.enclosure)
        rows = iter(reader)
        if not csv_options.include_header:
            next(rows, None)
        for row in rows:
            writer.writerow(row)
    return tmp_path


"""
        Main entrypoint
"""
if __name__ == "__main__":
    try:
        comp = Component()
        # this triggers the run method by default and is controlled by the configuration.action parameter
        comp.execute_action()
    except UserException:
        logger.exception("Component failed with a user error")
        sys.exit(1)
    except Exception:
        logger.exception("Component failed with an unexpected error")
        sys.exit(2)
