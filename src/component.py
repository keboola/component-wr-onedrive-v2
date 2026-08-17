"""
Component main class for keboola.wr-onedrive-v2.

Design spec: ``docs/superpowers/specs/2026-08-17-wr-onedrive-v2-design.md`` §3 (auth), §5
(configuration & sync actions), and §6 (run orchestration, error mapping).
"""

import csv
import json
import logging
import os
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
        worksheet_params = self.configuration.parameters.get("worksheet") or {}
        name = worksheet_params.get("name")
        if not name:
            raise UserException('To create worksheet please configure "parameters.worksheet.name".')
        try:
            drive_id, file_id, _workbook_created = resolve_workbook(client, account, workbook, create_if_missing=False)
            worksheet_id, created, _actual_name = resolve_worksheet(
                client, drive_id, file_id, Worksheet(name=name), session=None
            )
        except (AuthenticationError, GraphClientError) as exc:
            raise UserException(str(exc)) from exc
        if not created:
            raise UserException(f'Worksheet "{name}" already exists.')
        return {"worksheet": {"driveId": drive_id, "fileId": file_id, "worksheetId": worksheet_id}}

    def _require_workbook_path(self, missing_message: str) -> str:
        path = (self.configuration.parameters.get("workbook") or {}).get("path")
        if not path:
            raise UserException(missing_message)
        return path

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
