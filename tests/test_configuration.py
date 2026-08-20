import pytest
from pydantic import ValidationError

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
)


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
        assert config.append is False
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
        assert config.append is True
        assert config.batch_size == 2500

    def test_account_error_propagates_through_row_config(self):
        with pytest.raises(ValidationError, match="account.tenant_id is required"):
            RowConfig(mode="file", account={"account_type": "sharepoint"})


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
