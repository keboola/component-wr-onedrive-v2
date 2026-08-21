"""Cross-check: every `options.async.action` declared in the row/root schemas has a matching
`@sync_action(...)`-decorated method on `Component` (dropdown UX addition — the fastest way to
notice a typo'd action name in the schema, which would otherwise only surface as a runtime "action
not found" error inside the Keboola UI itself).
"""

import json
from pathlib import Path
from typing import Any

import pytest
from keboola.component.base import _SYNC_ACTION_MAPPING

import component  # noqa: F401 - importing populates `_SYNC_ACTION_MAPPING` (decorators run on import)

REPO_ROOT = Path(__file__).resolve().parents[1]
COMPONENT_CONFIG_DIR = REPO_ROOT / "component_config"

# The Docker test image copies only src/ and tests/ (canonical template), so the schema files are
# absent there — this cross-check then runs only in local/CI-repo contexts where they exist.
pytestmark = pytest.mark.skipif(
    not COMPONENT_CONFIG_DIR.exists(),
    reason="component_config/ not present in this environment (Docker test image)",
)


def _iter_async_actions(node: Any) -> list[str]:
    """Recursively collect every `options.async.action` value found anywhere in a schema dict."""
    actions: list[str] = []
    if isinstance(node, dict):
        async_options = node.get("options", {}).get("async") if isinstance(node.get("options"), dict) else None
        if isinstance(async_options, dict) and "action" in async_options:
            actions.append(async_options["action"])
        for value in node.values():
            actions.extend(_iter_async_actions(value))
    elif isinstance(node, list):
        for item in node:
            actions.extend(_iter_async_actions(item))
    return actions


def _schema_actions(filename: str) -> list[str]:
    schema = json.loads((COMPONENT_CONFIG_DIR / filename).read_text())
    return _iter_async_actions(schema)


class TestSchemaSyncActionsExist:
    def test_row_schema_async_actions_are_all_registered_sync_actions(self):
        actions = _schema_actions("configRowSchema.json")
        assert actions, "expected at least one options.async.action in configRowSchema.json"
        missing = [action for action in actions if action not in _SYNC_ACTION_MAPPING]
        assert not missing, f"configRowSchema.json references undefined sync action(s): {missing}"

    def test_root_schema_async_actions_are_all_registered_sync_actions(self):
        actions = _schema_actions("configSchema.json")
        assert actions, "expected at least one options.async.action in configSchema.json"
        missing = [action for action in actions if action not in _SYNC_ACTION_MAPPING]
        assert not missing, f"configSchema.json references undefined sync action(s): {missing}"

    def test_new_dropdown_actions_are_present_in_the_row_schema(self):
        # The three UX-addition actions this test module was written to guard against regressing.
        actions = set(_schema_actions("configRowSchema.json"))
        assert {"listLibraries", "listWorkbooks", "listWorksheets"} <= actions


class TestRowSchemaJsonValidity:
    """Change 2/3 UX additions: `configRowSchema.json` stays syntactically valid JSON and the new
    `destination.date` field / reworded `mode` description are present with the expected shape."""

    def _row_schema(self) -> dict:
        return json.loads((COMPONENT_CONFIG_DIR / "configRowSchema.json").read_text())

    def test_row_schema_is_valid_json(self):
        # `.read_text()` + `json.loads` above already raises on malformed JSON; this test exists
        # so a parse failure reports here with a clear name instead of only inside some other
        # test's setup.
        assert self._row_schema()["type"] == "object"

    def test_mode_description_reads_as_a_how_not_a_what(self):
        mode = self._row_schema()["properties"]["mode"]
        assert mode["description"] == (
            "How this row writes data — File uploads mapped files as-is and writes mapped tables "
            "as CSV files; Worksheet writes one mapped table into an Excel worksheet."
        )

    def test_destination_date_field_is_present_between_folder_path_and_conflict_behavior(self):
        destination_properties = self._row_schema()["properties"]["destination"]["properties"]
        assert "date" in destination_properties
        date_field = destination_properties["date"]
        assert date_field["title"] == "Date"
        assert date_field["propertyOrder"] > destination_properties["folder_path"]["propertyOrder"]
        assert date_field["propertyOrder"] < destination_properties["conflict_behavior"]["propertyOrder"]


class TestModeSchemaTwoModes:
    """Change A: only two output modes remain — `mode`'s enum shrinks to `file`/`worksheet`, and
    every section gated on the old three-way split now depends on the correct one of the two."""

    def _row_schema(self) -> dict:
        return json.loads((COMPONENT_CONFIG_DIR / "configRowSchema.json").read_text())

    def test_mode_enum_has_exactly_file_and_worksheet(self):
        mode = self._row_schema()["properties"]["mode"]
        assert mode["enum"] == ["file", "worksheet"]
        assert mode["options"]["enum_titles"] == ["File", "Worksheet"]
        assert mode["default"] == "file"

    def test_destination_depends_on_mode_file_only(self):
        destination = self._row_schema()["properties"]["destination"]
        assert destination["options"]["dependencies"] == {"mode": "file"}

    def test_csv_depends_on_mode_file_only_and_is_retitled(self):
        csv_section = self._row_schema()["properties"]["csv"]
        assert csv_section["options"]["dependencies"] == {"mode": "file"}
        assert csv_section["title"] == "CSV Options (for mapped tables)"

    def test_workbook_and_worksheet_depend_on_mode_worksheet(self):
        schema = self._row_schema()
        assert schema["properties"]["workbook"]["options"]["dependencies"] == {"mode": "worksheet"}
        assert schema["properties"]["worksheet"]["options"]["dependencies"] == {"mode": "worksheet"}

    def test_write_mode_and_batch_size_depend_on_mode_worksheet(self):
        schema = self._row_schema()
        assert schema["properties"]["write_mode"]["options"]["dependencies"] == {"mode": "worksheet"}
        assert schema["properties"]["batch_size"]["options"]["dependencies"] == {"mode": "worksheet"}


class TestWorkbookTargetingSchema:
    """Change B: `workbook.targeting` (pick vs. path) replaces the old "never combine with Path"
    tooltip with real conditional visibility — the ids and the path field are now shown/hidden by
    `options.dependencies` instead of relying on a warning sentence."""

    def _workbook_properties(self) -> dict:
        schema = json.loads((COMPONENT_CONFIG_DIR / "configRowSchema.json").read_text())
        return schema["properties"]["workbook"]["properties"]

    def test_targeting_is_the_first_field_with_the_expected_shape(self):
        properties = self._workbook_properties()
        targeting = properties["targeting"]
        assert targeting["enum"] == ["pick", "path"]
        assert targeting["options"]["enum_titles"] == ["Pick via dropdowns", "By path"]
        assert targeting["default"] == "pick"
        assert targeting["propertyOrder"] == 1
        assert all(targeting["propertyOrder"] < other["propertyOrder"] for name, other in properties.items() if name != "targeting")

    def test_ids_depend_on_targeting_pick(self):
        properties = self._workbook_properties()
        assert properties["drive_id"]["options"]["dependencies"] == {"targeting": "pick"}
        assert properties["file_id"]["options"]["dependencies"] == {"targeting": "pick"}

    def test_path_depends_on_targeting_path(self):
        properties = self._workbook_properties()
        assert properties["path"]["options"]["dependencies"] == {"targeting": "path"}

    def test_no_never_combine_tooltip_anywhere_in_the_row_schema(self):
        # The whole "never combine with Path"-style sentence is gone — targeting is now enforced
        # by real conditional visibility, not a warning the user has to notice and honor.
        schema_text = (COMPONENT_CONFIG_DIR / "configRowSchema.json").read_text().lower()
        assert "never combine" not in schema_text


class TestWorksheetSelectionSchema:
    """Change C: `worksheet.selection` (pick vs. name) replaces id/position inference; `position`
    is dropped from the UI schema entirely (kept only on the model, for v1-parity API configs).
    No UI tooltip mentions renaming anymore, since neither explicit branch renames."""

    def _worksheet_properties(self) -> dict:
        schema = json.loads((COMPONENT_CONFIG_DIR / "configRowSchema.json").read_text())
        return schema["properties"]["worksheet"]["properties"]

    def test_selection_is_the_first_field_with_the_expected_shape(self):
        properties = self._worksheet_properties()
        selection = properties["selection"]
        assert selection["enum"] == ["pick", "name"]
        assert selection["options"]["enum_titles"] == ["Pick existing", "By name (creates if missing)"]
        assert selection["default"] == "pick"
        assert selection["propertyOrder"] == 1
        assert all(selection["propertyOrder"] < other["propertyOrder"] for name, other in properties.items() if name != "selection")

    def test_id_depends_on_selection_pick(self):
        properties = self._worksheet_properties()
        assert properties["id"]["options"]["dependencies"] == {"selection": "pick"}

    def test_name_depends_on_selection_name(self):
        properties = self._worksheet_properties()
        assert properties["name"]["options"]["dependencies"] == {"selection": "name"}

    def test_position_field_is_absent_from_the_schema(self):
        # `position` stays on the `configuration.Worksheet` model for v1-parity/API configs, but
        # is deliberately not exposed as a UI field anymore.
        assert "position" not in self._worksheet_properties()

    def test_no_rename_mention_anywhere_in_the_worksheet_section(self):
        schema = json.loads((COMPONENT_CONFIG_DIR / "configRowSchema.json").read_text())
        worksheet_text = json.dumps(schema["properties"]["worksheet"]).lower()
        assert "rename" not in worksheet_text


class TestWriteModeSchema:
    """Change 3: `write_mode` (Full Load / Append / Upsert (Incremental)) replaces the old
    `append` checkbox; `key_columns` is a new array field shown only for Upsert."""

    def _row_schema(self) -> dict:
        return json.loads((COMPONENT_CONFIG_DIR / "configRowSchema.json").read_text())

    def test_append_field_is_gone(self):
        schema = self._row_schema()
        assert "append" not in schema["properties"]

    def test_write_mode_enum_and_titles(self):
        write_mode = self._row_schema()["properties"]["write_mode"]
        assert write_mode["enum"] == ["overwrite", "append", "upsert"]
        assert write_mode["options"]["enum_titles"] == ["Full Load", "Append", "Upsert (Incremental)"]
        assert write_mode["default"] == "overwrite"

    def test_key_columns_is_an_array_of_strings(self):
        key_columns = self._row_schema()["properties"]["key_columns"]
        assert key_columns["type"] == "array"
        assert key_columns["items"]["type"] == "string"
        assert key_columns["title"] == "Key Columns"

    def test_key_columns_only_shown_for_upsert(self):
        key_columns = self._row_schema()["properties"]["key_columns"]
        assert key_columns["options"]["dependencies"] == {"write_mode": "upsert"}


class TestInputAttributePlaceholders:
    """Change 2: ghost-text placeholders (`options.inputAttributes.placeholder`) on the text
    fields listed in the task — verified against
    `component-build-ui/references/advanced.md`'s "Placeholder Hints" pattern."""

    def _row_schema(self) -> dict:
        return json.loads((COMPONENT_CONFIG_DIR / "configRowSchema.json").read_text())

    def _root_schema(self) -> dict:
        return json.loads((COMPONENT_CONFIG_DIR / "configSchema.json").read_text())

    def test_root_account_fields_have_placeholders(self):
        account_properties = self._root_schema()["properties"]["account"]["properties"]
        assert account_properties["tenant_id"]["options"]["inputAttributes"] == {
            "placeholder": "00000000-0000-0000-0000-000000000000"
        }
        assert account_properties["site_url"]["options"]["inputAttributes"] == {
            "placeholder": "https://contoso.sharepoint.com/sites/DataTeam"
        }

    def test_destination_and_csv_fields_have_placeholders(self):
        schema = self._row_schema()
        destination = schema["properties"]["destination"]["properties"]
        assert destination["folder_path"]["options"]["inputAttributes"] == {"placeholder": "acme/reports/{{date}}/"}
        assert destination["date"]["options"]["inputAttributes"] == {"placeholder": "yesterday"}
        csv_section = schema["properties"]["csv"]["properties"]
        assert csv_section["file_name"]["options"]["inputAttributes"] == {"placeholder": "orders.csv"}

    def test_workbook_and_worksheet_fields_have_placeholders(self):
        schema = self._row_schema()
        workbook = schema["properties"]["workbook"]["properties"]
        assert workbook["path"]["options"]["inputAttributes"] == {
            "placeholder": "site://Site Name/folder/workbook.xlsx"
        }
        worksheet = schema["properties"]["worksheet"]["properties"]
        assert worksheet["name"]["options"]["inputAttributes"] == {"placeholder": "Sheet1"}


class TestNoStrftimeExamplesAnywhere:
    """Change 1: `{{date}}` replaces the strftime `{date:%Y-%m-%d}` placeholder everywhere the
    user can see it — the legacy form keeps *working* (silently), but must not appear in any
    schema tooltip/description/example, the README, or the configuration description anymore."""

    def test_no_strftime_style_examples_in_the_schemas(self):
        for filename in ("configRowSchema.json", "configSchema.json"):
            text = (COMPONENT_CONFIG_DIR / filename).read_text()
            assert "strftime" not in text.lower()
            assert "%Y" not in text
            assert "{date:" not in text

    def test_no_strftime_style_examples_in_the_readme_and_description(self):
        readme = (COMPONENT_CONFIG_DIR.parent / "README.md").read_text()
        description = (COMPONENT_CONFIG_DIR / "configuration_description.md").read_text()
        for text in (readme, description):
            assert "strftime" not in text.lower()
            assert "{date:" not in text


class TestRowSchemaHelperAccountType:
    """Change 6 UX addition: a hidden `destination.helper_account_type` mirrors the root config's
    `account.account_type` (root-watch pattern, keboola.ex-delta-lake precedent) so
    `destination.drive_id` can be shown only for SharePoint accounts — a row schema cannot
    reference the root config's fields directly in `options.dependencies`."""

    def _destination_properties(self) -> dict:
        schema = json.loads((COMPONENT_CONFIG_DIR / "configRowSchema.json").read_text())
        return schema["properties"]["destination"]["properties"]

    def test_helper_field_watches_the_root_account_type(self):
        helper = self._destination_properties()["helper_account_type"]
        assert helper["watch"] == {"val": "_metadata_.root.parameters.account.account_type"}
        assert helper["options"]["hidden"] is True

    def test_drive_id_depends_on_the_helper_being_sharepoint(self):
        drive_id = self._destination_properties()["drive_id"]
        assert drive_id["options"]["dependencies"] == {"helper_account_type": "sharepoint"}
