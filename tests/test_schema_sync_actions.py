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

REPO_ROOT = Path(__file__).resolve().parent.parent
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
            "How this row writes data to OneDrive/SharePoint — upload files as-is, write a table "
            "as CSV, or write a table into an Excel worksheet."
        )

    def test_destination_date_field_is_present_between_folder_path_and_conflict_behavior(self):
        destination_properties = self._row_schema()["properties"]["destination"]["properties"]
        assert "date" in destination_properties
        date_field = destination_properties["date"]
        assert date_field["title"] == "Date"
        assert date_field["propertyOrder"] > destination_properties["folder_path"]["propertyOrder"]
        assert date_field["propertyOrder"] < destination_properties["conflict_behavior"]["propertyOrder"]


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
