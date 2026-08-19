"""Cross-check: every `options.async.action` declared in the row/root schemas has a matching
`@sync_action(...)`-decorated method on `Component` (dropdown UX addition — the fastest way to
notice a typo'd action name in the schema, which would otherwise only surface as a runtime "action
not found" error inside the Keboola UI itself).
"""

import json
from pathlib import Path
from typing import Any

from keboola.component.base import _SYNC_ACTION_MAPPING

import component  # noqa: F401 - importing populates `_SYNC_ACTION_MAPPING` (decorators run on import)

REPO_ROOT = Path(__file__).resolve().parent.parent
COMPONENT_CONFIG_DIR = REPO_ROOT / "component_config"


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
