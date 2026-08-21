"""Component-level datadir-style tests for keboola.wr-onedrive-v2 (mocked HTTP, no network).

Merged from three retired modules into one flat file, in three sections below:

1. ``tests/component/test_component.py`` — ``Component`` sync actions (testConnection, listLibraries,
   listWorkbooks, listWorksheets, search, createWorkbook, createWorksheet, getWorksheets,
   listColumns) built against a ``KBC_DATADIR``-style fixture with a mocked ``GraphClient``.
2. ``tests/component/test_run_modes.py`` — ``Component.run()``'s orchestration: mode dispatch, file
   mode, CSV mode, error mapping, token-state persistence — same mocked-collaborator idiom.
3. ``tests/e2e/test_functional_http.py`` — the same production code path end-to-end
   (``Component.run()``, drive resolution, the uploader, the Excel writer) with only the HTTP
   transport boundary mocked (a small fake router, ``GraphFake``, patched onto
   ``requests.Session.request``) rather than the component's own collaborators. See that section's
   own module-level comment for why ``keboola.datadirtest``'s plain ``TestDataDir`` isn't used here
   (a ``SystemExit``-in-``unittest.TestCase`` hazard) — the VCR-cassette-replay flavor of that
   library IS used instead, in ``tests/test_functional.py``.

Shared fixtures below (``_oauth_credentials``/``_fake_token_provider``) are hoisted to the top,
deduplicated from the three originals; each section keeps its own ``_build_*_component`` helper
distinct (different signatures/fixture shapes), and one class name that existed in both section 2
and section 3 (``TestFileModeMergedTableInput``) is disambiguated as
``TestFileModeMergedTableInputHttp`` in section 3 to avoid a module-level name collision.
"""

import inspect
import json
import os
import re
import subprocess
import sys
import tempfile
import unittest
from dataclasses import dataclass, field
from datetime import UTC, datetime
from pathlib import Path
from unittest import mock
from unittest.mock import MagicMock

import dateparser
import pytest
import requests
from freezegun import freeze_time
from keboola.component.exceptions import UserException
from keboola.component.sync_actions import SelectElement

from client.excel_writer import XLSX_MIME_TYPE
from client.exceptions import GraphConnectionError, GraphPermissionError
from component import Component
from configuration import WriteMode


def _oauth_credentials(refresh_token: str = "refresh-config") -> dict:
    return {
        "id": "oauth-1",
        "created": "2026-01-01",
        "appKey": "client-1",
        "#appSecret": "secret-1",
        "oauthVersion": "2.0",
        "#data": json.dumps({"refresh_token": refresh_token}),
    }


def _fake_token_provider(rotated_refresh_token: str | None = None) -> MagicMock:
    """A `TokenProvider` double that never touches the network and reports a fixed rotation."""
    provider = MagicMock()
    provider.rotated_refresh_token = rotated_refresh_token
    return provider


# ----------------------------------------------------------------------------------------------
# 1. Component sync actions (tests/component/test_component.py)
# ----------------------------------------------------------------------------------------------


class TestComponent(unittest.TestCase):
    # set global time to 2010-10-10 - affects functions like datetime.now()
    @freeze_time("2010-10-10")
    # set KBC_DATADIR env to non-existing dir
    @mock.patch.dict(os.environ, {"KBC_DATADIR": "./non-existing-dir"})
    def test_run_no_cfg_fails(self):
        with self.assertRaises(ValueError):
            comp = Component()
            comp.run()


if __name__ == "__main__":
    unittest.main()


# --- Sync-action tests (plan Task 4) -----------------------------------------------------



def _write_config(data_dir, parameters: dict, oauth: dict | None) -> None:
    (data_dir / "in" / "tables").mkdir(parents=True, exist_ok=True)
    (data_dir / "in" / "files").mkdir(parents=True, exist_ok=True)
    (data_dir / "out").mkdir(parents=True, exist_ok=True)
    # `action: "run"` keeps `keboola.component.base.sync_action`'s wrapper in its non-sync-action
    # branch, so exceptions raised by the wrapped method propagate normally (as `raise e`)
    # instead of being caught, written to stderr, and turned into `sys.exit(1)` — the behavior
    # the wrapper applies whenever `configuration.action != "run"`, i.e. whenever no action is
    # set at all. That production behavior is exercised end-to-end in the datadir tests (a later
    # task); these unit tests exercise the sync-action *methods'* own logic directly.
    config: dict = {"parameters": parameters, "action": "run"}
    if oauth is not None:
        config["authorization"] = {"oauth_api": {"credentials": oauth}}
    (data_dir / "config.json").write_text(json.dumps(config))


_UNSET = object()


def _build_component(tmp_path, parameters: dict, oauth: dict | None = _UNSET, state: dict | None = None) -> Component:
    if oauth is _UNSET:
        oauth = _oauth_credentials()
    data_dir = tmp_path / "data"
    _write_config(data_dir, parameters, oauth)
    if state is not None:
        (data_dir / "in").mkdir(parents=True, exist_ok=True)
        (data_dir / "in" / "state.json").write_text(json.dumps(state))
    with mock.patch.dict(os.environ, {"KBC_DATADIR": str(data_dir)}):
        return Component()


def _fake_token_response(access_token="access-1", refresh_token="refresh-rotated-1") -> MagicMock:
    response = MagicMock()
    response.status_code = 200
    response.json.return_value = {
        "access_token": access_token,
        "refresh_token": refresh_token,
        "expires_in": 3599,
    }
    return response


class TestTestConnection:
    def test_happy_path_calls_me_endpoint(self, tmp_path):
        parameters = {"account": {"account_type": "private_onedrive"}}
        comp = _build_component(tmp_path, parameters)
        fake_client = MagicMock()

        with mock.patch("component.GraphClient", return_value=fake_client):
            result = comp.test_connection()

        assert result is None
        fake_client.get.assert_called_once_with("/me", params={"$select": "userPrincipalName"})

    def test_missing_oauth_raises_user_exception(self, tmp_path):
        parameters = {"account": {"account_type": "private_onedrive"}}
        comp = _build_component(tmp_path, parameters, oauth=None)

        with pytest.raises(UserException, match="not authorized"):
            comp.test_connection()

    def test_graph_error_is_mapped_to_user_exception(self, tmp_path):
        parameters = {"account": {"account_type": "private_onedrive"}}
        comp = _build_component(tmp_path, parameters)
        fake_client = MagicMock()
        fake_client.get.side_effect = GraphPermissionError("no access", status_code=403)

        with mock.patch("component.GraphClient", return_value=fake_client), pytest.raises(UserException):
            comp.test_connection()

    def test_invalid_account_params_raise_user_exception(self, tmp_path):
        # sharepoint requires tenant_id + site_url; neither is provided.
        parameters = {"account": {"account_type": "sharepoint"}}
        comp = _build_component(tmp_path, parameters)

        with pytest.raises(UserException, match="Invalid account configuration"):
            comp.test_connection()

    def test_ignores_row_parameters_that_would_fail_full_row_validation(self, tmp_path):
        # mode=table_excel would require workbook/worksheet for a full RowConfig; testConnection
        # must only look at `account` and must not attempt to validate the rest.
        parameters = {"mode": "table_excel", "account": {"account_type": "private_onedrive"}}
        comp = _build_component(tmp_path, parameters)
        fake_client = MagicMock()

        with mock.patch("component.GraphClient", return_value=fake_client):
            comp.test_connection()

        fake_client.get.assert_called_once()


class TestListLibraries:
    def test_returns_select_elements_with_drive_id_as_value(self, tmp_path):
        parameters = {
            "account": {
                "account_type": "sharepoint",
                "tenant_id": "tenant-1",
                "site_url": "https://contoso.sharepoint.com/sites/marketing",
            }
        }
        comp = _build_component(tmp_path, parameters)
        fake_client = MagicMock()
        drives = [
            {"id": "drive-1", "name": "Documents", "webUrl": "https://x/Documents"},
            {"id": "drive-2", "name": "Marketing Assets", "webUrl": "https://x/Marketing"},
        ]

        with (
            mock.patch("component.GraphClient", return_value=fake_client),
            mock.patch("component.get_site_id", return_value="site-1") as mock_get_site_id,
            mock.patch("component.list_drives", return_value=drives) as mock_list_drives,
        ):
            result = comp.list_libraries()

        mock_get_site_id.assert_called_once_with(fake_client, "https://contoso.sharepoint.com/sites/marketing")
        mock_list_drives.assert_called_once_with(fake_client, "site-1")
        assert result == [
            SelectElement(label="Documents", value="drive-1"),
            SelectElement(label="Marketing Assets", value="drive-2"),
        ]

    def test_business_account_returns_the_default_drive_as_a_single_element(self, tmp_path):
        # onedrive_for_business (and private_onedrive) have exactly one library: the account's
        # own default drive — no site/tenant needed, no more "requires SharePoint" error.
        parameters = {"account": {"account_type": "onedrive_for_business", "tenant_id": "tenant-1"}}
        comp = _build_component(tmp_path, parameters)
        fake_client = MagicMock()
        fake_client.get.return_value.json.return_value = {"id": "default-drive-1"}

        with mock.patch("component.GraphClient", return_value=fake_client):
            result = comp.list_libraries()

        fake_client.get.assert_called_once_with("/me/drive")
        assert result == [SelectElement(label="OneDrive (default)", value="default-drive-1")]

    def test_private_onedrive_account_returns_the_default_drive_as_a_single_element(self, tmp_path):
        parameters = {"account": {"account_type": "private_onedrive"}}
        comp = _build_component(tmp_path, parameters)
        fake_client = MagicMock()
        fake_client.get.return_value.json.return_value = {"id": "default-drive-2"}

        with mock.patch("component.GraphClient", return_value=fake_client):
            result = comp.list_libraries()

        fake_client.get.assert_called_once_with("/me/drive")
        assert result == [SelectElement(label="OneDrive (default)", value="default-drive-2")]

    def test_ignores_row_parameters_built_from_root_config_only(self, tmp_path):
        parameters = {
            "mode": "table_excel",  # would fail full RowConfig validation (no workbook/worksheet)
            "account": {
                "account_type": "sharepoint",
                "tenant_id": "tenant-1",
                "site_url": "https://contoso.sharepoint.com/sites/marketing",
            },
            "destination": {"drive_id": "should-be-ignored-by-listLibraries"},
        }
        comp = _build_component(tmp_path, parameters)
        fake_client = MagicMock()

        with (
            mock.patch("component.GraphClient", return_value=fake_client),
            mock.patch("component.get_site_id", return_value="site-1"),
            mock.patch("component.list_drives", return_value=[{"id": "d1", "name": "Documents"}]),
        ):
            result = comp.list_libraries()

        assert result == [SelectElement(label="Documents", value="d1")]


class TestListWorkbooks:
    """`listWorkbooks` (Excel workbook-picker UX addition): lists XLSX files in the row's
    target drive via Graph's drive-wide `search`, filtering out non-XLSX hits client-side (`q`
    matching is a loose text match, not a mime-type filter)."""

    def _items(self, count: int) -> list[dict]:
        return [
            {
                "id": f"file-{i}",
                "name": f"report{i}.xlsx",
                "file": {"mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
                "parentReference": {"path": "/drives/drive-1/root:/Reports"},
            }
            for i in range(count)
        ]

    def test_uses_row_drive_id_and_filters_by_mime_type_and_labels_by_path(self, tmp_path):
        parameters = {
            "account": {"account_type": "private_onedrive"},
            "workbook": {"drive_id": "drive-1"},
        }
        comp = _build_component(tmp_path, parameters)
        fake_client = MagicMock()
        fake_client.get_paged.return_value = iter(
            [
                {
                    "id": "file-xlsx",
                    "name": "budget.xlsx",
                    "file": {"mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
                    "parentReference": {"path": "/drives/drive-1/root:/Reports/2026"},
                },
                {
                    "id": "file-not-xlsx",
                    "name": "budget.xlsx.bak",
                    "file": {"mimeType": "application/octet-stream"},
                    "parentReference": {"path": "/drives/drive-1/root:/Reports"},
                },
            ]
        )

        with mock.patch("component.GraphClient", return_value=fake_client):
            result = comp.list_workbooks()

        called_url = fake_client.get_paged.call_args.args[0]
        assert called_url == "/drives/drive-1/root/search(q='.xlsx')"
        assert result == [SelectElement(label="Reports/2026/budget.xlsx", value="file-xlsx")]

    def test_falls_back_to_site_default_drive_for_sharepoint_when_drive_id_missing(self, tmp_path):
        parameters = {
            "account": {
                "account_type": "sharepoint",
                "tenant_id": "tenant-1",
                "site_url": "https://contoso.sharepoint.com/sites/marketing",
            },
        }
        comp = _build_component(tmp_path, parameters)
        fake_client = MagicMock()
        fake_client.get.return_value.json.return_value = {"id": "site-default-drive"}
        fake_client.get_paged.return_value = iter([])

        with (
            mock.patch("component.GraphClient", return_value=fake_client),
            mock.patch("component.get_site_id", return_value="site-1") as mock_get_site_id,
        ):
            comp.list_workbooks()

        mock_get_site_id.assert_called_once_with(fake_client, "https://contoso.sharepoint.com/sites/marketing")
        fake_client.get.assert_called_once_with("/sites/site-1/drive")
        called_url = fake_client.get_paged.call_args.args[0]
        assert called_url == "/drives/site-default-drive/root/search(q='.xlsx')"

    def test_falls_back_to_me_drive_for_business_and_private_when_drive_id_missing(self, tmp_path):
        parameters = {"account": {"account_type": "onedrive_for_business", "tenant_id": "tenant-1"}}
        comp = _build_component(tmp_path, parameters)
        fake_client = MagicMock()
        fake_client.get.return_value.json.return_value = {"id": "my-default-drive"}
        fake_client.get_paged.return_value = iter([])

        with mock.patch("component.GraphClient", return_value=fake_client):
            comp.list_workbooks()

        fake_client.get.assert_called_once_with("/me/drive")
        called_url = fake_client.get_paged.call_args.args[0]
        assert called_url == "/drives/my-default-drive/root/search(q='.xlsx')"

    def test_results_are_sorted_by_label(self, tmp_path):
        parameters = {"account": {"account_type": "private_onedrive"}, "workbook": {"drive_id": "drive-1"}}
        comp = _build_component(tmp_path, parameters)
        fake_client = MagicMock()
        fake_client.get_paged.return_value = iter(
            [
                {
                    "id": "file-b",
                    "name": "banana.xlsx",
                    "file": {"mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
                    "parentReference": {},
                },
                {
                    "id": "file-a",
                    "name": "apple.xlsx",
                    "file": {"mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
                    "parentReference": {},
                },
            ]
        )

        with mock.patch("component.GraphClient", return_value=fake_client):
            result = comp.list_workbooks()

        assert [element.value for element in result] == ["file-a", "file-b"]

    def test_truncates_and_warns_past_the_cap(self, tmp_path, caplog):
        parameters = {"account": {"account_type": "private_onedrive"}, "workbook": {"drive_id": "drive-1"}}
        comp = _build_component(tmp_path, parameters)
        fake_client = MagicMock()
        fake_client.get_paged.return_value = iter(self._items(250))

        with mock.patch("component.GraphClient", return_value=fake_client), caplog.at_level("WARNING"):
            result = comp.list_workbooks()

        assert len(result) == 200
        assert any("truncat" in message.lower() for message in caplog.messages)

    def test_no_truncation_warning_under_the_cap(self, tmp_path, caplog):
        parameters = {"account": {"account_type": "private_onedrive"}, "workbook": {"drive_id": "drive-1"}}
        comp = _build_component(tmp_path, parameters)
        fake_client = MagicMock()
        fake_client.get_paged.return_value = iter(self._items(5))

        with mock.patch("component.GraphClient", return_value=fake_client), caplog.at_level("WARNING"):
            result = comp.list_workbooks()

        assert len(result) == 5
        assert not any("truncat" in message.lower() for message in caplog.messages)


class TestListWorksheets:
    """`listWorksheets` (Excel worksheet-picker UX addition): lists a workbook's sheets (no
    per-sheet header read) via `client.excel_writer.list_worksheets`, labeling hidden sheets the
    same way `getWorksheets` does."""

    def test_returns_labels_with_hidden_suffix_sorted_by_position(self, tmp_path):
        parameters = {
            "account": {"account_type": "private_onedrive"},
            "workbook": {"drive_id": "drive-1", "file_id": "file-1"},
        }
        comp = _build_component(tmp_path, parameters)
        fake_client = MagicMock()
        worksheets = [
            {"id": "ws-1", "name": "Sheet1", "position": 0, "visibility": "Visible"},
            {"id": "ws-2", "name": "Sheet2", "position": 1, "visibility": "Hidden"},
        ]

        with (
            mock.patch("component.GraphClient", return_value=fake_client),
            mock.patch(
                "component.resolve_workbook", return_value=("drive-1", "file-1", False)
            ) as mock_resolve_workbook,
            mock.patch("component.list_worksheets_summary", return_value=worksheets) as mock_list_worksheets,
        ):
            result = comp.list_worksheets()

        mock_resolve_workbook.assert_called_once()
        assert mock_resolve_workbook.call_args.kwargs["create_if_missing"] is False
        mock_list_worksheets.assert_called_once_with(fake_client, "drive-1", "file-1")
        assert result == [
            SelectElement(label="Sheet1", value="ws-1"),
            SelectElement(label="Sheet2 (hidden)", value="ws-2"),
        ]

    def test_accepts_path_mode_workbook(self, tmp_path):
        parameters = {
            "account": {"account_type": "private_onedrive"},
            "workbook": {"path": "/book.xlsx"},
        }
        comp = _build_component(tmp_path, parameters)
        fake_client = MagicMock()

        with (
            mock.patch("component.GraphClient", return_value=fake_client),
            mock.patch("component.resolve_workbook", return_value=("drive-2", "file-2", False)) as mock_resolve,
            mock.patch("component.list_worksheets_summary", return_value=[]) as mock_list_worksheets,
        ):
            result = comp.list_worksheets()

        assert mock_resolve.call_args.args[2].path == "/book.xlsx"
        mock_list_worksheets.assert_called_once_with(fake_client, "drive-2", "file-2")
        assert result == []

    def test_missing_workbook_raises_invalid_workbook_configuration(self, tmp_path):
        parameters = {"account": {"account_type": "private_onedrive"}}
        comp = _build_component(tmp_path, parameters)

        with pytest.raises(UserException, match="Invalid workbook configuration"):
            comp.list_worksheets()


class TestSyncActionWorkbookTargetingCompatibility:
    """Change D: `listWorksheets`/`getWorksheets`/`createWorksheet` all read `parameters.workbook`
    through `Component._load_workbook_param` (`Workbook.model_validate(workbook_params or {})`),
    which already picks up `configuration.Workbook`'s new `targeting` switch and its
    ignore-the-other-form semantics for free — these tests are the sync-action-level proof that
    "pick" and "path" targeting, plus a stale hidden value from the form not currently selected,
    all resolve exactly as they do for a row-run (`tests/unit/test_configuration.py`'s model-level
    tests already cover every branch of that logic directly)."""

    def test_list_worksheets_targeting_pick_ignores_a_stale_hidden_path(self, tmp_path):
        parameters = {
            "account": {"account_type": "private_onedrive"},
            "workbook": {
                "targeting": "pick",
                "drive_id": "drive-1",
                "file_id": "file-1",
                "path": "/stale-from-path-mode.xlsx",
            },
        }
        comp = _build_component(tmp_path, parameters)
        fake_client = MagicMock()

        with (
            mock.patch("component.GraphClient", return_value=fake_client),
            mock.patch(
                "component.resolve_workbook", return_value=("drive-1", "file-1", False)
            ) as mock_resolve_workbook,
            mock.patch("component.list_worksheets_summary", return_value=[]),
        ):
            comp.list_worksheets()

        resolved_workbook = mock_resolve_workbook.call_args.args[2]
        assert resolved_workbook.path is None
        assert resolved_workbook.drive_id == "drive-1"
        assert resolved_workbook.file_id == "file-1"

    def test_get_worksheets_targeting_path_ignores_stale_hidden_ids(self, tmp_path):
        parameters = {
            "account": {"account_type": "private_onedrive"},
            "workbook": {
                "targeting": "path",
                "path": "/book.xlsx",
                "drive_id": "stale-drive-from-pick-mode",
                "file_id": "stale-file-from-pick-mode",
            },
        }
        comp = _build_component(tmp_path, parameters)
        fake_client = MagicMock()

        with (
            mock.patch("component.GraphClient", return_value=fake_client),
            mock.patch(
                "component.resolve_workbook", return_value=("drive-2", "file-2", False)
            ) as mock_resolve_workbook,
            mock.patch("component.list_worksheets_with_headers", return_value=[]),
        ):
            comp.get_worksheets()

        resolved_workbook = mock_resolve_workbook.call_args.args[2]
        assert resolved_workbook.path == "/book.xlsx"
        assert resolved_workbook.drive_id is None
        assert resolved_workbook.file_id is None

    def test_create_worksheet_targeting_pick_requires_both_ids(self, tmp_path):
        parameters = {
            "account": {"account_type": "private_onedrive"},
            "workbook": {"targeting": "pick", "drive_id": "drive-1"},
            "worksheet": {"name": "Sheet1"},
        }
        comp = _build_component(tmp_path, parameters)

        with pytest.raises(UserException, match="Invalid workbook configuration"):
            comp.create_worksheet()

    def test_missing_targeting_still_enforces_the_humanized_legacy_xor_message(self, tmp_path):
        # No `targeting` at all (row API payload/VCR cassette config predating Change B) keeps
        # today's mutual-exclusivity check, just humanized (Change B).
        parameters = {
            "account": {"account_type": "private_onedrive"},
            "workbook": {"path": "/book.xlsx", "drive_id": "drive-1", "file_id": "file-1"},
        }
        comp = _build_component(tmp_path, parameters)

        with pytest.raises(UserException, match="Choose one way to target the workbook"):
            comp.list_worksheets()


class TestBuildTokenProvider:
    def test_uses_common_authority_for_private_onedrive(self, tmp_path):
        parameters = {"account": {"account_type": "private_onedrive"}}
        comp = _build_component(tmp_path, parameters)
        account = comp._load_account()
        provider = comp._build_token_provider(account)
        fake_session = MagicMock()
        fake_session.post.return_value = _fake_token_response()
        provider._session = fake_session

        provider.get_access_token()

        assert fake_session.post.call_args.args[0] == "https://login.microsoftonline.com/common/oauth2/v2.0/token"

    def test_uses_tenant_id_authority_for_business_and_sharepoint(self, tmp_path):
        parameters = {"account": {"account_type": "onedrive_for_business", "tenant_id": "tenant-xyz"}}
        comp = _build_component(tmp_path, parameters)
        account = comp._load_account()
        provider = comp._build_token_provider(account)
        fake_session = MagicMock()
        fake_session.post.return_value = _fake_token_response()
        provider._session = fake_session

        provider.get_access_token()

        assert fake_session.post.call_args.args[0] == "https://login.microsoftonline.com/tenant-xyz/oauth2/v2.0/token"

    def test_state_refresh_token_is_tried_before_config_refresh_token(self, tmp_path):
        parameters = {"account": {"account_type": "private_onedrive"}}
        state = {"#refreshed_auth_data": json.dumps({"refresh_token": "refresh-state"})}
        comp = _build_component(
            tmp_path, parameters, oauth=_oauth_credentials(refresh_token="refresh-config"), state=state
        )
        account = comp._load_account()
        provider = comp._build_token_provider(account)
        fake_session = MagicMock()
        fake_session.post.return_value = _fake_token_response()
        provider._session = fake_session

        provider.get_access_token()

        assert fake_session.post.call_args.kwargs["data"]["refresh_token"] == "refresh-state"

    def test_missing_oauth_raises_user_exception(self, tmp_path):
        parameters = {"account": {"account_type": "private_onedrive"}}
        comp = _build_component(tmp_path, parameters, oauth=None)
        account = comp._load_account()

        with pytest.raises(UserException, match="not authorized"):
            comp._build_token_provider(account)


# --- MINOR-8: workbook/worksheet partial-model validation --------------------------------------


class TestRequireWorkbookPath:
    """`_require_workbook_path` (used by `search`/`createWorkbook`) validates through the
    `Workbook` partial model instead of poking raw `parameters.get()` values, while still
    preserving the exact v1-parity "missing" messages."""

    def test_search_missing_workbook_raises_v1_parity_message(self, tmp_path):
        parameters = {"account": {"account_type": "private_onedrive"}}
        comp = _build_component(tmp_path, parameters)

        with pytest.raises(
            UserException, match=re.escape('To search for a workbook please configure "parameters.workbook.path".')
        ):
            comp.search()

    def test_create_workbook_missing_workbook_raises_v1_parity_message(self, tmp_path):
        parameters = {"account": {"account_type": "private_onedrive"}}
        comp = _build_component(tmp_path, parameters)

        with pytest.raises(
            UserException, match=re.escape('To create workbook please configure "parameters.workbook.path".')
        ):
            comp.create_workbook()

    def test_create_workbook_empty_workbook_section_raises_v1_parity_message(self, tmp_path):
        # An empty `workbook: {}` section, not just a missing key entirely.
        parameters = {"account": {"account_type": "private_onedrive"}, "workbook": {}}
        comp = _build_component(tmp_path, parameters)

        with pytest.raises(
            UserException, match=re.escape('To create workbook please configure "parameters.workbook.path".')
        ):
            comp.create_workbook()

    def test_ids_only_workbook_still_raises_the_missing_path_message(self, tmp_path):
        # `search`/`createWorkbook` only ever accept a path (v1 parity) — a well-formed
        # ids-only `Workbook` must still be rejected with the path-specific message, not silently
        # accepted or given a generic validation error.
        parameters = {
            "account": {"account_type": "private_onedrive"},
            "workbook": {"drive_id": "drive-1", "file_id": "file-1"},
        }
        comp = _build_component(tmp_path, parameters)

        with pytest.raises(
            UserException, match=re.escape('To create workbook please configure "parameters.workbook.path".')
        ):
            comp.create_workbook()

    def test_stray_ids_alongside_a_valid_path_are_ignored_not_a_validation_error(self, tmp_path):
        # Change D fix: `_require_workbook_path` now builds the partial `Workbook` model from
        # `path` alone — `search`/`createWorkbook` only ever accept a path (no ids targeting,
        # design spec §5, and this class's own docstring), so any `drive_id`/`file_id` sitting
        # alongside a valid `path` (e.g. leftover from switching `workbook.targeting` back and
        # forth in the row's Worksheet-mode form) is simply irrelevant here, not a conflict.
        parameters = {
            "account": {"account_type": "private_onedrive"},
            "workbook": {"path": "/book.xlsx", "drive_id": "drive-1", "file_id": "file-1"},
        }
        comp = _build_component(tmp_path, parameters)
        fake_client = MagicMock()

        with (
            mock.patch("component.GraphClient", return_value=fake_client),
            mock.patch("component.search_workbook", return_value=None) as mock_search_workbook,
        ):
            result = comp.search()

        assert mock_search_workbook.call_args.args[2] == "/book.xlsx"
        assert result == {"file": None}

    def test_stale_hidden_ids_do_not_crash_when_targeting_is_pick(self, tmp_path):
        # The bug Change D's fix actually prevents: before it, `_require_workbook_path` fed the
        # *whole* `workbook_params` dict into `Workbook.model_validate`, which — for a row sitting
        # in `targeting: "pick"` — silently clears `path` back to `None` (by design: pick mode
        # ignores path entirely). The presence check above already passed on the raw, still-set
        # `path`, so the post-validation `assert workbook.path is not None` would then blow up
        # with an unhandled `AssertionError` (exit 2) instead of a clean result — even though
        # `search`/`createWorkbook` were always documented as path-only and should never care
        # about `targeting`/ids at all.
        parameters = {
            "account": {"account_type": "private_onedrive"},
            "workbook": {
                "targeting": "pick",
                "path": "/stale-leftover.xlsx",
                "drive_id": "drive-1",
                "file_id": "file-1",
            },
        }
        comp = _build_component(tmp_path, parameters)
        fake_client = MagicMock()

        with (
            mock.patch("component.GraphClient", return_value=fake_client),
            mock.patch("component.search_workbook", return_value=None) as mock_search_workbook,
        ):
            result = comp.search()  # must not raise (neither UserException nor AssertionError)

        assert mock_search_workbook.call_args.args[2] == "/stale-leftover.xlsx"
        assert result == {"file": None}

    def test_valid_path_is_returned_from_the_validated_model(self, tmp_path):
        parameters = {"account": {"account_type": "private_onedrive"}, "workbook": {"path": "/book.xlsx"}}
        comp = _build_component(tmp_path, parameters)
        fake_client = MagicMock()

        with (
            mock.patch("component.GraphClient", return_value=fake_client),
            mock.patch(
                "component.search_workbook",
                return_value=None,
            ) as mock_search_workbook,
        ):
            result = comp.search()

        mock_search_workbook.assert_called_once()
        assert mock_search_workbook.call_args.args[2] == "/book.xlsx"
        assert result == {"file": None}


class TestCreateWorksheetValidation:
    """`create_worksheet` validates `parameters.worksheet.name` through the `Worksheet` partial
    model instead of raw `parameters.get()`, while preserving the v1-parity "missing" message and
    the existing name-only targeting behavior (id/position are ignored, same as before)."""

    def test_missing_worksheet_section_raises_v1_parity_message(self, tmp_path):
        parameters = {"account": {"account_type": "private_onedrive"}, "workbook": {"path": "/book.xlsx"}}
        comp = _build_component(tmp_path, parameters)

        with pytest.raises(
            UserException, match=re.escape('To create worksheet please configure "parameters.worksheet.name".')
        ):
            comp.create_worksheet()

    def test_worksheet_section_without_name_raises_v1_parity_message(self, tmp_path):
        parameters = {
            "account": {"account_type": "private_onedrive"},
            "workbook": {"path": "/book.xlsx"},
            "worksheet": {"id": "some-id"},
        }
        comp = _build_component(tmp_path, parameters)

        with pytest.raises(
            UserException, match=re.escape('To create worksheet please configure "parameters.worksheet.name".')
        ):
            comp.create_worksheet()

    def test_uses_name_only_ignoring_id_and_position(self, tmp_path):
        parameters = {
            "account": {"account_type": "private_onedrive"},
            "workbook": {"path": "/book.xlsx"},
            "worksheet": {"name": "Sheet1", "id": "should-be-ignored", "position": 3},
        }
        comp = _build_component(tmp_path, parameters)
        fake_client = MagicMock()

        with (
            mock.patch("component.GraphClient", return_value=fake_client),
            mock.patch(
                "component.resolve_workbook", return_value=("drive-1", "file-1", False)
            ) as mock_resolve_workbook,
            mock.patch(
                "component.resolve_worksheet", return_value=("ws-1", True, "Sheet1")
            ) as mock_resolve_worksheet,
        ):
            result = comp.create_worksheet()

        mock_resolve_workbook.assert_called_once()
        called_worksheet = mock_resolve_worksheet.call_args.args[3]
        assert called_worksheet.name == "Sheet1"
        assert called_worksheet.id is None
        assert called_worksheet.position is None
        assert result == {"worksheet": {"driveId": "drive-1", "fileId": "file-1", "worksheetId": "ws-1"}}


# --- IMPORTANT-1/IMPORTANT-5 (phase 8 audit): shared `_sync_action_client` ----------------------


class TestSyncActionTokenPersistence:
    """Every sync action goes through `_sync_action_client`, which must persist a rotated
    refresh token to row state in `finally` — previously every sync action refreshed a token
    (Graph always rotates it) and silently discarded the rotation, forcing the next run to
    present an already-consumed refresh token to Microsoft."""

    def test_test_connection_persists_rotated_refresh_token(self, tmp_path):
        parameters = {"account": {"account_type": "private_onedrive"}}
        comp = _build_component(tmp_path, parameters)
        fake_client = MagicMock()

        with (
            mock.patch("component.GraphClient", return_value=fake_client),
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider("rotated-tc-1")),
        ):
            comp.test_connection()

        state = json.loads((tmp_path / "data" / "out" / "state.json").read_text())
        payload = json.loads(state["#refreshed_auth_data"])
        assert payload["refresh_token"] == "rotated-tc-1"

    def test_search_persists_rotated_refresh_token(self, tmp_path):
        # A different sync action than test_connection — proves the persistence lives in the
        # shared helper, not duplicated (and possibly missed) per action.
        parameters = {"account": {"account_type": "private_onedrive"}, "workbook": {"path": "/book.xlsx"}}
        comp = _build_component(tmp_path, parameters)
        fake_client = MagicMock()

        with (
            mock.patch("component.GraphClient", return_value=fake_client),
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider("rotated-search-1")),
            mock.patch("component.search_workbook", return_value=None),
        ):
            comp.search()

        state = json.loads((tmp_path / "data" / "out" / "state.json").read_text())
        payload = json.loads(state["#refreshed_auth_data"])
        assert payload["refresh_token"] == "rotated-search-1"

    def test_no_state_written_when_nothing_rotated(self, tmp_path):
        parameters = {"account": {"account_type": "private_onedrive"}}
        comp = _build_component(tmp_path, parameters)
        fake_client = MagicMock()

        with (
            mock.patch("component.GraphClient", return_value=fake_client),
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider(None)),
        ):
            comp.test_connection()

        assert not (tmp_path / "data" / "out" / "state.json").exists()

    def test_token_still_persisted_when_the_sync_action_raises(self, tmp_path):
        parameters = {"account": {"account_type": "private_onedrive"}}
        comp = _build_component(tmp_path, parameters)
        fake_client = MagicMock()
        fake_client.get.side_effect = GraphPermissionError("no access", status_code=403)

        with (
            mock.patch("component.GraphClient", return_value=fake_client),
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider("rotated-tc-2")),
            pytest.raises(UserException),
        ):
            comp.test_connection()

        state = json.loads((tmp_path / "data" / "out" / "state.json").read_text())
        payload = json.loads(state["#refreshed_auth_data"])
        assert payload["refresh_token"] == "rotated-tc-2"


class TestSyncActionRetryBudget:
    """IMPORTANT-5 (phase 8 audit): sync actions block the Keboola UI while they run, so they
    must use a reduced, fast-fail retry budget instead of `run()`'s full unattended-job defaults
    — otherwise a rate-limited Graph call could hang the UI for minutes."""

    def test_sync_action_client_uses_reduced_retry_budget(self, tmp_path):
        parameters = {"account": {"account_type": "private_onedrive"}}
        comp = _build_component(tmp_path, parameters)

        with (
            mock.patch("component.GraphClient") as mock_graph_client,
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider()),
        ):
            comp.test_connection()

        mock_graph_client.assert_called_once()
        _, kwargs = mock_graph_client.call_args
        assert kwargs["total_wait_cap_seconds"] == 15
        assert kwargs["max_retry_attempts"] == 2


class TestListColumnsSyncAction:
    """`listColumns` mirrors wr-delta-lake's list_table_columns (forwarded Storage token)."""

    @staticmethod
    def _component(tmp_path, monkeypatch, tables, url="https://connection.test", token="tok-123"):
        import json as _json
        from unittest import mock

        datadir = tmp_path / "data"
        (datadir / "in" / "tables").mkdir(parents=True)
        (datadir / "out" / "tables").mkdir(parents=True)
        config = {
            "parameters": {"mode": "worksheet", "account": {"account_type": "private_onedrive"}},
            "storage": {"input": {"tables": tables}},
            "action": "run",
        }
        (datadir / "config.json").write_text(_json.dumps(config))
        monkeypatch.setenv("KBC_DATADIR", str(datadir))
        if url:
            monkeypatch.setenv("KBC_URL", url)
        if token:
            monkeypatch.setenv("KBC_TOKEN", token)
        import component as component_module

        return component_module.Component(), mock

    def test_returns_columns_from_storage_api(self, tmp_path, monkeypatch):
        comp, mock = self._component(
            tmp_path, monkeypatch, [{"source": "in.c-x.orders", "destination": "orders.csv"}]
        )
        response = mock.MagicMock(status_code=200)
        response.json.return_value = {"columns": ["order_id", "status"]}
        with mock.patch("component.requests.get", return_value=response) as get:
            elements = comp.list_columns()
        assert [e.value for e in elements] == ["order_id", "status"]
        called_url = get.call_args.args[0]
        assert called_url.endswith("/v2/storage/tables/in.c-x.orders")
        assert get.call_args.kwargs["headers"]["X-StorageApi-Token"] == "tok-123"

    def test_no_input_table_is_user_error(self, tmp_path, monkeypatch):
        import pytest
        from keboola.component.exceptions import UserException

        comp, _ = self._component(tmp_path, monkeypatch, [])
        with pytest.raises(UserException, match="Map an input table"):
            comp.list_columns()

    def test_missing_forwarded_token_is_user_error(self, tmp_path, monkeypatch):
        import pytest
        from keboola.component.exceptions import UserException

        monkeypatch.delenv("KBC_TOKEN", raising=False)
        monkeypatch.delenv("KBC_URL", raising=False)
        comp, _ = self._component(
            tmp_path, monkeypatch, [{"source": "in.c-x.orders", "destination": "o.csv"}],
            url=None, token=None,
        )
        with pytest.raises(UserException, match="forwardToken"):
            comp.list_columns()


# ----------------------------------------------------------------------------------------------
# 2. Component.run() orchestration (tests/component/test_run_modes.py)
# ----------------------------------------------------------------------------------------------


def _build_run_component(
    tmp_path,
    parameters: dict,
    *,
    files: dict[str, bytes] | None = None,
    tables: dict[str, str] | None = None,
    oauth: dict | None = None,
) -> Component:
    """Build a `Component` from a `KBC_DATADIR`-style fixture (datadir pattern).

    `files`/`tables` map a file name (under `data/in/files` / `data/in/tables`) to its content.
    Table entries get a minimal input manifest (`{"id": ...}`) so `TableDefinition` resolves them
    as input tables, matching what the platform actually writes.
    """
    data_dir = tmp_path / "data"
    files_dir = data_dir / "in" / "files"
    tables_dir = data_dir / "in" / "tables"
    files_dir.mkdir(parents=True, exist_ok=True)
    tables_dir.mkdir(parents=True, exist_ok=True)
    (data_dir / "out").mkdir(parents=True, exist_ok=True)

    for name, content in (files or {}).items():
        (files_dir / name).write_bytes(content)

    for name, content in (tables or {}).items():
        (tables_dir / name).write_text(content)
        (tables_dir / f"{name}.manifest").write_text(json.dumps({"id": f"in.c-main.{name}"}))

    config = {
        "parameters": parameters,
        "action": "run",
        "authorization": {"oauth_api": {"credentials": oauth if oauth is not None else _oauth_credentials()}},
    }
    (data_dir / "config.json").write_text(json.dumps(config))

    with mock.patch.dict(os.environ, {"KBC_DATADIR": str(data_dir)}):
        return Component()


class TestRunOrchestratorShape:
    def test_run_is_a_thin_orchestrator_under_30_lines(self):
        source = inspect.getsource(Component.run)
        assert len(source.splitlines()) <= 30

    def test_run_uses_the_default_unattended_job_retry_budget(self, tmp_path):
        """IMPORTANT-5 (phase 8 audit): unlike a sync action, `run()` is an unattended job — it
        must keep `GraphClient`'s full default retry budget, not the sync action's fast-fail one."""
        parameters = {"mode": "file", "account": {"account_type": "private_onedrive"}, "destination": {}}
        comp = _build_run_component(tmp_path, parameters, files={"a.txt": b"aaa"})

        with (
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider()),
            mock.patch("component.GraphClient") as mock_graph_client,
            mock.patch("component.resolve_drive_id", return_value="drive-1"),
            mock.patch("component.ensure_folder", return_value="root-id"),
            mock.patch("component.upload_file", return_value={"id": "item-1"}),
        ):
            comp.run()

        mock_graph_client.assert_called_once_with(token_provider=mock.ANY)


class TestFileMode:
    @freeze_time("2026-08-17")
    def test_uploads_every_input_file_with_resolved_folder_path(self, tmp_path):
        parameters = {
            "mode": "file",
            "account": {"account_type": "private_onedrive"},
            "destination": {"folder_path": "reports/{date:%Y-%m-%d}", "conflict_behavior": "replace"},
        }
        comp = _build_run_component(tmp_path, parameters, files={"a.txt": b"aaa", "b.txt": b"bbb"})

        with (
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider()),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            mock.patch("component.resolve_drive_id", return_value="drive-1") as mock_resolve_drive,
            mock.patch("component.ensure_folder", return_value="parent-1") as mock_ensure_folder,
            mock.patch("component.upload_file", return_value={"id": "item-1"}) as mock_upload,
        ):
            comp.run()

        mock_resolve_drive.assert_called_once()
        mock_ensure_folder.assert_called_once_with(mock.ANY, "drive-1", "reports/2026-08-17")
        assert mock_upload.call_count == 2
        uploaded_names = {call_args.args[4] for call_args in mock_upload.call_args_list}
        assert uploaded_names == {"a.txt", "b.txt"}
        for call_args in mock_upload.call_args_list:
            assert call_args.args[1] == "drive-1"
            assert call_args.args[2] == "parent-1"
            assert call_args.args[5] == "replace"

    def test_zero_files_and_zero_tables_raises_user_exception(self, tmp_path):
        # Change A: mode 'file' now processes both input mappings, so the "nothing to do" check
        # covers both — zero files *and* zero tables is required to trigger it.
        parameters = {"mode": "file", "account": {"account_type": "private_onedrive"}, "destination": {}}
        comp = _build_run_component(tmp_path, parameters, files={}, tables={})

        with (
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider()),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            mock.patch("component.resolve_drive_id", return_value="drive-1"),
            pytest.raises(UserException, match="No files or tables found in the input mapping"),
        ):
            comp.run()


class TestFileModeMergedTableInput:
    """Change A: mode 'file' now processes *both* of the row's input mappings — every file from
    the file input mapping uploaded as-is, and every table from the table input mapping written
    as CSV (merged from the pre-merge 'file'/'table_csv' split)."""

    def test_tables_only_writes_each_table_as_csv(self, tmp_path):
        parameters = {"mode": "file", "account": {"account_type": "private_onedrive"}, "destination": {}}
        comp = _build_run_component(tmp_path, parameters, tables={"mytable": "id,name\n1,a\n"})

        with (
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider()),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            mock.patch("component.resolve_drive_id", return_value="drive-1"),
            mock.patch("component.ensure_folder", return_value="root-id"),
            mock.patch("component.upload_file", return_value={"id": "item-1"}) as mock_upload,
        ):
            comp.run()

        assert mock_upload.call_count == 1
        assert mock_upload.call_args.args[4] == "mytable.csv"

    def test_files_and_tables_together_are_both_uploaded(self, tmp_path):
        parameters = {"mode": "file", "account": {"account_type": "private_onedrive"}, "destination": {}}
        comp = _build_run_component(
            tmp_path, parameters, files={"a.txt": b"aaa"}, tables={"mytable": "id,name\n1,a\n"}
        )

        with (
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider()),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            mock.patch("component.resolve_drive_id", return_value="drive-1"),
            mock.patch("component.ensure_folder", return_value="root-id"),
            mock.patch("component.upload_file", return_value={"id": "item-1"}) as mock_upload,
        ):
            comp.run()

        assert mock_upload.call_count == 2
        uploaded_names = {call_args.args[4] for call_args in mock_upload.call_args_list}
        assert uploaded_names == {"a.txt", "mytable.csv"}

    def test_multiple_tables_each_become_their_own_csv(self, tmp_path):
        parameters = {"mode": "file", "account": {"account_type": "private_onedrive"}, "destination": {}}
        comp = _build_run_component(
            tmp_path, parameters, tables={"a.csv": "id\n1\n", "b.csv": "id\n2\n"}
        )

        with (
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider()),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            mock.patch("component.resolve_drive_id", return_value="drive-1"),
            mock.patch("component.ensure_folder", return_value="root-id"),
            mock.patch("component.upload_file", return_value={"id": "item-1"}) as mock_upload,
        ):
            comp.run()

        assert mock_upload.call_count == 2
        uploaded_names = {call_args.args[4] for call_args in mock_upload.call_args_list}
        assert uploaded_names == {"a.csv", "b.csv"}

    def test_csv_file_name_applies_when_exactly_one_table_is_mapped(self, tmp_path):
        parameters = {
            "mode": "file",
            "account": {"account_type": "private_onedrive"},
            "destination": {},
            "csv": {"file_name": "custom.csv"},
        }
        comp = _build_run_component(tmp_path, parameters, tables={"mytable": "id\n1\n"})

        with (
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider()),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            mock.patch("component.resolve_drive_id", return_value="drive-1"),
            mock.patch("component.ensure_folder", return_value="root-id"),
            mock.patch("component.upload_file", return_value={"id": "item-1"}) as mock_upload,
        ):
            comp.run()

        assert mock_upload.call_args.args[4] == "custom.csv"

    def test_csv_file_name_with_multiple_tables_raises_user_exception(self, tmp_path):
        parameters = {
            "mode": "file",
            "account": {"account_type": "private_onedrive"},
            "destination": {},
            "csv": {"file_name": "custom.csv"},
        }
        comp = _build_run_component(
            tmp_path, parameters, tables={"a.csv": "id\n1\n", "b.csv": "id\n2\n"}
        )

        with (
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider()),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            mock.patch("component.resolve_drive_id", return_value="drive-1"),
            pytest.raises(UserException, match="csv.file_name can only be used when exactly one table is mapped"),
        ):
            comp.run()

    @freeze_time("2026-08-17")
    def test_csv_file_name_resolves_the_double_brace_date_placeholder(self, tmp_path):
        # Change 1: `{{date}}` applies to `csv.file_name`, not just `destination.folder_path`,
        # resolved against the same `now` (job start here — `destination.date` unset).
        parameters = {
            "mode": "file",
            "account": {"account_type": "private_onedrive"},
            "destination": {},
            "csv": {"file_name": "orders-{{date}}.csv"},
        }
        comp = _build_run_component(tmp_path, parameters, tables={"mytable": "id\n1\n"})

        with (
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider()),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            mock.patch("component.resolve_drive_id", return_value="drive-1"),
            mock.patch("component.ensure_folder", return_value="root-id"),
            mock.patch("component.upload_file", return_value={"id": "item-1"}) as mock_upload,
        ):
            comp.run()

        assert mock_upload.call_args.args[4] == "orders-2026-08-17.csv"

    def test_csv_file_name_legacy_strftime_placeholder_still_resolves(self, tmp_path):
        # Silent legacy handling: an already-configured row's `csv.file_name` could still use the
        # pre-`{{date}}` strftime form — it must keep resolving unchanged.
        parameters = {
            "mode": "file",
            "account": {"account_type": "private_onedrive"},
            "destination": {},
            "csv": {"file_name": "orders-{date:%Y}.csv"},
        }
        comp = _build_run_component(tmp_path, parameters, tables={"mytable": "id\n1\n"})

        with (
            freeze_time("2026-08-17"),
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider()),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            mock.patch("component.resolve_drive_id", return_value="drive-1"),
            mock.patch("component.ensure_folder", return_value="root-id"),
            mock.patch("component.upload_file", return_value={"id": "item-1"}) as mock_upload,
        ):
            comp.run()

        assert mock_upload.call_args.args[4] == "orders-2026.csv"


class TestResolveNow:
    """`Component._resolve_now` (Change 2 UX addition): resolves the timestamp fed into
    `folder_path`'s `{date:...}` placeholders — job start (UTC) by default, or `destination.date`
    via `dateparser` when set."""

    @freeze_time("2026-08-17 12:34:56")
    def test_none_returns_job_start_utc(self, tmp_path):
        parameters = {"mode": "file", "account": {"account_type": "private_onedrive"}, "destination": {}}
        comp = _build_run_component(tmp_path, parameters, files={"a.txt": b"aaa"})

        now = comp._resolve_now(None)

        assert now == datetime(2026, 8, 17, 12, 34, 56, tzinfo=UTC)

    @freeze_time("2026-08-17 12:34:56")
    def test_blank_string_returns_job_start_utc(self, tmp_path):
        parameters = {"mode": "file", "account": {"account_type": "private_onedrive"}, "destination": {}}
        comp = _build_run_component(tmp_path, parameters, files={"a.txt": b"aaa"})

        now = comp._resolve_now("")

        assert now == datetime(2026, 8, 17, 12, 34, 56, tzinfo=UTC)

    def test_absolute_date_is_parsed_exactly(self, tmp_path):
        parameters = {"mode": "file", "account": {"account_type": "private_onedrive"}, "destination": {}}
        comp = _build_run_component(tmp_path, parameters, files={"a.txt": b"aaa"})

        now = comp._resolve_now("2026-01-31")

        assert now == datetime(2026, 1, 31, tzinfo=UTC)

    @pytest.mark.parametrize("value", ["yesterday", "3 days ago", "last week"])
    def test_relative_date_matches_dateparsers_own_output(self, tmp_path, value):
        parameters = {"mode": "file", "account": {"account_type": "private_onedrive"}, "destination": {}}
        comp = _build_run_component(tmp_path, parameters, files={"a.txt": b"aaa"})
        expected = dateparser.parse(value, settings={"TIMEZONE": "UTC", "RETURN_AS_TIMEZONE_AWARE": True})

        now = comp._resolve_now(value)

        # No frozen clock (per the design note: compare against dateparser's own output) — allow a
        # small slop for the two `dateparser.parse` calls landing a few seconds apart.
        assert abs((now - expected).total_seconds()) < 5

    def test_unparseable_value_raises_user_exception_naming_the_value_and_examples(self, tmp_path):
        parameters = {"mode": "file", "account": {"account_type": "private_onedrive"}, "destination": {}}
        comp = _build_run_component(tmp_path, parameters, files={"a.txt": b"aaa"})

        with pytest.raises(UserException) as exc_info:
            comp._resolve_now("not-a-real-date-xyz123")

        message = str(exc_info.value)
        assert "not-a-real-date-xyz123" in message
        assert "yesterday" in message


class TestDestinationDateFolderPath:
    """End-to-end: `destination.date` feeds `resolve_placeholders` via `_resolve_now`, in place
    of the job's own start time, for both file and CSV mode."""

    def test_relative_date_resolves_the_folder_path_placeholder(self, tmp_path):
        parameters = {
            "mode": "file",
            "account": {"account_type": "private_onedrive"},
            "destination": {"folder_path": "reports/{date:%Y-%m-%d}", "date": "yesterday"},
        }
        comp = _build_run_component(tmp_path, parameters, files={"a.txt": b"aaa"})
        expected_date = dateparser.parse("yesterday", settings={"TIMEZONE": "UTC", "RETURN_AS_TIMEZONE_AWARE": True})
        expected_path = f"reports/{expected_date:%Y-%m-%d}"

        with (
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider()),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            mock.patch("component.resolve_drive_id", return_value="drive-1"),
            mock.patch("component.ensure_folder", return_value="parent-1") as mock_ensure_folder,
            mock.patch("component.upload_file", return_value={"id": "item-1"}),
        ):
            comp.run()

        mock_ensure_folder.assert_called_once_with(mock.ANY, "drive-1", expected_path)

    def test_empty_destination_date_keeps_job_start_behavior(self, tmp_path):
        parameters = {
            "mode": "table_csv",
            "account": {"account_type": "private_onedrive"},
            "destination": {"folder_path": "reports/{date:%Y-%m-%d}"},
        }
        comp = _build_run_component(tmp_path, parameters, tables={"mytable": "id\n1\n"})

        with (
            freeze_time("2026-08-17"),
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider()),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            mock.patch("component.resolve_drive_id", return_value="drive-1"),
            mock.patch("component.ensure_folder", return_value="parent-1") as mock_ensure_folder,
            mock.patch("component.upload_file", return_value={"id": "item-1"}),
        ):
            comp.run()

        mock_ensure_folder.assert_called_once_with(mock.ANY, "drive-1", "reports/2026-08-17")

    def test_invalid_destination_date_raises_user_exception(self, tmp_path):
        parameters = {
            "mode": "file",
            "account": {"account_type": "private_onedrive"},
            "destination": {"folder_path": "reports", "date": "not-a-real-date-xyz123"},
        }
        comp = _build_run_component(tmp_path, parameters, files={"a.txt": b"aaa"})

        with (
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider()),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            pytest.raises(UserException, match="not-a-real-date-xyz123"),
        ):
            comp.run()


class TestIgnoredInputWarnings:
    """Change A: mode 'file' no longer warns about an ignored table input mapping — it now
    processes both input mappings, so a table mapping is never ignored there anymore (the old
    "table input mapping is ignored" warning is gone entirely). Mode 'worksheet' keeps its own
    "file input mapping is ignored" warning unchanged — it still only ever reads its single input
    table."""

    def test_worksheet_mode_warns_when_file_input_mapping_is_also_present(self, tmp_path, caplog):
        parameters = {
            "mode": "worksheet",
            "account": {"account_type": "onedrive_for_business", "tenant_id": "tenant-1"},
            "workbook": {"path": "/book.xlsx"},
            "worksheet": {"name": "Sheet1"},
        }
        comp = _build_run_component(
            tmp_path, parameters, tables={"mytable": "id,name\n1,a\n"}, files={"ignored.txt": b"x"}
        )

        with (
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider()),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            mock.patch("component.resolve_workbook", return_value=("wb-drive", "wb-file", False)),
            mock.patch("component.workbook_session", return_value=_fake_session_context_manager("session-1")),
            mock.patch("component.resolve_worksheet", return_value=("sheet-1", False, "Sheet1")),
            mock.patch("component.write_table", return_value=True),
            caplog.at_level("WARNING"),
        ):
            comp.run()

        assert any("file input mapping is ignored" in message for message in caplog.messages)

    def test_worksheet_mode_no_warning_when_file_input_mapping_is_empty(self, tmp_path, caplog):
        parameters = {
            "mode": "worksheet",
            "account": {"account_type": "onedrive_for_business", "tenant_id": "tenant-1"},
            "workbook": {"path": "/book.xlsx"},
            "worksheet": {"name": "Sheet1"},
        }
        comp = _build_run_component(tmp_path, parameters, tables={"mytable": "id,name\n1,a\n"})

        with (
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider()),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            mock.patch("component.resolve_workbook", return_value=("wb-drive", "wb-file", False)),
            mock.patch("component.workbook_session", return_value=_fake_session_context_manager("session-1")),
            mock.patch("component.resolve_worksheet", return_value=("sheet-1", False, "Sheet1")),
            mock.patch("component.write_table", return_value=True),
            caplog.at_level("WARNING"),
        ):
            comp.run()

        assert not any("file input mapping is ignored" in message for message in caplog.messages)

    def test_file_mode_never_warns_about_the_table_input_mapping(self, tmp_path, caplog):
        """The pre-merge "table input mapping is ignored" warning is gone (Change A) — mode
        'file' now uses a mapped table (writing it as CSV) rather than ignoring it."""
        parameters = {"mode": "file", "account": {"account_type": "private_onedrive"}, "destination": {}}
        comp = _build_run_component(
            tmp_path, parameters, files={"a.txt": b"aaa"}, tables={"used": "id\n1\n"}
        )

        with (
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider()),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            mock.patch("component.resolve_drive_id", return_value="drive-1"),
            mock.patch("component.ensure_folder", return_value="root-id"),
            mock.patch("component.upload_file", return_value={"id": "item-1"}),
            caplog.at_level("WARNING"),
        ):
            comp.run()

        assert not any("ignored" in message for message in caplog.messages)


class TestWorksheetModeCardinality:
    """Mode 'worksheet' (the pre-merge 'table_excel') still requires exactly one input table —
    `_require_single_input_table`'s v1-parity messages are unaffected by Change A's file-mode
    merge, which only touches mode 'file'."""

    def test_zero_tables_raises_v1_parity_message(self, tmp_path):
        parameters = {
            "mode": "worksheet",
            "account": {"account_type": "onedrive_for_business", "tenant_id": "tenant-1"},
            "workbook": {"path": "/book.xlsx"},
            "worksheet": {"name": "Sheet1"},
        }
        comp = _build_run_component(tmp_path, parameters, tables={})

        with (
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider()),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            pytest.raises(UserException, match=re.escape('No CSV file found in "/data/in/tables".')),
        ):
            comp.run()

    def test_multiple_tables_raises_v1_parity_message_comma_joined(self, tmp_path):
        parameters = {
            "mode": "worksheet",
            "account": {"account_type": "onedrive_for_business", "tenant_id": "tenant-1"},
            "workbook": {"path": "/book.xlsx"},
            "worksheet": {"name": "Sheet1"},
        }
        comp = _build_run_component(tmp_path, parameters, tables={"a.csv": "id\n1\n", "b.csv": "id\n2\n"})
        expected = re.escape('Expected one CSV file, found multiple: "a.csv", "b.csv".')

        with (
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider()),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            pytest.raises(UserException, match=expected),
        ):
            comp.run()


class TestCsvModeUpload:
    def test_passthrough_when_options_match_storage_defaults(self, tmp_path):
        parameters = {
            "mode": "table_csv",
            "account": {"account_type": "private_onedrive"},
            "destination": {"conflict_behavior": "fail"},
        }
        comp = _build_run_component(tmp_path, parameters, tables={"mytable": "id,name\n1,a\n2,b\n"})

        with (
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider()),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            mock.patch("component.resolve_drive_id", return_value="drive-1"),
            mock.patch("component.ensure_folder", return_value="root-id"),
            mock.patch("component.upload_file", return_value={"id": "item-1"}) as mock_upload,
        ):
            comp.run()

        expected_source_path = str(tmp_path / "data" / "in" / "tables" / "mytable")
        upload_path = mock_upload.call_args.args[3]
        file_name = mock_upload.call_args.args[4]
        assert upload_path == expected_source_path  # streamed as-is, no rewrite
        assert file_name == "mytable.csv"  # csv.file_name defaults to "<table name>.csv"
        assert os.path.exists(upload_path)  # the input file itself, never deleted

    def test_rewrite_when_delimiter_differs_and_header_disabled_drops_first_row(self, tmp_path):
        parameters = {
            "mode": "table_csv",
            "account": {"account_type": "private_onedrive"},
            "destination": {},
            "csv": {"delimiter": ";", "include_header": False},
        }
        comp = _build_run_component(tmp_path, parameters, tables={"mytable": "id,name\n1,a\n2,b\n"})
        captured: dict = {}

        def _fake_upload_file(client, drive_id, parent_id, local_path, file_name, conflict_behavior):
            # Read the rewritten temp file's content before the component's `finally` deletes it.
            with open(local_path, newline="") as file_handle:
                captured["content"] = file_handle.read()
            captured["path"] = local_path
            return {"id": "item-1"}

        with (
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider()),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            mock.patch("component.resolve_drive_id", return_value="drive-1"),
            mock.patch("component.ensure_folder", return_value="root-id"),
            mock.patch("component.upload_file", side_effect=_fake_upload_file),
        ):
            comp.run()

        assert captured["content"] == "1;a\r\n2;b\r\n"  # header dropped, ";"-delimited
        assert not os.path.exists(captured["path"])  # cleaned up in `finally`
        source_path = str(tmp_path / "data" / "in" / "tables" / "mytable")
        assert captured["path"] != source_path  # a genuinely different (temp) file
        # Never scratch under `data/out/` — only `state.json` is written there.
        data_out_dir = str(tmp_path / "data" / "out")
        assert not captured["path"].startswith(data_out_dir)
        assert captured["path"].startswith(tempfile.gettempdir())

    def test_rewrite_when_enclosure_differs_streams_row_by_row(self, tmp_path):
        parameters = {
            "mode": "table_csv",
            "account": {"account_type": "private_onedrive"},
            "destination": {},
            "csv": {"enclosure": "'"},
        }
        comp = _build_run_component(tmp_path, parameters, tables={"mytable": 'id,name\n1,"a,b"\n'})
        captured: dict = {}

        def _fake_upload_file(client, drive_id, parent_id, local_path, file_name, conflict_behavior):
            with open(local_path, newline="") as file_handle:
                captured["content"] = file_handle.read()
            return {"id": "item-1"}

        with (
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider()),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            mock.patch("component.resolve_drive_id", return_value="drive-1"),
            mock.patch("component.ensure_folder", return_value="root-id"),
            mock.patch("component.upload_file", side_effect=_fake_upload_file),
        ):
            comp.run()

        assert captured["content"] == "id,name\r\n1,'a,b'\r\n"


class TestTokenStatePersistence:
    def test_rotated_token_persisted_on_success(self, tmp_path):
        parameters = {"mode": "file", "account": {"account_type": "private_onedrive"}, "destination": {}}
        comp = _build_run_component(tmp_path, parameters, files={"a.txt": b"aaa"})

        with (
            mock.patch(
                "component.RefreshTokenProvider", return_value=_fake_token_provider("new-refresh-token")
            ),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            mock.patch("component.resolve_drive_id", return_value="drive-1"),
            mock.patch("component.ensure_folder", return_value="root-id"),
            mock.patch("component.upload_file", return_value={"id": "item-1"}),
        ):
            comp.run()

        state = json.loads((tmp_path / "data" / "out" / "state.json").read_text())
        payload = json.loads(state["#refreshed_auth_data"])
        assert payload["refresh_token"] == "new-refresh-token"

    def test_rotated_token_persisted_even_when_the_upload_raises(self, tmp_path):
        parameters = {"mode": "file", "account": {"account_type": "private_onedrive"}, "destination": {}}
        comp = _build_run_component(tmp_path, parameters, files={"a.txt": b"aaa"})

        with (
            mock.patch(
                "component.RefreshTokenProvider", return_value=_fake_token_provider("new-refresh-token-2")
            ),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            mock.patch("component.resolve_drive_id", return_value="drive-1"),
            mock.patch("component.ensure_folder", return_value="root-id"),
            mock.patch(
                "component.upload_file", side_effect=GraphPermissionError("no access", status_code=403)
            ),pytest.raises(UserException)
        ):
            comp.run()

        state_path = tmp_path / "data" / "out" / "state.json"
        assert state_path.exists()
        state = json.loads(state_path.read_text())
        payload = json.loads(state["#refreshed_auth_data"])
        assert payload["refresh_token"] == "new-refresh-token-2"

    def test_no_state_written_when_nothing_rotated(self, tmp_path):
        parameters = {"mode": "file", "account": {"account_type": "private_onedrive"}, "destination": {}}
        comp = _build_run_component(tmp_path, parameters, files={"a.txt": b"aaa"})

        with (
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider(None)),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            mock.patch("component.resolve_drive_id", return_value="drive-1"),
            mock.patch("component.ensure_folder", return_value="root-id"),
            mock.patch("component.upload_file", return_value={"id": "item-1"}),
        ):
            comp.run()

        assert not (tmp_path / "data" / "out" / "state.json").exists()


def _fake_session_context_manager(session_id: str | None) -> MagicMock:
    """A `workbook_session`-shaped context manager double yielding `session_id`."""
    context_manager = MagicMock()
    context_manager.__enter__.return_value = session_id
    context_manager.__exit__.return_value = False
    return context_manager


class TestExcelMode:
    """Excel mode wiring (plan Task 8): personal-account gate, empty CSV, happy-path plumbing.

    `resolve_workbook`/`workbook_session`/`resolve_worksheet`/`write_table` themselves are
    exercised at the unit level in ``tests/unit/test_excel_writer.py`` (plan Task 7); these tests only
    assert `Component._run_excel_mode` wires them together correctly.
    """

    def test_private_onedrive_account_raises_user_exception(self, tmp_path):
        parameters = {
            "mode": "table_excel",
            "account": {"account_type": "private_onedrive"},
            "workbook": {"path": "/book.xlsx"},
            "worksheet": {"name": "Sheet1"},
        }
        comp = _build_run_component(tmp_path, parameters, tables={"mytable": "id,name\n1,a\n"})

        with (
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider()),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            mock.patch("component.resolve_drive_id") as mock_resolve_drive_id,
            mock.patch("component.resolve_workbook") as mock_resolve_workbook,
            pytest.raises(UserException, match="private_onedrive"),
        ):
            comp.run()

        mock_resolve_workbook.assert_not_called()
        # IMPORTANT-2 (phase 8 audit): Excel mode never resolves a drive id at all.
        mock_resolve_drive_id.assert_not_called()

    def test_empty_csv_logs_v1_parity_warning_and_exits_cleanly(self, tmp_path, caplog):
        parameters = {
            "mode": "table_excel",
            "account": {"account_type": "onedrive_for_business", "tenant_id": "tenant-1"},
            "workbook": {"path": "/book.xlsx"},
            "worksheet": {"name": "Sheet1"},
        }
        comp = _build_run_component(tmp_path, parameters, tables={"empty": ""})

        with (
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider()),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            mock.patch("component.resolve_drive_id") as mock_resolve_drive_id,
            mock.patch("component.resolve_workbook", return_value=("wb-drive", "wb-file", False)),
            mock.patch("component.workbook_session", return_value=_fake_session_context_manager("session-1")),
            mock.patch("component.resolve_worksheet", return_value=("sheet-1", False, "Sheet1")),
            mock.patch("component.write_table", return_value=False) as mock_write,
            caplog.at_level("WARNING"),
        ):
            comp.run()  # must not raise — v1 parity: exit 0, sheet untouched.

        mock_write.assert_called_once()
        assert 'Ignored empty CSV file "empty".' in caplog.text
        mock_resolve_drive_id.assert_not_called()

    def test_happy_path_passes_configured_append_and_batch_size_to_write_table(self, tmp_path):
        parameters = {
            "mode": "table_excel",
            "account": {"account_type": "onedrive_for_business", "tenant_id": "tenant-1"},
            "workbook": {"path": "/book.xlsx"},
            "worksheet": {"name": "Sheet1"},
            "append": True,
            "batch_size": 1234,
        }
        comp = _build_run_component(tmp_path, parameters, tables={"mytable": "id,name\n1,a\n"})

        with (
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider()),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            mock.patch("component.resolve_drive_id") as mock_resolve_drive_id,
            mock.patch(
                "component.resolve_workbook", return_value=("wb-drive", "wb-file", False)
            ) as mock_resolve_workbook,
            mock.patch(
                "component.workbook_session", return_value=_fake_session_context_manager("session-1")
            ) as mock_session,
            mock.patch(
                "component.resolve_worksheet", return_value=("sheet-1", False, "Sheet1")
            ) as mock_resolve_worksheet,
            mock.patch("component.write_table", return_value=True) as mock_write,
        ):
            comp.run()

        mock_resolve_drive_id.assert_not_called()
        mock_resolve_workbook.assert_called_once()
        assert mock_resolve_workbook.call_args.args[1].account_type.value == "onedrive_for_business"
        mock_session.assert_called_once_with(mock.ANY, "wb-drive", "wb-file")
        assert mock_resolve_worksheet.call_args.args[1:3] == ("wb-drive", "wb-file")
        assert mock_resolve_worksheet.call_args.args[4] == "session-1"

        mock_write.assert_called_once()
        write_args = mock_write.call_args
        assert write_args.args[1:4] == ("wb-drive", "wb-file", "sheet-1")
        assert write_args.kwargs["write_mode"] == WriteMode.APPEND
        assert write_args.kwargs["key_columns"] == []
        assert write_args.kwargs["batch_size"] == 1234
        assert write_args.kwargs["is_new_sheet"] is False
        assert write_args.kwargs["session"] == "session-1"

    def test_upsert_write_mode_and_key_columns_are_passed_to_write_table(self, tmp_path):
        parameters = {
            "mode": "worksheet",
            "account": {"account_type": "onedrive_for_business", "tenant_id": "tenant-1"},
            "workbook": {"path": "/book.xlsx"},
            "worksheet": {"name": "Sheet1"},
            "write_mode": "upsert",
            "key_columns": ["id"],
        }
        comp = _build_run_component(tmp_path, parameters, tables={"mytable": "id,name\n1,a\n"})

        with (
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider()),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            mock.patch("component.resolve_drive_id"),
            mock.patch("component.resolve_workbook", return_value=("wb-drive", "wb-file", False)),
            mock.patch("component.workbook_session", return_value=_fake_session_context_manager("session-1")),
            mock.patch("component.resolve_worksheet", return_value=("sheet-1", False, "Sheet1")),
            mock.patch("component.write_table", return_value=True) as mock_write,
        ):
            comp.run()

        mock_write.assert_called_once()
        assert mock_write.call_args.kwargs["write_mode"] == WriteMode.UPSERT
        assert mock_write.call_args.kwargs["key_columns"] == ["id"]

    def test_warns_when_file_input_mapping_is_also_present(self, tmp_path, caplog):
        parameters = {
            "mode": "table_excel",
            "account": {"account_type": "onedrive_for_business", "tenant_id": "tenant-1"},
            "workbook": {"path": "/book.xlsx"},
            "worksheet": {"name": "Sheet1"},
        }
        comp = _build_run_component(
            tmp_path, parameters, tables={"mytable": "id,name\n1,a\n"}, files={"ignored.txt": b"x"}
        )

        with (
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider()),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            mock.patch("component.resolve_drive_id") as mock_resolve_drive_id,
            mock.patch("component.resolve_workbook", return_value=("wb-drive", "wb-file", False)),
            mock.patch("component.workbook_session", return_value=_fake_session_context_manager("session-1")),
            mock.patch("component.resolve_worksheet", return_value=("sheet-1", False, "Sheet1")),
            mock.patch("component.write_table", return_value=True),
            caplog.at_level("WARNING"),
        ):
            comp.run()

        assert any("file input mapping is ignored" in message for message in caplog.messages)
        mock_resolve_drive_id.assert_not_called()


class TestErrorMapping:
    def test_graph_permission_error_maps_to_user_exception(self, tmp_path):
        parameters = {"mode": "file", "account": {"account_type": "private_onedrive"}, "destination": {}}
        comp = _build_run_component(tmp_path, parameters, files={"a.txt": b"aaa"})

        with (
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider()),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            mock.patch("component.resolve_drive_id", return_value="drive-1"),
            mock.patch("component.ensure_folder", return_value="root-id"),
            mock.patch(
                "component.upload_file", side_effect=GraphPermissionError("no access", status_code=403)
            ),pytest.raises(UserException, match="no access")
        ):
            comp.run()

    def test_graph_connection_error_maps_to_user_exception(self, tmp_path):
        """IMPORTANT-1: a network failure that exhausts the client's retry budget surfaces as
        `GraphConnectionError`, which must map to a `UserException` (exit 1), not propagate as an
        unhandled exception (exit 2)."""
        parameters = {"mode": "file", "account": {"account_type": "private_onedrive"}, "destination": {}}
        comp = _build_run_component(tmp_path, parameters, files={"a.txt": b"aaa"})

        with (
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider()),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            mock.patch("component.resolve_drive_id", return_value="drive-1"),
            mock.patch("component.ensure_folder", return_value="root-id"),
            mock.patch(
                "component.upload_file",
                side_effect=GraphConnectionError("network error, retry budget exhausted"),
            ),
            pytest.raises(UserException, match="retry budget exhausted"),
        ):
            comp.run()

    def test_invalid_configuration_raises_user_exception(self, tmp_path):
        # mode=table_excel requires workbook/worksheet — neither is provided.
        parameters = {"mode": "table_excel", "account": {"account_type": "private_onedrive"}}
        comp = _build_run_component(tmp_path, parameters)

        with pytest.raises(UserException, match="Invalid configuration"):
            comp.run()


# ----------------------------------------------------------------------------------------------
# 3. End-to-end via a mocked HTTP transport boundary (tests/e2e/test_functional_http.py)
# ----------------------------------------------------------------------------------------------


BASE_URL = "https://graph.microsoft.com/v1.0"
_TOKEN_URL_RE = re.compile(r"^https://login\.microsoftonline\.com/[^/]+/oauth2/v2\.0/token$")

_SRC_DIR = Path(__file__).resolve().parents[1] / "src"
_COMPONENT_SCRIPT = _SRC_DIR / "component.py"


def _graph_url(path: str) -> str:
    return f"{BASE_URL}{path}"


@dataclass
class FakeResponse:
    """A minimal stand-in for ``requests.Response`` — just enough for ``GraphClient``."""

    status_code: int
    payload: object = None
    headers: dict = field(default_factory=dict)

    @property
    def ok(self) -> bool:
        return 200 <= self.status_code < 300

    def json(self):
        return self.payload

    @property
    def text(self) -> str:
        return json.dumps(self.payload) if self.payload is not None else ""


class GraphFake:
    """Fake HTTP transport routed through a patched ``requests.Session.request``.

    Routes are registered with :meth:`add` (``matcher`` is an exact URL string or a compiled
    regex, matched with ``re.fullmatch``); the Microsoft identity token endpoint is handled
    automatically — every scenario gets a working, rotating refresh token without registering it
    explicitly. Every call (including the token exchange) is recorded in :attr:`calls` for
    assertions on exactly which Graph calls a scenario made.
    """

    def __init__(self, rotated_refresh_token: str = "rotated-refresh-token"):
        self._rules: list[tuple[str, object, object]] = []
        self.calls: list[tuple[str, str, dict]] = []
        self.rotated_refresh_token = rotated_refresh_token

    def add(self, method: str, matcher, response) -> GraphFake:
        self._rules.append((method, matcher, response))
        return self

    def __call__(self, method: str, url: str, **kwargs) -> FakeResponse:
        self.calls.append((method, url, kwargs))
        if method == "POST" and _TOKEN_URL_RE.fullmatch(url):
            return FakeResponse(
                200,
                {
                    "access_token": "access-1",
                    "refresh_token": self.rotated_refresh_token,
                    "expires_in": 3599,
                },
            )
        for rule_method, matcher, response in self._rules:
            if rule_method != method:
                continue
            matched = matcher.fullmatch(url) if isinstance(matcher, re.Pattern) else matcher == url
            if not matched:
                continue
            return response(**kwargs) if callable(response) else response
        raise AssertionError(f"GraphFake: no rule registered for {method} {url} (kwargs={kwargs!r})")

    def calls_for(self, method: str) -> list[tuple[str, dict]]:
        return [(url, kwargs) for called_method, url, kwargs in self.calls if called_method == method]


_HTTP_UNSET = object()


def _build_data_dir(
    tmp_path,
    parameters: dict,
    *,
    files: dict[str, bytes] | None = None,
    tables: dict[str, str] | None = None,
    oauth=_HTTP_UNSET,
) -> Path:
    """Build a ``KBC_DATADIR``-style fixture directory (datadir idiom — merged ``config.json``,
    row-scoped state), returning the ``data`` directory path."""
    data_dir = tmp_path / "data"
    files_dir = data_dir / "in" / "files"
    tables_dir = data_dir / "in" / "tables"
    files_dir.mkdir(parents=True, exist_ok=True)
    tables_dir.mkdir(parents=True, exist_ok=True)
    (data_dir / "out").mkdir(parents=True, exist_ok=True)

    for name, content in (files or {}).items():
        (files_dir / name).write_bytes(content)
    for name, content in (tables or {}).items():
        (tables_dir / name).write_text(content)
        (tables_dir / f"{name}.manifest").write_text(json.dumps({"id": f"in.c-main.{name}"}))

    config: dict = {"parameters": parameters, "action": "run"}
    resolved_oauth = _oauth_credentials() if oauth is _HTTP_UNSET else oauth
    if resolved_oauth is not None:
        config["authorization"] = {"oauth_api": {"credentials": resolved_oauth}}
    (data_dir / "config.json").write_text(json.dumps(config))
    return data_dir


def _build_http_component(tmp_path, parameters: dict, **kwargs) -> Component:
    data_dir = _build_data_dir(tmp_path, parameters, **kwargs)
    with mock.patch.dict(os.environ, {"KBC_DATADIR": str(data_dir)}):
        return Component()


def _run(comp: Component, graph: GraphFake) -> None:
    """Run the real ``Component.run()`` with HTTP routed through ``graph`` (plan Task 10)."""
    with mock.patch.object(requests.Session, "request", side_effect=graph):
        comp.run()


def _xlsx_item(file_id: str, name: str) -> dict:
    return {"id": file_id, "name": name, "file": {"mimeType": XLSX_MIME_TYPE}, "parentReference": {}}


def _not_found(code: str = "itemNotFound", message: str = "The resource could not be found.") -> FakeResponse:
    return FakeResponse(404, {"error": {"code": code, "message": message}})


# ---------------------------------------------------------------------------------------------
# 1. Happy path per mode
# ---------------------------------------------------------------------------------------------


class TestHappyPathFileMode:
    def test_uploads_two_files_via_simple_put_and_persists_rotated_token(self, tmp_path):
        parameters = {
            "mode": "file",
            "account": {"account_type": "private_onedrive"},
            "destination": {"folder_path": "uploads", "conflict_behavior": "replace"},
        }
        comp = _build_http_component(tmp_path, parameters, files={"a.txt": b"aaa", "b.txt": b"bbb"})
        graph = GraphFake()
        graph.add("GET", _graph_url("/me/drive"), FakeResponse(200, {"id": "drive-1"}))
        graph.add("GET", _graph_url("/drives/drive-1/root:/uploads"), _not_found())
        graph.add("POST", _graph_url("/drives/drive-1/root/children"), FakeResponse(201, {"id": "folder-1"}))
        upload_pattern = re.compile(re.escape(_graph_url("/drives/drive-1/items/folder-1:/")) + r"[ab]\.txt:/content")
        graph.add("PUT", upload_pattern, FakeResponse(200, {"id": "item-x"}))

        _run(comp, graph)  # must not raise: exit 0

        put_calls = graph.calls_for("PUT")
        assert len(put_calls) == 2
        uploaded_names = {re.search(r":/(\w+\.txt):/content$", url).group(1) for url, _ in put_calls}
        assert uploaded_names == {"a.txt", "b.txt"}
        for _url, kwargs in put_calls:
            assert kwargs["params"] == {"@microsoft.graph.conflictBehavior": "replace"}

        state = json.loads((tmp_path / "data" / "out" / "state.json").read_text())
        payload = json.loads(state["#refreshed_auth_data"])
        assert payload["refresh_token"] == "rotated-refresh-token"


class TestHappyPathCsvMode:
    def test_uploads_single_table_as_passthrough_csv(self, tmp_path):
        parameters = {
            "mode": "table_csv",
            "account": {"account_type": "private_onedrive"},
            "destination": {"conflict_behavior": "fail"},
        }
        comp = _build_http_component(tmp_path, parameters, tables={"mytable": "id,name\n1,a\n2,b\n"})
        graph = GraphFake()
        graph.add("GET", _graph_url("/me/drive"), FakeResponse(200, {"id": "drive-1"}))
        graph.add("GET", _graph_url("/drives/drive-1/root"), FakeResponse(200, {"id": "root-1"}))
        upload_url = _graph_url("/drives/drive-1/items/root-1:/mytable.csv:/content")
        captured: dict = {}

        def _capture_put(**kwargs):
            captured["body"] = kwargs["data"].read()
            return FakeResponse(200, {"id": "item-1"})

        graph.add("PUT", upload_url, _capture_put)

        _run(comp, graph)  # must not raise: exit 0

        assert captured["body"] == b"id,name\n1,a\n2,b\n"  # streamed as-is, no rewrite


class TestHappyPathExcelModeOverwrite:
    """Excel mode, ``append=False`` (overwrite = clear + write from A1), business account."""

    def test_clears_and_writes_from_a1(self, tmp_path):
        parameters = {
            "mode": "table_excel",
            "account": {"account_type": "onedrive_for_business", "tenant_id": "tenant-1"},
            "workbook": {"path": "/book.xlsx"},
            "worksheet": {"name": "Sheet1"},
            "append": False,
        }
        comp = _build_http_component(tmp_path, parameters, tables={"mytable": "id,name\n1,a\n"})
        graph = _excel_graph_with_existing_sheet()

        _run(comp, graph)  # must not raise: exit 0

        clear_calls = graph.calls_for("POST")
        clear_urls = [url for url, _ in clear_calls if url.endswith("/range/clear")]
        assert len(clear_urls) == 1

        patch_calls = graph.calls_for("PATCH")
        assert len(patch_calls) == 1
        patch_url, patch_kwargs = patch_calls[0]
        assert patch_url.endswith("/range(address='A1:B2')")
        assert patch_kwargs["json"] == {"values": [["id", "name"], ["1", "a"]]}

        # session opened and closed exactly once around the write.
        session_urls = [url for url, _ in graph.calls_for("POST") if "workbook/createSession" in url]
        close_urls = [url for url, _ in graph.calls_for("POST") if "workbook/closeSession" in url]
        assert len(session_urls) == 1
        assert len(close_urls) == 1


def _excel_graph_with_existing_sheet(existing_header: list[str] | None = None) -> GraphFake:
    """Common Excel-mode routing: an existing workbook `/book.xlsx` with an existing `Sheet1`.

    ``existing_header`` (when given) also wires up `usedRange`/header-row responses for append
    scenarios; overwrite scenarios never call those endpoints (range/clear is unconditional).
    """
    graph = GraphFake()
    graph.add("GET", _graph_url("/me/drive"), FakeResponse(200, {"id": "drive-1"}))
    graph.add("GET", _graph_url("/drives/drive-1/root:/book.xlsx"), FakeResponse(200, _xlsx_item("file-1", "book.xlsx")))
    graph.add(
        "POST",
        _graph_url("/drives/drive-1/items/file-1/workbook/createSession"),
        FakeResponse(201, {"id": "session-1"}),
    )
    graph.add(
        "GET",
        _graph_url("/drives/drive-1/items/file-1/workbook/worksheets"),
        FakeResponse(200, {"value": [{"id": "sheet-1", "name": "Sheet1", "position": 0, "visibility": "Visible"}]}),
    )
    graph.add(
        "POST",
        _graph_url("/drives/drive-1/items/file-1/workbook/worksheets/sheet-1/range/clear"),
        FakeResponse(200, {}),
    )
    graph.add(
        "POST",
        _graph_url("/drives/drive-1/items/file-1/workbook/closeSession"),
        FakeResponse(204, {}),
    )
    if existing_header is not None:
        graph.add(
            "GET",
            _graph_url("/drives/drive-1/items/file-1/workbook/worksheets/sheet-1/range/usedRange(valuesOnly=true)"),
            FakeResponse(200, {"address": "Sheet1!A1:B3"}),
        )
        graph.add(
            "GET",
            _graph_url(
                "/drives/drive-1/items/file-1/workbook/worksheets/sheet-1/range/usedRange(valuesOnly=true)/row(row=0)"
            ),
            FakeResponse(200, {"text": [existing_header]}),
        )
    _wire_patch_catch_all(graph)
    return graph


def _wire_patch_catch_all(graph: GraphFake) -> None:
    """Any ``range(address=...)`` PATCH succeeds — the address itself varies per scenario and is
    asserted from ``graph.calls_for("PATCH")`` rather than pre-registered per exact address."""
    pattern = re.compile(re.escape(_graph_url("/drives/drive-1/items/file-1/workbook/worksheets/sheet-1/range(address='")) + r".+'\)")
    graph.add("PATCH", pattern, FakeResponse(200, {}))


# ---------------------------------------------------------------------------------------------
# 2. Excel append vs overwrite
# ---------------------------------------------------------------------------------------------


class TestExcelAppendOffsetsFromUsedRange:
    def test_append_skips_header_and_starts_after_used_range(self, tmp_path):
        parameters = {
            "mode": "table_excel",
            "account": {"account_type": "onedrive_for_business", "tenant_id": "tenant-1"},
            "workbook": {"path": "/book.xlsx"},
            "worksheet": {"name": "Sheet1"},
            "append": True,
        }
        comp = _build_http_component(tmp_path, parameters, tables={"mytable": "id,name\n1,a\n"})
        graph = _excel_graph_with_existing_sheet(existing_header=["id", "name"])

        _run(comp, graph)  # must not raise: exit 0

        clear_urls = [url for url, _ in graph.calls_for("POST") if url.endswith("/range/clear")]
        assert clear_urls == []  # append never clears

        patch_calls = graph.calls_for("PATCH")
        assert len(patch_calls) == 1
        patch_url, patch_kwargs = patch_calls[0]
        assert patch_url.endswith("/range(address='A4:B4')")  # offset from usedRange's A1:B3
        assert patch_kwargs["json"] == {"values": [["1", "a"]]}  # header skipped (headers match)

    def test_append_warns_on_header_mismatch_but_still_appends(self, tmp_path, caplog):
        parameters = {
            "mode": "table_excel",
            "account": {"account_type": "onedrive_for_business", "tenant_id": "tenant-1"},
            "workbook": {"path": "/book.xlsx"},
            "worksheet": {"name": "Sheet1"},
            "append": True,
        }
        comp = _build_http_component(tmp_path, parameters, tables={"mytable": "id,name\n1,a\n"})
        graph = _excel_graph_with_existing_sheet(existing_header=["identifier", "full_name"])

        with caplog.at_level("WARNING"):
            _run(comp, graph)  # must not raise: exit 0

        assert "Headers mismatch" in caplog.text
        patch_calls = graph.calls_for("PATCH")
        assert len(patch_calls) == 1
        assert patch_calls[0][1]["json"] == {"values": [["1", "a"]]}


# ---------------------------------------------------------------------------------------------
# 3. Empty CSV in Excel mode
# ---------------------------------------------------------------------------------------------


class TestExcelEmptyCsvInput:
    def test_empty_csv_logs_warning_exits_cleanly_and_issues_no_range_patch(self, tmp_path, caplog):
        parameters = {
            "mode": "table_excel",
            "account": {"account_type": "onedrive_for_business", "tenant_id": "tenant-1"},
            "workbook": {"path": "/book.xlsx"},
            "worksheet": {"name": "Sheet1"},
        }
        comp = _build_http_component(tmp_path, parameters, tables={"empty": ""})
        graph = _excel_graph_with_existing_sheet()

        with caplog.at_level("WARNING"):
            _run(comp, graph)  # must not raise: exit 0, sheet untouched (v1 parity)

        assert 'Ignored empty CSV file "empty".' in caplog.text
        assert graph.calls_for("PATCH") == []
        clear_urls = [url for url, _ in graph.calls_for("POST") if url.endswith("/range/clear")]
        assert clear_urls == []  # `write_table` returns before touching the sheet at all


# ---------------------------------------------------------------------------------------------
# 4. CSV cardinality
# ---------------------------------------------------------------------------------------------


class TestCsvCardinality:
    """Mode 'worksheet' (the pre-merge 'table_excel') still requires exactly one input table —
    unaffected by Change A's file-mode merge, which only touches mode 'file'."""

    def test_zero_tables_raises_v1_parity_user_exception(self, tmp_path):
        parameters = {
            "mode": "table_excel",
            "account": {"account_type": "onedrive_for_business", "tenant_id": "tenant-1"},
            "workbook": {"path": "/book.xlsx"},
            "worksheet": {"name": "Sheet1"},
        }
        comp = _build_http_component(tmp_path, parameters, tables={})
        graph = GraphFake()

        with pytest.raises(UserException, match=re.escape('No CSV file found in "/data/in/tables".')):
            _run(comp, graph)  # exit 1
        assert graph.calls_for("GET") == []  # fails before any Graph call (no drive_id resolution either)

    def test_multiple_tables_raises_v1_parity_user_exception_naming_both(self, tmp_path):
        parameters = {
            "mode": "table_excel",
            "account": {"account_type": "onedrive_for_business", "tenant_id": "tenant-1"},
            "workbook": {"path": "/book.xlsx"},
            "worksheet": {"name": "Sheet1"},
        }
        comp = _build_http_component(tmp_path, parameters, tables={"a.csv": "id\n1\n", "b.csv": "id\n2\n"})
        graph = GraphFake()
        expected = re.escape('Expected one CSV file, found multiple: "a.csv", "b.csv".')

        with pytest.raises(UserException, match=expected):
            _run(comp, graph)  # exit 1


class TestFileModeMergedTableInputHttp:
    """Change A: mode 'file' now processes both input mappings end-to-end (files uploaded as-is,
    mapped tables written as CSV) — a real HTTP-mocked (not just mocked-collaborator) proof that
    both actually reach Graph in the same run."""

    def test_files_and_tables_together_are_both_uploaded(self, tmp_path):
        parameters = {"mode": "file", "account": {"account_type": "private_onedrive"}, "destination": {}}
        comp = _build_http_component(
            tmp_path, parameters, files={"a.txt": b"aaa"}, tables={"mytable": "id,name\n1,a\n"}
        )
        graph = GraphFake()
        graph.add("GET", _graph_url("/me/drive"), FakeResponse(200, {"id": "drive-1"}))
        graph.add("GET", _graph_url("/drives/drive-1/root"), FakeResponse(200, {"id": "root-1"}))
        graph.add("PUT", _graph_url("/drives/drive-1/items/root-1:/a.txt:/content"), FakeResponse(200, {"id": "item-1"}))
        graph.add(
            "PUT", _graph_url("/drives/drive-1/items/root-1:/mytable.csv:/content"), FakeResponse(200, {"id": "item-2"})
        )

        _run(comp, graph)  # must not raise: exit 0

        put_urls = {url for url, _ in graph.calls_for("PUT")}
        assert put_urls == {
            _graph_url("/drives/drive-1/items/root-1:/a.txt:/content"),
            _graph_url("/drives/drive-1/items/root-1:/mytable.csv:/content"),
        }


# ---------------------------------------------------------------------------------------------
# 5. Missing OAuth
# ---------------------------------------------------------------------------------------------


class TestMissingOAuth:
    def test_missing_oauth_authorization_raises_user_exception_before_any_network_call(self, tmp_path):
        parameters = {"mode": "file", "account": {"account_type": "private_onedrive"}, "destination": {}}
        comp = _build_http_component(tmp_path, parameters, files={"a.txt": b"aaa"}, oauth=None)
        graph = GraphFake()

        with pytest.raises(UserException, match="not authorized"):
            _run(comp, graph)  # exit 1

        assert graph.calls == []  # fails before the token provider (and thus any HTTP call) exists
        assert not (tmp_path / "data" / "out" / "state.json").exists()


# ---------------------------------------------------------------------------------------------
# 6. Excel mode gated off for private_onedrive
# ---------------------------------------------------------------------------------------------


class TestExcelPrivateOnedriveGate:
    def test_private_onedrive_account_raises_user_exception(self, tmp_path):
        parameters = {
            "mode": "table_excel",
            "account": {"account_type": "private_onedrive"},
            "workbook": {"path": "/book.xlsx"},
            "worksheet": {"name": "Sheet1"},
        }
        comp = _build_http_component(tmp_path, parameters, tables={"mytable": "id,name\n1,a\n"})
        graph = GraphFake()
        graph.add("GET", _graph_url("/me/drive"), FakeResponse(200, {"id": "drive-1"}))

        with pytest.raises(UserException, match="private_onedrive"):
            _run(comp, graph)  # exit 1

        # IMPORTANT-2 (phase 8 audit): Excel mode never resolves a drive id at all (it targets
        # `workbook.{path,drive_id,file_id}` instead) — the account-type gate must raise before
        # any Graph call, not just before an Excel-specific one.
        assert graph.calls_for("GET") == []


# ---------------------------------------------------------------------------------------------
# 7. Conflict behavior `fail` on an existing file
# ---------------------------------------------------------------------------------------------


class TestConflictFailOnExistingFile:
    def test_existing_file_with_conflict_fail_raises_user_exception_naming_the_file(self, tmp_path):
        parameters = {
            "mode": "file",
            "account": {"account_type": "private_onedrive"},
            "destination": {},  # conflict_behavior default is "fail"
        }
        comp = _build_http_component(tmp_path, parameters, files={"report.csv": b"a,b\n1,2\n"})
        graph = GraphFake()
        graph.add("GET", _graph_url("/me/drive"), FakeResponse(200, {"id": "drive-1"}))
        graph.add("GET", _graph_url("/drives/drive-1/root"), FakeResponse(200, {"id": "root-1"}))
        graph.add(
            "PUT",
            _graph_url("/drives/drive-1/items/root-1:/report.csv:/content"),
            FakeResponse(409, {"error": {"code": "nameAlreadyExists", "message": "An item with the same name already exists."}}),
        )

        with pytest.raises(UserException, match=re.escape("'report.csv' already exists")):
            _run(comp, graph)  # exit 1


# ---------------------------------------------------------------------------------------------
# 8. Rotated refresh token persisted to state, including on a mid-run failure
# ---------------------------------------------------------------------------------------------


class TestTokenRotationPersistence:
    def test_rotated_token_persisted_even_when_a_later_upload_fails(self, tmp_path):
        parameters = {
            "mode": "file",
            "account": {"account_type": "private_onedrive"},
            "destination": {"conflict_behavior": "replace"},
        }
        comp = _build_http_component(tmp_path, parameters, files={"a.txt": b"aaa", "b.txt": b"bbb"})
        graph = GraphFake(rotated_refresh_token="rotated-after-failure")
        graph.add("GET", _graph_url("/me/drive"), FakeResponse(200, {"id": "drive-1"}))
        graph.add("GET", _graph_url("/drives/drive-1/root"), FakeResponse(200, {"id": "root-1"}))
        graph.add(
            "PUT",
            _graph_url("/drives/drive-1/items/root-1:/a.txt:/content"),
            FakeResponse(200, {"id": "item-a"}),
        )
        graph.add(
            "PUT",
            _graph_url("/drives/drive-1/items/root-1:/b.txt:/content"),
            FakeResponse(403, {"error": {"code": "accessDenied", "message": "Access denied."}}),
        )

        with pytest.raises(UserException):
            _run(comp, graph)  # exit 1 (mid-run failure)

        state = json.loads((tmp_path / "data" / "out" / "state.json").read_text())
        payload = json.loads(state["#refreshed_auth_data"])
        assert payload["refresh_token"] == "rotated-after-failure"


# ---------------------------------------------------------------------------------------------
# Literal process exit-code mapping (subprocess — deliberately independent of the in-process
# `pytest.raises(UserException)` idiom used above; see the module docstring for why a full
# `keboola.datadirtest`-style run of every scenario is not used instead).
# ---------------------------------------------------------------------------------------------


class TestEntrypointExitCodeMapping:
    def test_missing_oauth_exits_with_code_1(self, tmp_path):
        parameters = {"mode": "file", "account": {"account_type": "private_onedrive"}, "destination": {}}
        data_dir = _build_data_dir(tmp_path, parameters, files={"a.txt": b"aaa"}, oauth=None)

        result = subprocess.run(
            [sys.executable, str(_COMPONENT_SCRIPT)],
            env={**os.environ, "KBC_DATADIR": str(data_dir), "PYTHONPATH": str(_SRC_DIR)},
            capture_output=True,
            text=True,
            timeout=30,
            check=False,
        )

        assert result.returncode == 1, f"stdout={result.stdout!r} stderr={result.stderr!r}"
        assert "not authorized" in result.stderr
