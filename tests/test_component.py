import json
import os
import re
import unittest
from unittest import mock
from unittest.mock import MagicMock

import pytest
from freezegun import freeze_time
from keboola.component.exceptions import UserException
from keboola.component.sync_actions import SelectElement

from client.exceptions import GraphPermissionError
from component import Component


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


def _oauth_credentials(refresh_token: str = "refresh-config") -> dict:
    return {
        "id": "oauth-1",
        "created": "2026-01-01",
        "appKey": "client-1",
        "#appSecret": "secret-1",
        "oauthVersion": "2.0",
        "#data": json.dumps({"refresh_token": refresh_token}),
    }


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


def _fake_token_provider(rotated_refresh_token: str | None = None) -> MagicMock:
    """A `TokenProvider` double that never touches the network and reports a fixed rotation."""
    provider = MagicMock()
    provider.rotated_refresh_token = rotated_refresh_token
    return provider


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
    all resolve exactly as they do for a row-run (`tests/test_configuration.py`'s model-level
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
