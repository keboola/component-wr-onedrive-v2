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

    def test_missing_site_url_raises_user_exception_without_any_network_call(self, tmp_path):
        parameters = {"account": {"account_type": "onedrive_for_business", "tenant_id": "tenant-1"}}
        comp = _build_component(tmp_path, parameters)
        fake_client = MagicMock()

        with mock.patch("component.GraphClient", return_value=fake_client), pytest.raises(
            UserException, match="SharePoint"
        ):
            comp.list_libraries()

        fake_client.get.assert_not_called()

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

    def test_invalid_workbook_combo_raises_validation_error_not_raw_path(self, tmp_path):
        # `path` combined with `drive_id`/`file_id` is invalid per the `Workbook` model — this can
        # only be caught once actual model validation runs (raw `.get("path")` would silently
        # ignore the conflicting ids and let it through).
        parameters = {
            "account": {"account_type": "private_onedrive"},
            "workbook": {"path": "/book.xlsx", "drive_id": "drive-1", "file_id": "file-1"},
        }
        comp = _build_component(tmp_path, parameters)

        with pytest.raises(UserException, match="Invalid workbook configuration"):
            comp.create_workbook()

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
