import json
import os
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
