from unittest.mock import MagicMock

import pytest

from client.drives import get_site_id, list_drives, resolve_drive_id
from client.exceptions import GraphNotFoundError
from configuration import Account, AccountType


def _client_with_get(return_value=None, side_effect=None) -> MagicMock:
    client = MagicMock()
    if side_effect is not None:
        client.get.side_effect = side_effect
    else:
        response = MagicMock()
        response.json.return_value = return_value
        client.get.return_value = response
    return client


class TestGetSiteId:
    def test_builds_hostname_and_server_relative_path_lookup_url(self):
        client = _client_with_get(return_value={"id": "contoso.sharepoint.com,GUID1,GUID2"})

        site_id = get_site_id(client, "https://contoso.sharepoint.com/sites/marketing")

        client.get.assert_called_once_with("/sites/contoso.sharepoint.com:/sites/marketing")
        assert site_id == "contoso.sharepoint.com,GUID1,GUID2"

    def test_root_site_url_omits_the_colon_path_suffix(self):
        client = _client_with_get(return_value={"id": "contoso.sharepoint.com,GUID1,GUID2"})

        get_site_id(client, "https://contoso.sharepoint.com")

        client.get.assert_called_once_with("/sites/contoso.sharepoint.com")

    def test_trailing_slash_in_path_is_normalized(self):
        client = _client_with_get(return_value={"id": "site-id"})

        get_site_id(client, "https://contoso.sharepoint.com/sites/marketing/")

        client.get.assert_called_once_with("/sites/contoso.sharepoint.com:/sites/marketing")

    def test_missing_hostname_raises_without_any_network_call(self):
        client = MagicMock()

        with pytest.raises(GraphNotFoundError):
            get_site_id(client, "not-a-url")

        client.get.assert_not_called()

    def test_404_is_wrapped_with_a_helpful_message(self):
        client = _client_with_get(
            side_effect=GraphNotFoundError("Item not found", status_code=404, error_code="itemNotFound")
        )

        with pytest.raises(GraphNotFoundError) as exc_info:
            get_site_id(client, "https://contoso.sharepoint.com/sites/marketing")

        assert "https://contoso.sharepoint.com/sites/marketing" in str(exc_info.value)
        assert exc_info.value.status_code == 404
        assert exc_info.value.error_code == "itemNotFound"


class TestListDrives:
    def test_delegates_to_the_paging_helper_and_returns_a_list(self):
        client = MagicMock()
        client.get_paged.return_value = iter(
            [
                {"id": "drive-1", "name": "Documents", "webUrl": "https://x/Documents"},
                {"id": "drive-2", "name": "Marketing Assets", "webUrl": "https://x/Marketing"},
            ]
        )

        drives = list_drives(client, "site-1")

        client.get_paged.assert_called_once_with("/sites/site-1/drives")
        assert drives == [
            {"id": "drive-1", "name": "Documents", "webUrl": "https://x/Documents"},
            {"id": "drive-2", "name": "Marketing Assets", "webUrl": "https://x/Marketing"},
        ]


class TestResolveDriveId:
    @pytest.mark.parametrize("account_type", [AccountType.PRIVATE_ONEDRIVE, AccountType.ONEDRIVE_FOR_BUSINESS])
    def test_personal_and_business_accounts_without_configured_drive_id_resolve_via_me_drive(self, account_type):
        client = _client_with_get(return_value={"id": "my-drive-id"})
        account = Account(account_type=account_type, tenant_id="tenant-1")

        drive_id = resolve_drive_id(client, account, destination_drive_id=None)

        client.get.assert_called_once_with("/me/drive")
        assert drive_id == "my-drive-id"

    @pytest.mark.parametrize("account_type", [AccountType.PRIVATE_ONEDRIVE, AccountType.ONEDRIVE_FOR_BUSINESS])
    def test_personal_and_business_accounts_with_configured_drive_id_use_it_verbatim(self, account_type):
        # Drives are globally addressable in Graph — a configured `destination.drive_id` is now
        # honored for every account type, not just `sharepoint`, and never triggers a `/me/drive`
        # lookup.
        client = MagicMock()
        account = Account(account_type=account_type, tenant_id="tenant-1")

        drive_id = resolve_drive_id(client, account, destination_drive_id="drive-configured")

        assert drive_id == "drive-configured"
        client.get.assert_not_called()

    def test_sharepoint_with_configured_drive_id_makes_no_network_call(self):
        client = MagicMock()
        account = Account(
            account_type=AccountType.SHAREPOINT,
            tenant_id="tenant-1",
            site_url="https://contoso.sharepoint.com/sites/marketing",
        )

        drive_id = resolve_drive_id(client, account, destination_drive_id="drive-configured")

        assert drive_id == "drive-configured"
        client.get.assert_not_called()
        client.get_paged.assert_not_called()

    def test_sharepoint_without_configured_drive_id_resolves_site_default_drive(self):
        client = MagicMock()
        site_response = MagicMock()
        site_response.json.return_value = {"id": "site-1"}
        drive_response = MagicMock()
        drive_response.json.return_value = {"id": "default-drive-id"}
        client.get.side_effect = [site_response, drive_response]
        account = Account(
            account_type=AccountType.SHAREPOINT,
            tenant_id="tenant-1",
            site_url="https://contoso.sharepoint.com/sites/marketing",
        )

        drive_id = resolve_drive_id(client, account, destination_drive_id=None)

        assert drive_id == "default-drive-id"
        assert client.get.call_args_list[0].args[0] == "/sites/contoso.sharepoint.com:/sites/marketing"
        assert client.get.call_args_list[1].args[0] == "/sites/site-1/drive"
