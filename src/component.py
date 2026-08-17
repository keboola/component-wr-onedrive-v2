"""
Component main class for keboola.wr-onedrive-v2.

Design spec: ``docs/superpowers/specs/2026-08-17-wr-onedrive-v2-design.md`` §3 (auth) and §5
(configuration & sync actions).
"""

import json
import logging
import sys

from keboola.component.base import ComponentBase, sync_action
from keboola.component.exceptions import UserException
from keboola.component.sync_actions import SelectElement
from pydantic import ValidationError

from client.auth import AuthenticationError, RefreshTokenProvider, TokenProvider
from client.drives import get_site_id, list_drives
from client.exceptions import GraphClientError
from client.graph_client import GraphClient
from configuration import Account, AccountType, RowConfig

logger = logging.getLogger(__name__)

# v1-compatible state key (design spec §2/§3) — holds the JSON-encoded payload produced by
# `RefreshTokenProvider.rotated_refresh_token` (only `refresh_token` is actually read back).
STATE_KEY_REFRESHED_AUTH_DATA = "#refreshed_auth_data"


class Component(ComponentBase):
    """
    Extends base class for general Python components. Initializes the CommonInterface
    and performs configuration validation.

    For easier debugging the data folder is picked up by default from `../data` path,
    relative to working directory.
    """

    def __init__(self):
        super().__init__()

    def run(self):
        """
        Main execution code.

        NOTE: this is scaffolding only — the real orchestration (mode dispatch, client wiring,
        token rotation persistence) lands in a later implementation task (plan Task 6).
        """
        params = self._load_configuration()
        logger.info("Loaded configuration for mode: %s", params.mode)

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

    def _load_configuration(self) -> RowConfig:
        try:
            return RowConfig.model_validate(self.configuration.parameters)
        except ValidationError as e:
            raise UserException(f"Invalid configuration: {_format_validation_error(e)}") from e

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
