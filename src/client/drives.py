"""Site and drive resolution for keboola.wr-onedrive-v2's SharePoint account type.

Design spec: ``docs/superpowers/specs/2026-08-17-wr-onedrive-v2-design.md`` §3 (account types →
base drive) and §5 (``listLibraries`` sync action).

Kept as module-level functions taking a :class:`~client.graph_client.GraphClient` (rather than
methods on ``GraphClient`` itself) so ``graph_client.py`` stays transport-only, per the design
spec's code architecture (§6).

Deliberately does **not** port the sibling extractor's doubled ``/sites/{id}/sites/{id}/lists``
path — a drive's ``id`` is the canonical, stable identifier Graph uses everywhere
(``/drives/{drive-id}/...``), so every later task (uploader, Excel writer) addresses drives
uniformly through it, regardless of account type.
"""

from typing import Any
from urllib.parse import urlparse

from client.exceptions import GraphNotFoundError
from client.graph_client import GraphClient
from configuration import Account, AccountType


def get_site_id(client: GraphClient, site_url: str) -> str:
    """Resolve a SharePoint site URL to its Graph composite site id.

    Calls ``GET /sites/{hostname}:{server-relative-path}`` — Graph's documented way to address a
    site by its URL. The colon + path suffix is omitted entirely for the tenant's root site
    (empty server-relative path), which Graph only accepts as a bare ``/sites/{hostname}``.
    """
    parsed = urlparse(site_url)
    hostname = parsed.netloc
    server_relative_path = parsed.path.rstrip("/")

    if not hostname:
        raise GraphNotFoundError(
            f"'{site_url}' is not a valid SharePoint site URL: no hostname could be parsed from "
            "it. Check the account.site_url value in the configuration."
        )

    lookup_path = f"/sites/{hostname}:{server_relative_path}" if server_relative_path else f"/sites/{hostname}"

    try:
        response = client.get(lookup_path)
    except GraphNotFoundError as exc:
        raise GraphNotFoundError(
            f"SharePoint site not found for URL '{site_url}'. Check that account.site_url points "
            "to an existing site and that the authorized account has access to it.",
            status_code=exc.status_code,
            error_code=exc.error_code,
        ) from exc
    return response.json()["id"]


def list_drives(client: GraphClient, site_id: str) -> list[dict[str, Any]]:
    """List every document library (drive) on a site.

    Each item has (at least) ``id``, ``name``, and ``webUrl`` — paginated via
    :meth:`GraphClient.get_paged` (Graph pages the drives list for large sites).
    """
    return list(client.get_paged(f"/sites/{site_id}/drives"))


def resolve_drive_id(client: GraphClient, account: Account, destination_drive_id: str | None = None) -> str:
    """Resolve the concrete drive id later tasks address as ``/drives/{drive_id}/...``.

    Drives are globally addressable in Graph (``/drives/{drive-id}/...`` works regardless of
    which account "owns" them), so ``destination_drive_id`` — the row's ``listLibraries``-picked
    document library — is honored verbatim for **every** account type when it's set, not just
    ``sharepoint``. Only when it's empty does this fall back to the per-account-type default
    (design spec §3):

    - ``private_onedrive`` / ``onedrive_for_business``: the user's own default drive
      (``GET /me/drive``).
    - ``sharepoint``: the site's default document library (``GET /sites/{site_id}/drive``).

    Always returns a concrete drive id — never a bare ``/me/drive``-style path — so the uploader
    and Excel writer (later tasks) can address every account type identically via
    ``/drives/{drive_id}/...``.
    """
    if destination_drive_id:
        return destination_drive_id

    if account.account_type in (AccountType.PRIVATE_ONEDRIVE, AccountType.ONEDRIVE_FOR_BUSINESS):
        return client.get("/me/drive").json()["id"]

    site_id = get_site_id(client, account.site_url)
    return client.get(f"/sites/{site_id}/drive").json()["id"]
