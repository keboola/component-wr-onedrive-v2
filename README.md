OneDrive/SharePoint Writer
==========================

Writes files and tables from Keboola to Microsoft OneDrive and SharePoint document libraries via
Microsoft Graph. This component is the successor to `keboola.wr-onedrive` ("OneDrive Excel
Sheets") — it keeps full feature parity for Excel worksheet writes and adds two new output modes:
plain file upload and table-as-CSV.

**Table of Contents:**

[TOC]

Functionality Notes
====================

Each configuration row targets exactly one destination, in one of three output modes:

- **Upload Files** (`mode: file`) — every file in the row's file input mapping is uploaded as-is
  to a document library folder.
- **Table as CSV** (`mode: table_csv`) — the row's input table is written to the library as a CSV
  file.
- **Table as Excel Worksheet** (`mode: table_excel`) — the row's input table is written into a
  worksheet of an XLSX workbook, creating the workbook/worksheet when needed, with overwrite or
  append semantics.

Rows run sequentially. There is no results/audit output table — the component only writes to
OneDrive/SharePoint; it does not produce any Storage output tables. Column mapping is not
supported: file and CSV modes upload data as-is, and Excel mode uses v1's header-based semantics
(see below) rather than an explicit source/destination column mapping.

Authorization
=============

The component uses **delegated OAuth 2.0** — you authorize it once per configuration through the
Keboola UI's OAuth flow (the "Authorize" button), the same way you would authorize any OAuth-based
component. There is no vendor-side admin setup and no application/service-principal mode; every
request runs on behalf of the authorizing user's Microsoft account.

The root configuration also has an **Account** section that tells the component what kind of
Microsoft account it's connecting to:

| Account Type            | `tenant_id`  | `site_url`   | Notes                                                          |
|--------------------------|:------------:|:------------:|-----------------------------------------------------------------|
| Private OneDrive         | not used     | not used     | Personal Microsoft account. Excel mode and SharePoint document libraries are **not available**. |
| OneDrive for Business    | required     | not used     | Writes to the user's own OneDrive for Business drive.          |
| SharePoint               | required     | required     | Writes to a SharePoint site's document library(ies), including non-default libraries. |

- **Tenant ID** — the Azure AD tenant that owns the account (Azure Portal → Azure Active Directory
  → Overview → Tenant ID). Required for OneDrive for Business and SharePoint accounts.
- **Site URL** — the full URL of the SharePoint site to connect to, e.g.
  `https://contoso.sharepoint.com/sites/DataTeam`. Required for SharePoint accounts only.

Use the **Test Connection** button on the root configuration to verify the account/authorization
combination before adding rows.

Configuration
=============

The configuration has two levels:

- **Root configuration** — set once per configuration: the **Account** (account type, tenant ID,
  site URL as needed) and the OAuth authorization.
- **Configuration rows** — one row per destination. Each row selects an **Output Mode** and
  attaches its own input mapping (a file input mapping for `file` mode, exactly one input table for
  `table_csv`/`table_excel` mode — zero or more than one table in a CSV/Excel row is a
  configuration error).

Destination (file / table_csv modes)
-------------------------------------

- **Document Library** (`destination.drive_id`) — the target document library (drive) on a
  SharePoint site, selected from the "List Libraries" dropdown. Leave empty to use the
  site's/account's default library. Not applicable to Private OneDrive or OneDrive for Business
  accounts.
- **Folder Path** (`destination.folder_path`) — path inside the library, relative to the library
  root. Missing folders are created automatically. Supports `strftime`-style date placeholders
  resolved at run start (UTC), for example:

  ```
  reports/{date:%Y-%m-%d}
  ```

- **Conflict Behavior** (`destination.conflict_behavior`) — what happens when a file with the same
  name already exists at the destination: `fail` (default — stop with an error), `replace`
  (overwrite it), or `rename` (upload under a new, non-colliding name).

CSV options (table_csv mode)
------------------------------

- **File Name** (`csv.file_name`) — name of the uploaded file. Defaults to the input table's name
  with a `.csv` extension when left empty.
- **Delimiter** (`csv.delimiter`) — field delimiter character, default `,`.
- **Enclosure** (`csv.enclosure`) — field enclosure (quote) character, default `"`.
- **Include Header** (`csv.include_header`) — whether to include the column header row, default
  `true`.

Workbook and worksheet (table_excel mode)
-------------------------------------------

- **Workbook Path** (`workbook.path`) — where the target workbook lives; created automatically
  when missing. Accepts several forms:
  - a library-relative path, e.g. `/Reports/data.xlsx`
  - `drive://{driveId}/path`
  - `site://{siteName}/path`
  - an `https://` sharing link

  Alternatively, target an existing workbook by id using **Drive ID** (`workbook.drive_id`) and
  **File ID** (`workbook.file_id`) together — obtain both from the `search` sync action. `path` and
  the drive/file id pair are mutually exclusive; never combine them.
- **Worksheet Name** (`worksheet.name`), **Worksheet ID** (`worksheet.id`), **Worksheet Position**
  (`worksheet.position`) — select or create the target worksheet. Provide `name` alone to select or
  create a worksheet by that name; combine `id` or `position` with `name` to rename that worksheet.
  `id` and `position` are mutually exclusive.
- **Append** (`append`) — when `false` (default), the worksheet is cleared and rewritten from cell
  A1 (overwrite). When `true`, rows are appended below the existing used range; if the sheet
  already has a header, the CSV's own header row is skipped on append (a column-count mismatch
  between the new data and the existing header produces a warning, not a failure).
- **Batch Size** (`batch_size`) — number of rows written per Excel API request, default `5000`.
  Lower it if you hit throttling on very wide tables.

Sync Actions
============

The UI exposes several sync actions to help configure a row: **Test Connection** and **List
Libraries** (root configuration), and **Search**, **Get Worksheets**, **Create Workbook**, and
**Create Worksheet** (used while configuring `table_excel` mode, e.g. from the Workbook Path
field's search helper). Sync actions are normally read-only lookups — **Create Workbook** and
**Create Worksheet** are a deliberate exception: they perform a real write (creating an empty
workbook or worksheet), matching v1's own sync-action behavior. Only use them once you actually
want that workbook or worksheet created.

Output
======

This component does not create any Storage output tables — its only output is the data it writes
to OneDrive/SharePoint. The one exception is internal: rotated OAuth refresh tokens are persisted
to state (both after a row's job run and after any sync action that touches the API) so the
configuration keeps working past token expiry; this requires no action from you.

Development
===========

This repository uses [uv](https://docs.astral.sh/uv/) for dependency management.

Clone this repository, initialize the workspace, and run the component locally:

```
git clone https://github.com/keboola/component-wr-onedrive-v2
cd component-wr-onedrive-v2
uv sync
KBC_DATADIR=./data uv run python src/component.py
```

To customize the local data folder path, replace the `CUSTOM_FOLDER` placeholder with your desired
path in the `docker-compose.yml` file:

```yaml
    volumes:
      - ./:/code
      - ./CUSTOM_FOLDER:/data
```

Run the test suite:

```
uv run pytest
```

Run lint checks:

```
uv run ruff check .
```

Alternatively, using Docker:

```
docker-compose build
docker-compose run --rm dev
docker-compose run --rm test
```

Integration
===========

For details about deployment and integration with Keboola, refer to the
[deployment section of the developer
documentation](https://developers.keboola.com/extend/component/deployment/).
