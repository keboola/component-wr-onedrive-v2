OneDrive/SharePoint Writer
==========================

Writes files and tables from Keboola to Microsoft OneDrive and SharePoint document libraries via
Microsoft Graph. This component is the successor to `keboola.wr-onedrive` ("OneDrive Excel
Sheets") — it keeps full feature parity for Excel worksheet writes and adds a second output mode
for plain file/CSV upload.

**Table of Contents:**

[TOC]

Functionality Notes
====================

Each configuration row targets exactly one destination, in one of two output modes:

- **File** (`mode: file`) — uploads every file in the row's file input mapping as-is, *and*
  writes every table in the row's table input mapping as a CSV file, both to the same document
  library folder. Either input mapping may be empty, but not both.
- **Worksheet** (`mode: worksheet`) — the row's single input table is written into a worksheet of
  an XLSX workbook, creating the workbook/worksheet when needed, with overwrite or append
  semantics.

Rows run sequentially. There is no results/audit output table — the component only writes to
OneDrive/SharePoint; it does not produce any Storage output tables. Column mapping is not
supported: mode `file` uploads/writes data as-is, and mode `worksheet` uses v1's header-based
semantics (see below) rather than an explicit source/destination column mapping.

> **Note on older configurations:** the pre-2026-08 mode names `table_csv` and `table_excel` are
> still accepted (silently normalized to `file`/`worksheet` respectively) — nothing needs to be
> re-saved.

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
- **Configuration rows** — one row per destination. Each row selects an **Output Mode** and attaches
  its own input mapping:
  - mode `file` — any combination of a file input mapping (zero or more files) and a table input
    mapping (zero or more tables), as long as at least one of the two is non-empty.
  - mode `worksheet` — exactly one input table (zero or more than one is a configuration error);
    any file input mapping is ignored (with a warning).

Destination (mode `file`)
--------------------------

- **Document Library** (`destination.drive_id`) — the target document library (drive), selected
  from the "List Libraries" dropdown. Applies to every account type when set (drives are globally
  addressable in Microsoft Graph); leave empty to use the account's/site's default document
  library. The UI only shows this field for SharePoint accounts (Private OneDrive/OneDrive for
  Business have a single default drive), but any already-configured value is still honored
  verbatim for every account type.
- **Folder Path** (`destination.folder_path`) — path inside the library, relative to the library
  root. Missing folders are created automatically. Supports the `{{date}}` placeholder, for
  example:

  ```
  acme/reports/{{date}}/
  ```

  `{{date}}` resolves to the job's start time (UTC) by default, or to **Date**
  (`destination.date`, below) when set, formatted `YYYY-MM-DD`.
- **Date** (`destination.date`) — the date used to resolve `folder_path`'s (and `csv.file_name`'s)
  `{{date}}` placeholder, parsed via [`dateparser`](https://github.com/scrapinghub/dateparser).
  Accepts a relative expression (`yesterday`, `3 days ago`, `last week`) or an absolute date
  (`2026-01-31`). Leave empty to default to the job's start date (UTC).
- **Conflict Behavior** (`destination.conflict_behavior`) — what happens when a file with the same
  name already exists at the destination: `fail` (default — stop with an error), `replace`
  (overwrite it), or `rename` (upload under a new, non-colliding name).

CSV options for mapped tables (mode `file`)
---------------------------------------------

Only applies to tables from the row's table input mapping — mapped files are always uploaded
as-is, byte-for-byte.

- **File Name** (`csv.file_name`) — name of the uploaded CSV file. Only applies when exactly one
  table is mapped to the row (mapping more than one table while also setting `csv.file_name` is a
  configuration error). Defaults to the table's name with a `.csv` extension when left empty; with
  more than one table mapped, each is uploaded as `<table name>.csv`. Supports the `{{date}}`
  placeholder (resolved the same way as `destination.folder_path`'s).
- **Delimiter** (`csv.delimiter`) — field delimiter character, default `,`.
- **Enclosure** (`csv.enclosure`) — field enclosure (quote) character, default `"`.
- **Include Header** (`csv.include_header`) — whether to include the column header row, default
  `true`.

Workbook and worksheet (mode `worksheet`)
-------------------------------------------

**Workbook targeting** (`workbook.targeting`) picks how the target workbook is identified:

- **Pick via dropdowns** (`pick`, default) — target an existing workbook by id, using the
  **Library** (`workbook.drive_id`) and **Workbook** (`workbook.file_id`) dropdowns together.
- **By path** (`path`) — target (and auto-create, if missing) a workbook via **Workbook Path**
  (`workbook.path`), which accepts several forms:
  - a library-relative path, e.g. `/Reports/data.xlsx`
  - `drive://{driveId}/path`
  - `site://{siteName}/path`
  - an `https://` sharing link

  Use the `search` sync action to check whether a path already resolves to an existing workbook.

Only the fields for the selected targeting are used; the other form's value (if any is still
sitting in the row's configuration from before switching `workbook.targeting`) is ignored.

**Worksheet selection** (`worksheet.selection`) picks how the target worksheet is identified:

- **Pick existing** (`pick`, default) — target an existing sheet by id, using the **Worksheet ID**
  (`worksheet.id`) dropdown (populated from the workbook configured above). Never renames the
  sheet.
- **By name (creates if missing)** (`name`) — target **Worksheet Name** (`worksheet.name`) alone;
  the sheet is created if it doesn't exist yet.

As with workbook targeting, only the selected form's field is used.

- **Write Mode** (`write_mode`) — how the row's data is written into the worksheet:
  - **Full Load** (`overwrite`, default) — the worksheet is cleared and rewritten from cell A1.
  - **Append** (`append`) — rows are appended below the existing used range; if the sheet already
    has a header, the CSV's own header row is skipped on append (a column-count mismatch between
    the new data and the existing header produces a warning, not a failure).
  - **Upsert (Incremental)** (`upsert`) — existing rows are matched to CSV rows by **Key Columns**
    (`key_columns`, below); a match whose values differ is updated in place, a CSV row with no
    match is appended, and an unchanged match is left untouched. A brand-new or genuinely empty
    worksheet behaves like Full Load. Requires the existing used range to be at most 500,000
    cells — a larger sheet should use Full Load or Append instead.
- **Key Columns** (`key_columns`) — column name(s) from the input table's header that identify a
  row, e.g. `id`. Required (non-empty) when **Write Mode** is Upsert; unused otherwise.
- **Batch Size** (`batch_size`) — number of rows written per Excel API request, default `5000`.
  Lower it if you hit throttling on very wide tables.

> **Note on older configurations:** the pre-2026-08 `append: true`/`false` boolean is still
> accepted (silently normalized to `write_mode: append`/`overwrite`) — nothing needs to be
> re-saved.

Sync Actions
============

The UI exposes several sync actions to help configure a row: **Test Connection** and **List
Libraries** (root configuration), and **List Workbooks**, **List Worksheets**, **Search**, **Get
Worksheets**, **Create Workbook**, and **Create Worksheet** (used while configuring mode
`worksheet`, e.g. from the Workbook/Worksheet dropdowns or the Workbook Path field's search
helper). Sync actions are normally read-only lookups — **Create Workbook** and **Create Worksheet**
are a deliberate exception: they perform a real write (creating an empty workbook or worksheet),
matching v1's own sync-action behavior. Only use them once you actually want that workbook or
worksheet created.

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
