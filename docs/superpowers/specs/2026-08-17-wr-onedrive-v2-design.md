# wr-onedrive-v2 — Design Spec

> Type: writer
> Component ID: keboola.wr-onedrive-v2
> Status: draft
> Date: 2026-08-17

## 1. Overview & source system

A data destination that writes files and tables from Keboola into OneDrive and SharePoint document
libraries via Microsoft Graph, with three output modes: **file** (file input mapping → uploaded
as-is), **table as CSV** (table input mapping → CSV file), and **table as Excel worksheet**
(table input mapping → worksheet in an XLSX workbook, feature parity with `keboola.wr-onedrive`).

Target system: Microsoft Graph v1.0 — <https://learn.microsoft.com/en-us/graph/api/overview>.
Primary use case: P3 Group (SUPPORT-15935 / CFTL-488) and other customers landing Keboola outputs in
SharePoint document libraries, including non-default libraries — the specific gap no current
component covers. v2 is the successor to `keboola.wr-onedrive` (Excel-only, PHP).

Reuse sources (read, not modified): `kds-team.ex-onedrive` (`keboola/component-onedrive`) for the
auth/account-type/site-resolution design; `keboola.wr-onedrive` (`keboola/wr-onedrive`, PHP) for the
Excel worksheet write semantics and sync-action contracts.

## 2. Keboola mapping

- **Config rows** (Keboola convention): one row = one destination target — one input table (CSV or
  Excel mode) or one file-selection (file mode) written to one target. Rows run **sequentially by
  default** (platform behaviour); no `parallelism` requested — Excel writes must be serialized per
  workbook anyway (Microsoft guidance), and sequential rows keep refresh-token rotation simple.
- **Root config**: account (`account_type`, `tenant_id`, `site_url`) + OAuth (platform-injected
  `authorization.oauth_api.credentials`). **Row config**: mode, target, mode-specific options.
  Input mapping lives **on the row**, never the root (a root mapping stacks onto every row).
- **Direction**: input mapping → API writes. No output tables (results/audit table excluded by user
  decision, 2026-08-17), so output-mapping, native-types, and default-bucket concerns don't apply.
- **State (`state.json`, per row)**: only rotated OAuth refresh tokens, keyed
  `#refreshed_auth_data` (v1-compatible). Each row has its own state; the platform loads/saves it
  around each row. First run: missing state handled (fall back to config token). No incremental
  watermark — a writer pushes whatever the input mapping provides; "incremental" is the user's
  input-mapping choice, not component state.
- **Secrets**: `#appSecret` and `#data` arrive via the OAuth broker; the rotated token is stored
  under a `#`-prefixed state key so it's encrypted. No other secrets.
- **Sync actions**: `testConnection`, `listLibraries`, plus v1-parity `search`, `getWorksheets`,
  `createWorkbook`, `createWorksheet` (section 5).
- **Scratch files**: `/tmp` only — never `data/out/tables/` or `data/out/files/` (everything there
  is uploaded to Storage). The one legitimate `data/out/` write is `data/out/state.json`, for token
  rotation.

## 3. Authentication & connection

- **Delegated OAuth 2.0 (authorization-code + refresh-token grant) — mandatory, not a preference.**
  The Graph Excel API has *no application-permission mode* (range write and worksheet add are
  delegated-only), and label-protected file replacement is delegated-only. Service-principal
  (client-credentials) auth is deferred to CFTL-702 and isolated behind a `TokenProvider` interface
  so it can be added without reworking upload code.
- **Azure AD app**: reuse `keboola.wr-onedrive` v1's app registration in the platform OAuth broker
  (user decision 2026-08-17). Scopes: `offline_access User.Read Files.ReadWrite.All
  Sites.ReadWrite.All` — already granted to that app.
- **Token flow** (extractor pattern, no MSAL): raw
  `POST https://login.microsoftonline.com/{authority}/oauth2/v2.0/token` with
  `grant_type=refresh_token`. Authority: `common` for private OneDrive, `{tenant_id}` for
  business/SharePoint (guest/cross-tenant scenarios require the explicit tenant). Access token
  refreshed at startup and proactively before expiry (`expires_in` ≈ 3599 s) — not only reactively
  on 401 (fixes the extractor's 60-minute-runtime limitation).
- **Token rotation**: every refresh returns a *new* refresh token; persist it to row state
  (`#refreshed_auth_data`) after each run, and on startup try the state token first, config token
  as fallback (v1 + extractor pattern). `invalid_grant` on both → `UserException` telling the user
  to reauthorize the configuration.
- **Account types** (explicit enum, not inferred from empty strings like the extractor):
  `private_onedrive` (→ `/me/drive`), `onedrive_for_business` (tenant_id, → `/me/drive`),
  `sharepoint` (tenant_id + site_url, → site's drives). Excel mode and the sites API are
  **unsupported on personal Microsoft accounts** — gate with a clear `UserException`.
- **Provisioning**: the user authorizes in the Keboola UI OAuth flow; no vendor-side admin setup
  per configuration. One-time platform step: confirm the v1 Azure app's broker entry is available
  to component id `keboola.wr-onedrive-v2` (open item, section 9).
- **Access**: Keboola M365 test tenant with a SharePoint site is available (user-confirmed) — no
  blocker for VCR recording or the cf-dev smoke test.

## 4. Capability inventory & scope

| Capability | Verdict | Rationale |
|---|---|---|
| File upload, simple `PUT .../content` (≤ 250 MB) | In scope | core file mode |
| File upload via `createUploadSession` (chunked, resumable, > threshold) | In scope | large files; 10 MiB chunks |
| Conflict behavior `fail` / `replace` / `rename` | In scope | per row, file + CSV modes |
| Target: non-default document library (drive picker) | In scope | the gap that motivated the component |
| Target: folder path, auto-create missing folders (level by level) | In scope | — |
| Target: date placeholders in folder path | In scope | dynamic per-run folders |
| Table → CSV (file name, delimiter, enclosure, header toggle) | In scope | streams from disk |
| Table → Excel worksheet: overwrite (clear + write from A1) | In scope | v1 parity |
| Table → Excel worksheet: append (usedRange offset, header skip + mismatch warning) | In scope | v1 parity |
| Excel: batched range PATCH (`batch_size`, default 5000 — deliberate change from v1's 500) | In scope | user decision 2026-08-17 |
| Excel: create workbook when missing (path mode only; upload minimal XLSX) | In scope | v1 parity; no create-workbook endpoint exists |
| Excel: create/rename worksheet | In scope | v1 parity |
| Excel: workbook targeting by `driveId`+`fileId`, `/path`, `drive://`, `site://`, `https://` sharing link | In scope | v1 parity |
| Excel: formula escaping (`=` → `'=`) | In scope | v1 parity, injection safety |
| Account types: private OneDrive / Business / SharePoint | In scope | Excel + sites gated off for personal (Graph limitation) |
| Sync actions: `testConnection`, `listLibraries`, `search`, `getWorksheets`, `createWorkbook`, `createWorksheet` | In scope | v1-parity output shapes for the four ported ones |
| Refresh-token rotation persisted via state | In scope | keeps configs alive past token expiry |
| Results/audit output table (uploaded items → Storage) | Excluded | user opted out 2026-08-17 ("just upload file / table") |
| Sharing-link creation (`createLink`) | Excluded | user opted out 2026-08-17 |
| SharePoint list-item metadata fields on uploaded files | Excluded | user opted out 2026-08-17 |
| Gzip option for CSV mode | Excluded | user opted out 2026-08-17 |
| Service-principal (app-only) auth | Excluded (deferred) | CFTL-702; isolated behind `TokenProvider`; Excel API has no app-only mode at all |
| DriveItem delete / move / copy | Excluded | not writer-shaped; out of a destination's contract |
| SharePoint Embedded containers (`FileStorageContainer.Selected`) | Excluded | different product surface |
| `@microsoft.graph.sourceUrl` server-side fetch upload | Excluded | unsupported on OneDrive for Business / SharePoint Online |
| Excel mode on personal OneDrive | Excluded | Graph Excel API is business-only; gated with UserException |
| Docs, v1 migration tooling, v1 deprecation | Excluded | tracked separately per the Linear definition |

### Mechanics of the in-scope surface

- **Pagination**: only listings paginate (`@odata.nextLink` loop) — drives list, worksheets list.
- **Rate limits**: honor `Retry-After` on 429/503 (SharePoint does not support IETF RateLimit
  headers); exponential backoff when absent; cap total wait and fail with a clear message beyond it
  (fix v1's seconds-vs-milliseconds comparison bug, don't port it). Throttled requests still count
  against quota — no aggressive retry. Decorate traffic:
  `User-Agent: NONISV|Keboola|wr-onedrive-v2/<version>`. Excel hard limits: 5,000 req/10 s per app
  all-tenants, 1,500 req/10 s per app per tenant; never parallelize writes to one workbook.
- **Uploads**: simple PUT up to a 10 MiB threshold (limit is actually 250 MB; threshold follows
  Microsoft's session recommendation), `conflictBehavior` as a **URL query param** (its default is
  `replace` — always set explicitly). Above threshold: upload session, `conflictBehavior` in the
  **body under `item`** (its default is `fail` — always set explicitly), sequential 10 MiB chunks
  (= 32 × 320 KiB), **no Authorization header on chunk PUTs**, resume via `GET uploadUrl` →
  `nextExpectedRanges` on transient failure, session `404` → restart upload, `DELETE uploadUrl` on
  abort so no partial file is left. Late-conflict `409 nameAlreadyExists` on the final chunk is
  mapped per the row's conflict behavior.
- **Excel session**: `createSession` (`persistChanges: true`) with `Prefer: respond-async` +
  `202`/`Location` polling; `Workbook-Session-Id` header on subsequent calls; sessions expire after
  ~5 min inactivity → on session-`404` recreate and retry; `closeSession` in a `finally` block
  (v1 relied on a destructor — do not copy). Sessionless fallback is acceptable (persistence is
  guaranteed without a session; it's a performance optimization).
- **Path/naming**: per-segment percent-encoding (colon syntax `root:/{path}:/{action}`); validate
  against Graph's reserved characters (`/ \ * < > ? : |`, plus `# %` on Business/SharePoint),
  reserved names (`CON`, `_vti_`, leading `~$`, …), 400-char path / 255-char segment limits —
  violations are `UserException`s before any network call.

## 5. Configuration & schema

**Config-level (root) parameters**
- `account.account_type` — enum `private_onedrive` | `onedrive_for_business` | `sharepoint`
  (explicit; drives both UI field visibility and runtime dispatch — unlike the extractor, where the
  UI enum wasn't in the model).
- `account.tenant_id` — required for business + sharepoint.
- `account.site_url` — required for sharepoint.
- OAuth: `authorization.oauth_api.credentials.{appKey, #appSecret, #data}`; `#data` is a JSON
  string containing `refresh_token`. Missing authorization → `UserException` at startup.

**Row-level parameters**
- `mode` — enum `file` | `table_csv` | `table_excel` (UI mode selector; `options.dependencies`
  shows only the relevant target fields).
- `destination.drive_id` — document library, async dropdown via `listLibraries` (**value = drive
  id**, label = library name; not the extractor's URL-slug hack). For non-SharePoint accounts the
  default drive is used and this field is hidden.
- `destination.folder_path` — relative to library root; missing folders created; supports
  `strftime`-style date placeholders (e.g. `reports/{date:%Y-%m-%d}`) resolved at run start (UTC).
  File + CSV modes.
- `destination.conflict_behavior` — enum `fail` | `replace` | `rename`, default `fail`
  (file + CSV modes; Excel mode has its own append/overwrite semantics).
- CSV options (`mode=table_csv`): `csv.file_name` (default: input table name + `.csv`),
  `csv.delimiter` (default `,`), `csv.enclosure` (default `"`), `csv.include_header`
  (default true).
- Excel options (`mode=table_excel`), v1-compatible names inside the row:
  `workbook.{drive_id, file_id, path}` (path XOR ids, v1 validation rules; `path` supports
  `/path`, `drive://`, `site://`, `https://` forms), `worksheet.{id, name, position}` (id XOR
  position, name optionally combined = rename), `append` (default false), `batch_size`
  (default **5000**), plus opaque `workbook.metadata` / `worksheet.metadata` passthrough (UI file
  picker storage — accepted and ignored, v1 parity).
- File mode consumes **all** files from the row's file input mapping; CSV/Excel modes require
  exactly one table in the row's table input mapping (0 or >1 → `UserException`, v1-parity
  messages).

**Column mapping — recorded deviation:** the ui-schema checklist expects an explicit
`{source, destination}` column mapping for structured destinations. Excel mode instead keeps v1's
header-based semantics (append skips the CSV header when the sheet has one; mismatch → warning).
Reason: v1 feature parity is an acceptance criterion, and the user scoped v2 to "just upload
file / table" (2026-08-17). Revisit only if a customer asks for column-level control.

**Sync actions**
- `testConnection` — `GET /me?$select=userPrincipalName` with root credentials.
- `listLibraries` — `GET /sites/{site-id}/drives` (site id via
  `GET /sites/{hostname}:{server-relative-path}`); returns `SelectElement(label=drive.name,
  value=drive.id)`. Built from **root** config only (extractor regression: reading row params in
  the client factory broke the dropdown).
- v1-parity (byte-compatible output shapes — the UI parses them):
  `search` → `{"file": {driveId, fileId, name, path|null}}` or `{"file": null}` when not found;
  `createWorkbook` → `{"file": {driveId, fileId}}`; `createWorksheet` → `{"worksheet": {driveId,
  fileId, worksheetId}}`; `getWorksheets` → `{"worksheets": [{position, name, title (+" (hidden)"
  suffix), driveId, fileId, worksheetId, visible, header}]}` with v1's ASCII header normalization
  (NFD strip, non-`[A-Za-z0-9-.]` → `_`, blanks → `column-{i+1}`, duplicate suffixes `-1`, `-2`).

**Handoff:** the actual `configSchema.json` / `configRowSchema.json` (dependencies, async selects
with `enum: []`, grid sections, Title Case labels, autoload) is built by `component-build-ui`.
`uiOptions` will need `genericDockerUI-rows`, `genericDockerUI-authorization`, and per-row table
input mapping.

## 6. Code architecture

```
src/
  component.py            Component(ComponentBase): thin run(), sync actions, state handling
  configuration.py        Pydantic models: Account, Destination, CsvOptions, Workbook, Worksheet, RowConfig
  client/
    auth.py               TokenProvider (interface) + RefreshTokenProvider (rotation, state-first fallback)
    graph_client.py       GraphClient: session w/ retries (Retry-After), decorated User-Agent, paging helper
    uploader.py           DriveUploader: simple PUT / upload session, folder resolution+creation, path validation
    excel_writer.py       ExcelWriter: workbook resolution (all path forms), session lifecycle, batched range PATCH
```

- `run()` is a thin orchestrator (< 30 lines): load config → build client → dispatch on mode →
  persist rotated token. Logic lives in private methods. Clients are initialized in `__init__` /
  a factory, not inside `run()` — but **no network calls in constructors** (extractor anti-pattern:
  token fetch + site resolution in `__init__`); auth happens lazily on first request via
  `TokenProvider`.
- Pydantic v2 models: typed fields, explicit defaults, `extra="ignore"` (metadata passthrough and
  forward-compat), field aliases where names differ, no `debug` field (platform-handled), no raw
  `.parameters.get()` outside the model. `worksheet.position` coerces `"0"` → `0` (v1 configs hold
  both). Partial models for sync actions that need fewer fields (`createWorkbook` needs only
  `workbook.path`).
- Streaming: CSV mode streams the input file from disk in chunks (both to compute size for the
  threshold decision and to feed chunk PUTs); Excel mode reads the CSV row-batch by row-batch —
  nothing is buffered whole in memory.
- **Error mapping** — `UserException` (exit 1): missing/invalid OAuth (incl. `invalid_grant` →
  "reauthorize the configuration"), 401/403 permissions, 507 quota, library/folder/workbook/
  worksheet not found, invalid path characters/length, non-XLSX workbook target, workbook locked
  (`EditModeCannotAcquireLockTooManyRequests` after retries), Excel mode on `private_onedrive`,
  Retry-After exceeding the wait cap, config validation errors, wrong input-mapping cardinality.
  Everything else (unexpected 5xx after retries, bugs) propagates → exit 2. Failures fail the row's
  job — no silent skips (the extractor silently skipped a 429'd download; do not port).
- Retries: single strategy in `GraphClient` — 429/503/5xx with `Retry-After` honored, plus Excel's
  405/409 transient codes on workbook calls; **no automatic retry of non-idempotent chunk PUTs**
  (resume via `nextExpectedRanges` instead). Not three overlapping layers like the extractor.
- Dependencies: `keboola.component`, `requests`, `pydantic`; test-time `openpyxl` (build the
  minimal empty workbook fixture and verify written content in integration checks). No Graph SDK,
  no MSAL — parity with the extractor's deliberate raw-requests choice.

## 7. Testing

- **Unit**: path/placeholder resolution, reserved-name validation, chunk math (320 KiB multiples),
  Excel column addressing (A→1, AA→27), header normalization (v1 fixtures as golden values),
  formula escaping, Pydantic validation rules (XOR constraints, coercions).
- **Datadir tests** (single merged `config.json` shape — the platform merges root+row before the
  component sees it; state files row-scoped): happy path per mode; append vs overwrite; empty CSV
  input (warning, exit 0, sheet untouched — v1 parity); zero/multiple input tables → exit 1;
  missing OAuth → exit 1; Excel on personal account → exit 1.
- **VCR (greenfield — the extractor has no recorded-HTTP harness)**: record against the Keboola
  test tenant. Cassettes: token refresh; site+drive resolution; simple PUT (each conflict
  behavior); upload session (small file with forced small threshold, incl. a resumed chunk);
  folder creation; Excel session create/patch/close; worksheet create + rename; each sync action.
  Sanitizers: access/refresh tokens, `client_secret`, tenant/site/drive GUIDs, `uploadUrl`
  signatures, user principal names.
- **Sync-action tests**: output shapes asserted byte-for-byte against v1's expected-stdout fixtures
  (copied from the v1 repo's datadir tests as golden files).

## 8. Deployment & validation (CF test project)

- Image for the `initial-implementation` branch is built by CI on push (bootstrap `0.0.1` already
  deployed, so branch builds have a target).
- kbagent: create a config + rows in the **cf-dev** project with `runtime.tag` pinned to the branch
  build; OAuth-authorize against the test tenant; run one job per mode.
- Success = three green jobs: a file lands in a **non-default** library folder (with a date
  placeholder resolved), a table lands as CSV, and a table lands in a worksheet (append verified by
  a second run); conflict `fail` produces a clean exit-1 message on re-run.

## 9. Open risks & blockers

1. **OAuth broker wiring** — `keboola.wr-onedrive-v2` must be registered in the platform OAuth
   broker pointing at v1's Azure app before any platform-side OAuth authorization works (VCR
   recording can use a manually minted refresh token meanwhile). Owner: user/platform team. This is
   the only credential-dependent item.
2. **KDS code-lift agreement** — the Linear def says the auth-design lift from `kds-team.ex-onedrive`
   should be agreed with KDS Team. Design-level reuse (no code copied verbatim) likely moots this,
   but confirm before release. Owner: user.
3. **Excel API throttling under load** — Microsoft publishes no per-request row/size limit for range
   PATCH; `batch_size=5000` is a research-informed default, tuned empirically during Phase 7 smoke
   tests. Low risk: configurable per row.
4. **Sharing-link (`https://`) workbook targeting needs a real shared link** in the test tenant to
   record its cassette — minor test-setup task, not a design risk.
