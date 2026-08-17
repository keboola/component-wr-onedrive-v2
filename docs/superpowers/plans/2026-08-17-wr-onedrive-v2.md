# wr-onedrive-v2 — Implementation Plan

> Spec: `docs/superpowers/specs/2026-08-17-wr-onedrive-v2-design.md` (approved 2026-08-17)
> Branch: `initial-implementation` (never push to `main`; one PR at the end)
> Execution: one fresh subagent per task; implementation tasks owned by `component-develop`
> (schema/UI → `component-build-ui`), test tasks by `component-test` / `generate-vcr-tests`.
> Every task ends with `ruff check` clean and the test suite green.

Task status: `[ ]` open · `[x]` done (checked only after the task's verification passed).

## Task 1 — Configuration models · owner: component-develop
- [x] `src/configuration.py`: Pydantic v2 models per spec §5 — `Account` (explicit `account_type`
  enum `private_onedrive|onedrive_for_business|sharepoint`, `tenant_id`, `site_url`, cross-field
  requiredness), `Destination` (`drive_id`, `folder_path`, `conflict_behavior` enum default `fail`),
  `CsvOptions` (`file_name`, `delimiter`, `enclosure`, `include_header`), `Workbook`
  (`drive_id`+`file_id` XOR `path`, v1 validation rules, opaque `metadata`), `Worksheet`
  (`id` XOR `position`, optional `name`, str→int coercion for `position`, opaque `metadata`),
  `RowConfig` (`mode` enum, `append=False`, `batch_size=5000`). `extra="ignore"`, no `debug` field,
  explicit defaults, no `.parameters.get()` outside models.
- [x] Unit tests: XOR constraints, coercions, requiredness per account type, metadata passthrough.
- Verify: `pytest tests/test_configuration.py` green; `ruff check` clean.

## Task 2 — Token provider · owner: component-develop
- [x] `src/client/auth.py`: `TokenProvider` interface (swappable for CFTL-702) +
  `RefreshTokenProvider`: raw refresh-token grant against
  `login.microsoftonline.com/{common|tenant}/oauth2/v2.0/token`, proactive refresh before
  `expires_in` elapses, rotation (new refresh token captured), candidates tried state-first then
  config, `invalid_grant` on all candidates → `UserException` ("reauthorize"). State key
  `#refreshed_auth_data` (v1-compatible JSON payload).
- [x] Unit tests with mocked HTTP: refresh success, rotation, fallback order, invalid_grant,
  proactive re-refresh.
- Verify: `pytest tests/test_auth.py` green.

## Task 3 — Graph client core · owner: component-develop
- [x] `src/client/graph_client.py`: `requests.Session`-based client; `Authorization` from
  `TokenProvider`; `User-Agent: NONISV|Keboola|wr-onedrive-v2/<version>`; retry policy — honor
  `Retry-After` on 429/503, exponential backoff on 5xx, workbook-transient 405/409 opt-in per call,
  total-wait cap → `UserException`; no retry of non-idempotent chunk PUTs; `@odata.nextLink` paging
  helper; error mapping (401/403/404/507/400 → typed exceptions with response `error.code`/message).
- [x] Unit tests: Retry-After honored, cap exceeded → UserException, paging, error mapping.
- Verify: `pytest tests/test_graph_client.py` green.

## Task 4 — Site/drive resolution + testConnection/listLibraries · owner: component-develop
- [ ] Client methods: site URL → site id (`GET /sites/{hostname}:{path}`), list drives
  (`GET /sites/{site-id}/drives`), account-type dispatch to base drive (`/me/drive` vs site drive).
  Do NOT port the extractor's doubled `/sites/{id}/sites/{id}/lists` path; drive id is the
  canonical library identifier.
- [ ] `src/component.py`: `@sync_action("testConnection")` (`GET /me?$select=userPrincipalName`),
  `@sync_action("listLibraries")` → `SelectElement(label=name, value=drive.id)`, built from root
  config only.
- Verify: unit tests for dispatch + sync-action shapes; `ruff check` clean.

## Task 5 — Drive uploader · owner: component-develop
- [ ] `src/client/uploader.py`: path validation (reserved chars incl. `#%` on business, reserved
  names, 400/255 limits → `UserException` pre-network); per-segment URL encoding (colon syntax);
  date placeholders (`{date:%Y-%m-%d}`, UTC, resolved at run start); folder resolution + creation
  level-by-level (`POST /children`, `conflictBehavior` body key, tolerate existing); simple PUT
  ≤ 10 MiB threshold with `?@microsoft.graph.conflictBehavior=` query (always explicit); upload
  session above threshold — `createUploadSession` with `item.@microsoft.graph.conflictBehavior`,
  sequential 10 MiB chunks (32 × 320 KiB), no Authorization header on chunk PUTs, resume via
  `GET uploadUrl`/`nextExpectedRanges`, session 404 → single restart, `DELETE uploadUrl` on abort,
  final-chunk 409 `nameAlreadyExists` mapped per conflict behavior; streams from disk.
- [ ] Unit tests: chunk math, placeholder resolution, validation, conflict mapping, resume logic
  (mocked HTTP).
- Verify: `pytest tests/test_uploader.py` green.

## Task 6 — File mode + CSV mode orchestration · owner: component-develop
- [ ] `src/component.py`: thin `run()` (< 30 lines) — load merged config → token provider/client →
  dispatch on `mode` → persist rotated token to state (also on failure via `finally`). File mode:
  upload every file from the row's file input mapping to the target folder. CSV mode: exactly one
  input table (0/>1 → `UserException`, v1-parity messages); rewrite delimiter/enclosure/header only
  when options differ from Storage defaults (else stream the input file as-is); scratch in `/tmp`.
- [ ] Datadir-style tests for both modes with mocked client.
- Verify: `pytest` green; `run()` line count; scratch-path assertion.

## Task 7 — Excel writer · owner: component-develop
- [ ] `src/client/excel_writer.py`: workbook resolution for all v1 path forms (`/path`,
  `drive://{driveId}/path`, `site://{siteName}/path` via `GET /sites?search=` requiring exactly one
  hit, `https://` sharing link via `/shares/u!{base64url}`), XLSX MIME check; create-when-missing
  only in path mode (upload minimal valid xlsx built with openpyxl at build time or bundled
  fixture); workbook session lifecycle (`createSession` persistChanges + `Prefer: respond-async`
  202 polling, `Workbook-Session-Id`, recreate on session-404, `closeSession` in `finally`,
  sessionless fallback); worksheet resolve by id/name/position, create in name mode, rename when
  `name` differs; overwrite = `range/clear` + write from A1; append = offset from
  `usedRange(valuesOnly=true)`, header-skip when target has one + `Headers mismatch` warning;
  batched `PATCH .../range(address=...)` with `batch_size` rows, fixed column count from CSV
  header, formula escaping (`=` → `'=`), multi-letter column addressing; serialized writes (no
  concurrency).
- [ ] Unit tests: address math, path-form parsing, header skip semantics, escaping, batching
  boundaries (mocked HTTP).
- Verify: `pytest tests/test_excel_writer.py` green.

## Task 8 — Excel mode wiring + v1-parity sync actions · owner: component-develop
- [ ] Excel mode in `run()` dispatch, gated off for `private_onedrive` (`UserException`); empty CSV
  input → warning + exit 0, sheet untouched.
- [ ] Sync actions with byte-compatible v1 output shapes: `search` (`{"file": {...}}`/
  `{"file": null}`), `createWorkbook` (`{"file": {driveId, fileId}}`, exists → UserException
  "already exists"), `createWorksheet` (`{"worksheet": {...}}`), `getWorksheets` (position-sorted,
  `" (hidden)"` title suffix, ASCII header normalization: NFD strip, non-`[A-Za-z0-9-.]` → `_`,
  blanks `column-{i+1}`, duplicates `-1`/`-2`).
- [ ] Golden tests against v1 expected-stdout fixtures (copied from `keboola/wr-onedrive`
  `tests/datadir/*/expected-stdout`).
- Verify: golden tests green.

## Task 9 — configSchema + configRowSchema + uiOptions · owner: component-build-ui
- [ ] Root schema: account section (`account_type` enum + `enum_titles`, `tenant_id`/`site_url`
  via `options.dependencies`), test-connection widget. Row schema: `mode` selector; destination
  section (async `listLibraries` select with `enum: []` + autoload, folder path, conflict
  behavior); CSV section and Excel section shown via `options.dependencies` on `mode`; Excel
  workbook/worksheet pickers wired to `search`/`getWorksheets`/`createWorkbook`/`createWorksheet`.
  Title Case titles, sentence-case descriptions, tooltips for long help, grid sections.
- [ ] `component_config/uiOptions.md` → `["genericDockerUI", "genericDockerUI-rows",
  "genericDockerUI-authorization", "genericDockerUI-tableInput", "genericDockerUI-fileInput"]`
  (exact set confirmed against portal docs during the task).
- [ ] Update `component_config/*description*.md` (what the component does, per portal conventions)
  and `data/config.json` sample to the real schema.
- Verify: schema lints (JSON valid, every `options.async.action` has a matching `@sync_action`,
  async selects have `enum: []`, required arrays parent-level); schema-tester if available.

## Task 10 — Datadir + unit test completion · owner: component-test
- [ ] Datadir tests (merged single `config.json`, row-scoped state fixtures): happy path per mode;
  append vs overwrite; empty CSV; cardinality errors; missing OAuth; Excel-on-personal; conflict
  `fail` on existing file. Expected exit codes asserted (1 vs 2).
- Verify: full `pytest` suite green in the Docker `test` target (`docker compose` or
  `scripts/build_n_test.sh`).

## Task 11 — VCR harness + cassettes · owner: generate-vcr-tests / component-test
- [ ] Greenfield VCR setup (extractor has none): record against the Keboola M365 test tenant —
  token refresh, site+drive resolution, simple PUT per conflict behavior, upload session with
  forced small threshold incl. resume, folder creation, Excel session/patch/close, worksheet
  create+rename, every sync action. Sanitizers: tokens, client_secret, tenant/site/drive GUIDs,
  uploadUrl signatures, UPNs. Needs user-supplied test-tenant OAuth credentials at record time.
- Verify: cassette validation gate (`component-test` → `references/vcr-validation-gate.md`) —
  sanitization grep clean + recordings match intent; suite green offline.

## Task 12 — Repo hygiene + PR · owner: component-develop
- [ ] README (configuration walkthrough per mode), remove `src/*.egg-info` from tree/gitignore it,
  final `ruff check` + full suite, push `initial-implementation`, open PR to `main` (after Phase 8
  audit per the lifecycle tracker).
- Verify: CI green on the branch.

## Milestone gates (lifecycle tracker `docs/superpowers/wr-onedrive-v2-lifecycle.md`)
- Tasks 1–9 → tracker Phase 4 (scoped checklist-review gate).
- Tasks 10–11 → tracker Phase 5 (testing + credentials gate).
- Then Phase 6 (portal value setup, `component-dev-portal`), Phase 7 (cf-dev smoke test,
  `runtime.tag` override), Phase 8 (full audit) per the tracker.
