# VCR functional tests (plan Task 11)

This directory holds the `keboola.datadirtest` VCR functional test cases for
`keboola.wr-onedrive-v2` — one subdirectory per scenario declared in
[`tests/setup/configs.json`](../setup/configs.json), each containing a recorded, sanitized
Microsoft Graph HTTP cassette (`source/data/cassettes/requests.json`) plus captured logs / sync
action output / expected exit code.

Recording is **greenfield** (the sibling extractor, `kds-team.ex-onedrive`, has no VCR harness
to model this on) and is **deferred**: it requires real Microsoft 365 test-tenant OAuth
credentials that arrive later in a gitignored `secrets.json` at the repo root. Until then, this
directory is (deliberately) empty and every scenario in `tests/test_functional_vcr.py` reports as
a `pytest` **skip** — see that module's docstring for why the suite must stay green in all three
states (not recorded / recorded+replaying / actively recording).

## Recording — one command

```bash
uv run python scripts/record_vcr_cassettes.py
```

Add `--regenerate` to force re-recording every scenario from scratch (default: scenarios that
already have a cassette are left alone), and `--no-teardown` to skip the best-effort cleanup step
(useful if you want to inspect the created objects on the test tenant afterwards).

The script:

1. Reads `secrets.json` (repo root, **gitignored** — never commit it).
2. Templates a run-unique folder suffix (`keboola-vcr-tests/<UTC timestamp>`) into every
   scenario's destination path, so re-recording never collides with objects a previous recording
   run left behind (deliberately-conflicting scenarios instead rely on **file-order sequencing
   within one run** — e.g. `20_file_conflict_replace` uploads the target file that
   `21_file_conflict_fail`/`22_file_conflict_rename` then reuse, rather than each pre-seeding its
   own copy).
3. Records each scenario via `keboola.vcr`'s `TestScaffolder` (the same engine
   `python -m keboola.datadirtest scaffold` uses), applying `VCR_SANITIZERS`
   (`src/component.py`) to every recorded interaction.
4. For `23_upload_session_small_threshold` only, temporarily monkeypatches
   `client.uploader.SIMPLE_UPLOAD_THRESHOLD`/`CHUNK_SIZE` down to a few hundred bytes so a small
   committed fixture file exercises the chunked `createUploadSession` + `PUT` path instead of a
   simple `PUT` — restored immediately after that one scenario finishes recording.
5. Best-effort deletes `keboola-vcr-tests/<run id>` from every drive used (the Business account's
   own drive, and the SharePoint site's default document library) so the test tenant doesn't
   accumulate cruft across recording runs.

### `secrets.json` shape (repo root, gitignored)

```json
{
  "appKey": "<Azure app (client) id>",
  "#appSecret": "<Azure app client secret>",
  "refresh_token": "<a refresh token already granted offline_access, User.Read, "
                   "Files.ReadWrite.All, Sites.ReadWrite.All>",
  "tenant_id": "<test tenant id (GUID) — Business + SharePoint account types only>",
  "site_url": "https://<tenant>.sharepoint.com/sites/<site>"
}
```

Scope decision (2026-08-18): record against **Business + SharePoint account types only**.
Personal-OneDrive flows are already covered by the mocked `GraphFake` suite
(`tests/test_functional_http.py`) — no personal-account cassettes are recorded.

### After recording — run the validation gate

`pytest` passing after a recording is not sufficient proof the cassettes are good — a recording
can be green and still (a) have leaked a secret a sanitizer missed, or (b) have recorded a
failure and quietly called it a pass. **Before committing new/re-recorded cassettes**, run the
cassette validation gate described in the `component-test` skill
(`references/vcr-validation-gate.md`) — ideally via a fresh subagent with no recording bias — and
fix (add a sanitizer + `--regenerate`, or re-record against a working account) until it passes on
both axes.

## What gets sanitized (`VCR_SANITIZERS` in `src/component.py`)

| Sanitizer | Covers |
|---|---|
| `DefaultSanitizer` (with `"code"` dropped from the default field list — Graph's `error.code` error taxonomy is not a secret and redacting it would desync `logs.json` between record and replay) | `access_token`, `refresh_token`, `client_secret`, `client_id`, `client_assertion`, `id_token`, `password` in the token-refresh form body, the JSON token response, and query strings; **all headers except `content-type`/`content-length`/`accept`** — this is what strips `Authorization: Bearer …`, `Set-Cookie`, `Cookie`, and `WWW-Authenticate` |
| `QueryParamSanitizer(parameters=["tempauth"])` | The pre-signed `tempauth` query-string token embedded in upload-session `uploadUrl` values (in both the `createUploadSession` response body and every subsequent chunk `PUT`/`GET`/`DELETE` request URI that uses it) |
| `BodyFieldSanitizer(fields=["userPrincipalName", "mail", "displayName", "givenName", "surname"])` | User principal names / identity fields returned by `/me` and any `createdBy`/`lastModifiedBy` blocks |
| `_StableGuidSanitizer` (custom, in `component.py`) | Real GUIDs — the OAuth `tenant_id` and the two GUIDs embedded in a SharePoint composite site id (`hostname,guid,guid`) — mapped to small, stable, deterministic placeholders (`00000000-0000-4000-8000-000000000001`, `...002`, …) so cassette request/response matching still lines up on replay. Scoped to one cassette recording (the mapping resets per scenario); **not** `scrub_before_read`, since that flavor would scrub the response the *live* component itself reads mid-recording — breaking the real recording run the moment a round-tripped id gets sent back to Graph as a placeholder |

The cassette validation gate greps for exactly this coverage (unredacted `Authorization`/
`Bearer`, `access_token`/`refresh_token`/`client_secret`/`password`/`token` fields, `set-cookie`,
signed-URL params) — see `vcr-validation-gate.md`.

## Scenarios (`tests/setup/configs.json`)

| # | Scenario | Covers |
|---|---|---|
| 01 | `testConnection` | OAuth refresh-token grant (token refresh) + `GET /me` |
| 02 | `listLibraries` | Site resolution (`GET /sites/{hostname}:{path}`) + `GET /sites/{id}/drives` |
| 10 | `search` (missing) | v1-parity search sync action, no match → `{"file": null}` |
| 11 | `createWorkbook` | Create a brand-new empty workbook |
| 12 | `search` (found) | Same path 11 just created → resolves to a real file |
| 13 | `getWorksheets` | Lists the default worksheet of the workbook 11 created |
| 14 | `createWorksheet` | Adds a new named worksheet |
| 15 | `createWorksheet` (conflict) | Same name as 14 → `UserException`, exit 1 (deliberate failure) |
| 20 | `file` mode, `replace` | Simple `PUT`; creates the shared target 21/22 depend on |
| 21 | `file` mode, `fail` | Same target as 20 → Graph 409 → `UserException`, exit 1 (deliberate failure) |
| 22 | `file` mode, `rename` | Same target again → Graph auto-renames |
| 23 | `file` mode, forced small threshold | `createUploadSession` + chunked `PUT` (resumed-chunk sub-case is best-effort, only captured if a transient error happens to occur live) |
| 30 | `table_csv` mode | Upload the single input table as CSV |
| 31 | `table_excel` mode, overwrite | Workbook/worksheet session create → range `PATCH` → session close |
| 32 | `table_excel` mode, append | Same workbook/worksheet as 31, `append=true` |
| 33 | `table_excel` mode, rename | Targets a sheet by `position` with a different `name` → rename-before-write |

## Caveat: refresh-token reuse across scenarios

Every scenario authenticates independently using the **same** `refresh_token` from
`secrets.json` (state isn't chained between scenarios — each recording is a deterministic,
order-independent "first run"). Microsoft Graph refresh tokens are rotated on every use but
typically remain valid for a grace period, so re-recording the whole batch normally works
without needing a fresh token per scenario. If the test tenant enforces strict single-use
rotation and a batch recording starts failing with `invalid_grant` partway through, mint a fresh
refresh token and re-run.

## Re-recording an individual scenario

```bash
rm -rf tests/functional/14_createWorksheet
uv run python scripts/record_vcr_cassettes.py
```

(scenarios with dependents — e.g. 20/21/22, or 11/12/13/14/15 — should be re-recorded together;
delete the whole dependent chain's directories, or just pass `--regenerate` to redo everything).
