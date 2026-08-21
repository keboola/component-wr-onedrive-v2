# Tests

Run everything with `uv run pytest`. No credentials are needed to run the suite — the committed
VCR cassettes contain only sanitized data and replay with no network access.

## Layout

| Path | Kind | Network | Notes |
|------|------|---------|-------|
| `test_unit.py` | Unit | none | Per-module unit tests for `client.auth`, `client.graph_client`, `client.drives`, `client.uploader`, `client.headers`, and `configuration` — mocked HTTP, organized below by source module with a section banner + its own test classes. |
| `test_excel_writer.py` | Unit | none | `client.excel_writer` on its own (1,000+ lines — a distinct rendering engine: column/range math, workbook/worksheet resolution, session lifecycle, write/append/upsert). |
| `test_datadir.py` | Component / functional (in-process) | none | `Component`'s sync actions and `run()` orchestration against a `KBC_DATADIR`-style fixture, in three sections: mocked-`GraphClient` sync actions, mocked-collaborator `run()` orchestration, and the same code path end-to-end with only the HTTP transport boundary mocked (`GraphFake`). |
| `test_functional.py` | VCR replay | none (replay) | Auto-discovers and replays every case under `functional/`. Skips cleanly if no cassettes exist. |
| `test_recording_support.py` | Unit | none | Coverage for the VCR sanitizer mechanisms in `src/vcr_sanitizers.py` (see that module's docstring; this file currently has no tests of its own — noted deliberately in its docstring). |
| `test_security.py` | Unit | none | Secret-leak-prevention: `client.exceptions.sanitize_exception_text`, `client.auth._redact_identities`, and the query-string/upload-URL redaction regressions that used to live in `test_auth.py`/`test_uploader.py`. |
| `test_schema.py` | Cross-check | none | Every `options.async.action` in `component_config/*Schema.json` has a matching `@sync_action` on `Component`, plus assorted schema-shape/UX-copy assertions. Skipped when `component_config/` isn't present (the Docker test image copies only `src/`/`tests/`). |
| `test_v1_parity.py` | Golden | none | v1 (PHP `keboola.wr-onedrive`) byte-compatibility for the four ported sync actions, against fixtures in `fixtures/v1_parity/`. |
| `conftest.py` | — | — | Currently a placeholder — every module above is self-contained; see its docstring. |
| `fixtures/v1_parity/` | Golden fixtures | — | Verbatim `expected-stdout`/`expected-stderr` copied from v1's own datadir test suite — see `fixtures/v1_parity/README.md`. |
| `functional/` | VCR cases | none (replay) | The recorded cassette scenarios (see below). |
| `setup/configs.json` | Definitions | — | The scenario matrix `record_vcr_cassettes.py` records from. |
| `setup/input_files/` | Fixtures | — | Input files/tables the writer scenarios upload. |
| `setup/record_vcr_cassettes.py` | Recorder | live | Records every scenario in `setup/configs.json` against the live Microsoft 365 test tenant. |
| `setup/authorize_oauth.py` | Recorder prerequisite | live | One-time interactive OAuth helper that writes a `refresh_token` into `secrets.json` for the recorder above. |
| `setup/generate_empty_workbook_fixture.py` | Generator | none | One-off generator for `src/client/fixtures/empty.xlsx`. |

## Which module do I add a new test to?

- **Testing one function/class in isolation, with HTTP or collaborators mocked out** →
  `test_unit.py` (or `test_excel_writer.py` if it's the Excel writer). This is almost always the
  right place for a new test — cheapest to write, fastest to run, easiest to pin down a failure.
- **Testing `Component` behavior that spans several collaborators but doesn't need real
  HTTP-shaped payloads, or the whole component end-to-end with only HTTP mocked** →
  `test_datadir.py`.
- **A secret/PII must never leak into a log line, exception message, or job output** →
  `test_security.py`.
- **A VCR sanitizer's own guarantee (what gets scrubbed before a cassette is written)** →
  `test_recording_support.py`.
- **Asserting that a ported v1 sync action produces byte-identical output to the legacy PHP
  writer** → `test_v1_parity.py` (you're very unlikely to need a new one of these; the four
  ported actions are already covered).

Don't add new fixture data under `functional/` or `setup/` by hand — those are populated by the
VCR recording flow described next.

## The VCR flow (how `functional/` cassettes get there)

1. `setup/configs.json` declares every scenario (name, parameters, input files) — the source of
   truth for what's supposed to exist, whether or not it's been recorded yet.
2. `uv run python tests/setup/record_vcr_cassettes.py` records each declared scenario against the
   real Microsoft 365 test tenant (credentials from a gitignored `secrets.json`), sanitizing
   secrets/PII out of every request/response, and writes the result to `functional/<scenario>/`
   (config, cassette, expected stdout/exit code).
3. `test_functional.py` auto-discovers and replays every case already recorded under `functional/`,
   fully offline; if none exist yet, it skips cleanly with a single named reason.

Recording is a deliberate, manual, credentialed step — `pytest` itself never records, only
replays. See `functional/README.md` for the full record/re-record procedure, the `secrets.json`
shape, and what gets sanitized.

## Why `fixtures/v1_parity/` exists

`fixtures/v1_parity/` holds `expected-stdout`/`expected-stderr` files copied **verbatim** from
`keboola.wr-onedrive` (v1, PHP)'s own datadir test suite. Matching these byte-for-byte (modulo
v1's own `%s`/`%a`/`%A` wildcards for dynamic ids) is an acceptance criterion for the four sync
actions ported from v1: `search`, `createWorkbook`, `createWorksheet`, `getWorksheets` must keep
producing the exact output shape v1 did, so that existing Keboola configurations relying on that
output don't break when a project migrates from v1 to v2. See `test_v1_parity.py`'s module
docstring for the wildcard-matching details, and `fixtures/v1_parity/README.md` for the fixture
provenance.
