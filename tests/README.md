# Test suite map

470 tests, organized by layer. If you're adding a test and aren't sure where it goes, read the
"which layer?" line for each directory below — it's almost always `unit/` unless you're touching
`Component.run()`/the sync actions or an end-to-end flow.

```
tests/
  unit/         per-module unit tests, mocked HTTP (client.auth, client.graph_client,
                client.drives, client.uploader, client.excel_writer, client.headers,
                client.exceptions, configuration)
  component/    Component-level behavior: sync actions, run() per mode, schema/action
                cross-check
  e2e/          end-to-end through Component.run()/the entrypoint script
  parity/       v1 (PHP keboola.wr-onedrive) byte-compatibility golden tests
  functional/   recorded VCR scenario data dirs (UNCHANGED layout — do not restructure)
  setup/        VCR recording scenario definitions + input fixture files (UNCHANGED)
```

## Which layer do I add a new test to?

- **Testing one function/class in isolation, with HTTP or collaborators mocked out** →
  `unit/`. This is almost always the right place for a new test — cheapest to write, fastest to
  run, easiest to pin down a failure.
- **Testing `Component` behavior that spans several collaborators but doesn't need real
  HTTP-shaped payloads** (sync actions, per-mode `run()` behavior, the schema↔action
  cross-check) → `component/`.
- **Testing the whole component through `Component.run()`/the entrypoint script, with HTTP
  mocked at the transport boundary or replayed from a cassette** → `e2e/`.
- **Asserting that a ported v1 sync action produces byte-identical output to the legacy PHP
  writer** → `parity/` (see below — you're very unlikely to need a new one of these; the four
  ported actions are already covered).

Don't add new fixture data under `functional/` or `setup/` by hand — those are populated by the
VCR recording flow described next.

## The VCR flow (how `functional/` cassettes get there)

1. `setup/configs.json` declares every scenario (name, parameters, input files) — the source of
   truth for what's supposed to exist, whether or not it's been recorded yet.
2. `uv run python scripts/record_vcr_cassettes.py` records each declared scenario against the
   real Microsoft 365 test tenant (credentials from a gitignored `secrets.json`), sanitizing
   secrets/PII out of every request/response, and writes the result to
   `functional/<scenario>/` (config, cassette, expected stdout/exit code).
3. `e2e/test_functional_vcr.py` replays every scenario declared in `setup/configs.json` fully
   offline: cassette present → replay and assert; cassette missing → an explicit, named `pytest`
   skip (not silently dropped) telling you the exact recording command to run.

Recording is a deliberate, manual, credentialed step — `pytest` itself never records, only
replays. See `functional/README.md` for the full record/re-record procedure, the `secrets.json`
shape, and what gets sanitized.

## Why `parity/golden/` exists

`parity/golden/` (renamed from `fixtures/v1_parity/`) holds `expected-stdout`/`expected-stderr`
files copied **verbatim** from `keboola.wr-onedrive` (v1, PHP)'s own datadir test suite. Matching
these byte-for-byte (modulo v1's own `%s`/`%a`/`%A` wildcards for dynamic ids) is an acceptance
criterion for the four sync actions ported from v1: `search`, `createWorkbook`, `createWorksheet`,
`getWorksheets` must keep producing the exact output shape v1 did, so that existing Keboola
configurations relying on that output don't break when a project migrates from v1 to v2. See
`parity/test_v1_parity.py`'s module docstring for the wildcard-matching details.
