# Scripts

`build_n_test.sh` — the Docker `test` stage's entrypoint: runs `ruff check` (lint failures block
the build) then `pytest tests/ --tb=short -q`. Mirrors the `cookiecutter-python-component`
template this repo tracks — keep changes to it minimal and template-compatible.

`developer_portal/` — CI scripts (owned by the `cookiecutter-python-component` template) that push
`component_config/*` to the Keboola Developer Portal on release. Template-owned: do not hand-edit;
changes should flow from the template instead.

The VCR recording/authorization/fixture-generation helpers that used to live here
(`record_vcr_cassettes.py`, `authorize_oauth.py`, `generate_empty_workbook_fixture.py`) moved to
`tests/setup/` — they are test-recording tooling, not part of the component's own build/test/release
pipeline. See `tests/README.md` and `tests/functional/README.md` for the current record/re-record
procedure.
