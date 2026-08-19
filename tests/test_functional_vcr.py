"""VCR functional tests for keboola.wr-onedrive-v2 (plan Task 11 — "VCR harness + cassettes").

Complements, rather than replaces, ``tests/test_functional_http.py`` (the hand-rolled
``GraphFake``-backed functional layer built in Task 10, kept because
``keboola.datadirtest``'s plain ``DataDirTester``/``TestDataDir`` re-enters ``component.py``'s own
``if __name__ == "__main__":`` block via ``runpy.run_path(..., run_name="__main__")``, whose
``sys.exit(1)``/``sys.exit(2)`` handlers raise an uncaught ``SystemExit`` that isn't safe to run
inside a plain ``unittest.TestCase`` batch — see that module's docstring for the full analysis).
That hazard does **not** apply here: ``keboola.datadirtest.vcr``'s ``VCRRecorder``/
``VCRTestDataDir`` run the component through ``run_with_log_capture``, which explicitly catches
``SystemExit`` and converts it into a captured exit code (``keboola/vcr/log_capture.py``:
"Does NOT re-raise SystemExit — callers (record/replay) handle the exit code.") — so this module
is exactly the harness the plan asks for, and it's safe for a suite that mixes exit-0 and
deliberate exit-1 scenarios in the same run.

This module's job is real HTTP replay against **recorded, sanitized Microsoft Graph cassettes**
(business + SharePoint account types only — personal OneDrive stays covered by the existing
mocked/GraphFake suite per the 2026-08-18 scope decision). Recording is a separate, deliberate,
one-command step that requires real Microsoft 365 test-tenant credentials — never something
``pytest`` does implicitly:

    uv run python scripts/record_vcr_cassettes.py

See ``tests/functional/README.md`` for the full record / re-record procedure, the
``secrets.json`` shape, and what gets sanitized. Because recording is out-of-band, this module
must stay green in all three states the harness can be run in:

- **Cassette present** (current state — every scenario in ``tests/setup/configs.json`` has been
  recorded and its ``tests/functional/<name>/`` directory committed): the scenario **replays**
  fully offline — no network access, no credentials needed.
- **No cassette recorded yet** (e.g. a newly-declared scenario before anyone has run the
  recording script against it): it is reported as an explicit, visible ``pytest`` **skip** (not
  silently dropped from collection) naming the exact command to record it.
- **Recording** (a human runs the script above with real credentials in ``secrets.json``, e.g.
  to re-record a scenario after a behavior change): this module itself never records anything —
  it only ever replays what's already on disk.

Scenario **names** are parametrized from ``tests/setup/configs.json`` (the source of truth for
what's declared), not from ``keboola.datadirtest.vcr.get_test_cases()`` (which only enumerates
directories that already *have* a cassette) — parametrizing over the declared list is what makes
the "not recorded yet" state produce real, individually-named skips instead of just zero
collected tests, so ``pytest``'s summary line makes the pending-recording count observable.
"""

import contextlib
import json
import sys
from pathlib import Path

import pytest
from keboola.datadirtest.vcr import VCRDataDirTester

_TESTS_DIR = Path(__file__).parent
FUNCTIONAL_DIR = str(_TESTS_DIR / "functional")
_SRC_DIR = _TESTS_DIR.parent / "src"
COMPONENT_SCRIPT = str(_SRC_DIR / "component.py")
DEFINITIONS_PATH = _TESTS_DIR / "setup" / "configs.json"

# Must match scripts/record_vcr_cassettes.py's SMALL_THRESHOLD_SCENARIOS/forced values exactly:
# that scenario's cassette was recorded with client.uploader.SIMPLE_UPLOAD_THRESHOLD/CHUNK_SIZE
# monkeypatched down to a few hundred bytes (so a small fixture file goes through
# createUploadSession + chunked PUT instead of a simple PUT). Replay must apply the identical
# patch — otherwise, with the real ~10 MiB threshold, the small fixture takes the simple-PUT path
# instead and the outgoing request never matches the recorded chunked-upload interactions.
_SMALL_THRESHOLD_SCENARIOS = frozenset({"23_upload_session_small_threshold"})
_FORCED_SIMPLE_UPLOAD_THRESHOLD = 200
_FORCED_CHUNK_SIZE = 100


def _declared_scenario_names() -> list[str]:
    """Every scenario name declared in tests/setup/configs.json, in file order."""
    with open(DEFINITIONS_PATH) as f:
        definitions = json.load(f)
    return [definition["name"] for definition in definitions]


def _has_cassette(test_name: str) -> bool:
    return (Path(FUNCTIONAL_DIR) / test_name / "source" / "data" / "cassettes" / "requests.json").exists()


@contextlib.contextmanager
def _forced_small_upload_threshold_if_needed(test_name: str):
    """See ``_SMALL_THRESHOLD_SCENARIOS`` above — a no-op context for every other scenario."""
    if test_name not in _SMALL_THRESHOLD_SCENARIOS:
        yield
        return

    if str(_SRC_DIR) not in sys.path:
        sys.path.insert(0, str(_SRC_DIR))
    from client import uploader as uploader_module

    original_threshold = uploader_module.SIMPLE_UPLOAD_THRESHOLD
    original_chunk_size = uploader_module.CHUNK_SIZE
    uploader_module.SIMPLE_UPLOAD_THRESHOLD = _FORCED_SIMPLE_UPLOAD_THRESHOLD
    uploader_module.CHUNK_SIZE = _FORCED_CHUNK_SIZE
    try:
        yield
    finally:
        uploader_module.SIMPLE_UPLOAD_THRESHOLD = original_threshold
        uploader_module.CHUNK_SIZE = original_chunk_size


@pytest.mark.parametrize("test_name", _declared_scenario_names())
def test_vcr_functional(test_name):
    """Replay one recorded VCR scenario, or skip it if it hasn't been recorded yet.

    Never attempts a live network call in any state — recording only ever happens via
    ``scripts/record_vcr_cassettes.py``, run manually with real credentials.
    """
    if not _has_cassette(test_name):
        pytest.skip(
            f"'{test_name}' has no recorded cassette yet. Record it with "
            "`uv run python scripts/record_vcr_cassettes.py` (needs real Microsoft 365 "
            "test-tenant credentials in a gitignored secrets.json — see "
            "tests/functional/README.md), then commit the resulting "
            f"tests/functional/{test_name}/ directory."
        )

    tester = VCRDataDirTester(
        data_dir=FUNCTIONAL_DIR,
        component_script=COMPONENT_SCRIPT,
        selected_tests=[test_name],
        validate_snapshots=True,
    )
    with _forced_small_upload_threshold_if_needed(test_name):
        tester.run()
