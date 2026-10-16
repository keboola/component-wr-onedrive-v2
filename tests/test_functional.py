"""VCR functional replay for keboola.wr-onedrive-v2 (network-free).

Auto-discovers and replays every case under ``tests/functional/*``. Every scenario in that tree
has already been recorded against the real Microsoft 365 test tenant and sanitized at record time
(secrets/PII stripped — see ``tests/functional/README.md``), so replay here never touches the
network. When no cassettes are present at all (e.g. a fresh checkout before anything has been
recorded) this skips cleanly with a single named reason instead of collecting zero tests silently.

Deterministic logic (per mode, error mapping, token-state persistence) is covered independently
and network-free in ``tests/test_datadir.py``; this module is the replay-regression net over the
real recorded Microsoft Graph wire shapes.

Recording is a deliberate, manual, credentialed step — ``pytest`` itself never records, only
replays:

    uv run python tests/setup/record_vcr_cassettes.py

See ``tests/functional/README.md`` for the full record/re-record procedure, the ``secrets.json``
shape, and what gets sanitized.
"""

import contextlib
import sys
from pathlib import Path

import pytest
from keboola.datadirtest.vcr import VCRDataDirTester, get_test_cases

FUNCTIONAL_DIR = str(Path(__file__).parent / "functional")
_SRC_DIR = Path(__file__).parent.parent / "src"
COMPONENT_SCRIPT = str(_SRC_DIR / "component.py")

_ALL_CASES = get_test_cases(FUNCTIONAL_DIR) if Path(FUNCTIONAL_DIR).exists() else []

_NO_CASSETTES_REASON = "No VCR functional cases present — nothing to replay. See tests/README.md."

# Must match tests/setup/record_vcr_cassettes.py's SMALL_THRESHOLD_SCENARIOS/forced values exactly:
# that scenario's cassette was recorded with client.uploader.SIMPLE_UPLOAD_THRESHOLD/CHUNK_SIZE
# monkeypatched down to a few hundred bytes (so a small fixture file goes through
# createUploadSession + chunked PUT instead of a simple PUT). Replay must apply the identical
# patch — otherwise, with the real ~10 MiB threshold, the small fixture takes the simple-PUT path
# instead and the outgoing request never matches the recorded chunked-upload interactions.
_SMALL_THRESHOLD_SCENARIOS = frozenset({"23_upload_session_small_threshold"})
_FORCED_SIMPLE_UPLOAD_THRESHOLD = 200
_FORCED_CHUNK_SIZE = 100


def _parametrized_cases() -> list:
    if not _ALL_CASES:
        return [pytest.param("_no_cases_", marks=pytest.mark.skip(reason=_NO_CASSETTES_REASON))]
    return list(_ALL_CASES)


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


@pytest.mark.parametrize("test_name", _parametrized_cases())
def test_functional(test_name):
    """Replay a single recorded VCR functional case."""
    tester = VCRDataDirTester(
        data_dir=FUNCTIONAL_DIR,
        component_script=COMPONENT_SCRIPT,
        selected_tests=[test_name],
        validate_snapshots=True,
    )
    with _forced_small_upload_threshold_if_needed(test_name):
        tester.run()
