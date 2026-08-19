"""Single-command VCR recording entrypoint for keboola.wr-onedrive-v2 (plan Task 11).

Not part of the runtime image's import graph (``keboola.vcr``/``keboola.datadirtest`` are
*dev*-only dependencies — see ``pyproject.toml``'s ``[dependency-groups] dev``, and
``Dockerfile``'s production stage runs ``uv sync --no-dev``). This script is only ever invoked
manually by a human holding real Microsoft 365 test-tenant credentials, never at runtime.

Usage::

    uv run python scripts/record_vcr_cassettes.py

Reads ``secrets.json`` (repo root, gitignored) — shape::

    {
      "appKey": "<Azure app (client) id>",
      "#appSecret": "<Azure app client secret>",
      "refresh_token": "<a refresh token already granted the Sites.ReadWrite.All / "
                        "Files.ReadWrite.All / User.Read / offline_access scopes>",
      "tenant_id": "<test tenant id (GUID) — Business + SharePoint account types only>",
      "site_url": "<https://<tenant>.sharepoint.com/sites/<site> used by the SharePoint "
                  "scenarios (listLibraries, createWorkbook, search, ...)>"
    }

and records every scenario declared in ``tests/setup/configs.json`` into
``tests/functional/<name>/source/data/cassettes/`` via ``keboola.vcr``'s ``TestScaffolder`` —
the same engine ``python -m keboola.datadirtest scaffold`` uses. A thin wrapper (rather than the
bare CLI) is needed for two reasons this component's harness requires beyond the generic
``--secrets`` flag:

1. ``secrets.json``'s flat shape (fixed by an earlier project decision) doesn't line up
   1:1 with the nested ``authorization.oauth_api.credentials`` / ``parameters.account`` paths
   the component actually reads — see :func:`_build_secrets_override`.
2. One scenario (``23_upload_session_small_threshold``) needs
   ``client.uploader.SIMPLE_UPLOAD_THRESHOLD``/``CHUNK_SIZE`` monkeypatched down to a tiny value
   for *that recording only*, so a small fixture file exercises the chunked
   ``createUploadSession`` path instead of a simple ``PUT`` — see
   :data:`SMALL_THRESHOLD_SCENARIOS`.

Every scenario's destination path is namespaced under ``keboola-vcr-tests/<run id>`` (a
run-unique suffix — see :func:`_run_id`) so re-recording never collides with objects left behind
by a previous recording run (deliberately-conflicting scenarios rely on file-order sequencing
*within* one run, e.g. ``20_file_conflict_replace`` before ``21_file_conflict_fail``, not on the
target already existing from an old run). :func:`_teardown` then best-effort deletes that whole
folder tree from both drives used (the Business account's own drive, and the SharePoint site's
default document library) — see ``tests/functional/README.md`` for the full record/re-record
procedure and the cassette validation gate that must be run after recording.
"""

from __future__ import annotations

import contextlib
import json
import logging
import sys
import time
from datetime import UTC, datetime
from pathlib import Path

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")
logger = logging.getLogger("record_vcr_cassettes")

# vcrpy's own "vcr.stubs"/"vcr.cassette" loggers print full request URIs (and, for upload-session
# chunk PUT/GET/DELETE calls, the *live, unsanitized* pre-signed uploadUrl — including its
# tempauth JWT, which embeds the authorized user's UPN) at INFO/DEBUG level, e.g. "<Request ...>
# not in cassette, sending to real server". Sanitizers only ever touch what gets *written* to the
# cassette — never this live logging — so it must be silenced up front, before any recording
# happens, or a real credential can end up in a terminal/log capture outside the cassette file
# itself (see keboola.vcr.recorder.VCRRecorder.record_debug_run, which does the same for exactly
# this reason).
for _noisy_logger in ("vcr", "vcr.stubs", "vcr.cassette", "urllib3"):
    logging.getLogger(_noisy_logger).setLevel(logging.WARNING)

REPO_ROOT = Path(__file__).resolve().parent.parent
SECRETS_PATH = REPO_ROOT / "secrets.json"
DEFINITIONS_PATH = REPO_ROOT / "tests" / "setup" / "configs.json"
INPUT_FILES_DIR = REPO_ROOT / "tests" / "setup" / "input_files"
OUTPUT_DIR = REPO_ROOT / "tests" / "functional"
COMPONENT_SCRIPT = REPO_ROOT / "src" / "component.py"

RUN_FOLDER_TOKEN = "__RUN_ID__"
TEST_ROOT_FOLDER = "keboola-vcr-tests"

# Scenario names that need a tiny SIMPLE_UPLOAD_THRESHOLD/CHUNK_SIZE to exercise the chunked
# createUploadSession path instead of a simple PUT (design spec §7 / plan Task 11's "upload
# session with a forced small threshold").
SMALL_THRESHOLD_SCENARIOS = frozenset({"23_upload_session_small_threshold"})
FORCED_SIMPLE_UPLOAD_THRESHOLD = 200  # bytes
FORCED_CHUNK_SIZE = 100  # bytes — forces at least 2 chunk PUTs for the fixture file above


def _run_id() -> str:
    return datetime.now(UTC).strftime("%Y%m%d%H%M%S")


def _load_secrets() -> dict:
    if not SECRETS_PATH.exists():
        logger.error(
            "secrets.json not found at %s. This script requires real Microsoft 365 test-tenant "
            "credentials — see tests/functional/README.md for the required shape.",
            SECRETS_PATH,
        )
        sys.exit(1)
    with open(SECRETS_PATH) as f:
        return json.load(f)


def _build_secrets_override(flat_secrets: dict) -> dict:
    """Translate the flat secrets.json shape into the nested paths the config actually uses.

    ``TestScaffolder``/``VCRRecorder`` deep-merge this dict into each scenario's ``config`` dict
    (see ``keboola.vcr.scaffolder.TestScaffolder._deep_merge``) purely by matching key *paths* —
    a flat ``{"appKey": ...}`` would land as a new, unused top-level ``config["appKey"]`` instead
    of overriding ``config["authorization"]["oauth_api"]["credentials"]["appKey"]``. Building the
    matching nested shape here keeps ``secrets.json`` itself flat (an earlier, already-fixed
    project decision) without changing how the component reads its config.

    Only ``account.tenant_id`` is overridden here (not ``site_url``): ``tests/setup/configs.json``
    already carries the *real* SharePoint site URL directly (a non-secret resource identifier —
    see ``src/component.py``'s "deliberately NOT sanitized" note), used identically at record and
    replay time, so nothing needs merging/restoring for it. ``tenant_id`` is different: it's
    genuinely scrubbed from cassettes (the real tenant GUID must never be committed), so the real
    value is only ever used for the live token-refresh call during recording; the dummy
    placeholder committed to ``config.json`` afterward must exactly equal
    ``vcr_sanitizers._GuidRedactor.PLACEHOLDER`` for replay to match (see that class's docstring).
    """
    required = ["appKey", "#appSecret", "refresh_token", "tenant_id", "site_url"]
    missing = [key for key in required if not flat_secrets.get(key)]
    if missing:
        logger.error("secrets.json is missing required key(s): %s", ", ".join(missing))
        sys.exit(1)

    return {
        "authorization": {
            "oauth_api": {
                "credentials": {
                    "appKey": flat_secrets["appKey"],
                    "#appSecret": flat_secrets["#appSecret"],
                    "#data": json.dumps({"refresh_token": flat_secrets["refresh_token"]}),
                }
            }
        },
        "parameters": {
            "account": {
                "tenant_id": flat_secrets["tenant_id"],
            }
        },
    }


def _templated_definitions(run_id: str) -> list[dict]:
    """Load tests/setup/configs.json and substitute the run-unique folder suffix."""
    raw = DEFINITIONS_PATH.read_text()
    substituted = raw.replace(RUN_FOLDER_TOKEN, run_id)
    return json.loads(substituted)


# Short pause between scenario recordings. SharePoint/Excel has observable eventual-consistency
# lag right after a workbook is created via a raw file upload (not through the Excel session
# APIs) — the very next scenario's `/workbook/worksheets` call can 404 (ItemNotFound) even though
# the driveItem itself already exists and later scenarios succeed. One-scenario-per-call plus a
# fixed pause between every recording (not just the workbook-chain ones — cheap insurance for any
# other sequential dependency) trades ~1 minute of extra recording time for determinism.
INTER_SCENARIO_DELAY_SECONDS = 4.0

# Scenario names whose workbook needs one Excel session opened+closed on it (see
# _warm_up_workbook) immediately after recording, before any later scenario reuses that workbook
# through a sessionless call. A fixed settle *delay* alone was tried first and did not fix this
# (still 404ed after 20s) — it isn't a consistency-lag issue, a raw-uploaded workbook seemingly
# never becomes readable through sessionless usedRange calls until some session has opened it.
WARM_UP_AFTER = {"11_createWorkbook"}


def _record_all(definitions: list[dict], secrets_override: dict, flat_secrets: dict, *, regenerate: bool) -> None:
    """Record every scenario, one ``scaffold_from_json`` call per scenario.

    Deliberately uses ``TestScaffolder.scaffold_from_json`` (not the per-definition
    ``scaffold_from_dict``): only ``scaffold_from_json`` (via its private
    ``_scaffold_single_test``) threads ``input_files_dir`` through *before* recording, so a
    writer scenario's input CSV/file is already sitting in ``in/tables/``/``in/files/`` by the
    time the component actually runs — ``scaffold_from_dict`` has no ``input_files_dir``
    parameter at all and would record every writer scenario against an empty input mapping. One
    call per scenario (rather than a single batch call) is what makes the inter-scenario delay
    and post-creation workbook warm-up below possible, and keeps regenerate/skip-if-exists
    semantics identical either way.
    """
    from keboola.vcr.scaffolder import TestScaffolder

    freeze_time_at = datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%S")
    scaffolder = TestScaffolder()

    for index, definition in enumerate(definitions):
        name = definition["name"]
        if name in SMALL_THRESHOLD_SCENARIOS:
            logger.info("Recording %s with a forced small upload threshold.", name)
            with _forced_small_upload_threshold():
                _scaffold_batch(scaffolder, [definition], secrets_override, freeze_time_at, regenerate)
        else:
            logger.info("Recording %s.", name)
            _scaffold_batch(scaffolder, [definition], secrets_override, freeze_time_at, regenerate)

        if name in WARM_UP_AFTER:
            _warm_up_workbook(flat_secrets, name)

        _resanitize_sync_action_result(name)

        if index < len(definitions) - 1 and INTER_SCENARIO_DELAY_SECONDS:
            time.sleep(INTER_SCENARIO_DELAY_SECONDS)


def _resanitize_sync_action_result(test_dir_name: str) -> None:
    """Re-apply vcr_sanitizers._GuidRedactor's GUID collapsing to a recorded sync_action_result.json.

    That file is captured from the *live* (pre-sanitization) sync action return value —
    ``VCRRecorder.record()``'s own stdout-capture sanitization only substitutes exact known
    secrets (``LogSanitizer``), never ``VCR_SANITIZERS``' custom transforms. A sync action that
    echoes back a chained GUID Graph handed it (e.g. ``getWorksheets``/``createWorksheet``'s
    ``worksheetId``) would otherwise commit the *real* GUID here while the HTTP cassette (and
    hence what replay reconstructs from it) has already collapsed that same GUID to
    ``_GuidRedactor.PLACEHOLDER`` — a replay-time "sync action output mismatch" even though the
    HTTP layer replayed perfectly. Reusing the exact same regex/placeholder (imported directly,
    not re-implemented) keeps this a single source of truth. ``_GuidRedactor`` itself lives in
    ``vcr_sanitizers.py`` (extracted out of ``component.py`` — see that module's own docstring),
    not on ``component`` itself.
    """
    result_path = OUTPUT_DIR / test_dir_name / "source" / "data" / "cassettes" / "sync_action_result.json"
    if not result_path.exists():
        return

    sys.path.insert(0, str(REPO_ROOT / "src"))
    import vcr_sanitizers

    original = result_path.read_text()
    sanitized = vcr_sanitizers._GuidRedactor._GUID_RE.sub(vcr_sanitizers._GuidRedactor.PLACEHOLDER, original)
    if sanitized != original:
        result_path.write_text(sanitized)
        logger.info("Re-sanitized GUIDs in %s.", result_path)


def _scaffold_batch(scaffolder, definitions: list[dict], secrets_override: dict, freeze_time_at: str, regenerate: bool) -> None:
    """Write ``definitions``/``secrets_override`` to throwaway temp files and record via
    ``scaffold_from_json`` — the only public entrypoint that threads ``input_files_dir`` through
    *before* recording (see :func:`_record_all`). ``scaffold_from_json`` only accepts a secrets
    *file path* (``secrets_file``), not an in-memory dict, so the already-translated nested
    override from :func:`_build_secrets_override` is written to a temp file (outside the repo,
    never committed) purely to satisfy that signature.
    """
    import tempfile

    with tempfile.NamedTemporaryFile("w", suffix=".json", delete=False) as defs_tmp:
        json.dump(definitions, defs_tmp)
        defs_path = Path(defs_tmp.name)
    with tempfile.NamedTemporaryFile("w", suffix=".json", delete=False) as secrets_tmp:
        json.dump(secrets_override, secrets_tmp)
        secrets_path = Path(secrets_tmp.name)

    try:
        scaffolder.scaffold_from_json(
            definitions_file=defs_path,
            output_dir=OUTPUT_DIR,
            component_script=COMPONENT_SCRIPT,
            record=True,
            freeze_time_at=freeze_time_at,
            secrets_file=secrets_path,
            regenerate=regenerate,
            input_files_dir=INPUT_FILES_DIR,
        )
    finally:
        defs_path.unlink(missing_ok=True)
        secrets_path.unlink(missing_ok=True)


@contextlib.contextmanager
def _forced_small_upload_threshold():
    """Temporarily shrink client.uploader's simple-upload threshold and chunk size.

    Patches the *module* globals directly (not a bound reference) so every caller that looks up
    ``SIMPLE_UPLOAD_THRESHOLD``/``CHUNK_SIZE`` at call time — which is how ``client/uploader.py``
    itself reads them — sees the patched value for the duration of this context, then the
    original value is restored unconditionally.
    """
    sys.path.insert(0, str(REPO_ROOT / "src"))
    from client import uploader as uploader_module

    original_threshold = uploader_module.SIMPLE_UPLOAD_THRESHOLD
    original_chunk_size = uploader_module.CHUNK_SIZE
    uploader_module.SIMPLE_UPLOAD_THRESHOLD = FORCED_SIMPLE_UPLOAD_THRESHOLD
    uploader_module.CHUNK_SIZE = FORCED_CHUNK_SIZE
    try:
        yield
    finally:
        uploader_module.SIMPLE_UPLOAD_THRESHOLD = original_threshold
        uploader_module.CHUNK_SIZE = original_chunk_size


def _get_access_token(flat_secrets: dict) -> str | None:
    """Fetch a live access token directly (never through VCR — this traffic is not recorded).

    Returns ``None`` (never raises) on failure — used only by best-effort helpers
    (:func:`_teardown`, :func:`_warm_up_workbook`).
    """
    import requests

    try:
        token_url = f"https://login.microsoftonline.com/{flat_secrets['tenant_id']}/oauth2/v2.0/token"
        response = requests.post(
            token_url,
            data={
                "client_id": flat_secrets["appKey"],
                "client_secret": flat_secrets["#appSecret"],
                "grant_type": "refresh_token",
                "refresh_token": flat_secrets["refresh_token"],
                "scope": "offline_access User.Read Files.ReadWrite.All Sites.ReadWrite.All",
            },
            timeout=30,
        )
        response.raise_for_status()
        return response.json()["access_token"]
    except Exception:
        logger.warning("Could not obtain a live access token.", exc_info=True)
        return None


def _warm_up_workbook(flat_secrets: dict, test_dir_name: str) -> None:
    """Seed one cell in a just-created workbook's default sheet (best-effort, never raises).

    A workbook created via a raw file upload (``createUploadSession``/simple ``PUT`` — never
    through the Excel session APIs) has no ``usedRange`` yet: its *sessionless* ``usedRange``-based
    endpoints (``client/excel_writer.py``'s ``_read_header_row``, used by the ``getWorksheets``
    sync action) 404 with ``ItemNotFound`` until at least one cell has been written — observed
    empirically live; neither a 20s settle delay nor opening+closing an empty session (both tried
    first) fixed it, only an actual write does. Reads the driveId/fileId the just-recorded
    scenario's sync action returned from its own ``sync_action_result.json``.
    """
    import requests

    result_path = OUTPUT_DIR / test_dir_name / "source" / "data" / "cassettes" / "sync_action_result.json"
    if not result_path.exists():
        logger.warning("Warm-up: no sync_action_result.json for %s; skipping.", test_dir_name)
        return
    try:
        file_info = json.loads(result_path.read_text())["file"]
        drive_id, file_id = file_info["driveId"], file_info["fileId"]
    except (json.JSONDecodeError, KeyError, TypeError):
        logger.warning("Warm-up: could not parse %s; skipping.", result_path)
        return

    access_token = _get_access_token(flat_secrets)
    if access_token is None:
        return
    headers = {"Authorization": f"Bearer {access_token}"}
    base = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}/workbook"
    session_headers = dict(headers)
    try:
        session_id = requests.post(
            f"{base}/createSession", headers=headers, json={"persistChanges": True}, timeout=30
        ).json().get("id")
        if session_id:
            session_headers["workbook-session-id"] = session_id
        worksheets = requests.get(f"{base}/worksheets", headers=session_headers, timeout=30).json().get("value", [])
        if not worksheets:
            logger.warning("Warm-up: %s/%s has no worksheets; skipping cell seed.", drive_id, file_id)
            return
        first_sheet_id = worksheets[0]["id"]
        requests.patch(
            f"{base}/worksheets/{first_sheet_id}/range(address='A1')",
            headers=session_headers,
            json={"values": [["seed"]]},
            timeout=30,
        )
        if session_id:
            requests.post(f"{base}/closeSession", headers=session_headers, timeout=30)
        logger.info("Warm-up: seeded cell A1 in %s/%s.", drive_id, file_id)
    except Exception:
        logger.warning("Warm-up: failed to seed a cell; continuing anyway.", exc_info=True)


def _teardown(flat_secrets: dict, run_id: str) -> None:
    """Best-effort cleanup: delete keboola-vcr-tests/<run_id> from the SharePoint site's drive.

    Every scenario targets the SharePoint site's own drive only (the authorized test user has no
    provisioned personal OneDrive — every ``/me/drive`` call 403s with ``notAllowed`` — see
    ``tests/functional/README.md``), so there is no Business-account drive to clean up here.

    Never raises — a live test tenant accumulating a little test cruft from a failed cleanup is
    vastly preferable to a recording script that errors out *after* cassettes were already
    written.
    """
    logger.info("Tearing down keboola-vcr-tests/%s from the test tenant (best-effort).", run_id)
    access_token = _get_access_token(flat_secrets)
    if access_token is None:
        logger.warning("Teardown: could not obtain an access token; skipping cleanup.")
        return

    import requests

    headers = {"Authorization": f"Bearer {access_token}"}
    base = "https://graph.microsoft.com/v1.0"
    path = f"{TEST_ROOT_FOLDER}/{run_id}"

    # The SharePoint site's default document library — the only drive any scenario touches.
    try:
        from urllib.parse import urlparse

        parsed = urlparse(flat_secrets["site_url"])
        hostname = parsed.netloc
        server_relative_path = parsed.path.rstrip("/")
        lookup = f"{base}/sites/{hostname}:{server_relative_path}" if server_relative_path else f"{base}/sites/{hostname}"
        site_response = requests.get(lookup, headers=headers, timeout=30)
        site_response.raise_for_status()
        site_id = site_response.json()["id"]
        _delete_path_best_effort(f"{base}/sites/{site_id}/drive/root:/{path}", headers)
    except Exception:
        logger.warning("Teardown: could not resolve the SharePoint site drive; skipping.", exc_info=True)


def _delete_path_best_effort(url: str, headers: dict) -> None:
    import requests

    try:
        response = requests.delete(url, headers=headers, timeout=30)
        if response.status_code in (204, 404):
            logger.info("Teardown: deleted (or already absent) %s", url)
        else:
            logger.warning("Teardown: unexpected status %s deleting %s", response.status_code, url)
    except Exception:
        logger.warning("Teardown: failed to delete %s", url, exc_info=True)


def main() -> None:
    regenerate = "--regenerate" in sys.argv
    no_teardown = "--no-teardown" in sys.argv

    flat_secrets = _load_secrets()
    secrets_override = _build_secrets_override(flat_secrets)
    run_id = _run_id()
    logger.info("Recording run id: %s (destination folder: %s/%s)", run_id, TEST_ROOT_FOLDER, run_id)

    definitions = _templated_definitions(run_id)
    _record_all(definitions, secrets_override, flat_secrets, regenerate=regenerate)

    if no_teardown:
        logger.info("--no-teardown passed; leaving recorded objects on the test tenant.")
    else:
        _teardown(flat_secrets, run_id)

    logger.info(
        "Done. Recorded %d scenario(s) into %s. Next: run the cassette validation gate "
        "(component-test skill, references/vcr-validation-gate.md) before committing — "
        "grep every cassette for secrets and confirm each recording matches its declared intent.",
        len(definitions),
        OUTPUT_DIR,
    )


if __name__ == "__main__":
    main()
