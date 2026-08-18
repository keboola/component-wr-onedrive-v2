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
from datetime import UTC, datetime
from pathlib import Path

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")
logger = logging.getLogger("record_vcr_cassettes")

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
                "site_url": flat_secrets["site_url"],
            }
        },
    }


def _templated_definitions(run_id: str) -> list[dict]:
    """Load tests/setup/configs.json and substitute the run-unique folder suffix."""
    raw = DEFINITIONS_PATH.read_text()
    substituted = raw.replace(RUN_FOLDER_TOKEN, run_id)
    return json.loads(substituted)


def _record_all(definitions: list[dict], secrets_override: dict, *, regenerate: bool) -> None:
    """Record every scenario, in two batches split around the forced-small-threshold one.

    Deliberately uses ``TestScaffolder.scaffold_from_json`` (not the per-definition
    ``scaffold_from_dict``) for **both** batches: only ``scaffold_from_json`` (via its private
    ``_scaffold_single_test``) threads ``input_files_dir`` through *before* recording, so a
    writer scenario's input CSV/file is already sitting in ``in/tables/``/``in/files/`` by the
    time the component actually runs — ``scaffold_from_dict`` has no ``input_files_dir``
    parameter at all and would record every writer scenario against an empty input mapping.
    Splitting into two batches (rather than one call per scenario) keeps each batch's built-in
    skip-if-exists/regenerate/chained-state handling intact.
    """
    from keboola.vcr.scaffolder import TestScaffolder

    freeze_time_at = datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%S")
    scaffolder = TestScaffolder()

    normal = [d for d in definitions if d["name"] not in SMALL_THRESHOLD_SCENARIOS]
    small_threshold = [d for d in definitions if d["name"] in SMALL_THRESHOLD_SCENARIOS]

    if normal:
        logger.info("Recording %d scenario(s) with the real upload threshold.", len(normal))
        _scaffold_batch(scaffolder, normal, secrets_override, freeze_time_at, regenerate)

    if small_threshold:
        logger.info(
            "Recording %d scenario(s) with a forced small upload threshold (%s).",
            len(small_threshold),
            ", ".join(d["name"] for d in small_threshold),
        )
        with _forced_small_upload_threshold():
            _scaffold_batch(scaffolder, small_threshold, secrets_override, freeze_time_at, regenerate)


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


def _teardown(flat_secrets: dict, run_id: str) -> None:
    """Best-effort cleanup: delete keboola-vcr-tests/<run_id> from every drive used.

    Never raises — a live test tenant accumulating a little test cruft from a failed cleanup is
    vastly preferable to a recording script that errors out *after* cassettes were already
    written. Uses the real credentials directly (never through VCR — this traffic is not
    recorded).
    """
    import requests

    logger.info("Tearing down keboola-vcr-tests/%s from the test tenant (best-effort).", run_id)
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
        access_token = response.json()["access_token"]
    except Exception:
        logger.warning("Teardown: could not obtain an access token; skipping cleanup.", exc_info=True)
        return

    headers = {"Authorization": f"Bearer {access_token}"}
    base = "https://graph.microsoft.com/v1.0"
    path = f"{TEST_ROOT_FOLDER}/{run_id}"

    # The Business account's own default drive.
    _delete_path_best_effort(f"{base}/me/drive/root:/{path}", headers)

    # The SharePoint site's default document library.
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
    _record_all(definitions, secrets_override, regenerate=regenerate)

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
