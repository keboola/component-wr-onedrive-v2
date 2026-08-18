"""Path handling, folder resolution, and file upload for keboola.wr-onedrive-v2.

Design spec: ``docs/superpowers/specs/2026-08-17-wr-onedrive-v2-design.md`` §4 "Mechanics of the
in-scope surface" (Uploads, Path/naming).

Three independent concerns live here, kept as module-level functions taking an injected
:class:`~client.graph_client.GraphClient` (same style as ``client.drives`` — ``graph_client.py``
itself stays transport-only):

- **Path handling** — ``resolve_placeholders`` (date tokens) and ``validate_path`` (Graph's
  reserved-character/name/length rules) run entirely client-side, *before* any network call, so a
  bad ``destination.folder_path`` fails fast with a clear message instead of a confusing 400 deep
  into a folder-creation walk.
- **Folder resolution** — ``ensure_folder`` walks a (validated, placeholder-resolved) folder path
  level by level, creating any segment that doesn't exist yet.
- **Upload** — ``upload_file`` dispatches to a simple ``PUT`` or a chunked ``createUploadSession``
  based on file size, streaming from disk in both cases (never buffering a whole file in memory).

Callers (Task 6's ``run()``) are expected to call ``resolve_placeholders`` once at run start (with
a single UTC ``now``), then ``validate_path``, then ``ensure_folder``, then ``upload_file`` per
file — in that order, so validation always happens before the first request.
"""

import logging
import os
import re
from datetime import datetime
from urllib.parse import quote

import requests

from client.exceptions import (
    FileAlreadyExistsError,
    GraphClientError,
    GraphConnectionError,
    GraphNotFoundError,
    InvalidPathError,
    UploadSessionError,
    sanitize_exception_text,
)
from client.graph_client import GraphClient

logger = logging.getLogger(__name__)

# Above this size, use a resumable `createUploadSession` instead of a simple `PUT .../content`.
# Graph allows simple PUT up to 250 MB, but recommends switching to sessions well before that;
# 10 MiB follows Microsoft's own guidance rather than the hard limit.
SIMPLE_UPLOAD_THRESHOLD = 10 * 1024 * 1024

# Chunk size for upload-session PUTs: 32 * 320 KiB, Graph's recommended chunk-size multiple.
CHUNK_SIZE = 10 * 1024 * 1024

# Bounded resume attempts per upload session before giving up (aborting the session and raising).
MAX_RESUME_ATTEMPTS = 3

# HTTP statuses treated as transient during a chunk PUT: retried via the resume dance
# (`GET uploadUrl` -> `nextExpectedRanges`) rather than propagated immediately.
_TRANSIENT_CHUNK_STATUSES = frozenset({429, 500, 502, 503, 504})

# Only `{date:<strftime-format>}` placeholders are supported; anything else inside braces is a
# user configuration error caught before any network call.
_PLACEHOLDER_PATTERN = re.compile(r"\{([^{}]*)\}")
_DATE_PLACEHOLDER_PREFIX = "date:"

# Reserved characters within a single path segment (the `/` separator itself is never checked
# here — paths are already split on it before validation). `#` and `%` are additionally reserved
# on OneDrive for Business / SharePoint (but valid on personal OneDrive).
_RESERVED_CHARS_BASE = '"*:<>?\\|'
_RESERVED_CHARS_BUSINESS_EXTRA = "#%"

# Exact (case-insensitive) reserved file/folder names.
_RESERVED_EXACT_NAMES = frozenset({".lock", "desktop.ini"})
# Windows reserved device names: CON, PRN, AUX, NUL, COM0-9, LPT0-9.
_RESERVED_DEVICE_NAME_PATTERN = re.compile(r"^(CON|PRN|AUX|NUL|COM[0-9]|LPT[0-9])$", re.IGNORECASE)

MAX_SEGMENT_LENGTH = 255
MAX_PATH_LENGTH = 400


def resolve_placeholders(path: str, now: datetime) -> str:
    """Replace ``{date:<strftime-format>}`` tokens in ``path`` with ``now`` (caller-supplied UTC).

    ``now`` is resolved once by the caller at run start (design spec: "resolved at run start,
    UTC") and threaded through here so every file in a run lands under the same folder even if
    the run crosses a date boundary mid-execution.

    Any other ``{...}`` token is a configuration error: it either doesn't map to a supported
    placeholder syntax or would silently pass an unresolved literal through to Graph.
    """

    def _replace(match: re.Match) -> str:
        token = match.group(1)
        if token.startswith(_DATE_PLACEHOLDER_PREFIX):
            date_format = token[len(_DATE_PLACEHOLDER_PREFIX) :]
            if not date_format:
                raise InvalidPathError(
                    f"Empty date format in placeholder '{{{token}}}' in path '{path}'. Use e.g. '{{date:%Y-%m-%d}}'."
                )
            try:
                return now.strftime(date_format)
            except ValueError as exc:
                raise InvalidPathError(
                    f"Invalid strftime format '{date_format}' in placeholder '{{{token}}}' in path '{path}'."
                ) from exc
        raise InvalidPathError(
            f"Unknown placeholder '{{{token}}}' in destination.folder_path '{path}'. Only "
            "'{date:<strftime-format>}' placeholders are supported (e.g. '{date:%Y-%m-%d}')."
        )

    return _PLACEHOLDER_PATTERN.sub(_replace, path)


def validate_path(path: str, business: bool) -> None:
    """Validate a (placeholder-resolved) folder path against Graph's naming rules.

    Raises :class:`~client.exceptions.InvalidPathError` on the first violation found — always
    before any network call. ``business`` widens the reserved-character set with ``#``/``%``,
    which are only reserved on OneDrive for Business / SharePoint (design spec §4).
    """
    if len(path) > MAX_PATH_LENGTH:
        raise InvalidPathError(
            f"Destination path is {len(path)} characters long; the maximum is {MAX_PATH_LENGTH}. Path: '{path}'"
        )

    reserved_chars = _RESERVED_CHARS_BASE + (_RESERVED_CHARS_BUSINESS_EXTRA if business else "")
    segments = _split_segments(path)
    for index, segment in enumerate(segments):
        _validate_segment(segment, reserved_chars, is_root=(index == 0))


def _split_segments(path: str) -> list[str]:
    return [segment for segment in path.split("/") if segment != ""]


def _validate_segment(segment: str, reserved_chars: str, *, is_root: bool) -> None:
    if len(segment) > MAX_SEGMENT_LENGTH:
        raise InvalidPathError(
            f"Path segment '{segment}' is {len(segment)} characters long; the maximum is {MAX_SEGMENT_LENGTH}."
        )
    if segment != segment.strip():
        raise InvalidPathError(
            f"Path segment '{segment}' has leading/trailing spaces, which OneDrive/SharePoint does not allow."
        )
    for char in reserved_chars:
        if char in segment:
            raise InvalidPathError(f"Path segment '{segment}' contains the reserved character '{char}'.")
    if segment.endswith("."):
        raise InvalidPathError(f"Path segment '{segment}' ends with '.', which is not allowed.")
    if segment.startswith("~"):
        raise InvalidPathError(f"Path segment '{segment}' starts with '~', which is not allowed.")

    lowered = segment.lower()
    if lowered in _RESERVED_EXACT_NAMES:
        raise InvalidPathError(f"Path segment '{segment}' is a reserved name.")
    if _RESERVED_DEVICE_NAME_PATTERN.match(segment):
        raise InvalidPathError(f"Path segment '{segment}' is a reserved device name.")
    if "_vti_" in lowered:
        raise InvalidPathError(f"Path segment '{segment}' contains the reserved substring '_vti_'.")
    if is_root and lowered == "forms":
        raise InvalidPathError("A root-level folder named 'forms' is reserved by SharePoint.")


def encode_path_segments(path: str) -> str:
    """Percent-encode ``path`` segment by segment for Graph's colon-syntax addressing.

    Each segment is independently percent-encoded (``urllib.parse.quote`` with ``safe=""``, so
    even characters ``quote`` normally leaves alone, like ``/``, are encoded *within* a segment)
    before rejoining with literal ``/`` — encoding the whole path in one pass would also encode
    the segment separators themselves.
    """
    return "/".join(quote(segment, safe="") for segment in _split_segments(path))


def ensure_folder(client: GraphClient, drive_id: str, folder_path: str) -> str:
    """Resolve ``folder_path`` to a folder item id, creating any missing segment along the way.

    Walks the path level by level: for each prefix, ``GET /drives/{id}/root:/{path-so-far}``
    first (cheap, and correct when another run already created the folder); on 404, create the
    segment via ``POST .../children`` with ``conflictBehavior: fail``. A 409 ``nameAlreadyExists``
    on that create means a concurrent run won the race — re-``GET`` the same path instead of
    failing. An empty ``folder_path`` resolves to the drive's root item.
    """
    segments = _split_segments(folder_path or "")
    if not segments:
        return client.get(f"/drives/{drive_id}/root").json()["id"]

    parent_id: str | None = None
    encoded_segments: list[str] = []
    for segment in segments:
        encoded_segments.append(quote(segment, safe=""))
        encoded_path = "/".join(encoded_segments)
        try:
            response = client.get(f"/drives/{drive_id}/root:/{encoded_path}")
            parent_id = response.json()["id"]
            continue
        except GraphNotFoundError:
            pass

        parent_id = _create_or_reget_folder(client, drive_id, parent_id, segment, encoded_path)

    # `segments` is non-empty here (the empty case returns above), so the loop always runs at
    # least once and `parent_id` is always set by the time we get here.
    assert parent_id is not None
    return parent_id


def _create_or_reget_folder(
    client: GraphClient, drive_id: str, parent_id: str | None, segment: str, encoded_path: str
) -> str:
    children_url = (
        f"/drives/{drive_id}/items/{parent_id}/children"
        if parent_id is not None
        else f"/drives/{drive_id}/root/children"
    )
    body = {"name": segment, "folder": {}, "@microsoft.graph.conflictBehavior": "fail"}
    try:
        response = client.post(children_url, json=body)
        return response.json()["id"]
    except GraphClientError as exc:
        if not _is_name_conflict(exc):
            raise
        logger.info("Folder '%s' already existed (creation race); re-fetching its id.", segment)
        return client.get(f"/drives/{drive_id}/root:/{encoded_path}").json()["id"]


def upload_file(
    client: GraphClient,
    drive_id: str,
    parent_item_id: str,
    local_path: str,
    file_name: str,
    conflict_behavior: str,
) -> dict:
    """Upload ``local_path`` into the folder identified by ``parent_item_id`` as ``file_name``.

    ``parent_item_id`` is a folder item id — the drive root's own id (for uploads directly into
    the library root) or a folder id returned by :func:`ensure_folder`; never a raw path.

    Dispatches on ``os.path.getsize(local_path)``: files at or below
    :data:`SIMPLE_UPLOAD_THRESHOLD` use a single streamed ``PUT .../content``; larger files use a
    chunked ``createUploadSession``. ``conflict_behavior`` (``"fail" | "replace" | "rename"``) is
    always set explicitly in both paths — Graph's undocumented-by-default behavior differs between
    the two endpoints (``replace`` for simple PUT, ``fail`` for upload sessions), so relying on
    the default would be a silent, mode-dependent surprise.

    Returns the resulting driveItem dict (Graph's response body for the successful request).
    """
    size = os.path.getsize(local_path)
    if size <= SIMPLE_UPLOAD_THRESHOLD:
        return _simple_upload(client, drive_id, parent_item_id, local_path, file_name, conflict_behavior)
    return _session_upload(client, drive_id, parent_item_id, local_path, size, file_name, conflict_behavior)


def _simple_upload(
    client: GraphClient,
    drive_id: str,
    parent_item_id: str,
    local_path: str,
    file_name: str,
    conflict_behavior: str,
) -> dict:
    encoded_name = quote(file_name, safe="")
    url = f"/drives/{drive_id}/items/{parent_item_id}:/{encoded_name}:/content"
    with open(local_path, "rb") as file_handle:
        try:
            response = client.put(
                url,
                params={"@microsoft.graph.conflictBehavior": conflict_behavior},
                data=file_handle,
                headers={"Content-Type": "application/octet-stream"},
            )
        except GraphClientError as exc:
            if _is_name_conflict(exc):
                raise FileAlreadyExistsError(
                    f"'{file_name}' already exists at the destination and conflict behavior is 'fail'.",
                    status_code=exc.status_code,
                    error_code=exc.error_code,
                ) from exc
            raise
    return response.json()


class _SessionExpired(Exception):
    """Internal signal: the upload session's ``uploadUrl`` returned 404 (session gone).

    Caught only inside this module — never escapes :func:`upload_file` — to trigger the single
    whole-session restart the design spec allows.
    """


def _session_upload(
    client: GraphClient,
    drive_id: str,
    parent_item_id: str,
    local_path: str,
    size: int,
    file_name: str,
    conflict_behavior: str,
) -> dict:
    upload_url = _create_upload_session(client, drive_id, parent_item_id, file_name, conflict_behavior)
    already_restarted = False

    with open(local_path, "rb") as file_handle:
        while True:
            try:
                return _upload_chunks(client, file_handle, upload_url, size, file_name, conflict_behavior)
            except _SessionExpired as exc:
                if already_restarted:
                    raise UploadSessionError(
                        f"Upload session for '{file_name}' expired again immediately after being "
                        "restarted once; giving up. Please retry the job."
                    ) from exc
                already_restarted = True
                logger.info("Upload session for '%s' expired (404); restarting it once.", file_name)
                upload_url = _create_upload_session(client, drive_id, parent_item_id, file_name, conflict_behavior)
                file_handle.seek(0)


def _create_upload_session(
    client: GraphClient, drive_id: str, parent_item_id: str, file_name: str, conflict_behavior: str
) -> str:
    encoded_name = quote(file_name, safe="")
    body = {"item": {"@microsoft.graph.conflictBehavior": conflict_behavior, "name": file_name}}
    response = client.post(
        f"/drives/{drive_id}/items/{parent_item_id}:/{encoded_name}:/createUploadSession",
        json=body,
    )
    return response.json()["uploadUrl"]


def _upload_chunks(
    client: GraphClient,
    file_handle,
    upload_url: str,
    size: int,
    file_name: str,
    conflict_behavior: str,
) -> dict:
    """Stream ``file_handle`` to ``upload_url`` in :data:`CHUNK_SIZE` pieces from offset 0.

    Returns the final driveItem dict on success. Raises :class:`_SessionExpired` when the session
    itself is gone (404), for :func:`_session_upload` to restart once. On a transient chunk
    failure (connection error / 429 / 5xx), queries the session status
    (``GET uploadUrl`` -> ``nextExpectedRanges``) and resumes from the server's expected offset,
    bounded to :data:`MAX_RESUME_ATTEMPTS`; beyond that, aborts the session
    (``DELETE uploadUrl``, best-effort) and raises. A final-chunk 409 ``nameAlreadyExists`` is
    mapped per ``conflict_behavior``.
    """
    offset = 0
    resume_attempts = 0
    file_handle.seek(0)

    while offset < size:
        chunk_end = min(offset + CHUNK_SIZE, size) - 1
        length = chunk_end - offset + 1
        chunk = file_handle.read(length)
        headers = {
            "Content-Length": str(length),
            "Content-Range": f"bytes {offset}-{chunk_end}/{size}",
        }
        try:
            response = client.put(upload_url, data=chunk, headers=headers, absolute=True, auth=False, retry=False)
        except GraphClientError as exc:
            # `GraphConnectionError` (a `GraphClientError` subclass) is how a connection error /
            # timeout / DNS failure now surfaces from `client.put()` — it carries no status code,
            # so it falls through to `_resume_after_failure`, which treats it as transient.
            if exc.status_code == 404:
                raise _SessionExpired from exc
            if _is_name_conflict(exc):
                raise _map_late_conflict(exc, file_name, conflict_behavior) from exc
            offset, resume_attempts = _resume_after_failure(client, upload_url, exc, offset, resume_attempts, file_name)
            file_handle.seek(offset)
            continue
        except requests.exceptions.RequestException as exc:
            # Defense in depth only: `client.put()` wraps every `requests.RequestException` into
            # `GraphConnectionError` (caught above) at the transport boundary, so this branch
            # should be unreachable via `GraphClient` — kept in case a caller ever passes a raw
            # `requests`-based client instead.
            offset, resume_attempts = _resume_after_failure(client, upload_url, exc, offset, resume_attempts, file_name)
            file_handle.seek(offset)
            continue

        if response.status_code in (200, 201):
            return response.json()
        # 202 Accepted: more chunks expected (body carries `nextExpectedRanges`, which we don't
        # need on the happy path since we're already tracking the offset ourselves).
        offset = chunk_end + 1

    raise UploadSessionError(f"Upload session for '{file_name}' ended without a final driveItem response.")


def _resume_after_failure(
    client: GraphClient,
    upload_url: str,
    exc: Exception,
    offset: int,
    resume_attempts: int,
    file_name: str,
) -> tuple[int, int]:
    """Handle a transient chunk-PUT failure: resume via `nextExpectedRanges` or give up.

    Returns ``(resume_offset, resume_attempts + 1)`` to continue from. Raises
    :class:`_SessionExpired` if the status check itself finds the session gone (404), or
    :class:`~client.exceptions.UploadSessionError` (aborting the session first) once
    :data:`MAX_RESUME_ATTEMPTS` is exceeded.
    """
    # `GraphConnectionError` (connection error/timeout/DNS failure — no HTTP status code was ever
    # received) is always treated as transient here, same as 429/5xx; any other `GraphClientError`
    # subclass with a status code outside `_TRANSIENT_CHUNK_STATUSES` is not retried.
    if (
        isinstance(exc, GraphClientError)
        and not isinstance(exc, GraphConnectionError)
        and exc.status_code not in _TRANSIENT_CHUNK_STATUSES
    ):
        _abort_session(client, upload_url)
        raise UploadSessionError(
            f"Uploading '{file_name}' failed with a non-retryable error: {sanitize_exception_text(exc)}"
        ) from exc

    if resume_attempts >= MAX_RESUME_ATTEMPTS:
        _abort_session(client, upload_url)
        raise UploadSessionError(
            f"Uploading '{file_name}' failed after {MAX_RESUME_ATTEMPTS} resume attempts: "
            f"{sanitize_exception_text(exc)}"
        ) from exc

    logger.warning(
        "Chunk upload for '%s' failed transiently (%s); querying uploadUrl to resume.",
        file_name,
        sanitize_exception_text(exc),
    )
    try:
        status_response = client.get(upload_url, absolute=True, auth=False, retry=False)
    except GraphClientError as status_exc:
        if status_exc.status_code == 404:
            raise _SessionExpired from status_exc
        _abort_session(client, upload_url)
        raise UploadSessionError(
            f"Uploading '{file_name}' failed and the resume status check also failed: {status_exc}"
        ) from status_exc

    next_ranges = status_response.json().get("nextExpectedRanges") or []
    resumed_offset = int(next_ranges[0].split("-")[0]) if next_ranges else offset
    return resumed_offset, resume_attempts + 1


def _abort_session(client: GraphClient, upload_url: str) -> None:
    """Best-effort `DELETE uploadUrl` so an abandoned session leaves no partial file behind."""
    try:
        client.delete(upload_url, absolute=True, auth=False, retry=False)
    except (GraphClientError, requests.exceptions.RequestException):
        # `upload_url` is a pre-signed, credential-bearing URL — never log it (design spec §6
        # "logging"; a leaked `uploadUrl` grants unauthenticated write access to the session).
        logger.warning("Failed to delete the abandoned upload session.")


def _map_late_conflict(exc: GraphClientError, file_name: str, conflict_behavior: str) -> GraphClientError:
    if conflict_behavior == "fail":
        return FileAlreadyExistsError(
            f"'{file_name}' already exists at the destination and conflict behavior is 'fail'.",
            status_code=exc.status_code,
            error_code=exc.error_code,
        )
    # Graph has no server-side replace/rename recovery for a chunked upload session that hits a
    # name conflict on its final chunk (unlike `@microsoft.graph.sourceUrl`, which isn't supported
    # on OneDrive for Business / SharePoint Online anyway — design spec §4 capability inventory).
    # Surfacing a clear, retryable error is the documented simplification for this edge case.
    return UploadSessionError(
        f"'{file_name}' hit a late name conflict (HTTP 409 nameAlreadyExists) while finishing a "
        f"chunked upload with conflict_behavior='{conflict_behavior}'. Automatic replace/rename "
        "recovery is not supported for chunked uploads; please retry the job.",
        status_code=exc.status_code,
        error_code=exc.error_code,
    )


def _is_name_conflict(exc: GraphClientError) -> bool:
    return exc.error_code == "nameAlreadyExists" or exc.status_code == 409
