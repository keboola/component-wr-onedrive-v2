"""Typed error taxonomy for Microsoft Graph API responses.

Design spec: ``docs/superpowers/specs/2026-08-17-wr-onedrive-v2-design.md`` §6 "Error mapping".

``component.py`` (a later task) catches the subclasses below at the call site and re-raises the
user-actionable ones as ``keboola.component.exceptions.UserException`` (exit 1); everything else
(a bare ``GraphClientError``, or any other exception) is left to propagate to exit 2. This module
deliberately has no dependency on ``keboola.component`` so the ``client`` package stays usable
and testable on its own (same rule as ``auth.AuthenticationError``).
"""


class GraphClientError(Exception):
    """Base class for every error raised while talking to Microsoft Graph.

    Carries the parsed HTTP status code and Graph ``error.code`` (when available) so callers can
    inspect them without re-parsing the response body.
    """

    def __init__(self, message: str, *, status_code: int | None = None, error_code: str | None = None):
        super().__init__(message)
        self.status_code = status_code
        self.error_code = error_code


class GraphPermissionError(GraphClientError):
    """401 (still unauthorized after a single token re-fetch) or 403.

    Maps to a "reauthorize / grant permissions" ``UserException``.
    """


class GraphNotFoundError(GraphClientError):
    """404 — the requested drive, item, site, workbook, or worksheet does not exist."""


class GraphQuotaExceededError(GraphClientError):
    """507 Insufficient Storage — the OneDrive/SharePoint quota is exhausted."""


class GraphBadRequestError(GraphClientError):
    """400 — a malformed request (invalid path, invalid Excel range, etc.)."""


class GraphRateLimitCapExceededError(GraphClientError):
    """The retry policy gave up instead of waiting past the configured total-wait cap.

    Raised either when a single ``Retry-After`` value alone would exceed the remaining budget,
    or when the cumulative wait across several retries exhausts it. Always user-facing: Graph is
    actively throttling the app, and hammering it further only burns more quota.
    """


class InvalidPathError(GraphClientError):
    """A destination path/placeholder/segment is invalid — raised client-side, pre-network.

    Covers unknown ``{...}`` placeholders in ``destination.folder_path``, Graph's reserved
    characters/names, and the 255-char segment / 400-char total path limits (design spec §4
    "Mechanics of the in-scope surface" — Path/naming). Never carries a ``status_code``: nothing
    was sent to Graph yet.
    """


class FileAlreadyExistsError(GraphClientError):
    """409 ``nameAlreadyExists`` with ``conflict_behavior='fail'`` — a clean, user-facing conflict.

    Raised by both the simple-PUT and upload-session paths in :mod:`client.uploader` when the
    target file already exists and the row is configured not to replace or rename it.
    """


class UploadSessionError(GraphClientError):
    """A chunked upload session could not be completed.

    Covers: resume attempts exhausted after repeated transient chunk failures, a session that
    expired (404) immediately after already being restarted once, and a late final-chunk 409
    ``nameAlreadyExists`` under ``conflict_behavior`` ``replace``/``rename`` (Graph provides no
    server-side replace/rename recovery for chunked uploads — the caller must retry the job).
    """
