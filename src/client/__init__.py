"""Microsoft Graph client package for keboola.wr-onedrive-v2.

Sub-modules:
    auth            OAuth token acquisition/rotation (``TokenProvider``, ``RefreshTokenProvider``).
    graph_client    Session-based Graph HTTP client with retry policy, User-Agent, paging.
    exceptions      Typed Graph error taxonomy raised by ``graph_client``.
    uploader        Drive file upload logic (path validation, folder resolution, chunked upload).
    excel_writer    Excel worksheet write logic (workbook/session/worksheet resolution, batched writes).
"""

from client.auth import AuthenticationError, RefreshTokenProvider, TokenProvider
from client.exceptions import (
    GraphBadRequestError,
    GraphClientError,
    GraphNotFoundError,
    GraphPermissionError,
    GraphQuotaExceededError,
    GraphRateLimitCapExceededError,
)
from client.graph_client import GraphClient

__all__ = [
    "AuthenticationError",
    "GraphBadRequestError",
    "GraphClient",
    "GraphClientError",
    "GraphNotFoundError",
    "GraphPermissionError",
    "GraphQuotaExceededError",
    "GraphRateLimitCapExceededError",
    "RefreshTokenProvider",
    "TokenProvider",
]
