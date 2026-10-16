"""Microsoft Graph client package for keboola.wr-onedrive-v2.

Sub-modules:
    auth            OAuth token acquisition/rotation (``TokenProvider``, ``RefreshTokenProvider``).
    graph_client    Session-based Graph HTTP client with retry policy, User-Agent, paging.
    exceptions      Typed Graph error taxonomy raised by ``graph_client``.
    drives          Site/drive resolution for the SharePoint account type.
    uploader        Drive file upload logic (path validation, folder resolution, chunked upload).
    excel_writer    Excel worksheet write logic (workbook/session/worksheet resolution, batched writes).
    headers         Header-row normalization shared by the Excel writer.

No re-exports here (phase 8 audit) — every caller imports directly from the sub-module it needs
(e.g. ``from client.graph_client import GraphClient``), which is also what every module in this
package and ``component.py`` already does; the re-export list this docstring used to sit above
had drifted out of sync with that (missing ``drives``/``uploader``/``excel_writer`` exports) and
was never actually used anywhere.
"""
