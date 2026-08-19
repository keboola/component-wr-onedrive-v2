"""Golden tests: v1-parity sync-action output shapes (plan Task 8).

Fixtures under ``tests/fixtures/v1_parity/`` are ``expected-stdout``/``expected-stderr`` files
copied verbatim from ``keboola.wr-onedrive`` (v1, PHP)'s own datadir test suite
(``tests/datadir/<case>/``) — see the design spec §5 "byte-compatible v1 output shapes". Each
test here builds a `Component` with a mocked `GraphClient` whose ``get``/``post``/``put``
responses are shaped like real Microsoft Graph payloads, invokes the sync-action method
*directly* (bypassing `keboola.component.base.sync_action`'s stdout/exit(1) wrapper — the same
`action: "run"` trick ``tests/test_component.py`` already uses), and asserts the result against
the golden fixture.

v1's own placeholders (``%s``/``%a``/``%A``, from its PHPUnit-based ``assertStringMatchesFormat``
style datadir framework) stand in for values that are dynamic in v1's real test tenant (drive
ids, file ids, worksheet ids) — :func:`_golden_matches` and :func:`_stderr_matches_golden` below
treat them as wildcards rather than literal text.

**Known, intentional non-parity** (flagged per the task): ``get-worksheets-many-sheets`` (not
valid JSON — v1's own fixture embeds ``%A`` wildcards *inside* what would be object syntax) and
the ``get-worksheets-invalid-drive-id``/``invalid-file-id``/``invalid-file-type`` fixtures (v1's
generic Graph-error wording, which v2 deliberately does not replicate — Task 3 already
established v2's own error taxonomy/messages for non-sync-action-specific Graph errors) are not
exercised here. Every fixture that *is* exercised matches byte-for-byte (modulo v1's own
id/name wildcards).
"""

import json
import os
import re
from pathlib import Path
from unittest import mock
from unittest.mock import MagicMock

import pytest
from keboola.component.exceptions import UserException

from client.excel_writer import XLSX_MIME_TYPE
from client.exceptions import GraphBadRequestError, GraphNotFoundError
from client.graph_client import GraphClient
from component import Component

FIXTURES_DIR = Path(__file__).parent / "fixtures" / "v1_parity"


def _oauth_credentials() -> dict:
    return {
        "id": "oauth-1",
        "created": "2026-01-01",
        "appKey": "client-1",
        "#appSecret": "secret-1",
        "oauthVersion": "2.0",
        "#data": json.dumps({"refresh_token": "refresh-config"}),
    }


def _build_component(tmp_path, parameters: dict) -> Component:
    """Build a `Component` with `action: "run"` (see module docstring for why)."""
    data_dir = tmp_path / "data"
    (data_dir / "in" / "tables").mkdir(parents=True, exist_ok=True)
    (data_dir / "in" / "files").mkdir(parents=True, exist_ok=True)
    (data_dir / "out").mkdir(parents=True, exist_ok=True)
    config = {
        "parameters": parameters,
        "action": "run",
        "authorization": {"oauth_api": {"credentials": _oauth_credentials()}},
    }
    (data_dir / "config.json").write_text(json.dumps(config))
    with mock.patch.dict(os.environ, {"KBC_DATADIR": str(data_dir)}):
        return Component()


class _FakeResponse:
    def __init__(self, payload):
        self._payload = payload

    def json(self):
        return self._payload


def _route(rules, method: str, url: str):
    for matcher, response in rules:
        matched = matcher(url) if callable(matcher) else matcher == url
        if matched:
            if isinstance(response, BaseException):
                raise response
            return _FakeResponse(response)
    raise AssertionError(f"Unexpected {method} {url}")


def _fake_client(get_rules=(), post_rules=(), put_rules=()) -> MagicMock:
    """A `GraphClient` double routing on URL only (params/headers are ignored — irrelevant here).

    ``get_paged`` is the genuine ``GraphClient`` implementation bound onto this double (it only
    ever calls ``self.get(...)``, see ``client.graph_client``) — this lets URL-only routing via
    ``.get`` alone still work for callers that go through ``get_paged`` (e.g.
    ``client.excel_writer._list_worksheets``, phase 8 audit).
    """
    client = MagicMock()
    client.get.side_effect = lambda url, *a, **kw: _route(get_rules, "GET", url)
    client.post.side_effect = lambda url, *a, **kw: _route(post_rules, "POST", url)
    client.put.side_effect = lambda url, *a, **kw: _route(put_rules, "PUT", url)
    client.get_paged = GraphClient.get_paged.__get__(client)
    return client


def _xlsx_item(item_id: str, name: str, drive_id: str, parent_path: str | None = None) -> dict:
    item = {"id": item_id, "name": name, "file": {"mimeType": XLSX_MIME_TYPE}, "parentReference": {"driveId": drive_id}}
    if parent_path is not None:
        item["parentReference"]["path"] = parent_path
    return item


# --- golden-file comparison helpers ------------------------------------------------------------

_PLACEHOLDER_RE = re.compile(r"%([aAs])")


def _pattern_from_golden_string(golden: str) -> re.Pattern:
    """Compile a v1 `%s`/`%a`/`%A`-style format string into a regex (PHPUnit-ish semantics)."""
    parts = _PLACEHOLDER_RE.split(golden)
    pattern = []
    iterator = iter(parts)
    for literal in iterator:
        pattern.append(re.escape(literal))
        placeholder = next(iterator, None)
        if placeholder is None:
            break
        pattern.append(".*" if placeholder == "A" else ".+")
    return re.compile("^" + "".join(pattern) + "$", re.DOTALL)


def _golden_matches(actual, expected) -> bool:
    """Structural match: `expected` (parsed golden JSON) may use `%s`/`%a`/`%A` as string wildcards."""
    if isinstance(expected, dict):
        return (
            isinstance(actual, dict)
            and set(actual) == set(expected)
            and all(_golden_matches(actual[key], expected[key]) for key in expected)
        )
    if isinstance(expected, list):
        return (
            isinstance(actual, list)
            and len(actual) == len(expected)
            and all(_golden_matches(a, e) for a, e in zip(actual, expected, strict=True))
        )
    if expected is None:
        return actual is None
    if isinstance(expected, str):
        return isinstance(actual, str) and _pattern_from_golden_string(expected).match(actual) is not None
    return actual == expected


def _assert_matches_golden_stdout(actual: dict, fixture_name: str) -> None:
    expected = json.loads((FIXTURES_DIR / fixture_name / "expected-stdout").read_text())
    assert _golden_matches(actual, expected), f"{actual!r} does not match golden {fixture_name}: {expected!r}"


def _assert_matches_golden_stderr(message: str, fixture_name: str) -> None:
    expected = (FIXTURES_DIR / fixture_name / "expected-stderr").read_text().strip()
    assert _pattern_from_golden_string(expected).match(message) is not None, (
        f"{message!r} does not match golden {fixture_name} stderr: {expected!r}"
    )


# --- search --------------------------------------------------------------------------------


class TestSearchGolden:
    def test_drive_path_found(self, tmp_path):
        parameters = {
            "account": {"account_type": "onedrive_for_business", "tenant_id": "tenant-1"},
            "workbook": {"path": "drive://drive-id-1/__wr-onedrive-test-folder/valid/one_sheet.xlsx"},
        }
        comp = _build_component(tmp_path, parameters)
        client = _fake_client(
            get_rules=[
                (
                    "/drives/drive-id-1/root:/__wr-onedrive-test-folder/valid/one_sheet.xlsx",
                    _xlsx_item(
                        "file-id-1",
                        "one_sheet.xlsx",
                        "drive-id-1",
                        parent_path="/drives/drive-id-1/root:/__wr-onedrive-test-folder/valid",
                    ),
                )
            ]
        )

        with mock.patch("component.GraphClient", return_value=client):
            result = comp.search()

        _assert_matches_golden_stdout({"file": result["file"]}, "search-file-path-drive")

    def test_root_relative_path_in_me_drive_found(self, tmp_path):
        parameters = {
            "account": {"account_type": "onedrive_for_business", "tenant_id": "tenant-1"},
            "workbook": {"path": "/__wr-onedrive-test-folder/valid/one_sheet.xlsx"},
        }
        comp = _build_component(tmp_path, parameters)
        client = _fake_client(
            get_rules=[
                ("/me/drive", {"id": "me-drive-1"}),
                (
                    "/drives/me-drive-1/root:/__wr-onedrive-test-folder/valid/one_sheet.xlsx",
                    _xlsx_item(
                        "file-id-1",
                        "one_sheet.xlsx",
                        "me-drive-1",
                        parent_path="/drive/root:/__wr-onedrive-test-folder/valid",
                    ),
                ),
            ]
        )

        with mock.patch("component.GraphClient", return_value=client):
            result = comp.search()

        _assert_matches_golden_stdout(result, "search-file-path-drive-me")

    def test_site_path_found(self, tmp_path):
        parameters = {
            "account": {"account_type": "onedrive_for_business", "tenant_id": "tenant-1"},
            "workbook": {"path": "site://Test Site/__wr-onedrive-test-folder/valid/one_sheet.xlsx"},
        }
        comp = _build_component(tmp_path, parameters)
        client = _fake_client(
            get_rules=[
                ("/sites", {"value": [{"id": "site-id-1", "name": "Test Site"}]}),
                ("/sites/site-id-1/drive", {"id": "site-drive-1"}),
                (
                    "/drives/site-drive-1/root:/__wr-onedrive-test-folder/valid/one_sheet.xlsx",
                    _xlsx_item(
                        "file-id-1",
                        "one_sheet.xlsx",
                        "site-drive-1",
                        parent_path="/drives/site-drive-1/root:/__wr-onedrive-test-folder/valid",
                    ),
                ),
            ]
        )

        with mock.patch("component.GraphClient", return_value=client):
            result = comp.search()

        _assert_matches_golden_stdout(result, "search-file-path-drive-site")
        assert result["file"]["path"] == "sites/Test Site/__wr-onedrive-test-folder/valid"

    def test_path_not_found_returns_null_file(self, tmp_path):
        parameters = {
            "account": {"account_type": "onedrive_for_business", "tenant_id": "tenant-1"},
            "workbook": {"path": "/not/found/file.xlsx"},
        }
        comp = _build_component(tmp_path, parameters)
        client = _fake_client(
            get_rules=[
                ("/me/drive", {"id": "me-drive-1"}),
                ("/drives/me-drive-1/root:/not/found/file.xlsx", GraphNotFoundError("not found")),
            ]
        )

        with mock.patch("component.GraphClient", return_value=client):
            result = comp.search()

        _assert_matches_golden_stdout(result, "search-file-path-not-found")
        assert result == {"file": None}

    def test_sharing_link_found(self, tmp_path):
        parameters = {
            "account": {"account_type": "onedrive_for_business", "tenant_id": "tenant-1"},
            "workbook": {"path": "https://contoso.sharepoint.com/:x:/r/sites/x/share-token"},
        }
        comp = _build_component(tmp_path, parameters)
        client = _fake_client(
            get_rules=[
                (
                    lambda url: url.startswith("/shares/") and url.endswith("/driveItem"),
                    _xlsx_item(
                        "file-id-1",
                        "one_sheet.xlsx",
                        "drive-id-1",
                        parent_path="/drives/drive-id-1/root:/__wr-onedrive-test-folder/valid",
                    ),
                )
            ]
        )

        with mock.patch("component.GraphClient", return_value=client):
            result = comp.search()

        _assert_matches_golden_stdout(result, "search-sharing-link")

    def test_sharing_link_not_found_raises_user_exception_matching_v1_wording(self, tmp_path):
        link = "https://example.com/a/b/c/not/found"
        parameters = {
            "account": {"account_type": "onedrive_for_business", "tenant_id": "tenant-1"},
            "workbook": {"path": link},
        }
        comp = _build_component(tmp_path, parameters)
        client = _fake_client(
            get_rules=[
                (
                    lambda url: url.startswith("/shares/") and url.endswith("/driveItem"),
                    GraphBadRequestError("access denied", status_code=400),
                )
            ]
        )

        with mock.patch("component.GraphClient", return_value=client), pytest.raises(UserException) as exc_info:
            comp.search()

        _assert_matches_golden_stderr(str(exc_info.value), "search-sharing-link-not-found")
        # This one has no `%a`/`%s` wildcard at all in v1's fixture — assert the exact byte match.
        assert str(exc_info.value) == (
            'The sharing link "https://example.com/a/b/c/not/fo..." not exists, or you do not '
            "have permission to access it."
        )


# --- createWorkbook --------------------------------------------------------------------------


class TestCreateWorkbookGolden:
    def test_creates_missing_workbook(self, tmp_path):
        parameters = {
            "account": {"account_type": "onedrive_for_business", "tenant_id": "tenant-1"},
            "workbook": {"path": "drive://drive-id-1/newfolder/newfile.xlsx"},
        }
        comp = _build_component(tmp_path, parameters)
        client = _fake_client(
            get_rules=[
                ("/drives/drive-id-1/root:/newfolder/newfile.xlsx", GraphNotFoundError("not found")),
                ("/drives/drive-id-1/root:/newfolder", GraphNotFoundError("not found")),
            ],
            post_rules=[("/drives/drive-id-1/root/children", {"id": "folder-id-1"})],
            put_rules=[
                (
                    lambda url: url == "/drives/drive-id-1/items/folder-id-1:/newfile.xlsx:/content",
                    {"id": "created-file-id"},
                )
            ],
        )

        with mock.patch("component.GraphClient", return_value=client):
            result = comp.create_workbook()

        _assert_matches_golden_stdout(result, "create-workbook")
        assert result == {"file": {"driveId": "drive-id-1", "fileId": "created-file-id"}}

    def test_existing_workbook_raises_user_exception(self, tmp_path):
        path = "drive://drive-id-1/__wr-onedrive-test-folder/valid/one_sheet.xlsx"
        parameters = {
            "account": {"account_type": "onedrive_for_business", "tenant_id": "tenant-1"},
            "workbook": {"path": path},
        }
        comp = _build_component(tmp_path, parameters)
        client = _fake_client(
            get_rules=[
                (
                    "/drives/drive-id-1/root:/__wr-onedrive-test-folder/valid/one_sheet.xlsx",
                    _xlsx_item("file-id-1", "one_sheet.xlsx", "drive-id-1"),
                )
            ]
        )

        with mock.patch("component.GraphClient", return_value=client), pytest.raises(UserException) as exc_info:
            comp.create_workbook()

        assert str(exc_info.value) == f'Workbook "{path}" already exists.'
        _assert_matches_golden_stderr(str(exc_info.value), "create-workbook-already-exists")


# --- createWorksheet -------------------------------------------------------------------------


class TestCreateWorksheetGolden:
    def test_creates_missing_worksheet(self, tmp_path):
        parameters = {
            "account": {"account_type": "onedrive_for_business", "tenant_id": "tenant-1"},
            "workbook": {"drive_id": "drive-id-1", "file_id": "file-id-1"},
            "worksheet": {"name": "New Sheet"},
        }
        comp = _build_component(tmp_path, parameters)
        client = _fake_client(
            get_rules=[
                ("/drives/drive-id-1/items/file-id-1", _xlsx_item("file-id-1", "one_sheet.xlsx", "drive-id-1")),
                (
                    "/drives/drive-id-1/items/file-id-1/workbook/worksheets",
                    {"value": [{"id": "sheet-1", "name": "Only One Sheet", "position": 0, "visibility": "Visible"}]},
                ),
            ],
            post_rules=[
                (
                    "/drives/drive-id-1/items/file-id-1/workbook/worksheets/add",
                    {"id": "sheet-2", "name": "New Sheet"},
                )
            ],
        )

        with mock.patch("component.GraphClient", return_value=client):
            result = comp.create_worksheet()

        _assert_matches_golden_stdout(result, "create-worksheet")
        assert result == {"worksheet": {"driveId": "drive-id-1", "fileId": "file-id-1", "worksheetId": "sheet-2"}}

    def test_existing_worksheet_raises_user_exception(self, tmp_path):
        parameters = {
            "account": {"account_type": "onedrive_for_business", "tenant_id": "tenant-1"},
            "workbook": {"drive_id": "drive-id-1", "file_id": "file-id-1"},
            "worksheet": {"name": "Only One Sheet"},
        }
        comp = _build_component(tmp_path, parameters)
        client = _fake_client(
            get_rules=[
                ("/drives/drive-id-1/items/file-id-1", _xlsx_item("file-id-1", "one_sheet.xlsx", "drive-id-1")),
                (
                    "/drives/drive-id-1/items/file-id-1/workbook/worksheets",
                    {"value": [{"id": "sheet-1", "name": "Only One Sheet", "position": 0, "visibility": "Visible"}]},
                ),
            ]
        )

        with mock.patch("component.GraphClient", return_value=client), pytest.raises(UserException) as exc_info:
            comp.create_worksheet()

        assert str(exc_info.value) == 'Worksheet "Only One Sheet" already exists.'
        _assert_matches_golden_stderr(str(exc_info.value), "create-worksheet-already-exists")


# --- getWorksheets ---------------------------------------------------------------------------


def _header_route(drive_id: str, file_id: str, worksheet_id: str, cells: list[str] | None) -> tuple:
    url = f"/drives/{drive_id}/items/{file_id}/workbook/worksheets/{worksheet_id}/range/usedRange(valuesOnly=true)/row(row=0)"
    body = {"address": "Sheet1!A1:A1", "text": [cells] if cells is not None else [[""]]}
    return url, body


class TestGetWorksheetsGolden:
    def test_one_sheet_with_header(self, tmp_path):
        parameters = {
            "account": {"account_type": "onedrive_for_business", "tenant_id": "tenant-1"},
            "workbook": {"drive_id": "drive-id-1", "file_id": "file-id-1"},
        }
        comp = _build_component(tmp_path, parameters)
        client = _fake_client(
            get_rules=[
                ("/drives/drive-id-1/items/file-id-1", _xlsx_item("file-id-1", "one_sheet.xlsx", "drive-id-1")),
                (
                    "/drives/drive-id-1/items/file-id-1/workbook/worksheets",
                    {
                        "value": [
                            {"id": "sheet-1", "name": "Only One Sheet", "position": 0, "visibility": "Visible"}
                        ]
                    },
                ),
                _header_route("drive-id-1", "file-id-1", "sheet-1", ["Col 1", "Col 2", "Col 3"]),
            ]
        )

        with mock.patch("component.GraphClient", return_value=client):
            result = comp.get_worksheets()

        _assert_matches_golden_stdout(result, "get-worksheets-one-sheet")

    def test_empty_sheet_has_no_header(self, tmp_path):
        parameters = {
            "account": {"account_type": "onedrive_for_business", "tenant_id": "tenant-1"},
            "workbook": {"drive_id": "drive-id-1", "file_id": "file-id-1"},
        }
        comp = _build_component(tmp_path, parameters)
        client = _fake_client(
            get_rules=[
                ("/drives/drive-id-1/items/file-id-1", _xlsx_item("file-id-1", "empty.xlsx", "drive-id-1")),
                (
                    "/drives/drive-id-1/items/file-id-1/workbook/worksheets",
                    {"value": [{"id": "sheet-1", "name": "Sheet1", "position": 0, "visibility": "Visible"}]},
                ),
                _header_route("drive-id-1", "file-id-1", "sheet-1", [""]),
            ]
        )

        with mock.patch("component.GraphClient", return_value=client):
            result = comp.get_worksheets()

        _assert_matches_golden_stdout(result, "get-worksheets-empty-sheet")

    def test_by_path(self, tmp_path):
        parameters = {
            "account": {"account_type": "onedrive_for_business", "tenant_id": "tenant-1"},
            "workbook": {"path": "drive://drive-id-1/__wr-onedrive-test-folder/valid/one_sheet.xlsx"},
        }
        comp = _build_component(tmp_path, parameters)
        client = _fake_client(
            get_rules=[
                (
                    "/drives/drive-id-1/root:/__wr-onedrive-test-folder/valid/one_sheet.xlsx",
                    _xlsx_item("file-id-1", "one_sheet.xlsx", "drive-id-1"),
                ),
                (
                    "/drives/drive-id-1/items/file-id-1/workbook/worksheets",
                    {
                        "value": [
                            {"id": "sheet-1", "name": "Only One Sheet", "position": 0, "visibility": "Visible"}
                        ]
                    },
                ),
                _header_route("drive-id-1", "file-id-1", "sheet-1", ["Col 1", "Col 2", "Col 3"]),
            ]
        )

        with mock.patch("component.GraphClient", return_value=client):
            result = comp.get_worksheets()

        _assert_matches_golden_stdout(result, "get-worksheets-by-path")

    def test_hidden_sheet_among_several(self, tmp_path):
        parameters = {
            "account": {"account_type": "onedrive_for_business", "tenant_id": "tenant-1"},
            "workbook": {"drive_id": "drive-id-1", "file_id": "file-id-1"},
        }
        comp = _build_component(tmp_path, parameters)
        worksheets = {
            "value": [
                {"id": "sheet-0", "name": "Sheet1", "position": 0, "visibility": "Visible"},
                {"id": "sheet-1", "name": "Zošit 2", "position": 1, "visibility": "Visible"},
                {"id": "sheet-2", "name": "Hidden Sheet 3", "position": 2, "visibility": "Hidden"},
                {"id": "sheet-3", "name": "sheet=4", "position": 3, "visibility": "Visible"},
            ]
        }
        client = _fake_client(
            get_rules=[
                ("/drives/drive-id-1/items/file-id-1", _xlsx_item("file-id-1", "hidden.xlsx", "drive-id-1")),
                ("/drives/drive-id-1/items/file-id-1/workbook/worksheets", worksheets),
                _header_route("drive-id-1", "file-id-1", "sheet-0", ["Col 1", "Col 2", "Col 3"]),
                _header_route("drive-id-1", "file-id-1", "sheet-1", ["Col 1", "Col 2", "Col 3"]),
                _header_route("drive-id-1", "file-id-1", "sheet-2", ["Col 4", "Col 5", "Col 6"]),
                _header_route("drive-id-1", "file-id-1", "sheet-3", ["Col 1", "Col 2", "Col 3"]),
            ]
        )

        with mock.patch("component.GraphClient", return_value=client):
            result = comp.get_worksheets()

        _assert_matches_golden_stdout(result, "get-worksheets-hidden-sheet")
        titles = [sheet["title"] for sheet in result["worksheets"]]
        assert titles == ["Sheet1", "Zošit 2", "Hidden Sheet 3 (hidden)", "sheet=4"]

    def test_ids_not_found_raises_exact_v1_message(self, tmp_path):
        parameters = {
            "account": {"account_type": "onedrive_for_business", "tenant_id": "tenant-1"},
            "workbook": {"drive_id": "wrong-drive-id", "file_id": "file-id-1"},
        }
        comp = _build_component(tmp_path, parameters)
        client = _fake_client(
            get_rules=[("/drives/wrong-drive-id/items/file-id-1", GraphNotFoundError("not found"))]
        )

        with mock.patch("component.GraphClient", return_value=client), pytest.raises(UserException) as exc_info:
            comp.get_worksheets()

        assert str(exc_info.value) == "Configured workbook XLSX file not found."
        _assert_matches_golden_stderr(str(exc_info.value), "get-worksheets-not-found-drive-id")
