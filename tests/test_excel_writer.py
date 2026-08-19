import base64
import csv
import logging
from unittest.mock import MagicMock

import pytest

from client.excel_writer import _SESSION_POLL_MAX_ATTEMPTS as SESSION_POLL_MAX_ATTEMPTS
from client.excel_writer import (
    RangeAddress,
    column_int_to_str,
    column_str_to_int,
    parse_range_address,
    resolve_workbook,
    resolve_worksheet,
    workbook_session,
    write_table,
)
from client.exceptions import (
    GraphBadRequestError,
    GraphClientError,
    GraphNotFoundError,
    GraphPermissionError,
    InvalidWorkbookFormatError,
    InvalidWorkbookPathError,
    MultipleSitesFoundError,
    WorksheetNotFoundError,
)
from client.graph_client import GraphClient
from configuration import Account, AccountType, Workbook, Worksheet


def _mock_graph_client() -> MagicMock:
    """A `MagicMock` `GraphClient` double with a *real* ``get_paged`` bound on top.

    ``get_paged`` only ever calls ``self.get(...)`` (see ``client.graph_client``) — binding the
    genuine implementation lets these tests keep faking just ``.get()`` even now that
    ``_list_worksheets`` delegates to ``get_paged`` instead of hand-rolling its own pagination
    loop (phase 8 audit).
    """
    client = MagicMock()
    client.get_paged = GraphClient.get_paged.__get__(client)
    return client

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _response(status_code=200, json_body=None, headers=None) -> MagicMock:
    response = MagicMock()
    response.status_code = status_code
    response.headers = headers or {}
    if json_body is not None:
        response.json.return_value = json_body
    return response


def _xlsx_item(item_id="file-1", drive_id=None) -> dict:
    body = {"id": item_id, "name": "book.xlsx", "file": {"mimeType": XLSX_MIME}}
    if drive_id is not None:
        body["parentReference"] = {"driveId": drive_id}
    return body


def _write_csv(path, rows) -> str:
    with open(path, "w", newline="", encoding="utf-8") as fh:
        writer = csv.writer(fh)
        for row in rows:
            writer.writerow(row)
    return str(path)


PRIVATE_ACCOUNT = Account(account_type=AccountType.PRIVATE_ONEDRIVE)
BUSINESS_ACCOUNT = Account(
    account_type=AccountType.SHAREPOINT,
    tenant_id="tenant-1",
    site_url="https://contoso.sharepoint.com/sites/marketing",
)


# ---------------------------------------------------------------------------------------------
# Column / range address math
# ---------------------------------------------------------------------------------------------


class TestColumnMath:
    @pytest.mark.parametrize(
        ("column", "letters"),
        [(1, "A"), (26, "Z"), (27, "AA"), (52, "AZ"), (53, "BA")],
    )
    def test_round_trips(self, column, letters):
        assert column_int_to_str(column) == letters
        assert column_str_to_int(letters) == column

    def test_lowercase_letters_accepted(self):
        assert column_str_to_int("aa") == 27

    def test_parses_a_sheet_qualified_range(self):
        result = parse_range_address("Sheet1!C9:F20")

        assert result == RangeAddress(sheet="Sheet1", start_col=3, start_row=9, end_col=6, end_row=20)

    def test_parses_a_single_cell_range(self):
        result = parse_range_address("Sheet1!A1")

        assert result == RangeAddress(sheet="Sheet1", start_col=1, start_row=1, end_col=1, end_row=1)

    def test_parses_a_bare_range_without_sheet_prefix(self):
        result = parse_range_address("A1:B2")

        assert result.sheet == ""
        assert (result.start_col, result.start_row, result.end_col, result.end_row) == (1, 1, 2, 2)

    def test_invalid_address_raises(self):
        with pytest.raises(ValueError, match="Cannot parse range address"):
            parse_range_address("not-a-range")


# ---------------------------------------------------------------------------------------------
# Workbook resolution — ids mode
# ---------------------------------------------------------------------------------------------


class TestResolveWorkbookByIds:
    def test_success_returns_ids_and_created_false(self):
        client = MagicMock()
        client.get.return_value = _response(json_body=_xlsx_item())
        workbook = Workbook(drive_id="drive-1", file_id="file-1")

        drive_id, file_id, created = resolve_workbook(client, PRIVATE_ACCOUNT, workbook)

        assert (drive_id, file_id, created) == ("drive-1", "file-1", False)
        client.get.assert_called_once_with(
            "/drives/drive-1/items/file-1", params={"$select": "id,name,file,parentReference"}
        )

    def test_missing_file_raises_user_facing_not_found_and_never_creates(self):
        client = MagicMock()
        client.get.side_effect = GraphNotFoundError("boom", status_code=404, error_code="itemNotFound")
        workbook = Workbook(drive_id="drive-1", file_id="file-1")

        with pytest.raises(GraphNotFoundError, match="Configured workbook XLSX file not found."):
            resolve_workbook(client, PRIVATE_ACCOUNT, workbook)

        client.post.assert_not_called()

    def test_wrong_mime_type_raises_with_message(self):
        client = MagicMock()
        client.get.return_value = _response(json_body={"id": "file-1", "file": {"mimeType": "text/plain"}})
        workbook = Workbook(drive_id="drive-1", file_id="file-1")

        with pytest.raises(InvalidWorkbookFormatError) as exc_info:
            resolve_workbook(client, PRIVATE_ACCOUNT, workbook)

        assert 'File is not in the "XLSX" Excel format. Mime type: "text/plain"' == str(exc_info.value)


# ---------------------------------------------------------------------------------------------
# Workbook resolution — path forms
# ---------------------------------------------------------------------------------------------


class TestResolveWorkbookPathForms:
    def test_root_relative_path_resolves_via_me_drive(self):
        client = MagicMock()
        client.get.side_effect = [
            _response(json_body={"id": "my-drive"}),
            _response(json_body=_xlsx_item()),
        ]
        workbook = Workbook(path="/reports/book.xlsx")

        drive_id, file_id, created = resolve_workbook(client, PRIVATE_ACCOUNT, workbook)

        assert (drive_id, file_id, created) == ("my-drive", "file-1", False)
        assert client.get.call_args_list[0].args[0] == "/me/drive"
        assert client.get.call_args_list[1].args[0] == "/drives/my-drive/root:/reports/book.xlsx"

    def test_drive_path_url_decodes_the_drive_id_segment(self):
        client = MagicMock()
        client.get.return_value = _response(json_body=_xlsx_item())
        workbook = Workbook(path="drive://b%21AbCdEf1234567890/reports/book.xlsx")

        drive_id, _file_id, _created = resolve_workbook(client, BUSINESS_ACCOUNT, workbook)

        assert drive_id == "b!AbCdEf1234567890"
        assert client.get.call_args_list[0].args[0] == "/drives/b!AbCdEf1234567890/root:/reports/book.xlsx"

    def test_site_path_requires_exactly_one_site_hit(self):
        client = MagicMock()
        client.get.side_effect = [
            _response(json_body={"value": [{"id": "site-1", "name": "Marketing"}]}),
            _response(json_body={"id": "site-drive-1"}),
            _response(json_body=_xlsx_item()),
        ]
        workbook = Workbook(path="site://Marketing/reports/book.xlsx")

        drive_id, file_id, _created = resolve_workbook(client, BUSINESS_ACCOUNT, workbook)

        assert (drive_id, file_id) == ("site-drive-1", "file-1")
        assert client.get.call_args_list[0].args[0] == "/sites"
        assert client.get.call_args_list[0].kwargs["params"] == {"search": "Marketing", "$select": "id,name"}
        assert client.get.call_args_list[1].args[0] == "/sites/site-1/drive"
        assert client.get.call_args_list[2].args[0] == "/drives/site-drive-1/root:/reports/book.xlsx"

    def test_site_path_zero_hits_raises_not_found(self):
        client = MagicMock()
        client.get.return_value = _response(json_body={"value": []})
        workbook = Workbook(path="site://Nowhere/book.xlsx")

        with pytest.raises(GraphNotFoundError, match="not found"):
            resolve_workbook(client, BUSINESS_ACCOUNT, workbook)

    def test_site_path_multiple_hits_raises_multiple_sites_found(self):
        client = MagicMock()
        client.get.return_value = _response(
            json_body={"value": [{"id": "s1", "name": "Marketing"}, {"id": "s2", "name": "Marketing US"}]}
        )
        workbook = Workbook(path="site://Marketing/book.xlsx")

        with pytest.raises(MultipleSitesFoundError, match="Multiple sites found"):
            resolve_workbook(client, BUSINESS_ACCOUNT, workbook)

    def test_sharing_link_encodes_to_expected_u_bang_token(self):
        client = MagicMock()
        client.get.return_value = _response(json_body=_xlsx_item(drive_id="share-drive-1"))
        link = "https://contoso.sharepoint.com/:x:/g/somelongtoken=="
        expected_token = base64.b64encode(link.encode("utf-8")).decode("ascii").rstrip("=")
        expected_token = expected_token.replace("+", "-").replace("/", "_")
        workbook = Workbook(path=link)

        drive_id, file_id, created = resolve_workbook(client, BUSINESS_ACCOUNT, workbook)

        assert (drive_id, file_id, created) == ("share-drive-1", "file-1", False)
        called_url = client.get.call_args_list[0].args[0]
        assert called_url == f"/shares/u!{expected_token}/driveItem"

    @pytest.mark.parametrize(
        "raised",
        [
            GraphNotFoundError("nope", status_code=404),
            GraphBadRequestError("bad token", status_code=400),
            GraphPermissionError("forbidden", status_code=403),
        ],
    )
    def test_sharing_link_invalid_or_inaccessible_raises_user_facing_error(self, raised):
        client = MagicMock()
        client.get.side_effect = raised
        link = "https://contoso.sharepoint.com/:x:/g/badtoken"
        workbook = Workbook(path=link)

        with pytest.raises(GraphNotFoundError, match="not exists, or you do not have permission"):
            resolve_workbook(client, BUSINESS_ACCOUNT, workbook)

    def test_unexpected_path_format_raises_without_any_network_call(self):
        client = MagicMock()
        workbook = Workbook(path="relative/no/leading/slash.xlsx")

        with pytest.raises(InvalidWorkbookPathError, match="Unexpected path format"):
            resolve_workbook(client, PRIVATE_ACCOUNT, workbook)

        client.get.assert_not_called()


# ---------------------------------------------------------------------------------------------
# Create-when-missing
# ---------------------------------------------------------------------------------------------


class TestCreateWhenMissing:
    def test_ids_mode_never_creates_on_404(self):
        client = MagicMock()
        client.get.side_effect = GraphNotFoundError("boom", status_code=404)
        workbook = Workbook(drive_id="drive-1", file_id="file-1")

        with pytest.raises(GraphNotFoundError):
            resolve_workbook(client, PRIVATE_ACCOUNT, workbook)

        client.post.assert_not_called()
        client.put.assert_not_called()

    def test_path_mode_creates_a_minimal_workbook_on_404(self, monkeypatch, caplog):
        client = MagicMock()
        client.get.side_effect = [
            _response(json_body={"id": "my-drive"}),
            GraphNotFoundError("not found", status_code=404),
        ]

        ensure_folder_mock = MagicMock(return_value="parent-folder-id")
        upload_file_mock = MagicMock(return_value={"id": "new-file-id"})
        monkeypatch.setattr("client.excel_writer.ensure_folder", ensure_folder_mock)
        monkeypatch.setattr("client.excel_writer.upload_file", upload_file_mock)

        workbook = Workbook(path="/reports/book.xlsx")
        with caplog.at_level(logging.INFO, logger="client.excel_writer"):
            drive_id, file_id, created = resolve_workbook(client, PRIVATE_ACCOUNT, workbook)

        assert (drive_id, file_id, created) == ("my-drive", "new-file-id", True)
        ensure_folder_mock.assert_called_once_with(client, "my-drive", "reports")
        call = upload_file_mock.call_args
        assert call.args[0] is client
        assert call.args[1] == "my-drive"
        assert call.args[2] == "parent-folder-id"
        assert call.args[4] == "book.xlsx"
        assert call.kwargs["conflict_behavior"] == "replace"
        assert 'New workbook "reports/book.xlsx" created.' in caplog.text

    def test_empty_workbook_fixture_is_a_valid_xlsx_with_sheet_named_new(self):
        openpyxl = pytest.importorskip("openpyxl")
        from importlib import resources

        with resources.as_file(resources.files("client.fixtures").joinpath("empty.xlsx")) as fixture_path:
            wb = openpyxl.load_workbook(fixture_path)

        assert wb.sheetnames == ["New"]


# ---------------------------------------------------------------------------------------------
# Session lifecycle
# ---------------------------------------------------------------------------------------------


class TestWorkbookSession:
    def test_201_response_yields_session_id_and_closes_it(self):
        client = MagicMock()
        client.post.return_value = _response(201, json_body={"id": "sess-1"})

        with workbook_session(client, "drive-1", "file-1") as session_id:
            assert session_id == "sess-1"

        assert client.post.call_count == 2
        create_call, close_call = client.post.call_args_list
        assert create_call.args[0] == "/drives/drive-1/items/file-1/workbook/createSession"
        assert create_call.kwargs["json"] == {"persistChanges": True}
        assert create_call.kwargs["headers"] == {"Prefer": "respond-async"}
        assert close_call.args[0] == "/drives/drive-1/items/file-1/workbook/closeSession"
        assert close_call.kwargs["headers"] == {"workbook-session-id": "sess-1"}

    def test_202_response_polls_location_until_succeeded(self, monkeypatch):
        client = MagicMock()
        client.post.return_value = _response(202, headers={"Location": "https://graph/status/1"})
        client.get.side_effect = [
            _response(json_body={"status": "running"}),
            _response(json_body={"status": "succeeded", "resourceLocation": "https://graph/result/1"}),
            _response(json_body={"id": "sess-polled"}),
        ]
        sleeps = []
        monkeypatch.setattr("client.excel_writer.time.sleep", lambda seconds: sleeps.append(seconds))

        with workbook_session(client, "drive-1", "file-1") as session_id:
            assert session_id == "sess-polled"

        assert sleeps == [2.0]
        assert client.get.call_args_list[0].args[0] == "https://graph/status/1"
        assert client.get.call_args_list[0].kwargs["absolute"] is True
        assert client.get.call_args_list[2].args[0] == "https://graph/result/1"

    def test_poll_deadline_reached_returns_none_sessionless(self, monkeypatch, caplog):
        """MINOR-4: a session stuck "running" forever must not poll forever — it falls back to
        sessionless (like any other session-creation failure) once the poll deadline is hit."""
        client = MagicMock()
        client.post.return_value = _response(202, headers={"Location": "https://graph/status/1"})
        client.get.return_value = _response(json_body={"status": "running"})
        sleeps = []
        monkeypatch.setattr("client.excel_writer.time.sleep", lambda seconds: sleeps.append(seconds))

        with caplog.at_level(logging.WARNING, logger="client.excel_writer"), workbook_session(
            client, "drive-1", "file-1"
        ) as session_id:
            assert session_id is None

        assert len(sleeps) == SESSION_POLL_MAX_ATTEMPTS
        assert client.get.call_count == SESSION_POLL_MAX_ATTEMPTS
        assert "workbook session could not be created" in caplog.text

    def test_creation_failure_yields_none_sessionless(self):
        client = MagicMock()
        client.post.side_effect = GraphClientError("boom", status_code=500)

        with workbook_session(client, "drive-1", "file-1") as session_id:
            assert session_id is None

        client.post.assert_called_once()  # only the failed createSession attempt, no closeSession

    def test_poll_failure_yields_none(self, monkeypatch):
        client = MagicMock()
        client.post.return_value = _response(202, headers={"Location": "https://graph/status/1"})
        client.get.side_effect = GraphClientError("boom", status_code=500)
        monkeypatch.setattr("client.excel_writer.time.sleep", lambda seconds: None)

        with workbook_session(client, "drive-1", "file-1") as session_id:
            assert session_id is None

    def test_close_session_runs_in_finally_even_if_the_block_raises(self):
        client = MagicMock()
        client.post.return_value = _response(201, json_body={"id": "sess-1"})

        with pytest.raises(ValueError, match="boom"), workbook_session(client, "drive-1", "file-1") as session_id:
            assert session_id == "sess-1"
            raise ValueError("boom")

        assert client.post.call_count == 2
        assert client.post.call_args_list[1].args[0] == "/drives/drive-1/items/file-1/workbook/closeSession"

    def test_close_session_exceptions_are_swallowed(self, caplog):
        client = MagicMock()
        client.post.side_effect = [_response(201, json_body={"id": "sess-1"}), GraphClientError("close failed")]

        with (
            caplog.at_level(logging.DEBUG, logger="client.excel_writer"),
            workbook_session(client, "drive-1", "file-1") as session_id,
        ):
            assert session_id == "sess-1"

        assert "Failed to close workbook session" in caplog.text


# ---------------------------------------------------------------------------------------------
# Worksheet resolution
# ---------------------------------------------------------------------------------------------


def _worksheets_response(items):
    return _response(json_body={"value": items})


class TestResolveWorksheet:
    def test_resolve_by_id_success(self):
        client = _mock_graph_client()
        client.get.return_value = _worksheets_response([{"id": "id-1", "name": "Sheet1", "position": 0}])
        worksheet = Worksheet(id="id-1")

        worksheet_id, is_new, actual_name = resolve_worksheet(client, "drive-1", "file-1", worksheet, None)

        assert (worksheet_id, is_new, actual_name) == ("id-1", False, "Sheet1")
        assert client.get.call_args.kwargs["headers"] == {}

    def test_resolve_by_id_not_found_raises(self):
        client = _mock_graph_client()
        client.get.return_value = _worksheets_response([{"id": "other", "name": "Sheet1", "position": 0}])
        worksheet = Worksheet(id="missing")

        with pytest.raises(WorksheetNotFoundError, match='id "missing"'):
            resolve_worksheet(client, "drive-1", "file-1", worksheet, None)

    def test_resolve_by_position_including_hidden(self):
        client = _mock_graph_client()
        client.get.return_value = _worksheets_response(
            [
                {"id": "id-1", "name": "Visible", "position": 0, "visibility": "Visible"},
                {"id": "id-2", "name": "Hidden", "position": 1, "visibility": "Hidden"},
            ]
        )
        worksheet = Worksheet(position=1)

        worksheet_id, _is_new, actual_name = resolve_worksheet(client, "drive-1", "file-1", worksheet, None)

        assert (worksheet_id, actual_name) == ("id-2", "Hidden")

    def test_resolve_by_position_not_found_raises(self):
        client = _mock_graph_client()
        client.get.return_value = _worksheets_response([{"id": "id-1", "name": "Sheet1", "position": 0}])
        worksheet = Worksheet(position=5)

        with pytest.raises(WorksheetNotFoundError, match="position 5"):
            resolve_worksheet(client, "drive-1", "file-1", worksheet, None)

    def test_resolve_by_name_found(self):
        client = _mock_graph_client()
        client.get.return_value = _worksheets_response([{"id": "id-1", "name": "Data", "position": 0}])
        worksheet = Worksheet(name="Data")

        worksheet_id, is_new, actual_name = resolve_worksheet(client, "drive-1", "file-1", worksheet, "sess-1")

        assert (worksheet_id, is_new, actual_name) == ("id-1", False, "Data")
        assert client.get.call_args.kwargs["headers"] == {"workbook-session-id": "sess-1"}

    def test_resolve_by_name_missing_creates_sheet(self, caplog):
        client = _mock_graph_client()
        client.get.return_value = _worksheets_response([])
        client.post.return_value = _response(json_body={"id": "new-id", "name": "Data"})
        worksheet = Worksheet(name="Data")

        with caplog.at_level(logging.INFO, logger="client.excel_writer"):
            worksheet_id, is_new, actual_name = resolve_worksheet(client, "drive-1", "file-1", worksheet, None)

        assert (worksheet_id, is_new, actual_name) == ("new-id", True, "Data")
        client.post.assert_called_once_with(
            "/drives/drive-1/items/file-1/workbook/worksheets/add",
            json={"name": "Data"},
            headers={},
            retry_transient_workbook=True,
        )
        assert 'New sheet "Data" created.' in caplog.text

    def test_rename_when_resolved_by_id_and_name_differs(self):
        client = _mock_graph_client()
        client.get.return_value = _worksheets_response([{"id": "{abc-123}", "name": "Sheet1", "position": 0}])
        worksheet = Worksheet(id="{abc-123}", name="Renamed")

        worksheet_id, is_new, actual_name = resolve_worksheet(client, "drive-1", "file-1", worksheet, None)

        assert (worksheet_id, is_new, actual_name) == ("{abc-123}", False, "Renamed")
        client.patch.assert_called_once_with(
            "/drives/drive-1/items/file-1/workbook/worksheets/%7Babc-123%7D",
            json={"name": "Renamed"},
            headers={},
            retry_transient_workbook=True,
        )

    def test_no_rename_when_name_matches_resolved_name(self):
        client = _mock_graph_client()
        client.get.return_value = _worksheets_response([{"id": "id-1", "name": "Sheet1", "position": 0}])
        worksheet = Worksheet(id="id-1", name="Sheet1")

        resolve_worksheet(client, "drive-1", "file-1", worksheet, None)

        client.patch.assert_not_called()


# ---------------------------------------------------------------------------------------------
# Write algorithm
# ---------------------------------------------------------------------------------------------


class TestWriteTableEmptyCsv:
    def test_empty_csv_returns_false_and_touches_nothing(self, tmp_path):
        client = MagicMock()
        csv_path = _write_csv(tmp_path / "empty.csv", [])

        result = write_table(
            client, "drive-1", "file-1", "ws-1", csv_path,
            append=False, batch_size=5000, is_new_sheet=False, session=None,
        )

        assert result is False
        client.get.assert_not_called()
        client.post.assert_not_called()
        client.patch.assert_not_called()


class TestWriteTableOverwrite:
    def test_overwrite_clears_then_writes_from_a1_with_header(self, tmp_path, caplog):
        client = MagicMock()
        client.patch.return_value = _response()
        csv_path = _write_csv(tmp_path / "t.csv", [["a", "b"], ["1", "2"], ["3", "4"]])

        with caplog.at_level(logging.INFO, logger="client.excel_writer"):
            result = write_table(
                client, "drive-1", "file-1", "ws-1", csv_path,
                append=False, batch_size=5000, is_new_sheet=False, session=None,
            )

        assert result is True
        client.post.assert_called_once_with(
            "/drives/drive-1/items/file-1/workbook/worksheets/ws-1/range/clear",
            json={"applyTo": "all"},
            headers={},
            retry_transient_workbook=True,
        )
        client.patch.assert_called_once_with(
            "/drives/drive-1/items/file-1/workbook/worksheets/ws-1/range(address='A1:B3')",
            json={"values": [["a", "b"], ["1", "2"], ["3", "4"]]},
            headers={},
            retry_transient_workbook=True,
        )
        assert "Inserted 3 rows." in caplog.text

    def test_overwrite_of_a_new_sheet_skips_the_clear_call(self, tmp_path):
        client = MagicMock()
        client.patch.return_value = _response()
        csv_path = _write_csv(tmp_path / "t.csv", [["a", "b"], ["1", "2"]])

        write_table(
            client, "drive-1", "file-1", "ws-1", csv_path,
            append=False, batch_size=5000, is_new_sheet=True, session=None,
        )

        client.post.assert_not_called()
        client.patch.assert_called_once()
        assert "range(address='A1:B2')" in client.patch.call_args.args[0]


class TestWriteTableAppend:
    def test_append_to_new_sheet_writes_header_without_used_range_calls(self, tmp_path, caplog):
        client = MagicMock()
        client.patch.return_value = _response()
        csv_path = _write_csv(tmp_path / "t.csv", [["a", "b"], ["1", "2"]])

        with caplog.at_level(logging.INFO, logger="client.excel_writer"):
            write_table(
                client, "drive-1", "file-1", "ws-1", csv_path,
                append=True, batch_size=5000, is_new_sheet=True, session=None,
            )

        client.get.assert_not_called()
        assert "Sheet is empty." in caplog.text
        assert "range(address='A1:B2')" in client.patch.call_args.args[0]
        assert client.patch.call_args.kwargs["json"] == {"values": [["a", "b"], ["1", "2"]]}

    def test_append_skips_csv_header_when_existing_header_matches(self, tmp_path):
        client = MagicMock()
        client.get.side_effect = [
            _response(json_body={"address": "Sheet1!A1:B9"}),
            _response(json_body={"address": "Sheet1!A1:B1", "text": [["a", "b"]]}),
        ]
        client.patch.return_value = _response()
        csv_path = _write_csv(tmp_path / "t.csv", [["a", "b"], ["1", "2"]])

        write_table(
            client, "drive-1", "file-1", "ws-1", csv_path,
            append=True, batch_size=5000, is_new_sheet=False, session=None,
        )

        assert client.patch.call_args.kwargs["json"] == {"values": [["1", "2"]]}
        assert "range(address='A10:B10')" in client.patch.call_args.args[0]

    def test_append_warns_on_header_mismatch_but_still_skips_csv_header(self, tmp_path, caplog):
        client = MagicMock()
        client.get.side_effect = [
            _response(json_body={"address": "Sheet1!A1:C9"}),
            _response(json_body={"address": "Sheet1!A1:C1", "text": [["x", "y", "z"]]}),
        ]
        client.patch.return_value = _response()
        csv_path = _write_csv(tmp_path / "t.csv", [["a", "b", "c"], ["1", "2", "3"]])

        with caplog.at_level(logging.WARNING, logger="client.excel_writer"):
            write_table(
                client, "drive-1", "file-1", "ws-1", csv_path,
                append=True, batch_size=5000, is_new_sheet=False, session=None,
            )

        assert "Headers mismatch. Ignored new header:" in caplog.text
        assert client.patch.call_args.kwargs["json"] == {"values": [["1", "2", "3"]]}

    def test_append_to_an_empty_existing_sheet_writes_header(self, tmp_path, caplog):
        client = MagicMock()
        client.get.side_effect = [
            _response(json_body={"address": "Sheet1!A1:A1"}),
            _response(json_body={"address": "Sheet1!A1:A1", "text": [[""]]}),
        ]
        client.patch.return_value = _response()
        csv_path = _write_csv(tmp_path / "t.csv", [["a", "b"], ["1", "2"]])

        with caplog.at_level(logging.INFO, logger="client.excel_writer"):
            write_table(
                client, "drive-1", "file-1", "ws-1", csv_path,
                append=True, batch_size=5000, is_new_sheet=False, session=None,
            )

        assert "Sheet is empty." in caplog.text
        assert client.patch.call_args.kwargs["json"] == {"values": [["a", "b"], ["1", "2"]]}
        assert "range(address='A1:B2')" in client.patch.call_args.args[0]

    def test_append_start_column_follows_existing_ranges_start_column(self, tmp_path):
        client = MagicMock()
        client.get.side_effect = [
            _response(json_body={"address": "Sheet1!C9:D20"}),
            _response(json_body={"address": "Sheet1!C9:D9", "text": [["a", "b"]]}),
        ]
        client.patch.return_value = _response()
        csv_path = _write_csv(tmp_path / "t.csv", [["a", "b"], ["1", "2"]])

        write_table(
            client, "drive-1", "file-1", "ws-1", csv_path,
            append=True, batch_size=5000, is_new_sheet=False, session=None,
        )

        assert "range(address='C21:D21')" in client.patch.call_args.args[0]


class TestWriteTableBatchingAndFormatting:
    def test_batching_boundary_exact_multiple_is_a_single_batch(self, tmp_path):
        client = MagicMock()
        client.patch.return_value = _response()
        rows = [["a"]] + [[str(i)] for i in range(2)]  # header + 2 data rows
        csv_path = _write_csv(tmp_path / "t.csv", rows)

        write_table(
            client, "drive-1", "file-1", "ws-1", csv_path,
            append=False, batch_size=3, is_new_sheet=True, session=None,
        )

        assert client.patch.call_count == 1
        assert "range(address='A1:A3')" in client.patch.call_args.args[0]

    def test_batching_boundary_one_over_creates_a_second_smaller_batch(self, tmp_path):
        client = MagicMock()
        client.patch.return_value = _response()
        rows = [["a"]] + [[str(i)] for i in range(3)]  # header + 3 data rows = 4 items, batch_size=3
        csv_path = _write_csv(tmp_path / "t.csv", rows)

        write_table(
            client, "drive-1", "file-1", "ws-1", csv_path,
            append=False, batch_size=3, is_new_sheet=True, session=None,
        )

        assert client.patch.call_count == 2
        first_call, second_call = client.patch.call_args_list
        assert "range(address='A1:A3')" in first_call.args[0]
        assert first_call.kwargs["json"] == {"values": [["a"], ["0"], ["1"]]}
        assert "range(address='A4:A4')" in second_call.args[0]
        assert second_call.kwargs["json"] == {"values": [["2"]]}

    def test_short_rows_are_padded_and_long_rows_are_truncated(self, tmp_path):
        client = MagicMock()
        client.patch.return_value = _response()
        csv_path = _write_csv(tmp_path / "t.csv", [["a", "b", "c"], ["1"], ["2", "3", "4", "5"]])

        write_table(
            client, "drive-1", "file-1", "ws-1", csv_path,
            append=False, batch_size=5000, is_new_sheet=True, session=None,
        )

        values = client.patch.call_args.kwargs["json"]["values"]
        assert values == [["a", "b", "c"], ["1", "", ""], ["2", "3", "4"]]

    def test_formula_injection_is_escaped(self, tmp_path):
        client = MagicMock()
        client.patch.return_value = _response()
        csv_path = _write_csv(tmp_path / "t.csv", [["a"], ["=SUM(A1:A2)"]])

        write_table(
            client, "drive-1", "file-1", "ws-1", csv_path,
            append=False, batch_size=5000, is_new_sheet=True, session=None,
        )

        values = client.patch.call_args.kwargs["json"]["values"]
        assert values == [["a"], ["'=SUM(A1:A2)"]]

    def test_inserted_rows_logged_per_batch(self, tmp_path, caplog):
        client = MagicMock()
        client.patch.return_value = _response()
        rows = [["a"]] + [[str(i)] for i in range(4)]
        csv_path = _write_csv(tmp_path / "t.csv", rows)

        with caplog.at_level(logging.INFO, logger="client.excel_writer"):
            write_table(
                client, "drive-1", "file-1", "ws-1", csv_path,
                append=False, batch_size=2, is_new_sheet=True, session=None,
            )

        assert caplog.text.count("Inserted 2 rows.") == 2


class TestWriteTableSessionExpiry:
    def test_session_expired_mid_write_recreates_once_and_retries(self, tmp_path):
        client = MagicMock()
        client.patch.side_effect = [GraphNotFoundError("session gone", status_code=404), _response()]
        client.post.return_value = _response(201, json_body={"id": "sess-new"})
        csv_path = _write_csv(tmp_path / "t.csv", [["a"], ["1"]])

        result = write_table(
            client, "drive-1", "file-1", "ws-1", csv_path,
            append=True, batch_size=5000, is_new_sheet=True, session="sess-old",
        )

        assert result is True
        assert client.patch.call_count == 2
        client.post.assert_called_once_with(
            "/drives/drive-1/items/file-1/workbook/createSession",
            json={"persistChanges": True},
            headers={"Prefer": "respond-async"},
            retry_transient_workbook=True,
        )
        first_call, second_call = client.patch.call_args_list
        assert first_call.kwargs["headers"] == {"workbook-session-id": "sess-old"}
        assert second_call.kwargs["headers"] == {"workbook-session-id": "sess-new"}

    def test_session_expired_without_a_session_propagates(self, tmp_path):
        client = MagicMock()
        client.patch.side_effect = GraphNotFoundError("gone", status_code=404)
        csv_path = _write_csv(tmp_path / "t.csv", [["a"], ["1"]])

        with pytest.raises(GraphNotFoundError):
            write_table(
                client, "drive-1", "file-1", "ws-1", csv_path,
                append=True, batch_size=5000, is_new_sheet=True, session=None,
            )

        client.post.assert_not_called()
