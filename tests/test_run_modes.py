"""Tests for `Component.run`'s orchestration (plan Task 6): mode dispatch, file mode, CSV mode,
error mapping, and token-state persistence.

Design spec: ``docs/superpowers/specs/2026-08-17-wr-onedrive-v2-design.md`` §2 (scratch files,
state) and §6 (run orchestration, error mapping). Follows the datadir-style `KBC_DATADIR` fixture
pattern already used in ``tests/test_component.py``, extended with `data/in/files` and
`data/in/tables` content for file/CSV mode.
"""

import inspect
import json
import os
import re
import tempfile
from unittest import mock
from unittest.mock import MagicMock

import pytest
from freezegun import freeze_time
from keboola.component.exceptions import UserException

from client.exceptions import GraphConnectionError, GraphPermissionError
from component import Component


def _oauth_credentials(refresh_token: str = "refresh-config") -> dict:
    return {
        "id": "oauth-1",
        "created": "2026-01-01",
        "appKey": "client-1",
        "#appSecret": "secret-1",
        "oauthVersion": "2.0",
        "#data": json.dumps({"refresh_token": refresh_token}),
    }


def _build_component(
    tmp_path,
    parameters: dict,
    *,
    files: dict[str, bytes] | None = None,
    tables: dict[str, str] | None = None,
    oauth: dict | None = None,
) -> Component:
    """Build a `Component` from a `KBC_DATADIR`-style fixture (datadir pattern).

    `files`/`tables` map a file name (under `data/in/files` / `data/in/tables`) to its content.
    Table entries get a minimal input manifest (`{"id": ...}`) so `TableDefinition` resolves them
    as input tables, matching what the platform actually writes.
    """
    data_dir = tmp_path / "data"
    files_dir = data_dir / "in" / "files"
    tables_dir = data_dir / "in" / "tables"
    files_dir.mkdir(parents=True, exist_ok=True)
    tables_dir.mkdir(parents=True, exist_ok=True)
    (data_dir / "out").mkdir(parents=True, exist_ok=True)

    for name, content in (files or {}).items():
        (files_dir / name).write_bytes(content)

    for name, content in (tables or {}).items():
        (tables_dir / name).write_text(content)
        (tables_dir / f"{name}.manifest").write_text(json.dumps({"id": f"in.c-main.{name}"}))

    config = {
        "parameters": parameters,
        "action": "run",
        "authorization": {"oauth_api": {"credentials": oauth if oauth is not None else _oauth_credentials()}},
    }
    (data_dir / "config.json").write_text(json.dumps(config))

    with mock.patch.dict(os.environ, {"KBC_DATADIR": str(data_dir)}):
        return Component()


def _fake_token_provider(rotated_refresh_token: str | None = None) -> MagicMock:
    """A `TokenProvider` double that never touches the network and reports a fixed rotation."""
    provider = MagicMock()
    provider.rotated_refresh_token = rotated_refresh_token
    return provider


class TestRunOrchestratorShape:
    def test_run_is_a_thin_orchestrator_under_30_lines(self):
        source = inspect.getsource(Component.run)
        assert len(source.splitlines()) <= 30

    def test_run_uses_the_default_unattended_job_retry_budget(self, tmp_path):
        """IMPORTANT-5 (phase 8 audit): unlike a sync action, `run()` is an unattended job — it
        must keep `GraphClient`'s full default retry budget, not the sync action's fast-fail one."""
        parameters = {"mode": "file", "account": {"account_type": "private_onedrive"}, "destination": {}}
        comp = _build_component(tmp_path, parameters, files={"a.txt": b"aaa"})

        with (
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider()),
            mock.patch("component.GraphClient") as mock_graph_client,
            mock.patch("component.resolve_drive_id", return_value="drive-1"),
            mock.patch("component.ensure_folder", return_value="root-id"),
            mock.patch("component.upload_file", return_value={"id": "item-1"}),
        ):
            comp.run()

        mock_graph_client.assert_called_once_with(token_provider=mock.ANY)


class TestFileMode:
    @freeze_time("2026-08-17")
    def test_uploads_every_input_file_with_resolved_folder_path(self, tmp_path):
        parameters = {
            "mode": "file",
            "account": {"account_type": "private_onedrive"},
            "destination": {"folder_path": "reports/{date:%Y-%m-%d}", "conflict_behavior": "replace"},
        }
        comp = _build_component(tmp_path, parameters, files={"a.txt": b"aaa", "b.txt": b"bbb"})

        with (
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider()),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            mock.patch("component.resolve_drive_id", return_value="drive-1") as mock_resolve_drive,
            mock.patch("component.ensure_folder", return_value="parent-1") as mock_ensure_folder,
            mock.patch("component.upload_file", return_value={"id": "item-1"}) as mock_upload,
        ):
            comp.run()

        mock_resolve_drive.assert_called_once()
        mock_ensure_folder.assert_called_once_with(mock.ANY, "drive-1", "reports/2026-08-17")
        assert mock_upload.call_count == 2
        uploaded_names = {call_args.args[4] for call_args in mock_upload.call_args_list}
        assert uploaded_names == {"a.txt", "b.txt"}
        for call_args in mock_upload.call_args_list:
            assert call_args.args[1] == "drive-1"
            assert call_args.args[2] == "parent-1"
            assert call_args.args[5] == "replace"

    def test_zero_files_raises_user_exception(self, tmp_path):
        parameters = {"mode": "file", "account": {"account_type": "private_onedrive"}, "destination": {}}
        comp = _build_component(tmp_path, parameters, files={})

        with (
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider()),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            mock.patch("component.resolve_drive_id", return_value="drive-1"),
            pytest.raises(UserException, match="No files found in the input mapping"),
        ):
            comp.run()


class TestCsvModeCardinality:
    def test_zero_tables_raises_v1_parity_message(self, tmp_path):
        parameters = {"mode": "table_csv", "account": {"account_type": "private_onedrive"}, "destination": {}}
        comp = _build_component(tmp_path, parameters, tables={})

        with (
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider()),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            mock.patch("component.resolve_drive_id", return_value="drive-1"),
            pytest.raises(UserException, match=re.escape('No CSV file found in "/data/in/tables".')),
        ):
            comp.run()

    def test_multiple_tables_raises_v1_parity_message_comma_joined(self, tmp_path):
        parameters = {"mode": "table_csv", "account": {"account_type": "private_onedrive"}, "destination": {}}
        comp = _build_component(tmp_path, parameters, tables={"a.csv": "id\n1\n", "b.csv": "id\n2\n"})
        expected = re.escape('Expected one CSV file, found multiple: "a.csv", "b.csv".')

        with (
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider()),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            mock.patch("component.resolve_drive_id", return_value="drive-1"),
            pytest.raises(UserException, match=expected),
        ):
            comp.run()


class TestCsvModeUpload:
    def test_passthrough_when_options_match_storage_defaults(self, tmp_path):
        parameters = {
            "mode": "table_csv",
            "account": {"account_type": "private_onedrive"},
            "destination": {"conflict_behavior": "fail"},
        }
        comp = _build_component(tmp_path, parameters, tables={"mytable": "id,name\n1,a\n2,b\n"})

        with (
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider()),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            mock.patch("component.resolve_drive_id", return_value="drive-1"),
            mock.patch("component.ensure_folder", return_value="root-id"),
            mock.patch("component.upload_file", return_value={"id": "item-1"}) as mock_upload,
        ):
            comp.run()

        expected_source_path = str(tmp_path / "data" / "in" / "tables" / "mytable")
        upload_path = mock_upload.call_args.args[3]
        file_name = mock_upload.call_args.args[4]
        assert upload_path == expected_source_path  # streamed as-is, no rewrite
        assert file_name == "mytable.csv"  # csv.file_name defaults to "<table name>.csv"
        assert os.path.exists(upload_path)  # the input file itself, never deleted

    def test_rewrite_when_delimiter_differs_and_header_disabled_drops_first_row(self, tmp_path):
        parameters = {
            "mode": "table_csv",
            "account": {"account_type": "private_onedrive"},
            "destination": {},
            "csv": {"delimiter": ";", "include_header": False},
        }
        comp = _build_component(tmp_path, parameters, tables={"mytable": "id,name\n1,a\n2,b\n"})
        captured: dict = {}

        def _fake_upload_file(client, drive_id, parent_id, local_path, file_name, conflict_behavior):
            # Read the rewritten temp file's content before the component's `finally` deletes it.
            with open(local_path, newline="") as file_handle:
                captured["content"] = file_handle.read()
            captured["path"] = local_path
            return {"id": "item-1"}

        with (
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider()),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            mock.patch("component.resolve_drive_id", return_value="drive-1"),
            mock.patch("component.ensure_folder", return_value="root-id"),
            mock.patch("component.upload_file", side_effect=_fake_upload_file),
        ):
            comp.run()

        assert captured["content"] == "1;a\r\n2;b\r\n"  # header dropped, ";"-delimited
        assert not os.path.exists(captured["path"])  # cleaned up in `finally`
        source_path = str(tmp_path / "data" / "in" / "tables" / "mytable")
        assert captured["path"] != source_path  # a genuinely different (temp) file
        # Never scratch under `data/out/` — only `state.json` is written there.
        data_out_dir = str(tmp_path / "data" / "out")
        assert not captured["path"].startswith(data_out_dir)
        assert captured["path"].startswith(tempfile.gettempdir())

    def test_rewrite_when_enclosure_differs_streams_row_by_row(self, tmp_path):
        parameters = {
            "mode": "table_csv",
            "account": {"account_type": "private_onedrive"},
            "destination": {},
            "csv": {"enclosure": "'"},
        }
        comp = _build_component(tmp_path, parameters, tables={"mytable": 'id,name\n1,"a,b"\n'})
        captured: dict = {}

        def _fake_upload_file(client, drive_id, parent_id, local_path, file_name, conflict_behavior):
            with open(local_path, newline="") as file_handle:
                captured["content"] = file_handle.read()
            return {"id": "item-1"}

        with (
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider()),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            mock.patch("component.resolve_drive_id", return_value="drive-1"),
            mock.patch("component.ensure_folder", return_value="root-id"),
            mock.patch("component.upload_file", side_effect=_fake_upload_file),
        ):
            comp.run()

        assert captured["content"] == "id,name\r\n1,'a,b'\r\n"


class TestTokenStatePersistence:
    def test_rotated_token_persisted_on_success(self, tmp_path):
        parameters = {"mode": "file", "account": {"account_type": "private_onedrive"}, "destination": {}}
        comp = _build_component(tmp_path, parameters, files={"a.txt": b"aaa"})

        with (
            mock.patch(
                "component.RefreshTokenProvider", return_value=_fake_token_provider("new-refresh-token")
            ),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            mock.patch("component.resolve_drive_id", return_value="drive-1"),
            mock.patch("component.ensure_folder", return_value="root-id"),
            mock.patch("component.upload_file", return_value={"id": "item-1"}),
        ):
            comp.run()

        state = json.loads((tmp_path / "data" / "out" / "state.json").read_text())
        payload = json.loads(state["#refreshed_auth_data"])
        assert payload["refresh_token"] == "new-refresh-token"

    def test_rotated_token_persisted_even_when_the_upload_raises(self, tmp_path):
        parameters = {"mode": "file", "account": {"account_type": "private_onedrive"}, "destination": {}}
        comp = _build_component(tmp_path, parameters, files={"a.txt": b"aaa"})

        with (
            mock.patch(
                "component.RefreshTokenProvider", return_value=_fake_token_provider("new-refresh-token-2")
            ),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            mock.patch("component.resolve_drive_id", return_value="drive-1"),
            mock.patch("component.ensure_folder", return_value="root-id"),
            mock.patch(
                "component.upload_file", side_effect=GraphPermissionError("no access", status_code=403)
            ),pytest.raises(UserException)
        ):
            comp.run()

        state_path = tmp_path / "data" / "out" / "state.json"
        assert state_path.exists()
        state = json.loads(state_path.read_text())
        payload = json.loads(state["#refreshed_auth_data"])
        assert payload["refresh_token"] == "new-refresh-token-2"

    def test_no_state_written_when_nothing_rotated(self, tmp_path):
        parameters = {"mode": "file", "account": {"account_type": "private_onedrive"}, "destination": {}}
        comp = _build_component(tmp_path, parameters, files={"a.txt": b"aaa"})

        with (
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider(None)),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            mock.patch("component.resolve_drive_id", return_value="drive-1"),
            mock.patch("component.ensure_folder", return_value="root-id"),
            mock.patch("component.upload_file", return_value={"id": "item-1"}),
        ):
            comp.run()

        assert not (tmp_path / "data" / "out" / "state.json").exists()


def _fake_session_context_manager(session_id: str | None) -> MagicMock:
    """A `workbook_session`-shaped context manager double yielding `session_id`."""
    context_manager = MagicMock()
    context_manager.__enter__.return_value = session_id
    context_manager.__exit__.return_value = False
    return context_manager


class TestExcelMode:
    """Excel mode wiring (plan Task 8): personal-account gate, empty CSV, happy-path plumbing.

    `resolve_workbook`/`workbook_session`/`resolve_worksheet`/`write_table` themselves are
    exercised at the unit level in ``tests/test_excel_writer.py`` (plan Task 7); these tests only
    assert `Component._run_excel_mode` wires them together correctly.
    """

    def test_private_onedrive_account_raises_user_exception(self, tmp_path):
        parameters = {
            "mode": "table_excel",
            "account": {"account_type": "private_onedrive"},
            "workbook": {"path": "/book.xlsx"},
            "worksheet": {"name": "Sheet1"},
        }
        comp = _build_component(tmp_path, parameters, tables={"mytable": "id,name\n1,a\n"})

        with (
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider()),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            mock.patch("component.resolve_drive_id") as mock_resolve_drive_id,
            mock.patch("component.resolve_workbook") as mock_resolve_workbook,
            pytest.raises(UserException, match="private_onedrive"),
        ):
            comp.run()

        mock_resolve_workbook.assert_not_called()
        # IMPORTANT-2 (phase 8 audit): Excel mode never resolves a drive id at all.
        mock_resolve_drive_id.assert_not_called()

    def test_empty_csv_logs_v1_parity_warning_and_exits_cleanly(self, tmp_path, caplog):
        parameters = {
            "mode": "table_excel",
            "account": {"account_type": "onedrive_for_business", "tenant_id": "tenant-1"},
            "workbook": {"path": "/book.xlsx"},
            "worksheet": {"name": "Sheet1"},
        }
        comp = _build_component(tmp_path, parameters, tables={"empty": ""})

        with (
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider()),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            mock.patch("component.resolve_drive_id") as mock_resolve_drive_id,
            mock.patch("component.resolve_workbook", return_value=("wb-drive", "wb-file", False)),
            mock.patch("component.workbook_session", return_value=_fake_session_context_manager("session-1")),
            mock.patch("component.resolve_worksheet", return_value=("sheet-1", False, "Sheet1")),
            mock.patch("component.write_table", return_value=False) as mock_write,
            caplog.at_level("WARNING"),
        ):
            comp.run()  # must not raise — v1 parity: exit 0, sheet untouched.

        mock_write.assert_called_once()
        assert 'Ignored empty CSV file "empty".' in caplog.text
        mock_resolve_drive_id.assert_not_called()

    def test_happy_path_passes_configured_append_and_batch_size_to_write_table(self, tmp_path):
        parameters = {
            "mode": "table_excel",
            "account": {"account_type": "onedrive_for_business", "tenant_id": "tenant-1"},
            "workbook": {"path": "/book.xlsx"},
            "worksheet": {"name": "Sheet1"},
            "append": True,
            "batch_size": 1234,
        }
        comp = _build_component(tmp_path, parameters, tables={"mytable": "id,name\n1,a\n"})

        with (
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider()),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            mock.patch("component.resolve_drive_id") as mock_resolve_drive_id,
            mock.patch(
                "component.resolve_workbook", return_value=("wb-drive", "wb-file", False)
            ) as mock_resolve_workbook,
            mock.patch(
                "component.workbook_session", return_value=_fake_session_context_manager("session-1")
            ) as mock_session,
            mock.patch(
                "component.resolve_worksheet", return_value=("sheet-1", False, "Sheet1")
            ) as mock_resolve_worksheet,
            mock.patch("component.write_table", return_value=True) as mock_write,
        ):
            comp.run()

        mock_resolve_drive_id.assert_not_called()
        mock_resolve_workbook.assert_called_once()
        assert mock_resolve_workbook.call_args.args[1].account_type.value == "onedrive_for_business"
        mock_session.assert_called_once_with(mock.ANY, "wb-drive", "wb-file")
        assert mock_resolve_worksheet.call_args.args[1:3] == ("wb-drive", "wb-file")
        assert mock_resolve_worksheet.call_args.args[4] == "session-1"

        mock_write.assert_called_once()
        write_args = mock_write.call_args
        assert write_args.args[1:4] == ("wb-drive", "wb-file", "sheet-1")
        assert write_args.kwargs["append"] is True
        assert write_args.kwargs["batch_size"] == 1234
        assert write_args.kwargs["is_new_sheet"] is False
        assert write_args.kwargs["session"] == "session-1"


class TestErrorMapping:
    def test_graph_permission_error_maps_to_user_exception(self, tmp_path):
        parameters = {"mode": "file", "account": {"account_type": "private_onedrive"}, "destination": {}}
        comp = _build_component(tmp_path, parameters, files={"a.txt": b"aaa"})

        with (
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider()),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            mock.patch("component.resolve_drive_id", return_value="drive-1"),
            mock.patch("component.ensure_folder", return_value="root-id"),
            mock.patch(
                "component.upload_file", side_effect=GraphPermissionError("no access", status_code=403)
            ),pytest.raises(UserException, match="no access")
        ):
            comp.run()

    def test_graph_connection_error_maps_to_user_exception(self, tmp_path):
        """IMPORTANT-1: a network failure that exhausts the client's retry budget surfaces as
        `GraphConnectionError`, which must map to a `UserException` (exit 1), not propagate as an
        unhandled exception (exit 2)."""
        parameters = {"mode": "file", "account": {"account_type": "private_onedrive"}, "destination": {}}
        comp = _build_component(tmp_path, parameters, files={"a.txt": b"aaa"})

        with (
            mock.patch("component.RefreshTokenProvider", return_value=_fake_token_provider()),
            mock.patch("component.GraphClient", return_value=MagicMock()),
            mock.patch("component.resolve_drive_id", return_value="drive-1"),
            mock.patch("component.ensure_folder", return_value="root-id"),
            mock.patch(
                "component.upload_file",
                side_effect=GraphConnectionError("network error, retry budget exhausted"),
            ),
            pytest.raises(UserException, match="retry budget exhausted"),
        ):
            comp.run()

    def test_invalid_configuration_raises_user_exception(self, tmp_path):
        # mode=table_excel requires workbook/worksheet — neither is provided.
        parameters = {"mode": "table_excel", "account": {"account_type": "private_onedrive"}}
        comp = _build_component(tmp_path, parameters)

        with pytest.raises(UserException, match="Invalid configuration"):
            comp.run()
