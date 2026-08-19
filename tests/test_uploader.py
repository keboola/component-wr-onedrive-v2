import logging
from datetime import UTC, datetime
from unittest.mock import MagicMock, call

import pytest

from client.exceptions import (
    FileAlreadyExistsError,
    GraphClientError,
    GraphConnectionError,
    GraphNotFoundError,
    InvalidPathError,
    UploadSessionError,
)
from client.uploader import (
    CHUNK_SIZE,
    MAX_RESUME_ATTEMPTS,
    SIMPLE_UPLOAD_THRESHOLD,
    encode_path_segments,
    ensure_folder,
    resolve_placeholders,
    upload_file,
    validate_path,
)


def _response(status_code: int, json_body=None, headers: dict | None = None) -> MagicMock:
    response = MagicMock()
    response.status_code = status_code
    response.ok = 200 <= status_code < 400
    response.headers = headers or {}
    if json_body is not None:
        response.json.return_value = json_body
    return response


def _graph_error(status_code: int, error_code: str | None = None) -> GraphClientError:
    return GraphClientError(f"HTTP {status_code}", status_code=status_code, error_code=error_code)


def _make_sparse_file(path, size: int) -> str:
    """Create a file of exactly ``size`` bytes without allocating real content (sparse file)."""
    with open(path, "wb") as fh:
        if size:
            fh.truncate(size)
    return str(path)


class TestResolvePlaceholders:
    def test_resolves_a_single_date_token(self):
        now = datetime(2026, 8, 17, 12, 30, tzinfo=UTC)

        result = resolve_placeholders("reports/{date:%Y-%m-%d}", now)

        assert result == "reports/2026-08-17"

    def test_resolves_multiple_tokens_with_arbitrary_strftime_formats(self):
        now = datetime(2026, 8, 17, 9, 5, tzinfo=UTC)

        result = resolve_placeholders("{date:%Y}/{date:%m}/{date:%d}-report", now)

        assert result == "2026/08/17-report"

    def test_path_without_placeholders_is_unchanged(self):
        now = datetime(2026, 8, 17, tzinfo=UTC)

        assert resolve_placeholders("static/folder", now) == "static/folder"

    def test_unknown_token_raises_invalid_path_error(self):
        now = datetime(2026, 8, 17, tzinfo=UTC)

        with pytest.raises(InvalidPathError, match="Unknown placeholder"):
            resolve_placeholders("reports/{table_name}", now)

    def test_empty_date_format_raises(self):
        now = datetime(2026, 8, 17, tzinfo=UTC)

        with pytest.raises(InvalidPathError, match="Empty date format"):
            resolve_placeholders("reports/{date:}", now)


class TestValidatePath:
    @pytest.mark.parametrize("char", list('"*:<>?\\|'))
    def test_reserved_character_rejected(self, char):
        with pytest.raises(InvalidPathError, match="reserved character"):
            validate_path(f"folder/na{char}me", business=False)

    def test_hash_and_percent_allowed_on_personal_onedrive(self):
        validate_path("folder/na#me%2", business=False)  # does not raise

    @pytest.mark.parametrize("char", ["#", "%"])
    def test_hash_and_percent_rejected_on_business(self, char):
        with pytest.raises(InvalidPathError, match="reserved character"):
            validate_path(f"folder/na{char}me", business=True)

    @pytest.mark.parametrize("name", ["CON", "con", "PRN", "AUX", "NUL", "COM1", "com9", "LPT0", "LPT5"])
    def test_reserved_device_names_rejected(self, name):
        with pytest.raises(InvalidPathError, match="reserved device name"):
            validate_path(f"folder/{name}", business=False)

    @pytest.mark.parametrize("name", [".lock", "desktop.ini", "DESKTOP.INI"])
    def test_reserved_exact_names_rejected(self, name):
        with pytest.raises(InvalidPathError, match="reserved name"):
            validate_path(f"folder/{name}", business=False)

    def test_name_starting_with_tilde_dollar_rejected(self):
        with pytest.raises(InvalidPathError, match="starts with '~'"):
            validate_path("folder/~$temp.docx", business=False)

    def test_vti_substring_anywhere_rejected(self):
        with pytest.raises(InvalidPathError, match="_vti_"):
            validate_path("folder/my_vti_stuff", business=False)

    def test_forms_at_root_rejected(self):
        with pytest.raises(InvalidPathError, match="forms"):
            validate_path("forms", business=False)

    def test_forms_not_at_root_is_allowed(self):
        validate_path("folder/forms", business=False)  # does not raise

    def test_segment_ending_with_dot_rejected(self):
        with pytest.raises(InvalidPathError, match="ends with '.'"):
            validate_path("folder/name.", business=False)

    def test_segment_with_leading_trailing_spaces_rejected(self):
        with pytest.raises(InvalidPathError, match="leading/trailing spaces"):
            validate_path("folder/ name", business=False)

    def test_segment_too_long_rejected(self):
        with pytest.raises(InvalidPathError, match="256 characters"):
            validate_path("a" * 256, business=False)

    def test_segment_at_max_length_is_allowed(self):
        validate_path("a" * 255, business=False)  # does not raise

    def test_total_path_too_long_rejected(self):
        long_path = "/".join(["seg"] * 150)
        assert len(long_path) > 400
        with pytest.raises(InvalidPathError, match="maximum is 400"):
            validate_path(long_path, business=False)

    def test_valid_path_does_not_raise(self):
        validate_path("reports/2026-08-17/monthly export", business=True)


class TestEncodePathSegments:
    def test_space_is_percent_encoded(self):
        assert encode_path_segments("my folder/report") == "my%20folder/report"

    def test_hash_is_percent_encoded(self):
        assert encode_path_segments("folder#1") == "folder%231"

    def test_diacritics_are_percent_encoded(self):
        assert encode_path_segments("faktury/účetní") == "faktury/%C3%BA%C4%8Detn%C3%AD"

    def test_slash_within_a_conceptual_segment_never_occurs_but_separators_are_preserved(self):
        assert encode_path_segments("a/b/c") == "a/b/c"

    def test_empty_path_encodes_to_empty_string(self):
        assert encode_path_segments("") == ""


class TestEnsureFolder:
    def test_empty_path_returns_root_id(self):
        client = MagicMock()
        client.get.return_value = _response(200, {"id": "root-id"})

        folder_id = ensure_folder(client, "drive-1", "")

        client.get.assert_called_once_with("/drives/drive-1/root")
        assert folder_id == "root-id"

    def test_existing_folders_are_found_via_get(self):
        client = MagicMock()
        client.get.side_effect = [
            _response(200, {"id": "a-id"}),
            _response(200, {"id": "b-id"}),
        ]

        folder_id = ensure_folder(client, "drive-1", "a/b")

        assert client.get.call_args_list == [
            call("/drives/drive-1/root:/a"),
            call("/drives/drive-1/root:/a/b"),
        ]
        client.post.assert_not_called()
        assert folder_id == "b-id"

    def test_missing_segment_is_created(self):
        client = MagicMock()
        client.get.side_effect = GraphNotFoundError("not found", status_code=404)
        client.post.return_value = _response(201, {"id": "new-id"})

        folder_id = ensure_folder(client, "drive-1", "new")

        client.post.assert_called_once_with(
            "/drives/drive-1/root/children",
            json={"name": "new", "folder": {}, "@microsoft.graph.conflictBehavior": "fail"},
        )
        assert folder_id == "new-id"

    def test_second_level_folder_is_created_under_the_first_levels_id(self):
        client = MagicMock()
        client.get.side_effect = [
            _response(200, {"id": "a-id"}),
            GraphNotFoundError("not found", status_code=404),
        ]
        client.post.return_value = _response(201, {"id": "b-id"})

        folder_id = ensure_folder(client, "drive-1", "a/b")

        client.post.assert_called_once_with(
            "/drives/drive-1/items/a-id/children",
            json={"name": "b", "folder": {}, "@microsoft.graph.conflictBehavior": "fail"},
        )
        assert folder_id == "b-id"

    def test_409_creation_race_reregets_the_folder(self):
        client = MagicMock()
        client.get.side_effect = [
            GraphNotFoundError("not found", status_code=404),
            _response(200, {"id": "existing-id"}),
        ]
        client.post.side_effect = _graph_error(409, "nameAlreadyExists")

        folder_id = ensure_folder(client, "drive-1", "raced")

        assert folder_id == "existing-id"
        assert client.get.call_args_list == [
            call("/drives/drive-1/root:/raced"),
            call("/drives/drive-1/root:/raced"),
        ]

    def test_non_conflict_creation_error_propagates(self):
        client = MagicMock()
        client.get.side_effect = GraphNotFoundError("not found", status_code=404)
        client.post.side_effect = _graph_error(403)

        with pytest.raises(GraphClientError):
            ensure_folder(client, "drive-1", "denied")


class TestUploadDispatch:
    def test_small_file_uses_simple_put_with_conflict_behavior_query_param(self, tmp_path):
        local_path = _make_sparse_file(tmp_path / "small.csv", 1024)
        client = MagicMock()
        client.put.return_value = _response(201, {"id": "item-1", "name": "small.csv"})

        result = upload_file(client, "drive-1", "parent-1", local_path, "small.csv", "fail")

        client.put.assert_called_once()
        args, kwargs = client.put.call_args
        assert args[0] == "/drives/drive-1/items/parent-1:/small.csv:/content"
        assert kwargs["params"] == {"@microsoft.graph.conflictBehavior": "fail"}
        assert kwargs["headers"]["Content-Type"] == "application/octet-stream"
        assert kwargs["data"].name == local_path
        client.post.assert_not_called()
        assert result == {"id": "item-1", "name": "small.csv"}

    def test_simple_put_conflict_behavior_is_always_present_even_for_replace(self, tmp_path):
        local_path = _make_sparse_file(tmp_path / "small.csv", 10)
        client = MagicMock()
        client.put.return_value = _response(200, {"id": "item-1"})

        upload_file(client, "drive-1", "parent-1", local_path, "small.csv", "replace")

        _, kwargs = client.put.call_args
        assert kwargs["params"] == {"@microsoft.graph.conflictBehavior": "replace"}

    def test_file_exactly_at_threshold_uses_simple_put(self, tmp_path):
        local_path = _make_sparse_file(tmp_path / "exact.bin", SIMPLE_UPLOAD_THRESHOLD)
        client = MagicMock()
        client.put.return_value = _response(201, {"id": "item-1"})

        upload_file(client, "drive-1", "parent-1", local_path, "exact.bin", "fail")

        client.put.assert_called_once()
        client.post.assert_not_called()

    def test_simple_put_name_conflict_raises_file_already_exists(self, tmp_path):
        local_path = _make_sparse_file(tmp_path / "small.csv", 10)
        client = MagicMock()
        client.put.side_effect = _graph_error(409, "nameAlreadyExists")

        with pytest.raises(FileAlreadyExistsError, match="small.csv"):
            upload_file(client, "drive-1", "parent-1", local_path, "small.csv", "fail")

    def test_simple_put_other_error_propagates_unwrapped(self, tmp_path):
        local_path = _make_sparse_file(tmp_path / "small.csv", 10)
        client = MagicMock()
        client.put.side_effect = _graph_error(403)

        with pytest.raises(GraphClientError):
            upload_file(client, "drive-1", "parent-1", local_path, "small.csv", "fail")

    def test_file_over_threshold_uses_upload_session(self, tmp_path):
        size = SIMPLE_UPLOAD_THRESHOLD + 1
        local_path = _make_sparse_file(tmp_path / "big.bin", size)
        client = MagicMock()
        client.post.return_value = _response(200, {"uploadUrl": "https://upload.example/session"})
        client.put.side_effect = [
            _response(202),
            _response(201, {"id": "item-big"}),
        ]

        result = upload_file(client, "drive-1", "parent-1", local_path, "big.bin", "fail")

        client.post.assert_called_once_with(
            "/drives/drive-1/items/parent-1:/big.bin:/createUploadSession",
            json={"item": {"@microsoft.graph.conflictBehavior": "fail", "name": "big.bin"}},
        )
        assert result == {"id": "item-big"}
        assert client.put.call_count == 2
        first_headers = client.put.call_args_list[0].kwargs["headers"]
        second_headers = client.put.call_args_list[1].kwargs["headers"]
        assert first_headers["Content-Range"] == f"bytes 0-{SIMPLE_UPLOAD_THRESHOLD - 1}/{size}"
        assert second_headers["Content-Range"] == f"bytes {SIMPLE_UPLOAD_THRESHOLD}-{size - 1}/{size}"
        for kwargs in (client.put.call_args_list[0].kwargs, client.put.call_args_list[1].kwargs):
            assert kwargs["absolute"] is True
            assert kwargs["auth"] is False
            assert kwargs["retry"] is False


class TestChunkMath:
    def test_25_mib_file_uploads_in_three_chunks_with_a_partial_final_chunk(self, tmp_path):
        size = 25 * 1024 * 1024
        local_path = _make_sparse_file(tmp_path / "big.bin", size)
        client = MagicMock()
        client.post.return_value = _response(200, {"uploadUrl": "https://upload.example/session"})
        client.put.side_effect = [
            _response(202),
            _response(202),
            _response(201, {"id": "item-big"}),
        ]

        result = upload_file(client, "drive-1", "parent-1", local_path, "big.bin", "fail")

        assert result == {"id": "item-big"}
        assert client.put.call_count == 3
        ranges = [c.kwargs["headers"]["Content-Range"] for c in client.put.call_args_list]
        lengths = [c.kwargs["headers"]["Content-Length"] for c in client.put.call_args_list]
        assert ranges == [
            f"bytes 0-{CHUNK_SIZE - 1}/{size}",
            f"bytes {CHUNK_SIZE}-{2 * CHUNK_SIZE - 1}/{size}",
            f"bytes {2 * CHUNK_SIZE}-{size - 1}/{size}",
        ]
        assert lengths == [str(CHUNK_SIZE), str(CHUNK_SIZE), str(size - 2 * CHUNK_SIZE)]
        # Final chunk is smaller than a full CHUNK_SIZE.
        assert int(lengths[-1]) < CHUNK_SIZE


class TestUploadSessionResume:
    def test_transient_failure_resumes_from_next_expected_ranges(self, tmp_path):
        size = 25 * 1024 * 1024
        local_path = _make_sparse_file(tmp_path / "big.bin", size)
        client = MagicMock()
        client.post.return_value = _response(200, {"uploadUrl": "https://upload.example/session"})
        resumed_offset = CHUNK_SIZE + 2048
        client.put.side_effect = [
            _response(202),  # chunk 1 ok
            _graph_error(503),  # chunk 2 fails transiently
            _response(202),  # resumed partial chunk accepted
            _response(201, {"id": "item-big"}),  # final chunk
        ]
        client.get.return_value = _response(200, {"nextExpectedRanges": [f"{resumed_offset}-"]})

        result = upload_file(client, "drive-1", "parent-1", local_path, "big.bin", "fail")

        assert result == {"id": "item-big"}
        client.get.assert_called_once_with("https://upload.example/session", absolute=True, auth=False, retry=False)
        resumed_range = client.put.call_args_list[2].kwargs["headers"]["Content-Range"]
        expected_chunk_end = min(resumed_offset + CHUNK_SIZE, size) - 1
        assert resumed_range == f"bytes {resumed_offset}-{expected_chunk_end}/{size}"

    def test_graph_connection_error_resumes_from_next_expected_ranges(self, tmp_path):
        """A connection error/timeout at the `GraphClient` transport boundary now surfaces as
        `GraphConnectionError` (a `GraphClientError` subclass with no `status_code`) instead of a
        raw `requests.RequestException` — the resume dance must treat it as transient, same as a
        503, rather than aborting immediately because `None not in _TRANSIENT_CHUNK_STATUSES`."""
        size = 25 * 1024 * 1024
        local_path = _make_sparse_file(tmp_path / "big.bin", size)
        client = MagicMock()
        client.post.return_value = _response(200, {"uploadUrl": "https://upload.example/session"})
        resumed_offset = CHUNK_SIZE + 2048
        client.put.side_effect = [
            _response(202),  # chunk 1 ok
            GraphConnectionError("connection refused"),  # chunk 2: connection error
            _response(202),  # resumed partial chunk accepted
            _response(201, {"id": "item-big"}),  # final chunk
        ]
        client.get.return_value = _response(200, {"nextExpectedRanges": [f"{resumed_offset}-"]})

        result = upload_file(client, "drive-1", "parent-1", local_path, "big.bin", "fail")

        assert result == {"id": "item-big"}
        client.get.assert_called_once_with("https://upload.example/session", absolute=True, auth=False, retry=False)
        resumed_range = client.put.call_args_list[2].kwargs["headers"]["Content-Range"]
        expected_chunk_end = min(resumed_offset + CHUNK_SIZE, size) - 1
        assert resumed_range == f"bytes {resumed_offset}-{expected_chunk_end}/{size}"

    def test_session_404_restarts_once(self, tmp_path):
        size = 5
        local_path = _make_sparse_file(tmp_path / "small.bin", SIMPLE_UPLOAD_THRESHOLD + size)
        client = MagicMock()
        client.post.side_effect = [
            _response(200, {"uploadUrl": "https://upload.example/session-1"}),
            _response(200, {"uploadUrl": "https://upload.example/session-2"}),
        ]
        client.put.side_effect = [
            _graph_error(404),  # session-1's first chunk: session is gone
            _response(202),  # session-2's first chunk
            _response(201, {"id": "item-1"}),  # session-2's final chunk
        ]

        result = upload_file(client, "drive-1", "parent-1", local_path, "small.bin", "fail")

        assert result == {"id": "item-1"}
        assert client.post.call_count == 2
        urls_used = [c.args[0] for c in client.put.call_args_list]
        assert urls_used == [
            "https://upload.example/session-1",
            "https://upload.example/session-2",
            "https://upload.example/session-2",
        ]

    def test_session_404_twice_gives_up(self, tmp_path):
        size = 5
        local_path = _make_sparse_file(tmp_path / "small.bin", SIMPLE_UPLOAD_THRESHOLD + size)
        client = MagicMock()
        client.post.return_value = _response(200, {"uploadUrl": "https://upload.example/session"})
        client.put.side_effect = _graph_error(404)

        with pytest.raises(UploadSessionError):
            upload_file(client, "drive-1", "parent-1", local_path, "small.bin", "fail")

        assert client.post.call_count == 2

    def test_unrecoverable_failure_deletes_session_and_raises(self, tmp_path):
        size = 5
        local_path = _make_sparse_file(tmp_path / "small.bin", SIMPLE_UPLOAD_THRESHOLD + size)
        client = MagicMock()
        client.post.return_value = _response(200, {"uploadUrl": "https://upload.example/session"})
        # Every chunk PUT attempt fails transiently, and every resume status check succeeds but
        # never advances beyond the resume-attempt budget.
        client.put.side_effect = _graph_error(503)
        client.get.return_value = _response(200, {"nextExpectedRanges": ["0-"]})

        with pytest.raises(UploadSessionError):
            upload_file(client, "drive-1", "parent-1", local_path, "small.bin", "fail")

        assert client.put.call_count == MAX_RESUME_ATTEMPTS + 1
        client.delete.assert_called_once_with("https://upload.example/session", absolute=True, auth=False, retry=False)

    def test_resume_status_check_failure_message_has_query_strings_redacted(self, tmp_path):
        """IMPORTANT-6 (phase 8 audit): when the resume status check itself fails (not a 404,
        so not a session restart), the raised `UploadSessionError` must sanitize that error's
        message — a `GraphClientError` can otherwise reproduce a query string verbatim."""
        size = 5
        local_path = _make_sparse_file(tmp_path / "small.bin", SIMPLE_UPLOAD_THRESHOLD + size)
        client = MagicMock()
        client.post.return_value = _response(200, {"uploadUrl": "https://upload.example/session"})
        client.put.side_effect = _graph_error(503)
        client.get.side_effect = GraphClientError(
            "status check failed: https://upload.example/session?tempauth=super-secret-token",
            status_code=500,
        )

        with pytest.raises(UploadSessionError) as exc_info:
            upload_file(client, "drive-1", "parent-1", local_path, "small.bin", "fail")

        assert "super-secret-token" not in str(exc_info.value)
        assert "?<redacted>" in str(exc_info.value)
        client.delete.assert_called_once_with("https://upload.example/session", absolute=True, auth=False, retry=False)

    def test_abort_session_failure_does_not_log_the_upload_url(self, tmp_path, caplog):
        """`uploadUrl` is a pre-signed, credential-bearing URL — a failed best-effort cleanup
        DELETE must never log it (IMPORTANT-2)."""
        size = 5
        local_path = _make_sparse_file(tmp_path / "small.bin", SIMPLE_UPLOAD_THRESHOLD + size)
        secret_url = "https://upload.example/session?token=super-secret-signed-token"
        client = MagicMock()
        client.post.return_value = _response(200, {"uploadUrl": secret_url})
        client.put.side_effect = _graph_error(400)  # non-retryable -> triggers _abort_session
        client.delete.side_effect = _graph_error(500)  # cleanup DELETE itself fails

        with caplog.at_level(logging.WARNING, logger="client.uploader"), pytest.raises(UploadSessionError):
            upload_file(client, "drive-1", "parent-1", local_path, "small.bin", "fail")

        assert secret_url not in caplog.text
        assert "abandoned upload session" in caplog.text

    def test_non_retryable_status_aborts_immediately_without_resuming(self, tmp_path):
        size = 5
        local_path = _make_sparse_file(tmp_path / "small.bin", SIMPLE_UPLOAD_THRESHOLD + size)
        client = MagicMock()
        client.post.return_value = _response(200, {"uploadUrl": "https://upload.example/session"})
        client.put.side_effect = _graph_error(400)

        with pytest.raises(UploadSessionError):
            upload_file(client, "drive-1", "parent-1", local_path, "small.bin", "fail")

        client.get.assert_not_called()
        client.delete.assert_called_once()

    def test_final_chunk_conflict_fail_raises_file_already_exists(self, tmp_path):
        size = 5
        local_path = _make_sparse_file(tmp_path / "small.bin", SIMPLE_UPLOAD_THRESHOLD + size)
        client = MagicMock()
        client.post.return_value = _response(200, {"uploadUrl": "https://upload.example/session"})
        client.put.side_effect = [
            _response(202),
            _graph_error(409, "nameAlreadyExists"),
        ]

        with pytest.raises(FileAlreadyExistsError, match="small.bin"):
            upload_file(client, "drive-1", "parent-1", local_path, "small.bin", "fail")

        client.delete.assert_not_called()

    @pytest.mark.parametrize("conflict_behavior", ["replace", "rename"])
    def test_final_chunk_conflict_replace_or_rename_raises_clear_retryable_error(self, tmp_path, conflict_behavior):
        size = 5
        local_path = _make_sparse_file(tmp_path / "small.bin", SIMPLE_UPLOAD_THRESHOLD + size)
        client = MagicMock()
        client.post.return_value = _response(200, {"uploadUrl": "https://upload.example/session"})
        client.put.side_effect = [
            _response(202),
            _graph_error(409, "nameAlreadyExists"),
        ]

        with pytest.raises(UploadSessionError, match="late name conflict"):
            upload_file(client, "drive-1", "parent-1", local_path, "small.bin", conflict_behavior)
