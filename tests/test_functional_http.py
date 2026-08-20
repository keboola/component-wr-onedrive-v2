"""Functional (datadir-style) tests for keboola.wr-onedrive-v2 (plan Task 10).

Design spec: ``docs/superpowers/specs/2026-08-17-wr-onedrive-v2-design.md`` §7 "Testing" —
"Datadir tests (single merged ``config.json`` shape ... state files row-scoped)". Unlike every
other test module in this suite (which mocks the component's own collaborators — ``GraphClient``,
``resolve_drive_id``, ``upload_file``, etc. — see ``tests/test_run_modes.py``), these tests run
the **real** production code path (``Component.run()``, drive resolution, the uploader, the Excel
writer) end-to-end and mock only the HTTP transport boundary: ``requests.Session.request`` is
patched with a small fake router (:class:`GraphFake`) that serves realistic Microsoft Graph JSON
payloads, keyed by method + URL. This is possible because every HTTP call in the production code
goes through exactly one seam — ``GraphClient.request`` calls ``self._session.request(...)``
directly (never the ``session.get``/``.post`` convenience wrappers), and
``RefreshTokenProvider._request_token`` calls ``self._session.post(...)``, which — in the real
``requests`` library — itself delegates to ``Session.request`` internally. Patching
``requests.Session.request`` at the class level therefore intercepts both, without needing either
class to accept an injected session for these tests.

**Why not ``keboola.datadirtest``** (a dev dependency, and the idiom the plan text suggested
checking for first): that library's ``TestDataDir.run_component`` executes ``src/component.py``
via ``runpy.run_path(..., run_name="__main__")``, i.e. it re-enters the script's own
``if __name__ == "__main__":`` block, whose ``except UserException: sys.exit(1)`` /
``except Exception: sys.exit(2)`` handlers raise ``SystemExit``. Verified empirically here: a
``SystemExit`` raised *inside* a ``unittest.TestCase`` test method is not caught by
``unittest``'s own ``_Outcome.testPartExecutor`` (it only catches ``Exception``, not every
``BaseException``), so it propagates out of ``TestCase.run()`` and aborts the *entire*
``TestSuite`` batch — not just the one failing-by-design test directory — the moment any single
scenario is supposed to exit non-zero. That makes it unsuitable for a suite where most of the
required scenarios are deliberate ``UserException``/exit-1 cases sitting side by side with
exit-0 happy paths in the same run. Instead, ``comp.run()`` is called in-process and asserted via
``pytest.raises(UserException)`` (exit 1) or a clean return (exit 0) — the same idiom already
used by ``tests/test_run_modes.py``/``tests/test_component.py``, just with HTTP mocked instead of
the component's own collaborators. ``TestEntrypointExitCodeMapping`` below additionally runs
``src/component.py`` as a real subprocess for one scenario, to assert the literal process exit
code (1) end-to-end without touching the ``keboola.datadirtest``/``SystemExit`` hazard at all.
"""

import json
import os
import re
import subprocess
import sys
from dataclasses import dataclass, field
from pathlib import Path
from unittest import mock

import pytest
import requests
from keboola.component.exceptions import UserException

from client.excel_writer import XLSX_MIME_TYPE
from component import Component

BASE_URL = "https://graph.microsoft.com/v1.0"
_TOKEN_URL_RE = re.compile(r"^https://login\.microsoftonline\.com/[^/]+/oauth2/v2\.0/token$")

_SRC_DIR = Path(__file__).resolve().parents[1] / "src"
_COMPONENT_SCRIPT = _SRC_DIR / "component.py"


def _graph_url(path: str) -> str:
    return f"{BASE_URL}{path}"


@dataclass
class FakeResponse:
    """A minimal stand-in for ``requests.Response`` — just enough for ``GraphClient``."""

    status_code: int
    payload: object = None
    headers: dict = field(default_factory=dict)

    @property
    def ok(self) -> bool:
        return 200 <= self.status_code < 300

    def json(self):
        return self.payload

    @property
    def text(self) -> str:
        return json.dumps(self.payload) if self.payload is not None else ""


class GraphFake:
    """Fake HTTP transport routed through a patched ``requests.Session.request``.

    Routes are registered with :meth:`add` (``matcher`` is an exact URL string or a compiled
    regex, matched with ``re.fullmatch``); the Microsoft identity token endpoint is handled
    automatically — every scenario gets a working, rotating refresh token without registering it
    explicitly. Every call (including the token exchange) is recorded in :attr:`calls` for
    assertions on exactly which Graph calls a scenario made.
    """

    def __init__(self, rotated_refresh_token: str = "rotated-refresh-token"):
        self._rules: list[tuple[str, object, object]] = []
        self.calls: list[tuple[str, str, dict]] = []
        self.rotated_refresh_token = rotated_refresh_token

    def add(self, method: str, matcher, response) -> GraphFake:
        self._rules.append((method, matcher, response))
        return self

    def __call__(self, method: str, url: str, **kwargs) -> FakeResponse:
        self.calls.append((method, url, kwargs))
        if method == "POST" and _TOKEN_URL_RE.fullmatch(url):
            return FakeResponse(
                200,
                {
                    "access_token": "access-1",
                    "refresh_token": self.rotated_refresh_token,
                    "expires_in": 3599,
                },
            )
        for rule_method, matcher, response in self._rules:
            if rule_method != method:
                continue
            matched = matcher.fullmatch(url) if isinstance(matcher, re.Pattern) else matcher == url
            if not matched:
                continue
            return response(**kwargs) if callable(response) else response
        raise AssertionError(f"GraphFake: no rule registered for {method} {url} (kwargs={kwargs!r})")

    def calls_for(self, method: str) -> list[tuple[str, dict]]:
        return [(url, kwargs) for called_method, url, kwargs in self.calls if called_method == method]


def _oauth_credentials(refresh_token: str = "refresh-config") -> dict:
    return {
        "id": "oauth-1",
        "created": "2026-01-01",
        "appKey": "client-1",
        "#appSecret": "secret-1",
        "oauthVersion": "2.0",
        "#data": json.dumps({"refresh_token": refresh_token}),
    }


_UNSET = object()


def _build_data_dir(
    tmp_path,
    parameters: dict,
    *,
    files: dict[str, bytes] | None = None,
    tables: dict[str, str] | None = None,
    oauth=_UNSET,
) -> Path:
    """Build a ``KBC_DATADIR``-style fixture directory (datadir idiom — merged ``config.json``,
    row-scoped state), returning the ``data`` directory path."""
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

    config: dict = {"parameters": parameters, "action": "run"}
    resolved_oauth = _oauth_credentials() if oauth is _UNSET else oauth
    if resolved_oauth is not None:
        config["authorization"] = {"oauth_api": {"credentials": resolved_oauth}}
    (data_dir / "config.json").write_text(json.dumps(config))
    return data_dir


def _build_component(tmp_path, parameters: dict, **kwargs) -> Component:
    data_dir = _build_data_dir(tmp_path, parameters, **kwargs)
    with mock.patch.dict(os.environ, {"KBC_DATADIR": str(data_dir)}):
        return Component()


def _run(comp: Component, graph: GraphFake) -> None:
    """Run the real ``Component.run()`` with HTTP routed through ``graph`` (plan Task 10)."""
    with mock.patch.object(requests.Session, "request", side_effect=graph):
        comp.run()


def _xlsx_item(file_id: str, name: str) -> dict:
    return {"id": file_id, "name": name, "file": {"mimeType": XLSX_MIME_TYPE}, "parentReference": {}}


def _not_found(code: str = "itemNotFound", message: str = "The resource could not be found.") -> FakeResponse:
    return FakeResponse(404, {"error": {"code": code, "message": message}})


# ---------------------------------------------------------------------------------------------
# 1. Happy path per mode
# ---------------------------------------------------------------------------------------------


class TestHappyPathFileMode:
    def test_uploads_two_files_via_simple_put_and_persists_rotated_token(self, tmp_path):
        parameters = {
            "mode": "file",
            "account": {"account_type": "private_onedrive"},
            "destination": {"folder_path": "uploads", "conflict_behavior": "replace"},
        }
        comp = _build_component(tmp_path, parameters, files={"a.txt": b"aaa", "b.txt": b"bbb"})
        graph = GraphFake()
        graph.add("GET", _graph_url("/me/drive"), FakeResponse(200, {"id": "drive-1"}))
        graph.add("GET", _graph_url("/drives/drive-1/root:/uploads"), _not_found())
        graph.add("POST", _graph_url("/drives/drive-1/root/children"), FakeResponse(201, {"id": "folder-1"}))
        upload_pattern = re.compile(re.escape(_graph_url("/drives/drive-1/items/folder-1:/")) + r"[ab]\.txt:/content")
        graph.add("PUT", upload_pattern, FakeResponse(200, {"id": "item-x"}))

        _run(comp, graph)  # must not raise: exit 0

        put_calls = graph.calls_for("PUT")
        assert len(put_calls) == 2
        uploaded_names = {re.search(r":/(\w+\.txt):/content$", url).group(1) for url, _ in put_calls}
        assert uploaded_names == {"a.txt", "b.txt"}
        for _url, kwargs in put_calls:
            assert kwargs["params"] == {"@microsoft.graph.conflictBehavior": "replace"}

        state = json.loads((tmp_path / "data" / "out" / "state.json").read_text())
        payload = json.loads(state["#refreshed_auth_data"])
        assert payload["refresh_token"] == "rotated-refresh-token"


class TestHappyPathCsvMode:
    def test_uploads_single_table_as_passthrough_csv(self, tmp_path):
        parameters = {
            "mode": "table_csv",
            "account": {"account_type": "private_onedrive"},
            "destination": {"conflict_behavior": "fail"},
        }
        comp = _build_component(tmp_path, parameters, tables={"mytable": "id,name\n1,a\n2,b\n"})
        graph = GraphFake()
        graph.add("GET", _graph_url("/me/drive"), FakeResponse(200, {"id": "drive-1"}))
        graph.add("GET", _graph_url("/drives/drive-1/root"), FakeResponse(200, {"id": "root-1"}))
        upload_url = _graph_url("/drives/drive-1/items/root-1:/mytable.csv:/content")
        captured: dict = {}

        def _capture_put(**kwargs):
            captured["body"] = kwargs["data"].read()
            return FakeResponse(200, {"id": "item-1"})

        graph.add("PUT", upload_url, _capture_put)

        _run(comp, graph)  # must not raise: exit 0

        assert captured["body"] == b"id,name\n1,a\n2,b\n"  # streamed as-is, no rewrite


class TestHappyPathExcelModeOverwrite:
    """Excel mode, ``append=False`` (overwrite = clear + write from A1), business account."""

    def test_clears_and_writes_from_a1(self, tmp_path):
        parameters = {
            "mode": "table_excel",
            "account": {"account_type": "onedrive_for_business", "tenant_id": "tenant-1"},
            "workbook": {"path": "/book.xlsx"},
            "worksheet": {"name": "Sheet1"},
            "append": False,
        }
        comp = _build_component(tmp_path, parameters, tables={"mytable": "id,name\n1,a\n"})
        graph = _excel_graph_with_existing_sheet()

        _run(comp, graph)  # must not raise: exit 0

        clear_calls = graph.calls_for("POST")
        clear_urls = [url for url, _ in clear_calls if url.endswith("/range/clear")]
        assert len(clear_urls) == 1

        patch_calls = graph.calls_for("PATCH")
        assert len(patch_calls) == 1
        patch_url, patch_kwargs = patch_calls[0]
        assert patch_url.endswith("/range(address='A1:B2')")
        assert patch_kwargs["json"] == {"values": [["id", "name"], ["1", "a"]]}

        # session opened and closed exactly once around the write.
        session_urls = [url for url, _ in graph.calls_for("POST") if "workbook/createSession" in url]
        close_urls = [url for url, _ in graph.calls_for("POST") if "workbook/closeSession" in url]
        assert len(session_urls) == 1
        assert len(close_urls) == 1


def _excel_graph_with_existing_sheet(existing_header: list[str] | None = None) -> GraphFake:
    """Common Excel-mode routing: an existing workbook `/book.xlsx` with an existing `Sheet1`.

    ``existing_header`` (when given) also wires up `usedRange`/header-row responses for append
    scenarios; overwrite scenarios never call those endpoints (range/clear is unconditional).
    """
    graph = GraphFake()
    graph.add("GET", _graph_url("/me/drive"), FakeResponse(200, {"id": "drive-1"}))
    graph.add("GET", _graph_url("/drives/drive-1/root:/book.xlsx"), FakeResponse(200, _xlsx_item("file-1", "book.xlsx")))
    graph.add(
        "POST",
        _graph_url("/drives/drive-1/items/file-1/workbook/createSession"),
        FakeResponse(201, {"id": "session-1"}),
    )
    graph.add(
        "GET",
        _graph_url("/drives/drive-1/items/file-1/workbook/worksheets"),
        FakeResponse(200, {"value": [{"id": "sheet-1", "name": "Sheet1", "position": 0, "visibility": "Visible"}]}),
    )
    graph.add(
        "POST",
        _graph_url("/drives/drive-1/items/file-1/workbook/worksheets/sheet-1/range/clear"),
        FakeResponse(200, {}),
    )
    graph.add(
        "POST",
        _graph_url("/drives/drive-1/items/file-1/workbook/closeSession"),
        FakeResponse(204, {}),
    )
    if existing_header is not None:
        graph.add(
            "GET",
            _graph_url("/drives/drive-1/items/file-1/workbook/worksheets/sheet-1/range/usedRange(valuesOnly=true)"),
            FakeResponse(200, {"address": "Sheet1!A1:B3"}),
        )
        graph.add(
            "GET",
            _graph_url(
                "/drives/drive-1/items/file-1/workbook/worksheets/sheet-1/range/usedRange(valuesOnly=true)/row(row=0)"
            ),
            FakeResponse(200, {"text": [existing_header]}),
        )
    _wire_patch_catch_all(graph)
    return graph


def _wire_patch_catch_all(graph: GraphFake) -> None:
    """Any ``range(address=...)`` PATCH succeeds — the address itself varies per scenario and is
    asserted from ``graph.calls_for("PATCH")`` rather than pre-registered per exact address."""
    pattern = re.compile(re.escape(_graph_url("/drives/drive-1/items/file-1/workbook/worksheets/sheet-1/range(address='")) + r".+'\)")
    graph.add("PATCH", pattern, FakeResponse(200, {}))


# ---------------------------------------------------------------------------------------------
# 2. Excel append vs overwrite
# ---------------------------------------------------------------------------------------------


class TestExcelAppendOffsetsFromUsedRange:
    def test_append_skips_header_and_starts_after_used_range(self, tmp_path):
        parameters = {
            "mode": "table_excel",
            "account": {"account_type": "onedrive_for_business", "tenant_id": "tenant-1"},
            "workbook": {"path": "/book.xlsx"},
            "worksheet": {"name": "Sheet1"},
            "append": True,
        }
        comp = _build_component(tmp_path, parameters, tables={"mytable": "id,name\n1,a\n"})
        graph = _excel_graph_with_existing_sheet(existing_header=["id", "name"])

        _run(comp, graph)  # must not raise: exit 0

        clear_urls = [url for url, _ in graph.calls_for("POST") if url.endswith("/range/clear")]
        assert clear_urls == []  # append never clears

        patch_calls = graph.calls_for("PATCH")
        assert len(patch_calls) == 1
        patch_url, patch_kwargs = patch_calls[0]
        assert patch_url.endswith("/range(address='A4:B4')")  # offset from usedRange's A1:B3
        assert patch_kwargs["json"] == {"values": [["1", "a"]]}  # header skipped (headers match)

    def test_append_warns_on_header_mismatch_but_still_appends(self, tmp_path, caplog):
        parameters = {
            "mode": "table_excel",
            "account": {"account_type": "onedrive_for_business", "tenant_id": "tenant-1"},
            "workbook": {"path": "/book.xlsx"},
            "worksheet": {"name": "Sheet1"},
            "append": True,
        }
        comp = _build_component(tmp_path, parameters, tables={"mytable": "id,name\n1,a\n"})
        graph = _excel_graph_with_existing_sheet(existing_header=["identifier", "full_name"])

        with caplog.at_level("WARNING"):
            _run(comp, graph)  # must not raise: exit 0

        assert "Headers mismatch" in caplog.text
        patch_calls = graph.calls_for("PATCH")
        assert len(patch_calls) == 1
        assert patch_calls[0][1]["json"] == {"values": [["1", "a"]]}


# ---------------------------------------------------------------------------------------------
# 3. Empty CSV in Excel mode
# ---------------------------------------------------------------------------------------------


class TestExcelEmptyCsvInput:
    def test_empty_csv_logs_warning_exits_cleanly_and_issues_no_range_patch(self, tmp_path, caplog):
        parameters = {
            "mode": "table_excel",
            "account": {"account_type": "onedrive_for_business", "tenant_id": "tenant-1"},
            "workbook": {"path": "/book.xlsx"},
            "worksheet": {"name": "Sheet1"},
        }
        comp = _build_component(tmp_path, parameters, tables={"empty": ""})
        graph = _excel_graph_with_existing_sheet()

        with caplog.at_level("WARNING"):
            _run(comp, graph)  # must not raise: exit 0, sheet untouched (v1 parity)

        assert 'Ignored empty CSV file "empty".' in caplog.text
        assert graph.calls_for("PATCH") == []
        clear_urls = [url for url, _ in graph.calls_for("POST") if url.endswith("/range/clear")]
        assert clear_urls == []  # `write_table` returns before touching the sheet at all


# ---------------------------------------------------------------------------------------------
# 4. CSV cardinality
# ---------------------------------------------------------------------------------------------


class TestCsvCardinality:
    """Mode 'worksheet' (the pre-merge 'table_excel') still requires exactly one input table —
    unaffected by Change A's file-mode merge, which only touches mode 'file'."""

    def test_zero_tables_raises_v1_parity_user_exception(self, tmp_path):
        parameters = {
            "mode": "table_excel",
            "account": {"account_type": "onedrive_for_business", "tenant_id": "tenant-1"},
            "workbook": {"path": "/book.xlsx"},
            "worksheet": {"name": "Sheet1"},
        }
        comp = _build_component(tmp_path, parameters, tables={})
        graph = GraphFake()

        with pytest.raises(UserException, match=re.escape('No CSV file found in "/data/in/tables".')):
            _run(comp, graph)  # exit 1
        assert graph.calls_for("GET") == []  # fails before any Graph call (no drive_id resolution either)

    def test_multiple_tables_raises_v1_parity_user_exception_naming_both(self, tmp_path):
        parameters = {
            "mode": "table_excel",
            "account": {"account_type": "onedrive_for_business", "tenant_id": "tenant-1"},
            "workbook": {"path": "/book.xlsx"},
            "worksheet": {"name": "Sheet1"},
        }
        comp = _build_component(tmp_path, parameters, tables={"a.csv": "id\n1\n", "b.csv": "id\n2\n"})
        graph = GraphFake()
        expected = re.escape('Expected one CSV file, found multiple: "a.csv", "b.csv".')

        with pytest.raises(UserException, match=expected):
            _run(comp, graph)  # exit 1


class TestFileModeMergedTableInput:
    """Change A: mode 'file' now processes both input mappings end-to-end (files uploaded as-is,
    mapped tables written as CSV) — a real HTTP-mocked (not just mocked-collaborator) proof that
    both actually reach Graph in the same run."""

    def test_files_and_tables_together_are_both_uploaded(self, tmp_path):
        parameters = {"mode": "file", "account": {"account_type": "private_onedrive"}, "destination": {}}
        comp = _build_component(
            tmp_path, parameters, files={"a.txt": b"aaa"}, tables={"mytable": "id,name\n1,a\n"}
        )
        graph = GraphFake()
        graph.add("GET", _graph_url("/me/drive"), FakeResponse(200, {"id": "drive-1"}))
        graph.add("GET", _graph_url("/drives/drive-1/root"), FakeResponse(200, {"id": "root-1"}))
        graph.add("PUT", _graph_url("/drives/drive-1/items/root-1:/a.txt:/content"), FakeResponse(200, {"id": "item-1"}))
        graph.add(
            "PUT", _graph_url("/drives/drive-1/items/root-1:/mytable.csv:/content"), FakeResponse(200, {"id": "item-2"})
        )

        _run(comp, graph)  # must not raise: exit 0

        put_urls = {url for url, _ in graph.calls_for("PUT")}
        assert put_urls == {
            _graph_url("/drives/drive-1/items/root-1:/a.txt:/content"),
            _graph_url("/drives/drive-1/items/root-1:/mytable.csv:/content"),
        }


# ---------------------------------------------------------------------------------------------
# 5. Missing OAuth
# ---------------------------------------------------------------------------------------------


class TestMissingOAuth:
    def test_missing_oauth_authorization_raises_user_exception_before_any_network_call(self, tmp_path):
        parameters = {"mode": "file", "account": {"account_type": "private_onedrive"}, "destination": {}}
        comp = _build_component(tmp_path, parameters, files={"a.txt": b"aaa"}, oauth=None)
        graph = GraphFake()

        with pytest.raises(UserException, match="not authorized"):
            _run(comp, graph)  # exit 1

        assert graph.calls == []  # fails before the token provider (and thus any HTTP call) exists
        assert not (tmp_path / "data" / "out" / "state.json").exists()


# ---------------------------------------------------------------------------------------------
# 6. Excel mode gated off for private_onedrive
# ---------------------------------------------------------------------------------------------


class TestExcelPrivateOnedriveGate:
    def test_private_onedrive_account_raises_user_exception(self, tmp_path):
        parameters = {
            "mode": "table_excel",
            "account": {"account_type": "private_onedrive"},
            "workbook": {"path": "/book.xlsx"},
            "worksheet": {"name": "Sheet1"},
        }
        comp = _build_component(tmp_path, parameters, tables={"mytable": "id,name\n1,a\n"})
        graph = GraphFake()
        graph.add("GET", _graph_url("/me/drive"), FakeResponse(200, {"id": "drive-1"}))

        with pytest.raises(UserException, match="private_onedrive"):
            _run(comp, graph)  # exit 1

        # IMPORTANT-2 (phase 8 audit): Excel mode never resolves a drive id at all (it targets
        # `workbook.{path,drive_id,file_id}` instead) — the account-type gate must raise before
        # any Graph call, not just before an Excel-specific one.
        assert graph.calls_for("GET") == []


# ---------------------------------------------------------------------------------------------
# 7. Conflict behavior `fail` on an existing file
# ---------------------------------------------------------------------------------------------


class TestConflictFailOnExistingFile:
    def test_existing_file_with_conflict_fail_raises_user_exception_naming_the_file(self, tmp_path):
        parameters = {
            "mode": "file",
            "account": {"account_type": "private_onedrive"},
            "destination": {},  # conflict_behavior default is "fail"
        }
        comp = _build_component(tmp_path, parameters, files={"report.csv": b"a,b\n1,2\n"})
        graph = GraphFake()
        graph.add("GET", _graph_url("/me/drive"), FakeResponse(200, {"id": "drive-1"}))
        graph.add("GET", _graph_url("/drives/drive-1/root"), FakeResponse(200, {"id": "root-1"}))
        graph.add(
            "PUT",
            _graph_url("/drives/drive-1/items/root-1:/report.csv:/content"),
            FakeResponse(409, {"error": {"code": "nameAlreadyExists", "message": "An item with the same name already exists."}}),
        )

        with pytest.raises(UserException, match=re.escape("'report.csv' already exists")):
            _run(comp, graph)  # exit 1


# ---------------------------------------------------------------------------------------------
# 8. Rotated refresh token persisted to state, including on a mid-run failure
# ---------------------------------------------------------------------------------------------


class TestTokenRotationPersistence:
    def test_rotated_token_persisted_even_when_a_later_upload_fails(self, tmp_path):
        parameters = {
            "mode": "file",
            "account": {"account_type": "private_onedrive"},
            "destination": {"conflict_behavior": "replace"},
        }
        comp = _build_component(tmp_path, parameters, files={"a.txt": b"aaa", "b.txt": b"bbb"})
        graph = GraphFake(rotated_refresh_token="rotated-after-failure")
        graph.add("GET", _graph_url("/me/drive"), FakeResponse(200, {"id": "drive-1"}))
        graph.add("GET", _graph_url("/drives/drive-1/root"), FakeResponse(200, {"id": "root-1"}))
        graph.add(
            "PUT",
            _graph_url("/drives/drive-1/items/root-1:/a.txt:/content"),
            FakeResponse(200, {"id": "item-a"}),
        )
        graph.add(
            "PUT",
            _graph_url("/drives/drive-1/items/root-1:/b.txt:/content"),
            FakeResponse(403, {"error": {"code": "accessDenied", "message": "Access denied."}}),
        )

        with pytest.raises(UserException):
            _run(comp, graph)  # exit 1 (mid-run failure)

        state = json.loads((tmp_path / "data" / "out" / "state.json").read_text())
        payload = json.loads(state["#refreshed_auth_data"])
        assert payload["refresh_token"] == "rotated-after-failure"


# ---------------------------------------------------------------------------------------------
# Literal process exit-code mapping (subprocess — deliberately independent of the in-process
# `pytest.raises(UserException)` idiom used above; see the module docstring for why a full
# `keboola.datadirtest`-style run of every scenario is not used instead).
# ---------------------------------------------------------------------------------------------


class TestEntrypointExitCodeMapping:
    def test_missing_oauth_exits_with_code_1(self, tmp_path):
        parameters = {"mode": "file", "account": {"account_type": "private_onedrive"}, "destination": {}}
        data_dir = _build_data_dir(tmp_path, parameters, files={"a.txt": b"aaa"}, oauth=None)

        result = subprocess.run(
            [sys.executable, str(_COMPONENT_SCRIPT)],
            env={**os.environ, "KBC_DATADIR": str(data_dir), "PYTHONPATH": str(_SRC_DIR)},
            capture_output=True,
            text=True,
            timeout=30,
            check=False,
        )

        assert result.returncode == 1, f"stdout={result.stdout!r} stderr={result.stderr!r}"
        assert "not authorized" in result.stderr
