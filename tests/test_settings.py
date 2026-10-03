from pathlib import Path

import pytest

from xlsx_tools_mcp.errors import FileNotConfiguredError
from xlsx_tools_mcp.settings import CONFIGURED_FILES, _parse_configured_files, resolve_path


def _resolved(p):
    return str(Path(p).expanduser().resolve())


def test_parse_configured_files_named():
    parsed = _parse_configured_files("a=/x/y/z.xlsx,b=/m/n.xlsx")
    assert parsed == {"a": _resolved("/x/y/z.xlsx"), "b": _resolved("/m/n.xlsx")}


def test_parse_configured_files_bare_path_alias_is_filename():
    parsed = _parse_configured_files("/p/q.xlsx")
    assert parsed == {"q.xlsx": _resolved("/p/q.xlsx")}


def test_parse_configured_files_whitespace_tolerance():
    parsed = _parse_configured_files(" a = /x/b.xlsx ")
    assert parsed == {"a": _resolved("/x/b.xlsx")}


def test_resolve_omitted_single_configured(monkeypatch):
    monkeypatch.setattr("xlsx_tools_mcp.settings.CONFIGURED_FILES", {"a": "/x/y.xlsx"})
    assert resolve_path(None) == "/x/y.xlsx"


def test_resolve_omitted_none_configured(monkeypatch):
    monkeypatch.setattr("xlsx_tools_mcp.settings.CONFIGURED_FILES", {})
    with pytest.raises(FileNotConfiguredError):
        resolve_path(None)


def test_resolve_omitted_multiple_configured(monkeypatch):
    monkeypatch.setattr("xlsx_tools_mcp.settings.CONFIGURED_FILES", {"a": "/x/a.xlsx", "b": "/x/b.xlsx"})
    with pytest.raises(FileNotConfiguredError) as excinfo:
        resolve_path(None)
    assert "a" in str(excinfo.value) and "b" in str(excinfo.value)


def test_resolve_alias_match(monkeypatch):
    monkeypatch.setattr("xlsx_tools_mcp.settings.CONFIGURED_FILES", {"report": "/data/report.xlsx"})
    assert resolve_path("report") == "/data/report.xlsx"


def test_resolve_filename_match(monkeypatch):
    monkeypatch.setattr("xlsx_tools_mcp.settings.CONFIGURED_FILES", {"report.xlsx": "/data/report.xlsx"})
    assert resolve_path("report.xlsx") == "/data/report.xlsx"


def test_resolve_unknown_path_returned_as_is(monkeypatch):
    monkeypatch.setattr("xlsx_tools_mcp.settings.CONFIGURED_FILES", {"a": "/x/a.xlsx"})
    assert resolve_path("/some/other.xlsx") == "/some/other.xlsx"
from xlsx_tools_mcp.errors import AccessDeniedError
import os

def test_resolve_path_allowed_dirs(monkeypatch, tmp_path):
    allowed_dir = str(tmp_path / "allowed")
    os.makedirs(allowed_dir, exist_ok=True)
    monkeypatch.setattr("xlsx_tools_mcp.settings.ALLOWED_DIRS", [allowed_dir])
    
    file_inside = str(tmp_path / "allowed" / "file.xlsx")
    assert resolve_path(file_inside) == file_inside

    file_outside = str(tmp_path / "outside" / "file.xlsx")
    with pytest.raises(AccessDeniedError):
        resolve_path(file_outside)

def test_resolve_path_allowed_dirs_empty(monkeypatch, tmp_path):
    monkeypatch.setattr("xlsx_tools_mcp.settings.ALLOWED_DIRS", [])
    file_outside = str(tmp_path / "outside" / "file.xlsx")
    # should just return as is
    assert resolve_path(file_outside) == file_outside


def test_server_tools_enforce_allowed_dirs(monkeypatch, tmp_path):
    from xlsx_tools_mcp import server
    allowed_dir = str(tmp_path / "allowed")
    os.makedirs(allowed_dir, exist_ok=True)
    monkeypatch.setattr("xlsx_tools_mcp.settings.ALLOWED_DIRS", [allowed_dir])

    # create_workbook outside allowed dir must raise
    with pytest.raises(ValueError, match="Access denied"):
        server.create_workbook(str(tmp_path / "outside" / "new.xlsx"))

    # restore_backup outside allowed dir must raise
    with pytest.raises(ValueError, match="Access denied"):
        server.restore_backup(str(tmp_path / "outside" / "a.bak"), str(tmp_path / "outside" / "target.xlsx"))

