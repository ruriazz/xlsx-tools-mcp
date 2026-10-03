import os
import shutil
import subprocess
from pathlib import Path
from types import SimpleNamespace

import openpyxl
import pytest

from xlsx_tools_mcp import recalc


def _make_fixture(path: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["item", "qty"])
    ws.append(["apple", 10])
    wb.save(path)
    return path


@pytest.fixture
def workbook_path(tmp_path):
    return _make_fixture(str(tmp_path / "book.xlsx"))


def _fake_run(make_output):
    """A subprocess.run fake with an optional output-file writer.

    Returns (fn, proc). For success-style calls proc.returncode is 0; errors
    can be simulated by passing a monkeypatched proc instead of `make_output`.
    """
    def run(cmd, capture_output=False, timeout=None, **kwargs):
        outdir = Path(cmd[cmd.index("--outdir") + 1])
        src = Path(cmd[-1])
        if make_output is not None:
            make_output(outdir, src)
        return SimpleNamespace(returncode=0, stderr=b"")

    return run


def test_no_soffice_not_found(workbook_path, monkeypatch):
    monkeypatch.setattr(recalc, "find_soffice", lambda: None)
    result = recalc.recalculate(workbook_path)
    assert result.success is False
    assert "not found" in result.message
    assert "winget install TheDocumentFoundation.LibreOffice" in result.message


def test_success_output_created(workbook_path, monkeypatch):
    fixture = _make_fixture(str(Path(workbook_path).with_name("fixture.xlsx")))

    def make_output(outdir, src):
        shutil.copy(fixture, outdir / f"{src.stem}.xlsx")

    monkeypatch.setattr(recalc, "find_soffice", lambda: "/fake/soffice")
    monkeypatch.setattr(recalc.subprocess, "run", _fake_run(make_output))

    result = recalc.recalculate(workbook_path)
    assert result.success is True
    assert result.errors_found == []
    assert "Recalculated" in result.message


def test_scan_formula_errors_roundtrips(workbook_path, monkeypatch):
    def make_output(outdir, src):
        shutil.copy(workbook_path, outdir / f"{src.stem}.xlsx")

    monkeypatch.setattr(recalc, "find_soffice", lambda: "/fake/soffice")
    monkeypatch.setattr(recalc.subprocess, "run", _fake_run(make_output))
    monkeypatch.setattr(
        recalc, "scan_formula_errors", lambda _: [{"sheet": "Data", "cell": "A1", "error": "#DIV/0!"}]
    )

    result = recalc.recalculate(workbook_path)
    assert result.success is True
    assert result.errors_found == [{"sheet": "Data", "cell": "A1", "error": "#DIV/0!"}]


def test_nonzero_returncode(workbook_path, monkeypatch):
    def run(cmd, capture_output=False, timeout=None, **kwargs):
        return SimpleNamespace(returncode=7, stderr=b"boom")

    monkeypatch.setattr(recalc, "find_soffice", lambda: "/fake/soffice")
    monkeypatch.setattr(recalc.subprocess, "run", run)

    result = recalc.recalculate(workbook_path)
    assert result.success is False
    assert "exited with code 7" in result.message


def test_timeout(workbook_path, monkeypatch):
    def run(cmd, capture_output=False, timeout=None, **kwargs):
        raise subprocess.TimeoutExpired(cmd, timeout)

    monkeypatch.setattr(recalc, "find_soffice", lambda: "/fake/soffice")
    monkeypatch.setattr(recalc.subprocess, "run", run)

    result = recalc.recalculate(workbook_path)
    assert result.success is False
    assert "timed out" in result.message


def test_find_soffice_windows_candidate(monkeypatch):
    monkeypatch.setattr(recalc, "find_soffice", recalc._orig_find_soffice)
    monkeypatch.setattr(recalc.shutil, "which", lambda _: None)
    monkeypatch.setattr(recalc.sys, "platform", "win32")
    prog_files = "C:\\Program Files"
    monkeypatch.setenv("ProgramFiles", prog_files)
    expected_path = os.path.join(prog_files, "LibreOffice", "program", "soffice.exe")

    def mock_isfile(path):
        return path == expected_path

    monkeypatch.setattr(recalc.os.path, "isfile", mock_isfile)
    assert recalc.find_soffice() == expected_path


def test_find_soffice_windows_shutil_which(monkeypatch):
    monkeypatch.setattr(recalc, "find_soffice", recalc._orig_find_soffice)

    def mock_which(name):
        if name == "soffice.exe":
            return r"C:\Custom\soffice.exe"
        return None

    monkeypatch.setattr(recalc.shutil, "which", mock_which)
    assert recalc.find_soffice() == r"C:\Custom\soffice.exe"


def test_recalculate_user_installation_and_window_flags(workbook_path, monkeypatch):
    fixture = _make_fixture(str(Path(workbook_path).with_name("fixture2.xlsx")))
    captured_calls = []

    def mock_run(cmd, capture_output=False, timeout=None, **kwargs):
        captured_calls.append({"cmd": cmd, "kwargs": kwargs})
        outdir = Path(cmd[cmd.index("--outdir") + 1])
        src = Path(cmd[-1])
        shutil.copy(fixture, outdir / f"{src.stem}.xlsx")
        return SimpleNamespace(returncode=0, stderr=b"")

    monkeypatch.setattr(recalc, "find_soffice", lambda: "/fake/soffice")
    monkeypatch.setattr(recalc.sys, "platform", "win32")
    monkeypatch.setattr(recalc.subprocess, "run", mock_run)

    result = recalc.recalculate(workbook_path)
    assert result.success is True
    assert len(captured_calls) == 1
    call = captured_calls[0]

    cmd_str = " ".join(call["cmd"])
    assert "--nofirststartwizard" in call["cmd"]
    assert "-env:UserInstallation=file://" in cmd_str
    assert "creationflags" in call["kwargs"]


def test_recalculate_path_with_spaces(tmp_path, monkeypatch):
    folder_with_space = tmp_path / "My Folder With Spaces"
    folder_with_space.mkdir()
    book_path = folder_with_space / "Report 2026.xlsx"
    _make_fixture(str(book_path))

    captured_cmds = []

    def mock_run(cmd, capture_output=False, timeout=None, **kwargs):
        captured_cmds.append(cmd)
        outdir = Path(cmd[cmd.index("--outdir") + 1])
        src = Path(cmd[-1])
        shutil.copy(book_path, outdir / f"{src.stem}.xlsx")
        return SimpleNamespace(returncode=0, stderr=b"")

    monkeypatch.setattr(recalc, "find_soffice", lambda: "/fake/soffice")
    monkeypatch.setattr(recalc.subprocess, "run", mock_run)

    result = recalc.recalculate(str(book_path))
    assert result.success is True
    assert len(captured_cmds) == 1
    # Check that profile URI was properly encoded and input file was passed intact
    cmd = captured_cmds[0]
    profile_arg = [arg for arg in cmd if arg.startswith("-env:UserInstallation=")][0]
    assert profile_arg.startswith("-env:UserInstallation=file://")
    assert str(Path(cmd[-1])) == str(Path(cmd[cmd.index("--outdir") + 1]) / "Report 2026.xlsx")