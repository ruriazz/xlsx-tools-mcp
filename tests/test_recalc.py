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
    def run(cmd, capture_output=False, timeout=None):
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
    def run(cmd, capture_output=False, timeout=None):
        return SimpleNamespace(returncode=7, stderr=b"boom")

    monkeypatch.setattr(recalc, "find_soffice", lambda: "/fake/soffice")
    monkeypatch.setattr(recalc.subprocess, "run", run)

    result = recalc.recalculate(workbook_path)
    assert result.success is False
    assert "exited with code 7" in result.message


def test_timeout(workbook_path, monkeypatch):
    def run(cmd, capture_output=False, timeout=None):
        raise subprocess.TimeoutExpired(cmd, timeout)

    monkeypatch.setattr(recalc, "find_soffice", lambda: "/fake/soffice")
    monkeypatch.setattr(recalc.subprocess, "run", run)

    result = recalc.recalculate(workbook_path)
    assert result.success is False
    assert "timed out" in result.message