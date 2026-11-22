import pytest

from xlsx_tools_mcp import ripple, server, settings
from xlsx_tools_mcp.io import writer

from .helpers import use_fake_libreoffice


@pytest.fixture
def book(tmp_path, monkeypatch):
    use_fake_libreoffice(monkeypatch)
    path = str(tmp_path / "r.xlsx")
    writer.create_workbook(path, ["S"])
    writer.write_cells(path, "S", [
        {"cell": "A1", "value": 2},
        {"cell": "B1", "formula": "=A1*10"},
        {"cell": "C1", "formula": "=B1+1"},
        {"cell": "D1", "formula": "=5"},
        {"cell": "E1", "value": "label"},
    ], backup=False)
    return path


def test_ripple_reports_dependents_not_written_cells(book):
    res = server.write_cells("S", [{"cell": "A1", "value": 3}], path=book, backup=False)
    rp = res["ripple"]
    assert rp["changed_count"] == 2
    cells = {(c["cell"], c["before"], c["after"]) for c in rp["sample"]}
    assert cells == {("B1", 20, 30), ("C1", 21, 31)}
    assert all(c["cell"] != "A1" for c in rp["sample"])
    assert rp["sample"][0]["delta"] == 10


def test_written_formula_cell_is_excluded(book):
    res = server.write_cells("S", [{"cell": "D1", "formula": "=6"}], path=book, backup=False)
    assert res["ripple"]["changed_count"] == 0


def test_structural_write_skips_ripple(book):
    res = server.insert_rows("S", 1, path=book, backup=False)
    assert "structural" in res["ripple"]["skipped"]


def test_ripple_skipped_without_libreoffice(book, monkeypatch):
    monkeypatch.setattr("xlsx_tools_mcp.recalc.find_soffice", lambda: None)
    res = server.write_cells("S", [{"cell": "A1", "value": 3}], path=book, backup=False)
    assert "LibreOffice" in res["ripple"]["skipped"]


def test_ripple_disabled(book, monkeypatch):
    monkeypatch.setattr(settings, "RIPPLE_ENABLED", False)
    res = server.write_cells("S", [{"cell": "A1", "value": 3}], path=book, backup=False)
    assert "ripple" not in res


def test_ripple_cell_cap(book, monkeypatch):
    monkeypatch.setattr(settings, "RIPPLE_MAX_CELLS", 1)
    res = server.write_cells("S", [{"cell": "A1", "value": 3}], path=book, backup=False)
    assert "RIPPLE_MAX_CELLS" in res["ripple"]["skipped"]


def test_no_ripple_for_style_only_change(book):
    res = server.set_cell_style("S", "A1", {"bold": True}, path=book, backup=False)
    assert "ripple" not in res


def test_diff_unknown_before_and_errors():
    before = {("S", "A1"): None, ("S", "B1"): ("n", "1"), ("S", "C1"): ("n", "1")}
    after = {("S", "A1"): ("n", "5"), ("S", "B1"): ("e", "#DIV/0!"), ("S", "C1"): ("n", "1.0000000000001")}
    out = ripple.diff(before, after, set())
    assert out["unknown_before"] == 1
    assert out["changed_count"] == 1
    assert out["sample"][0] == {"sheet": "S", "cell": "B1", "before": 1, "after": "#DIV/0!"}


def test_ripple_failure_never_fails_the_write(book, monkeypatch):
    def broken(path):
        raise ValueError("unreadable")

    monkeypatch.setattr(ripple, "snapshot", broken)
    res = server.write_cells("S", [{"cell": "A1", "value": 3}], path=book, backup=False)
    assert res["saved"] is True
    assert "could not read" in res["ripple"]["skipped"]
