import openpyxl
import pytest

from xlsx_tools_mcp import ledger, server, settings
from xlsx_tools_mcp.io import writer


@pytest.fixture
def book(tmp_path):
    path = str(tmp_path / "log.xlsx")
    writer.create_workbook(path, ["Data"])
    return path


@pytest.fixture
def enabled(monkeypatch):
    monkeypatch.setattr(settings, "CHANGE_LOG_ENABLED", True)


def test_disabled_by_default_creates_no_sheet(book):
    server.write_cells("Data", [{"cell": "A1", "value": 1}], path=book)
    assert ledger.LOG_SHEET not in openpyxl.load_workbook(book).sheetnames
    assert server.get_change_log(path=book)["enabled_in_file"] is False


def test_write_cells_logged_with_reason(book, enabled):
    server.write_cells("Data", [{"cell": "A1", "formula": "=1+1"}], path=book)
    server.write_cells("Data", [{"cell": "a1", "value": 5}], reason="fix tax", path=book)
    log = server.get_change_log(path=book)
    latest = log["entries"][0]
    assert latest["tool"] == "write_cells"
    assert latest["ref"] == "A1"
    assert latest["before"] == "=1+1"
    assert latest["after"] == 5
    assert latest["reason"] == "fix tax"
    assert len(log["entries"]) == 2

    wb = openpyxl.load_workbook(book)
    ws = wb[ledger.LOG_SHEET]
    assert ws.sheet_state == "veryHidden"
    before_cells = [c for c in ws["F"][1:] if c.value]
    assert all(c.data_type == "s" for c in before_cells)  # "=1+1" stored as text
    assert openpyxl.load_workbook(book, data_only=True)[ledger.LOG_SHEET]["F3"].value == '"=1+1"'


def test_reason_with_formula_text_stays_text(book, enabled):
    server.write_cells("Data", [{"cell": "A1", "value": 1}], reason="=HYPERLINK(\"http://x\")", path=book)
    ws = openpyxl.load_workbook(book)[ledger.LOG_SHEET]
    assert ws["H2"].data_type == "s"


def test_log_sheet_hidden_and_reserved(book, enabled):
    server.write_cells("Data", [{"cell": "A1", "value": 1}], path=book)
    assert [s["name"] for s in server.list_sheets(path=book)] == ["Data"]
    assert [s["name"] for s in server.get_workbook_info(path=book)["sheets"]] == ["Data"]
    assert server.search_workbook("write_cells", path=book) == []
    for call in (
        lambda: server.read_sheet(ledger.LOG_SHEET, path=book),
        lambda: server.get_cell(ledger.LOG_SHEET, "A1", path=book),
        lambda: server.write_cells(ledger.LOG_SHEET, [{"cell": "A1", "value": 1}], path=book),
        lambda: server.delete_sheet(ledger.LOG_SHEET, path=book),
        lambda: server.rename_sheet(ledger.LOG_SHEET, "x", path=book),
        lambda: server.rename_sheet("Data", "_XLSX_MCP_LOG", path=book),
        lambda: server.create_sheet(ledger.LOG_SHEET, path=book),
        lambda: server.copy_sheet("Data", ledger.LOG_SHEET, path=book),
        lambda: server.clear_range(ledger.LOG_SHEET, "A1:B2", path=book),
    ):
        with pytest.raises(ValueError, match="reserved"):
            call()


def test_delete_last_visible_sheet_refused_with_log(book, enabled):
    server.write_cells("Data", [{"cell": "A1", "value": 1}], path=book)
    with pytest.raises(ValueError, match="only sheet"):
        server.delete_sheet("Data", path=book)


def test_structural_and_range_entries_match_cell(book, enabled):
    server.write_cells("Data", [{"cell": "B5", "value": 1}], path=book)
    server.insert_rows("Data", 4, count=3, path=book)
    server.clear_range("Data", "A1:C10", path=book)
    entries = server.get_change_log(cell="B5", path=book)["entries"]
    tools = [e["tool"] for e in entries]
    assert tools[:2] == ["clear_range", "insert_rows"]
    assert "write_cells" in tools
    assert server.get_change_log(sheet="Nope", path=book)["entries"] == []


def test_trim_keeps_newest(book, enabled, monkeypatch):
    monkeypatch.setattr(settings, "CHANGE_LOG_MAX_ROWS", 3)
    for i in range(5):
        server.write_cells("Data", [{"cell": "A1", "value": i}], path=book)
    entries = server.get_change_log(path=book)["entries"]
    assert [e["after"] for e in entries] == [4, 3, 2]


def test_per_op_cap_adds_summary(book, enabled, monkeypatch):
    monkeypatch.setattr(settings, "CHANGE_LOG_MAX_CELLS_PER_OP", 2)
    server.append_rows("Data", [[1, 2, 3]], path=book)
    refs = [e["ref"] for e in server.get_change_log(path=book)["entries"]]
    assert "(+1 more cells)" in refs
    assert len(refs) == 3


def test_limit_validation(book):
    with pytest.raises(ValueError):
        server.get_change_log(limit=0, path=book)


def test_create_workbook_rejects_reserved_name(tmp_path):
    with pytest.raises(ValueError, match="reserved"):
        server.create_workbook(str(tmp_path / "x.xlsx"), sheets=["Data", ledger.LOG_SHEET])


def test_sanitize_removes_change_log(book, enabled, tmp_path):
    server.write_cells("Data", [{"cell": "A1", "value": 1}], path=book)
    found = {f["category"] for f in server.inspect_for_sharing(path=book)["findings"]}
    assert "change_log" in found
    out = str(tmp_path / "shared.xlsx")
    server.sanitize_copy(out, path=book)
    assert ledger.LOG_SHEET not in openpyxl.load_workbook(out).sheetnames
