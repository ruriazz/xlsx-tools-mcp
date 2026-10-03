import openpyxl
import pytest

from xlsx_tools_mcp.errors import SheetNotFoundError
from xlsx_tools_mcp.io import reader


@pytest.fixture
def workbook_path(tmp_path):
    path = tmp_path / "book.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["item", "qty"])
    ws.append(["apple", 10])
    ws.append(["banana", 5])
    wb.save(path)
    return str(path)


def test_read_sheet_returns_absolute_grid(workbook_path):
    result = reader.read_sheet(workbook_path, "Data")
    assert result["rows"] == [["item", "qty"], ["apple", 10], ["banana", 5]]
    assert result["row_count"] == 3
    assert result["column_count"] == 2


def test_read_sheet_slices_cell_range(workbook_path):
    result = reader.read_sheet(workbook_path, "Data", cell_range="A2:A3")
    assert result["rows"] == [["apple"], ["banana"]]


def test_read_sheet_max_rows(workbook_path):
    result = reader.read_sheet(workbook_path, "Data", max_rows=2)
    assert result["rows"] == [["item", "qty"], ["apple", 10]]
    assert result["row_count"] == 2
    assert reader.read_sheet(workbook_path, "Data")["row_count"] == 3


def test_read_sheet_missing_sheet_raises(workbook_path):
    with pytest.raises(SheetNotFoundError):
        reader.read_sheet(workbook_path, "NoSuchSheet")


def test_list_sheets(workbook_path):
    sheets = reader.list_sheets(workbook_path)
    assert sheets == [{"name": "Data", "rows": 3, "columns": 2}]


def test_get_cell_plain_value(workbook_path):
    info = reader.get_cell(workbook_path, "Data", "A1")
    assert info["value"] == "item"
    assert info["formula"] is None
    for key in ("value", "formula", "data_type", "number_format", "font", "fill_color", "is_merged", "comment"):
        assert key in info


def test_get_cell_merged(tmp_path):
    path = tmp_path / "merged.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws["A1"] = "x"
    ws.merge_cells("A1:B1")
    wb.save(path)

    info = reader.get_cell(str(path), "Data", "A1")
    assert info["is_merged"] is True


def test_workbook_info(workbook_path):
    info = reader.workbook_info(workbook_path)
    assert "Data" in [s["name"] for s in info["sheets"]]
    for key in ("name", "dimensions", "max_row", "max_column", "sheet_state"):
        assert key in info["sheets"][0]
    assert info["active_sheet"] == "Data"
    assert isinstance(info["defined_names"], list)


def test_search_workbook_finds_match(workbook_path):
    matches = reader.search_workbook(workbook_path, "apple")
    assert matches == [{"sheet": "Data", "cell": "A2", "value": "apple"}]


def test_search_workbook_limit(workbook_path):
    matches = reader.search_workbook(workbook_path, "", limit=1)
    assert len(matches) == 1


def test_get_cell_formula_value(tmp_path):
    path = tmp_path / "formula.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws["A1"] = "=1+1"
    wb.save(path)

    info = reader.get_cell(str(path), "Data", "A1")
    assert info["formula"] == "=1+1"
    assert "value" in info


def test_search_workbook_missing_sheet_raises(workbook_path):
    with pytest.raises(SheetNotFoundError):
        reader.search_workbook(workbook_path, "apple", sheet="NoSuchSheet")


def test_read_sheet_format_records(workbook_path):
    result = reader.read_sheet(workbook_path, "Data", format="records")
    assert result["rows"] == [{"item": "apple", "qty": 10}, {"item": "banana", "qty": 5}]
    assert result["row_count"] == 2


def test_read_sheet_format_markdown(workbook_path):
    result = reader.read_sheet(workbook_path, "Data", format="markdown")
    lines = result["rows"].split("\n")
    assert "| item | qty |" in lines[0]
    assert "|---|---|" in lines[1]
    assert "| apple | 10.0 |" in lines[2]


def test_read_sheet_offset_row_array(workbook_path):
    result = reader.read_sheet(workbook_path, "Data", offset_row=2)
    assert result["rows"] == [["banana", 5]]


def test_read_sheet_offset_row_records(workbook_path):
    # Should use the original first row as headers, then skip 1 data row
    result = reader.read_sheet(workbook_path, "Data", format="records", offset_row=1)
    assert result["rows"] == [{"item": "banana", "qty": 5}]


def test_read_sheet_format_markdown_pipe_escaped(tmp_path):
    path = tmp_path / "pipe.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["col|a", "col|b"])
    ws.append(["val|1", "val\n2"])
    wb.save(path)

    result = reader.read_sheet(str(path), "Data", format="markdown")
    lines = result["rows"].split("\n")
    assert r"col\|a" in lines[0]
    assert r"val\|1" in lines[2]


def test_read_non_existent_file_raises_workbook_not_found(tmp_path):
    from xlsx_tools_mcp.errors import WorkbookNotFoundError
    missing = str(tmp_path / "missing.xlsx")
    with pytest.raises(WorkbookNotFoundError):
        reader.read_sheet(missing, "Sheet1")
    with pytest.raises(WorkbookNotFoundError):
        reader.list_sheets(missing)
    with pytest.raises(WorkbookNotFoundError):
        reader.workbook_info(missing)
    with pytest.raises(WorkbookNotFoundError):
        reader.get_cell(missing, "Sheet1", "A1")
    with pytest.raises(WorkbookNotFoundError):
        reader.search_workbook(missing, "query")


def test_corrupted_file_raises_invalid_workbook_error(tmp_path):
    from xlsx_tools_mcp.errors import InvalidWorkbookError
    corrupt = tmp_path / "corrupt.xlsx"
    corrupt.write_text("not a real excel file")
    with pytest.raises(InvalidWorkbookError):
        reader.read_sheet(str(corrupt), "Sheet1")
    with pytest.raises(InvalidWorkbookError):
        reader.list_sheets(str(corrupt))
    with pytest.raises(InvalidWorkbookError):
        reader.workbook_info(str(corrupt))


def test_reader_permission_error_not_masked(monkeypatch, tmp_path):
    p = tmp_path / "test.xlsx"
    wb = openpyxl.Workbook()
    wb.save(p)
    wb.close()

    def mock_load(*args, **kwargs):
        exc = PermissionError("File locked")
        exc.winerror = 32
        raise exc

    monkeypatch.setattr("openpyxl.load_workbook", mock_load)
    with pytest.raises(PermissionError):
        reader.workbook_info(str(p))
    with pytest.raises(PermissionError):
        reader.get_cell(str(p), "Sheet", "A1")




