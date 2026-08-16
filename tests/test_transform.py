import openpyxl
import pytest

from xlsx_mcp.io import transform


@pytest.fixture
def workbook_path(tmp_path):
    path = tmp_path / "book.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["item", "qty"])
    ws.append(["apple", 10])
    ws.append(["banana", 5])
    ws.append(["apple", 3])
    wb.save(path)
    return str(path)


def test_aggregate_with_header(workbook_path):
    result = transform.aggregate_sheet(workbook_path, "Data", group_by=["item"], agg={"qty": "sum"})
    records = {r["item"]: r["qty"] for r in result["records"]}
    assert result["row_count"] == 2
    assert records == {"apple": 13, "banana": 5}
    assert result["columns"] == ["item", "qty"]


def test_aggregate_without_header(workbook_path):
    result = transform.aggregate_sheet(
        workbook_path, "Data", group_by=["col_0"], agg={"col_1": "sum"}, has_header=False
    )
    assert result["columns"] == ["col_0", "col_1"]
    assert result["row_count"] == 3


def test_aggregate_with_cell_range(workbook_path):
    """Range A2:B3 excludes the header and the third data row ("apple", 3)."""
    result = transform.aggregate_sheet(
        workbook_path,
        "Data",
        group_by=["col_0"],
        agg={"col_1": "sum"},
        cell_range="A2:B3",
        has_header=False,
    )
    records = {r["col_0"]: r["col_1"] for r in result["records"]}
    assert result["row_count"] == 2
    assert records == {"apple": 10, "banana": 5}