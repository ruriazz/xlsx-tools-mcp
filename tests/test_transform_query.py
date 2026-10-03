import pytest
from xlsx_tools_mcp.io.transform import query_sheet
from xlsx_tools_mcp.io.writer import create_workbook, append_rows

def test_query_sheet(tmp_path):
    path = str(tmp_path / "test_query.xlsx")
    create_workbook(path, sheets=["Sheet1"])
    append_rows(path, "Sheet1", [
        ["Name", "Age", "Score"],
        ["Alice", 25, 90.5],
        ["Bob", 35, 80.0],
        ["Charlie", 30, 85.0],
        ["Dave", 22, 99.9]
    ])
    
    result = query_sheet(path, "Sheet1", "Age > 25 and Score < 90")
    assert result["row_count"] == 2
    assert len(result["records"]) == 2
    assert result["records"][0]["Name"] == "Bob"
    
    # Test columns filter
    res2 = query_sheet(path, "Sheet1", "Age < 25", columns=["Name"])
    assert res2["columns"] == ["Name"]
    assert res2["records"][0]["Name"] == "Dave"
    assert "Age" not in res2["records"][0]
    
    # Test max_rows
    res3 = query_sheet(path, "Sheet1", "Age > 0", max_rows=2)
    assert res3["row_count"] == 4
    assert res3["limited"] is True
    assert len(res3["records"]) == 2
