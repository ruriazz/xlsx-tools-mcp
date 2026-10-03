import pytest
from xlsx_tools_mcp.io.reader import profile_sheet
from xlsx_tools_mcp.io.writer import create_workbook, append_rows

def test_profile_sheet(tmp_path):
    path = str(tmp_path / "test_profile.xlsx")
    create_workbook(path, sheets=["Sheet1"])
    append_rows(path, "Sheet1", [
        ["Name", "Age", "Score"],
        ["Alice", 25, 90.5],
        ["Bob", None, 80.0],
        ["Charlie", 30, None],
        ["Dave", 22, 99.9]
    ])
    
    result = profile_sheet(path, "Sheet1", sample_rows=2)
    assert result["row_count"] == 4
    
    cols = {c["column"]: c for c in result["columns"]}
    assert cols["Name"]["type"] == "str"
    assert cols["Name"]["null_count"] == 0
    assert cols["Age"]["type"] == "float"
    assert cols["Age"]["null_count"] == 1
    assert cols["Age"]["min"] == 22
    assert cols["Age"]["max"] == 30
    assert cols["Score"]["type"] == "float"
    assert cols["Score"]["null_count"] == 1
    assert cols["Score"]["min"] == 80.0
    assert cols["Score"]["max"] == 99.9
    assert len(cols["Name"]["sample"]) == 2
