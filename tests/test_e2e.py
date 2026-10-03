import os
import pytest
from mcp.server import MCPServer
from xlsx_tools_mcp import server

@pytest.fixture
def test_file(tmp_path):
    path = tmp_path / "e2e.xlsx"
    server.create_workbook(str(path), ["Sheet1"])
    return str(path)


def test_e2e_sheet_workflow(test_file):
    # 1. Rename sheet
    server.rename_sheet("Sheet1", "Primary", path=test_file)
    
    # 2. Write data
    server.write_cells("Primary", [
        {"cell": "A1", "value": "id"},
        {"cell": "B1", "value": "name"},
        {"cell": "A2", "value": 1},
        {"cell": "B2", "value": "Alice"},
        {"cell": "A3", "value": 2},
        {"cell": "B3", "value": "Bob"}
    ], path=test_file)
    
    # 3. Copy sheet
    server.copy_sheet("Primary", "Backup", path=test_file)
    
    # 4. Read sheet with formats and offset
    read_array = server.read_sheet("Backup", path=test_file)
    assert read_array["rows"] == [["id", "name"], [1, "Alice"], [2, "Bob"]]
    
    read_records = server.read_sheet("Backup", format="records", path=test_file)
    assert read_records["rows"] == [{"id": 1, "name": "Alice"}, {"id": 2, "name": "Bob"}]
    
    read_records_offset = server.read_sheet("Backup", format="records", offset_row=1, path=test_file)
    assert read_records_offset["rows"] == [{"id": 2, "name": "Bob"}]
    
    read_md = server.read_sheet("Backup", format="markdown", path=test_file)
    assert "| id | name |" in read_md["rows"]
    
    # 5. Autofit columns
    server.autofit_columns("Primary", path=test_file)
    
    # 6. Clear range
    server.clear_range("Backup", "A3:B3", path=test_file)
    read_array_after_clear = server.read_sheet("Backup", path=test_file)
    # The cleared row at the end of the sheet may be omitted from the bounding box entirely
    assert len(read_array_after_clear["rows"]) == 2
    assert read_array_after_clear["rows"][0] == ["id", "name"]
    assert read_array_after_clear["rows"][1] == [1.0, "Alice"]


def test_e2e_backup_and_restore(test_file):
    # Enable backup
    server.write_cells("Sheet1", [{"cell": "A1", "value": "first"}], backup=True, path=test_file)
    
    # Check if a backup was created
    import glob
    import os
    backups = glob.glob(test_file.replace(".xlsx", "*.bak"))
    assert len(backups) == 1
    
    # Restore it
    server.restore_backup(backups[0], test_file)
    
    # Ensure it's empty again
    read_array = server.read_sheet("Sheet1", path=test_file)
    assert read_array["rows"] == []

