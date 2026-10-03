import openpyxl
import pytest

from xlsx_tools_mcp.errors import SheetNotFoundError
from xlsx_tools_mcp.io import reader, writer


@pytest.fixture
def workbook_path(tmp_path):
    path = tmp_path / "book.xlsx"
    return str(writer.create_workbook(str(path), sheets=["Data"])["path"])


def test_create_workbook_refuses_overwrite_by_default(tmp_path):
    path = str(tmp_path / "book.xlsx")
    writer.create_workbook(path)
    with pytest.raises(FileExistsError):
        writer.create_workbook(path)


def test_create_workbook_refuses_non_excel_path(tmp_path):
    with pytest.raises(ValueError):
        writer.create_workbook(str(tmp_path / "data.txt"))


def test_create_workbook_response_shape(tmp_path):
    result = writer.create_workbook(str(tmp_path / "new.xlsx"), sheets=["A"])
    for key in ("saved", "recalculated", "errors_found", "message", "path", "sheets"):
        assert key in result
    assert result["recalculated"] is False
    assert result["errors_found"] == []


def test_write_cells_then_read_sheet_roundtrip(workbook_path):
    writer.write_cells(workbook_path, "Data", [{"cell": "A1", "value": "hello"}])
    result = reader.read_sheet(workbook_path, "Data")
    assert result["rows"][0][0] == "hello"


def test_write_cells_missing_cell_key_raises(workbook_path):
    with pytest.raises(ValueError):
        writer.write_cells(workbook_path, "Data", [{"value": "oops"}])


def test_write_cells_missing_sheet_raises(workbook_path):
    with pytest.raises(SheetNotFoundError):
        writer.write_cells(workbook_path, "NoSuchSheet", [{"cell": "A1", "value": 1}])


def test_write_cells_survives_process_crash_mid_save(workbook_path, monkeypatch):
    """A crash between the temp-file save and the atomic rename must never
    leave the original file missing or truncated (see `_atomic_save`)."""
    original_save = openpyxl.Workbook.save

    def boom(self, path):
        original_save(self, path)
        raise RuntimeError("simulated crash after temp file write")

    monkeypatch.setattr(openpyxl.Workbook, "save", boom)

    with pytest.raises(RuntimeError):
        writer.write_cells(workbook_path, "Data", [{"cell": "A1", "value": "should not apply"}])

    # Original file must still be intact and readable — crash happened only on the temp copy.
    result = reader.read_sheet(workbook_path, "Data")
    assert result["rows"] == []


def test_delete_sheet_refuses_last_sheet(workbook_path):
    with pytest.raises(ValueError):
        writer.delete_sheet(workbook_path, "Data")


def test_merge_cells(workbook_path):
    result = writer.merge_cells(workbook_path, "Data", "A1:B1")
    assert result["saved"] is True


def test_write_cells_empty_is_noop(workbook_path):
    result = writer.write_cells(workbook_path, "Data", [])
    assert result == {
        "saved": True,
        "recalculated": False,
        "errors_found": [],
        "message": "Nothing to write; file unchanged.",
    }
    assert reader.read_sheet(workbook_path, "Data")["rows"] == []


def test_append_rows_empty_is_noop(workbook_path):
    result = writer.append_rows(workbook_path, "Data", [])
    assert result["saved"] is True
    assert result["recalculated"] is False
    assert reader.read_sheet(workbook_path, "Data")["rows"] == []


def test_write_cells_recalculate_false(workbook_path):
    result = writer.write_cells(workbook_path, "Data", [{"cell": "A1", "value": 42}], recalculate=False)
    assert result["saved"] is True
    assert result["recalculated"] is False


def test_insert_then_delete_rows(workbook_path):
    writer.write_cells(workbook_path, "Data", [{"cell": "A1", "value": "x"}], recalculate=False)
    result = writer.insert_rows(workbook_path, "Data", 2, count=1, recalculate=False)
    assert result["saved"] is True
    result = writer.delete_rows(workbook_path, "Data", 2, count=1, recalculate=False)
    assert result["saved"] is True
    assert reader.read_sheet(workbook_path, "Data")["rows"][0][0] == "x"


def test_insert_then_delete_columns(workbook_path):
    writer.write_cells(workbook_path, "Data", [{"cell": "A1", "value": "x"}], recalculate=False)
    writer.insert_columns(workbook_path, "Data", 2, count=1, recalculate=False)
    result = writer.delete_columns(workbook_path, "Data", 2, count=1, recalculate=False)
    assert result["saved"] is True
    assert reader.read_sheet(workbook_path, "Data")["rows"][0][0] == "x"


def test_merge_then_unmerge(workbook_path):
    writer.merge_cells(workbook_path, "Data", "A1:B1")
    wb = openpyxl.load_workbook(workbook_path)
    assert "A1:B1" in [str(r) for r in wb["Data"].merged_cells.ranges]
    wb.close()

    writer.unmerge_cells(workbook_path, "Data", "A1:B1")
    wb = openpyxl.load_workbook(workbook_path)
    assert len(wb["Data"].merged_cells.ranges) == 0
    wb.close()


def test_set_cell_style_applies_bold_and_fill(workbook_path):
    writer.write_cells(workbook_path, "Data", [{"cell": "A1", "value": "hi"}], recalculate=False)
    writer.set_cell_style(workbook_path, "Data", "A1", {"bold": True, "bg_color": "FF0000"})

    wb = openpyxl.load_workbook(workbook_path)
    cell = wb["Data"]["A1"]
    assert cell.font.bold is True
    assert cell.fill.start_color.rgb.endswith("FF0000")
    wb.close()


def test_rename_sheet(workbook_path):
    writer.rename_sheet(workbook_path, "Data", "NewData")
    wb = openpyxl.load_workbook(workbook_path)
    assert "NewData" in wb.sheetnames
    assert "Data" not in wb.sheetnames
    wb.close()


def test_rename_sheet_invalid(workbook_path):
    with pytest.raises(ValueError):
        writer.rename_sheet(workbook_path, "Data", "Invalid[]Name")


def test_copy_sheet(workbook_path):
    writer.write_cells(workbook_path, "Data", [{"cell": "A1", "value": "test"}])
    writer.copy_sheet(workbook_path, "Data", "DataCopy")
    wb = openpyxl.load_workbook(workbook_path)
    assert "DataCopy" in wb.sheetnames
    assert wb["DataCopy"]["A1"].value == "test"
    wb.close()


def test_autofit_columns(workbook_path):
    writer.write_cells(workbook_path, "Data", [{"cell": "A1", "value": "Very long text to autofit"}], recalculate=False)
    writer.autofit_columns(workbook_path, "Data")
    wb = openpyxl.load_workbook(workbook_path)
    assert wb["Data"].column_dimensions["A"].width > 20
    wb.close()


def test_clear_range(workbook_path):
    writer.write_cells(workbook_path, "Data", [{"cell": "A1", "value": "test"}, {"cell": "B1", "value": "keep"}])
    writer.set_cell_style(workbook_path, "Data", "A1", {"bold": True})
    
    writer.clear_range(workbook_path, "Data", "A1:A1", clear_values=True, clear_styles=True)
    wb = openpyxl.load_workbook(workbook_path)
    assert wb["Data"]["A1"].value is None
    assert wb["Data"]["A1"].font.bold is False
    assert wb["Data"]["B1"].value == "keep"
    wb.close()


def test_clear_range_full_column_unbounded(workbook_path):
    # "A:A" covers openpyxl 1 million rows, must safely bound to max_row without OOM/hang
    writer.write_cells(workbook_path, "Data", [{"cell": "A1", "value": "top"}, {"cell": "A2", "value": "bottom"}], recalculate=False)
    writer.clear_range(workbook_path, "Data", "A:A", clear_values=True)
    wb = openpyxl.load_workbook(workbook_path)
    assert wb["Data"]["A1"].value is None
    assert wb["Data"]["A2"].value is None
    wb.close()



from xlsx_tools_mcp.errors import UnsafeFormulaError

def test_formula_injection_guard(workbook_path):
    with pytest.raises(UnsafeFormulaError):
        writer.write_cells(workbook_path, "Data", [{"cell": "A1", "formula": "=WEBSERVICE(\"http://evil.com\")"}])
    
    with pytest.raises(UnsafeFormulaError):
        writer.write_cells(workbook_path, "Data", [{"cell": "A1", "formula": "+WEBSERVICE(\"http://evil.com\")"}])

    with pytest.raises(UnsafeFormulaError):
        writer.write_cells(workbook_path, "Data", [{"cell": "A1", "value": "=indirect(\"A1\")"}])

    with pytest.raises(UnsafeFormulaError):
        writer.write_cells(workbook_path, "Data", [{"cell": "A1", "formula": "=hyperlink(\"http://evil.com\")"}])

    # Should not raise when allow_external_formulas=True
    writer.write_cells(workbook_path, "Data", [{"cell": "A1", "formula": "=WEBSERVICE(\"http://evil.com\")"}], allow_external_formulas=True, recalculate=False)
    
def test_auto_backup_creation_and_restore(workbook_path, monkeypatch, tmp_path):
    import os
    import glob
    from pathlib import Path
    
    # write initial state
    writer.write_cells(workbook_path, "Data", [{"cell": "A1", "value": "original"}], recalculate=False)
    
    # Set backup dir to a tmp dir to avoid clutter
    backup_dir = tmp_path / "backups"
    monkeypatch.setattr("xlsx_tools_mcp.io.writer.BACKUP_DIR", str(backup_dir))
    
    # Write something with backup
    writer.write_cells(workbook_path, "Data", [{"cell": "A1", "value": "mutated"}], recalculate=False, backup=True)
    
    backups = list(backup_dir.glob("*.bak"))
    assert len(backups) == 1
    
    # Check that mutated is in the file
    from xlsx_tools_mcp.io import reader
    assert reader.read_sheet(workbook_path, "Data")["rows"][0][0] == "mutated"
    
    # Test extension validations in restore_backup
    with pytest.raises(ValueError, match="non-backup file"):
        writer.restore_backup(workbook_path, workbook_path)
    with pytest.raises(ValueError, match="non-Excel file"):
        writer.restore_backup(str(backups[0]), str(tmp_path / "notes.txt"))

    # restore
    writer.restore_backup(str(backups[0]), workbook_path)
    
    # Check that original is restored
    assert reader.read_sheet(workbook_path, "Data")["rows"][0][0] == "original"


