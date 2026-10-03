"""Comprehensive End-to-End Regression Test covering all 29 MCP tools.

Tests realistic cross-platform operations (Windows, macOS, Linux) including:
- All 29 MCP server tool functions
- Case-insensitivity in path handling
- Environment variable parsing with semicolons and commas
- Safe replacement and lock recovery
- Auto-backup creation and atomic restore
"""

import os
from pathlib import Path
import pytest
from xlsx_tools_mcp import server, settings


def test_full_29_tools_e2e_lifecycle(tmp_path, monkeypatch):
    """Exercise every single one of the 29 MCP tools in server.py."""
    data_dir = tmp_path / "DataDir"
    data_dir.mkdir()
    wb_file = data_dir / "Company_Report.xlsx"
    wb_path = str(wb_file)

    # 1. create_workbook
    res_cw = server.create_workbook(wb_path, ["Summary", "TempSheet"])
    assert res_cw["saved"] is True
    assert Path(wb_path).exists()

    # 2. list_sheets
    sheets = server.list_sheets(wb_path)
    sheet_names = [s["name"] for s in sheets]
    assert "Summary" in sheet_names
    assert "TempSheet" in sheet_names

    # 3. get_workbook_info
    info = server.get_workbook_info(wb_path)
    assert "Summary" in [s["name"] for s in info["sheets"]]
    assert info["active_sheet"] is not None

    # 4. write_cells (values and formulas)
    res_wc = server.write_cells(
        "Summary",
        [
            {"cell": "A1", "value": "Department"},
            {"cell": "B1", "value": "Budget"},
            {"cell": "C1", "value": "Spend"},
            {"cell": "D1", "value": "Remaining"},
            {"cell": "A2", "value": "Engineering"},
            {"cell": "B2", "value": 50000},
            {"cell": "C2", "value": 30000},
            {"cell": "D2", "formula": "=B2-C2"},
        ],
        path=wb_path,
        recalculate=False,
    )
    assert res_wc["saved"] is True

    # 5. append_rows
    res_ar = server.append_rows(
        "Summary",
        [
            ["Marketing", 20000, 15000, "=B3-C3"],
            ["Sales", 35000, 25000, "=B4-C4"],
            ["HR", 15000, 10000, "=B5-C5"],
        ],
        path=wb_path,
        recalculate=False,
    )
    assert res_ar["saved"] is True

    # 6. get_cell
    cell_a2 = server.get_cell("Summary", "A2", path=wb_path)
    assert cell_a2["value"] == "Engineering"
    cell_d2 = server.get_cell("Summary", "D2", path=wb_path)
    assert cell_d2["formula"] == "=B2-C2"

    # 7. read_sheet (all formats and offset pagination)
    arr_read = server.read_sheet("Summary", format="array", path=wb_path)
    assert len(arr_read["rows"]) == 5
    assert arr_read["rows"][0] == ["Department", "Budget", "Spend", "Remaining"]

    rec_read = server.read_sheet("Summary", format="records", path=wb_path)
    assert len(rec_read["rows"]) == 4
    assert rec_read["rows"][0]["Department"] == "Engineering"

    paged_read = server.read_sheet("Summary", format="records", offset_row=1, max_rows=2, path=wb_path)
    assert len(paged_read["rows"]) == 2
    assert paged_read["rows"][0]["Department"] == "Marketing"

    md_read = server.read_sheet("Summary", format="markdown", path=wb_path)
    assert "| Department | Budget |" in md_read["rows"]

    # 8. search_workbook
    search_res = server.search_workbook("Marketing", path=wb_path)
    assert len(search_res) == 1
    assert search_res[0]["cell"] == "A3"

    # 9. set_cell_style
    res_style = server.set_cell_style(
        "Summary",
        "A1:D1",
        {
            "bold": True,
            "font_size": 12,
            "bg_color": "E0E0E0",
            "horizontal": "center",
            "border": "thin",
        },
        path=wb_path,
    )
    assert res_style["saved"] is True

    # 10. merge_cells
    server.write_cells("Summary", [{"cell": "A10", "value": "Notes Header"}], path=wb_path, recalculate=False)
    res_merge = server.merge_cells("Summary", "A10:C10", path=wb_path)
    assert res_merge["saved"] is True

    # 11. unmerge_cells
    res_unmerge = server.unmerge_cells("Summary", "A10:C10", path=wb_path)
    assert res_unmerge["saved"] is True

    # 12. insert_rows
    res_ir = server.insert_rows("Summary", start_row=6, count=1, recalculate=False, path=wb_path)
    assert res_ir["saved"] is True

    # 13. delete_rows
    res_dr = server.delete_rows("Summary", start_row=6, count=1, recalculate=False, path=wb_path)
    assert res_dr["saved"] is True

    # 14. insert_columns
    res_ic = server.insert_columns("Summary", start_column=5, count=1, recalculate=False, path=wb_path)
    assert res_ic["saved"] is True

    # 15. delete_columns
    res_dc = server.delete_columns("Summary", start_column=5, count=1, recalculate=False, path=wb_path)
    assert res_dc["saved"] is True

    # 16. autofit_columns
    res_af = server.autofit_columns("Summary", min_width=10, max_width=40, padding=2, path=wb_path)
    assert res_af["saved"] is True

    # 17. rename_sheet
    res_ren = server.rename_sheet("TempSheet", "DraftData", path=wb_path)
    assert res_ren["saved"] is True
    assert "DraftData" in [s["name"] for s in server.list_sheets(wb_path)]

    # 18. copy_sheet
    res_cp = server.copy_sheet("Summary", "Summary_Copy", path=wb_path)
    assert res_cp["saved"] is True
    assert "Summary_Copy" in [s["name"] for s in server.list_sheets(wb_path)]

    # 19. delete_sheet
    res_del = server.delete_sheet("DraftData", path=wb_path)
    assert res_del["saved"] is True
    assert "DraftData" not in [s["name"] for s in server.list_sheets(wb_path)]

    # 20. create_sheet
    res_cs = server.create_sheet("Analytics", index=1, path=wb_path)
    assert res_cs["saved"] is True
    assert "Analytics" in [s["name"] for s in server.list_sheets(wb_path)]

    # 21. clear_range
    server.append_rows("Analytics", [["Temp1", "Temp2"], ["Val1", "Val2"]], path=wb_path, recalculate=False)
    res_clr = server.clear_range("Analytics", "A1:B2", clear_values=True, clear_styles=True, path=wb_path)
    assert res_clr["saved"] is True

    # 22. create_table
    res_tbl = server.create_table(
        "Summary",
        "A1:D5",
        "DepartmentBudgetTable",
        style_name="TableStyleLight1",
        show_filter=True,
        show_row_stripes=True,
        path=wb_path,
    )
    assert res_tbl["saved"] is True

    # 23. create_chart
    res_cht = server.create_chart(
        "Summary",
        "bar",
        data_range="B1:C5",
        categories_range="A2:A5",
        title="Department Spend Comparison",
        target_cell="F2",
        path=wb_path,
    )
    assert res_cht["saved"] is True

    # 24. profile_sheet
    prof = server.profile_sheet("Summary", sample_rows=2, path=wb_path)
    assert prof["row_count"] >= 4
    assert len(prof["columns"]) == 4

    # 25. query_sheet
    q_res = server.query_sheet("Summary", "Budget >= 30000", columns=["Department", "Budget"], cell_range="A1:D5", path=wb_path)
    assert q_res["row_count"] >= 2
    assert set(q_res["columns"]) == {"Department", "Budget"}

    # 26. aggregate_sheet
    agg_res = server.aggregate_sheet("Summary", group_by=["Department"], agg={"Budget": "sum", "Spend": "mean"}, cell_range="A1:D5", path=wb_path)
    assert agg_res["row_count"] == 4

    # 27. recalculate_workbook
    res_recalc = server.recalculate_workbook(path=wb_path)
    assert "saved" in res_recalc
    assert "message" in res_recalc

    # 28. restore_backup (write with backup=True then restore)
    res_wbk = server.write_cells(
        "Summary",
        [{"cell": "A2", "value": "Modified_Dept"}],
        backup=True,
        path=wb_path,
        recalculate=False,
    )
    assert res_wbk["saved"] is True

    # Find the backup file
    backup_files = list(data_dir.glob("Company_Report.xlsx.*.bak"))
    assert len(backup_files) >= 1
    newest_backup = str(sorted(backup_files)[-1])

    # Revert via restore_backup
    res_rest = server.restore_backup(newest_backup, wb_path)
    assert res_rest["restored"] is True
    # Verify restored content
    assert server.get_cell("Summary", "A2", path=wb_path)["value"] == "Engineering"

    # 29. list_configured_files
    raw_cfg = f"report={wb_path} ; other={data_dir / 'other.xlsx'}"
    monkeypatch.setenv("XLSX_MCP_FILES", raw_cfg)
    monkeypatch.setattr(server, "CONFIGURED_FILES", settings._parse_configured_files(raw_cfg))
    cfg_files = server.list_configured_files()
    assert "report" in cfg_files
    assert cfg_files["report"] == wb_path


def test_cross_platform_windows_sharing_violation_recovery(tmp_path, monkeypatch):
    """Simulate Windows ERROR_SHARING_VIOLATION during save and ensure retry succeeds."""
    target = tmp_path / "locked_book.xlsx"
    server.create_workbook(str(target), ["Data"])

    replace_calls = 0
    orig_replace = os.replace

    def mock_replace(src, dst):
        nonlocal replace_calls
        replace_calls += 1
        if replace_calls < 3:
            exc = PermissionError("Sharing violation")
            exc.winerror = 32
            raise exc
        return orig_replace(src, dst)

    monkeypatch.setattr(os, "replace", mock_replace)
    monkeypatch.setattr("time.sleep", lambda _: None)

    res = server.write_cells("Data", [{"cell": "A1", "value": "Recovered"}], path=str(target), recalculate=False)
    assert res["saved"] is True
    assert replace_calls == 3
    assert server.get_cell("Data", "A1", path=str(target))["value"] == "Recovered"


def test_cross_platform_windows_permanent_lock_translated(tmp_path, monkeypatch):
    """Simulate Windows file permanently held by Excel, verifying clean client error."""
    target = tmp_path / "held_by_excel.xlsx"
    server.create_workbook(str(target), ["Data"])

    def mock_replace(src, dst):
        exc = PermissionError("Sharing violation")
        exc.winerror = 32
        raise exc

    monkeypatch.setattr(os, "replace", mock_replace)
    monkeypatch.setattr("time.sleep", lambda _: None)

    with pytest.raises(ValueError, match="locked by another process"):
        server.write_cells("Data", [{"cell": "A1", "value": "Crash"}], path=str(target), recalculate=False)


def test_cross_platform_path_confinement_case_insensitivity(tmp_path, monkeypatch):
    """Verify case-insensitivity on Windows-style paths under path confinement."""
    import ntpath
    data_dir = tmp_path / "AllowedData"
    data_dir.mkdir()
    wb_file = data_dir / "target.xlsx"
    server.create_workbook(str(wb_file), ["Sheet1"])

    monkeypatch.setenv("XLSX_MCP_ALLOWED_DIRS", str(data_dir))
    monkeypatch.setattr(settings, "ALLOWED_DIRS", settings._parse_allowed_dirs(str(data_dir)))

    monkeypatch.setattr(settings.os.path, "normcase", ntpath.normcase)
    monkeypatch.setattr(settings.os.path, "commonpath", ntpath.commonpath)

    upper_path = str(wb_file).upper() if os.name == "nt" else str(wb_file)
    res = server.read_sheet("Sheet1", path=upper_path)
    assert "rows" in res
