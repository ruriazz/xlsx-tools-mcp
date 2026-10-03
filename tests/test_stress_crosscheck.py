import os
import tempfile
import pytest
from xlsx_tools_mcp import server, settings
from xlsx_tools_mcp.errors import AccessDeniedError, UnsafeFormulaError, SheetNotFoundError


@pytest.fixture
def clean_wb(tmp_path):
    path = str(tmp_path / "stress_test.xlsx")
    server.create_workbook(path, ["Main", "Data"])
    return path


# ── 1. Sheet operations & validations (Phase 1) ──────────────────────────────


def test_rename_sheet_stress(clean_wb):
    # Common case
    server.rename_sheet("Main", "Primary", path=clean_wb)
    sheets = [s["name"] for s in server.list_sheets(clean_wb)]
    assert "Primary" in sheets
    assert "Main" not in sheets

    # Idempotent rename (same name)
    res = server.rename_sheet("Primary", "Primary", path=clean_wb)
    assert res["saved"] is True

    # Non-existent source
    with pytest.raises(ValueError, match="not found"):
        server.rename_sheet("NonExistent", "NewName", path=clean_wb)

    # Collision with existing sheet
    with pytest.raises(ValueError, match="already exists"):
        server.rename_sheet("Primary", "Data", path=clean_wb)

    # Invalid names
    for bad in ["", "A" * 32, "Invalid/Name", "Bad?Name", "Test:Name", "[Sheet]", "Sheet*"]:
        with pytest.raises(ValueError):
            server.rename_sheet("Primary", bad, path=clean_wb)


def test_copy_sheet_stress(clean_wb):
    server.append_rows("Data", [["Col1", "Col2"], [10, 20]], path=clean_wb)

    # Common copy
    server.copy_sheet("Data", "DataCopy", path=clean_wb)
    sheets = [s["name"] for s in server.list_sheets(clean_wb)]
    assert "DataCopy" in sheets

    # Content copied
    data = server.read_sheet("DataCopy", path=clean_wb)
    assert len(data["rows"]) == 2

    # Non-existent source
    with pytest.raises(ValueError, match="not found"):
        server.copy_sheet("Ghost", "NewGhost", path=clean_wb)

    # Target already exists
    with pytest.raises(ValueError, match="already exists"):
        server.copy_sheet("Data", "DataCopy", path=clean_wb)

    # Invalid target name
    with pytest.raises(ValueError):
        server.copy_sheet("Data", "Bad/Name", path=clean_wb)


def test_read_sheet_formats_and_pagination(clean_wb):
    server.append_rows("Main", [
        ["Name", "Score", "Notes"],
        ["Alice", 95, "Pass | Good"],
        ["Bob", 80, "Needs\nImprovement"],
        ["Charlie", 70, None],
        ["Dave", 60, "Borderline"],
    ], path=clean_wb)

    # Format array (default)
    arr = server.read_sheet("Main", path=clean_wb)
    assert len(arr["rows"]) == 5
    assert arr["rows"][0] == ["Name", "Score", "Notes"]

    # Format records
    rec = server.read_sheet("Main", format="records", path=clean_wb)
    assert len(rec["rows"]) == 4
    assert rec["rows"][0]["Name"] == "Alice"
    assert rec["rows"][0]["Score"] == 95

    # Format records with offset_row and max_rows
    rec_paged = server.read_sheet("Main", format="records", offset_row=1, max_rows=2, path=clean_wb)
    assert len(rec_paged["rows"]) == 2
    assert rec_paged["rows"][0]["Name"] == "Bob"
    assert rec_paged["rows"][1]["Name"] == "Charlie"

    # Format markdown escaping pipes and newlines
    md = server.read_sheet("Main", format="markdown", path=clean_wb)
    assert "\\|" in md["rows"]  # pipe escaped
    assert "\nImprovement" not in md["rows"]  # newline replaced with space in cell

    # Offset beyond row count
    rec_empty = server.read_sheet("Main", format="records", offset_row=100, path=clean_wb)
    assert rec_empty["rows"] == []
    assert rec_empty["row_count"] == 0

    # Negative offset rejected
    with pytest.raises(ValueError, match="offset_row must be non-negative"):
        server.read_sheet("Main", offset_row=-1, path=clean_wb)

    # Unsupported format rejected
    with pytest.raises(ValueError, match="Unsupported format"):
        server.read_sheet("Main", format="xml", path=clean_wb)

    # Empty sheet read
    server.create_workbook(clean_wb.replace(".xlsx", "_empty.xlsx"), ["Empty"])
    p_empty = clean_wb.replace(".xlsx", "_empty.xlsx")
    assert server.read_sheet("Empty", path=p_empty)["rows"] == []
    assert server.read_sheet("Empty", format="records", path=p_empty)["rows"] == []
    assert server.read_sheet("Empty", format="markdown", path=p_empty)["rows"] == ""


def test_records_format_duplicate_and_empty_headers(clean_wb):
    server.append_rows("Data", [
        ["Col", "Col", "", None],
        [1, 2, 3, 4]
    ], path=clean_wb)

    rec = server.read_sheet("Data", format="records", path=clean_wb)
    assert len(rec["rows"]) == 1
    row = rec["rows"][0]
    # Header deduplication preserves all column values
    assert "Col" in row and "Col_1" in row
    assert row["Col"] == 1
    assert row["Col_1"] == 2


def test_autofit_columns_stress(clean_wb):
    # Empty sheet autofit
    res = server.autofit_columns("Main", path=clean_wb)
    assert res["saved"] is True

    # Data autofit
    server.append_rows("Main", [
        ["Short", "A" * 100],
        [1, 2]
    ], path=clean_wb)
    res = server.autofit_columns("Main", min_width=15, max_width=40, path=clean_wb)
    assert res["saved"] is True

    # min_width > max_width rejected
    with pytest.raises(ValueError, match="cannot be greater"):
        server.autofit_columns("Main", min_width=50, max_width=10, path=clean_wb)

    # Negative bounds rejected
    with pytest.raises(ValueError):
        server.autofit_columns("Main", min_width=-5, path=clean_wb)


def test_clear_range_stress(clean_wb):
    server.append_rows("Main", [
        ["A", "B", "C"],
        [1, 2, 3],
        [4, 5, 6]
    ], path=clean_wb)

    # Clear single cell values
    server.clear_range("Main", "B2", path=clean_wb)
    res = server.read_sheet("Main", cell_range="B2:B2", path=clean_wb)
    assert res["rows"][0][0] in (None, "")

    # Clear range with column slice "A:B" (bounded by max rows, no OOM)
    server.clear_range("Main", "A:B", clear_values=True, clear_styles=True, path=clean_wb)
    res_all = server.read_sheet("Main", path=clean_wb)
    assert res_all["rows"][0][2] == "C"  # C untouched

    # Invalid range syntax rejected
    with pytest.raises(ValueError):
        server.clear_range("Main", "INVALID", path=clean_wb)


# ── 2. Enterprise Safety & Reliability (Phase 2) ─────────────────────────────


def test_formula_injection_guard(clean_wb):
    # Safe formula
    server.write_cells("Main", [{"cell": "A1", "formula": "=SUM(1, 2)"}], path=clean_wb)

    # Unsafe external formulas blocked across variations
    blocked = [
        "=WEBSERVICE('http://evil.com')",
        "=HYPERLINK('http://evil.com')",
        "=INDIRECT('A' & '1')",
        "=+WEBSERVICE('http://evil.com')",
        " -INDIRECT('A1')",
        "@HYPERLINK('http://evil.com')",
        "+RTD('progid',,'topic')",
        "=CALL('kernel32', 'WinExec')",
        "=REGISTER('kernel32', 'WinExec')"
    ]
    for unsafe in blocked:
        with pytest.raises(ValueError, match="unsafe function"):
            server.write_cells("Main", [{"cell": "B1", "formula": unsafe}], path=clean_wb)
        # Also when placed in string value field
        with pytest.raises(ValueError, match="unsafe function"):
            server.write_cells("Main", [{"cell": "B1", "value": unsafe}], path=clean_wb)

    # Allowed when explicitly overridden
    res = server.write_cells(
        "Main",
        [{"cell": "B1", "formula": "=INDIRECT('A1')"}],
        allow_external_formulas=True,
        path=clean_wb
    )
    assert res["saved"] is True


def test_auto_backup_and_restore(clean_wb):
    server.write_cells("Main", [{"cell": "A1", "value": "Original"}], path=clean_wb)

    # Backup on write
    server.write_cells("Main", [{"cell": "A1", "value": "Mutated"}], backup=True, path=clean_wb)
    import glob
    backups = glob.glob(clean_wb.replace(".xlsx", "*.bak"))
    assert len(backups) == 1
    backup_file = backups[0]

    # Restore backup
    res = server.restore_backup(backup_file, clean_wb)
    assert res["restored"] is True
    assert res["target"] == clean_wb

    # Verify original value restored
    read_back = server.read_sheet("Main", cell_range="A1:A1", path=clean_wb)
    assert read_back["rows"][0][0] == "Original"

    # Restore with invalid backup extension
    with pytest.raises(ValueError, match="Must have .bak extension"):
        server.restore_backup(clean_wb, clean_wb)


def test_path_confinement_stress(tmp_path):
    allowed_dir = tmp_path / "allowed"
    outside_dir = tmp_path / "outside"
    allowed_dir.mkdir()
    outside_dir.mkdir()

    allowed_wb = str(allowed_dir / "safe.xlsx")
    outside_wb = str(outside_dir / "evil.xlsx")

    os.environ["XLSX_MCP_ALLOWED_DIRS"] = str(allowed_dir)
    settings.ALLOWED_DIRS = [os.path.realpath(str(allowed_dir))]

    try:
        # Inside allowed works
        server.create_workbook(allowed_wb, ["Sheet1"])
        assert os.path.exists(allowed_wb)

        # Outside allowed blocked
        with pytest.raises(ValueError, match="Access denied"):
            server.create_workbook(outside_wb, ["Sheet1"])

        with pytest.raises(ValueError, match="Access denied"):
            server.read_sheet("Sheet1", path=outside_wb)

        # Path traversal blocked
        traversal = str(allowed_dir / ".." / "outside" / "evil.xlsx")
        with pytest.raises(ValueError, match="Access denied"):
            server.read_sheet("Sheet1", path=traversal)
    finally:
        os.environ.pop("XLSX_MCP_ALLOWED_DIRS", None)
        settings.ALLOWED_DIRS = []


# ── 3. Native Excel Objects & Analytics (Phase 3) ────────────────────────────


def test_profile_sheet_stress(clean_wb):
    # Empty sheet
    prof_empty = server.profile_sheet("Main", path=clean_wb)
    assert prof_empty["row_count"] == 0
    assert prof_empty["columns"] == []

    # Mixed data types and nulls
    server.append_rows("Main", [
        ["ID", "Name", "Score", "Active", "MixedCol"],
        [1, "Alice", 92.5, True, 100],
        [2, "Bob", None, False, "TextValue"],
        [3, "Charlie", 78.0, True, 3.14],
        [4, None, None, True, None],
    ], path=clean_wb)

    prof = server.profile_sheet("Main", sample_rows=2, path=clean_wb)
    assert prof["row_count"] == 4
    cols = {c["column"]: c for c in prof["columns"]}

    assert cols["ID"]["type"] == "float"  # calamine numeric
    assert cols["ID"]["min"] == 1.0
    assert cols["ID"]["max"] == 4.0

    assert cols["Name"]["type"] == "str"
    assert cols["Name"]["null_count"] == 1
    assert len(cols["Name"]["sample"]) == 2

    assert cols["Active"]["type"] == "bool"
    assert cols["Active"]["min"] is None  # booleans excluded from numeric min/max

    assert cols["MixedCol"]["type"] == "mixed"

    # Negative sample_rows rejected
    with pytest.raises(ValueError, match="sample_rows must be non-negative"):
        server.profile_sheet("Main", sample_rows=-1, path=clean_wb)


def test_query_sheet_stress(clean_wb):
    server.append_rows("Data", [
        ["Product", "Category", "Price", "Stock"],
        ["Laptop", "Tech", 1200, 10],
        ["Mouse", "Tech", 25, 100],
        ["Desk", "Office", 300, 5],
        ["Chair", "Office", 150, 0],
    ], path=clean_wb)

    # Standard query
    q = server.query_sheet("Data", "Price > 100 and Stock > 0", path=clean_wb)
    assert q["row_count"] == 2
    prods = [r["Product"] for r in q["records"]]
    assert prods == ["Laptop", "Desk"]

    # Query with column projection
    q_proj = server.query_sheet("Data", "Category == 'Tech'", columns=["Product", "Price"], path=clean_wb)
    assert q_proj["columns"] == ["Product", "Price"]
    assert len(q_proj["records"]) == 2

    # Query with max_rows pagination
    q_lim = server.query_sheet("Data", "Price > 0", max_rows=2, path=clean_wb)
    assert len(q_lim["records"]) == 2
    assert q_lim["limited"] is True

    # Query matching 0 rows
    q_none = server.query_sheet("Data", "Price > 99999", path=clean_wb)
    assert q_none["row_count"] == 0
    assert q_none["records"] == []

    # Invalid query syntax
    with pytest.raises(ValueError, match="Invalid query"):
        server.query_sheet("Data", "Price >> && == 10", path=clean_wb)

    # Unknown column in query
    with pytest.raises(ValueError, match="Invalid query"):
        server.query_sheet("Data", "UnknownCol > 10", path=clean_wb)

    # Unknown column in projection
    with pytest.raises(ValueError, match="Unknown column"):
        server.query_sheet("Data", "Price > 0", columns=["NotThere"], path=clean_wb)

    # Negative max_rows rejected
    with pytest.raises(ValueError, match="max_rows must be non-negative"):
        server.query_sheet("Data", "Price > 0", max_rows=-1, path=clean_wb)


def test_create_table_stress(clean_wb):
    server.append_rows("Data", [
        ["Item", "Qty"],
        ["Apples", 10],
        ["Bananas", 20],
    ], path=clean_wb)

    # Valid table
    res = server.create_table("Data", "A1:B3", "FruitTable", path=clean_wb)
    assert res["saved"] is True

    # Duplicate table name in same sheet
    with pytest.raises(ValueError, match="already exists"):
        server.create_table("Data", "A1:B3", "FruitTable", path=clean_wb)

    # Duplicate table name across sheets in workbook
    server.append_rows("Main", [["X", "Y"], [1, 2]], path=clean_wb)
    with pytest.raises(ValueError, match="already exists in sheet 'Data'"):
        server.create_table("Main", "A1:B2", "FruitTable", path=clean_wb)

    # Invalid table names
    for bad_name in ["", "Table With Space", "123NumberStart", "A1", "R1C1", "T@ble", "A" * 256]:
        with pytest.raises(ValueError, match="Invalid table name"):
            server.create_table("Data", "A1:B3", bad_name, path=clean_wb)

    # Range must have at least 2 rows (header + 1 data row)
    with pytest.raises(ValueError, match="must span at least 2 rows"):
        server.create_table("Data", "A1:B1", "HeaderOnlyTable", path=clean_wb)

    # Invalid range coordinate
    with pytest.raises(ValueError, match="not a valid coordinate or range|Invalid cell_range"):
        server.create_table("Data", "INVALID", "ValidName", path=clean_wb)


def test_create_chart_stress(clean_wb):
    server.append_rows("Data", [
        ["Month", "Revenue"],
        ["Jan", 100],
        ["Feb", 150],
        ["Mar", 200]
    ], path=clean_wb)

    # Supported chart types
    for chart_type in ["bar", "line", "pie", "scatter"]:
        res = server.create_chart(
            "Data",
            chart_type,
            data_range="B1:B4",
            categories_range="A2:A4",
            title=f"{chart_type.title()} Chart",
            target_cell="E2",
            path=clean_wb
        )
        assert res["saved"] is True

    # Unsupported chart type
    with pytest.raises(ValueError, match="Unsupported chart type"):
        server.create_chart("Data", "donut", data_range="B1:B4", path=clean_wb)

    # Invalid data_range
    with pytest.raises(ValueError, match="not a valid coordinate or range|Invalid data_range"):
        server.create_chart("Data", "bar", data_range="INVALID", path=clean_wb)

    # Invalid categories_range
    with pytest.raises(ValueError, match="not a valid coordinate or range|Invalid categories_range"):
        server.create_chart("Data", "bar", data_range="B1:B4", categories_range="INVALID", path=clean_wb)

    # Invalid target_cell coordinate
    for bad_cell in ["INVALID", "123", "A", ""]:
        with pytest.raises(ValueError, match="Invalid target_cell"):
            server.create_chart("Data", "bar", data_range="B1:B4", target_cell=bad_cell, path=clean_wb)


# ── 4. Cell Inspection & Editing Tools ────────────────────────────────────────


def test_get_cell_stress(clean_wb):
    server.write_cells("Main", [
        {"cell": "A1", "value": "Hello"},
        {"cell": "A2", "formula": "=10+20"}
    ], path=clean_wb)

    # Normal get
    c1 = server.get_cell("Main", "A1", path=clean_wb)
    assert c1["value"] == "Hello"
    assert c1["cell"] == "A1"

    # Invalid cell coordinate
    for bad_coord in ["INVALID", "A", "123", "", "A1:B2"]:
        with pytest.raises(ValueError, match="Invalid cell coordinate"):
            server.get_cell("Main", bad_coord, path=clean_wb)


def test_row_and_column_bounds(clean_wb):
    # insert_rows index bounds
    with pytest.raises(ValueError, match="start_row must be >= 1"):
        server.insert_rows("Main", start_row=0, path=clean_wb)
    with pytest.raises(ValueError, match="count must be >= 1"):
        server.insert_rows("Main", start_row=1, count=0, path=clean_wb)

    # delete_rows bounds
    with pytest.raises(ValueError, match="start_row must be >= 1"):
        server.delete_rows("Main", start_row=-1, path=clean_wb)

    # insert_columns bounds
    with pytest.raises(ValueError, match="start_column must be >= 1"):
        server.insert_columns("Main", start_column=0, path=clean_wb)

    # delete_columns bounds
    with pytest.raises(ValueError, match="start_column must be >= 1"):
        server.delete_columns("Main", start_column=0, path=clean_wb)


# ── 5. Audit Hardening Verification ──────────────────────────────────────────


def test_query_sheet_rce_prevention(clean_wb):
    server.append_rows("Data", [
        ["Col1", "Col2"],
        [10, 20]
    ], path=clean_wb)

    # RCE payloads must be strictly blocked
    malicious_queries = [
        "@__builtins__.__import__('os').system('calc')",
        "__import__('os').system('calc')",
        "eval('1+1')",
        "os.system('calc')",
        "__class__.__base__",
        "@variable_ref",
    ]
    for bad_query in malicious_queries:
        with pytest.raises(ValueError, match="Invalid query"):
            server.query_sheet("Data", bad_query, path=clean_wb)


def test_dde_formula_injection_prevention(clean_wb):
    dde_payloads = [
        "=cmd|'/C calc'!A0",
        "+cmd|'/C calc'!A0",
        "-powershell|'/C calc'!A0",
        "@cmd|'/C calc'!A0",
    ]
    for dde in dde_payloads:
        with pytest.raises(ValueError, match="unsafe DDE command execution syntax"):
            server.write_cells("Main", [{"cell": "A1", "formula": dde}], path=clean_wb)


def test_table_header_validation(clean_wb):
    # Empty header cell in table range
    server.append_rows("Data", [
        ["Col1", None, "Col3"],
        [1, 2, 3],
    ], path=clean_wb)
    with pytest.raises(ValueError, match="Table header cell.*is empty"):
        server.create_table("Data", "A1:C2", "EmptyHeaderTable", path=clean_wb)

    # Duplicate header in table range
    server.append_rows("Main", [
        ["ColA", "ColA"],
        [1, 2],
    ], path=clean_wb)
    with pytest.raises(ValueError, match="Duplicate table column header"):
        server.create_table("Main", "A1:B2", "DuplicateHeaderTable", path=clean_wb)


def test_backup_dir_in_allowed_dirs(tmp_path):
    allowed_dir = tmp_path / "app_data"
    backup_dir = tmp_path / "app_backups"
    allowed_dir.mkdir()
    backup_dir.mkdir()

    wb_path = str(allowed_dir / "doc.xlsx")
    os.environ["XLSX_MCP_ALLOWED_DIRS"] = str(allowed_dir)
    os.environ["XLSX_MCP_BACKUP_DIR"] = str(backup_dir)

    # Re-evaluate settings
    settings.ALLOWED_DIRS = settings._parse_allowed_dirs(str(allowed_dir))
    settings.BACKUP_DIR = str(backup_dir.resolve())
    if settings.BACKUP_DIR and settings.ALLOWED_DIRS and settings.BACKUP_DIR not in settings.ALLOWED_DIRS:
        settings.ALLOWED_DIRS.append(settings.BACKUP_DIR)

    try:
        server.create_workbook(wb_path, ["S1"])
        server.write_cells("S1", [{"cell": "A1", "value": "Initial"}], backup=False, path=wb_path)

        # Mutate with backup=True -> backup captures 'Initial'
        server.write_cells("S1", [{"cell": "A1", "value": "Changed"}], backup=True, path=wb_path)

        # Check backup created in custom BACKUP_DIR
        import glob
        backups = glob.glob(str(backup_dir / "*.bak"))
        assert len(backups) == 1
        backup_file = backups[0]

        # Verify currently changed
        curr_res = server.read_sheet("S1", path=wb_path)
        assert curr_res["rows"][0][0] == "Changed"

        # Restore from backup_dir must succeed and NOT be blocked by ALLOWED_DIRS
        res = server.restore_backup(backup_file, wb_path)
        assert res["restored"] is True

        read_res = server.read_sheet("S1", path=wb_path)
        assert read_res["rows"][0][0] == "Initial"
    finally:
        os.environ.pop("XLSX_MCP_ALLOWED_DIRS", None)
        os.environ.pop("XLSX_MCP_BACKUP_DIR", None)
        settings.ALLOWED_DIRS = []
        settings.BACKUP_DIR = ""

