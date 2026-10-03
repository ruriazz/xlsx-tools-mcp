from __future__ import annotations

import os
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import range_boundaries
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from ..recalc import recalculate

import re
import shutil
import datetime
from ..errors import SheetNotFoundError, UnsafeFormulaError
from .. import settings
from ..settings import AUTO_BACKUP, BACKUP_DIR




UNSAFE_FORMULA_PATTERN = re.compile(r"\b(WEBSERVICE|HYPERLINK|INDIRECT|RTD|CALL|REGISTER)\s*\(", re.IGNORECASE)

def _validate_formula(formula: str, allow_external_formulas: bool = False) -> None:
    if not allow_external_formulas and isinstance(formula, str):
        text = formula.strip()
        if text.startswith(("=", "+", "-", "@")):
            if "|" in text:
                raise UnsafeFormulaError(f"Formula contains unsafe DDE command execution syntax: {formula}")
            if UNSAFE_FORMULA_PATTERN.search(text):
                raise UnsafeFormulaError(f"Formula contains unsafe function: {formula}")

def _create_backup(path: str) -> None:
    p = Path(path)
    if not p.exists():
        return
    timestamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    target_backup_dir = BACKUP_DIR or getattr(settings, "BACKUP_DIR", "")
    if target_backup_dir:
        backup_dir = Path(target_backup_dir)
        backup_dir.mkdir(parents=True, exist_ok=True)
        backup_path = backup_dir / f"{p.name}.{timestamp}.bak"
    else:
        backup_path = p.with_name(f"{p.name}.{timestamp}.bak")
    shutil.copy2(path, backup_path)

def restore_backup(backup_path: str, target_path: str) -> dict[str, Any]:
    b = Path(backup_path)
    if not b.exists():
        raise FileNotFoundError(f"Backup file not found: {backup_path}")
    if b.suffix.lower() != ".bak":
        raise ValueError(f"Refusing to restore from non-backup file: {backup_path}. Must have .bak extension.")
    
    t = Path(target_path)
    if t.suffix.lower() not in _EXCEL_EXTENSIONS:
        raise ValueError(f"Refusing to restore to non-Excel file: {target_path}. Must be .xlsx/.xlsm.")

    tmp_path = f"{target_path}.tmp-{os.getpid()}"
    try:
        shutil.copy2(backup_path, tmp_path)
        os.replace(tmp_path, target_path)
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
    
    return {"restored": True, "target": target_path, "backup": backup_path, "message": "Backup restored successfully."}

def _load(path: str) -> Workbook:
    return openpyxl.load_workbook(path, data_only=False)


def _ws(wb: Workbook, sheet: str) -> Worksheet:
    if sheet not in wb.sheetnames:
        raise SheetNotFoundError(f"Sheet '{sheet}' not found. Available: {wb.sheetnames}")
    return wb[sheet]


def _atomic_save(wb: Workbook, path: str, backup: bool = False) -> None:
    """Save to a temp file in the same directory, then atomically replace `path`.

    A plain `wb.save(path)` writes straight onto the target — if the process dies
    mid-write (crash, OOM kill, disk full), the file is left half-written and
    unrecoverable. `os.replace` only swaps the two once the temp file is complete.
    """
    if backup:
        _create_backup(path)
    tmp_path = f"{path}.tmp-{os.getpid()}"
    try:
        wb.save(tmp_path)
        os.replace(tmp_path, path)
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)


def _finalize(wb: Workbook, path: str, recalc: bool = True, backup: bool = False) -> dict[str, Any]:
    """Save via openpyxl (preserves everything it didn't touch), then recalculate.

    `recalc=False` skips the LibreOffice round-trip for structural edits (merge,
    style, sheet add/remove) that can't produce a stale formula result — recalc is
    still cheap to run for anything that touches cell values or formulas.
    """
    _atomic_save(wb, path, backup=backup)
    wb.close()
    if not recalc:
        return {"saved": True, "recalculated": False, "errors_found": [], "message": "Saved (no recalculation needed)."}

    result = recalculate(path)
    return {
        "saved": True,
        "recalculated": result.success,
        "errors_found": result.errors_found,
        "message": result.message,
    }


_EXCEL_EXTENSIONS = (".xlsx", ".xlsm")


def create_workbook(path: str, sheets: list[str] | None = None, overwrite: bool = False) -> dict[str, Any]:
    p = Path(path)
    if not p.suffix.lower() in _EXCEL_EXTENSIONS:
        raise ValueError(f"Refusing to write to a non-Excel file: {path}. Use a .xlsx/.xlsm path.")
    if p.exists() and not overwrite:
        raise FileExistsError(f"File already exists: {path}. Pass overwrite=True to replace it.")

    wb = openpyxl.Workbook()
    names = sheets or ["Sheet1"]
    wb.active.title = names[0]
    for name in names[1:]:
        wb.create_sheet(name)

    p.parent.mkdir(parents=True, exist_ok=True)
    _atomic_save(wb, path, backup=False)
    wb.close()
    return {
        "saved": True,
        "path": str(p),
        "sheets": names,
        "recalculated": False,
        "errors_found": [],
        "message": "Workbook created.",
    }


def write_cells(
    path: str,
    sheet: str,
    cells: list[dict[str, Any]],
    create_sheet_if_missing: bool = False,
    recalculate: bool = True,
    allow_external_formulas: bool = False,
    backup: bool = AUTO_BACKUP,
) -> dict[str, Any]:
    if not cells:
        return {"saved": True, "recalculated": False, "errors_found": [], "message": "Nothing to write; file unchanged."}
    wb = _load(path)
    if sheet not in wb.sheetnames and create_sheet_if_missing:
        wb.create_sheet(sheet)
    ws = _ws(wb, sheet)

    for item in cells:
        coord = item.get("cell")
        if not coord:
            raise ValueError(f"Each item in `cells` requires a 'cell' key, got: {item}")
        formula = item.get("formula")
        value = item.get("value")
        if formula is not None:
            _validate_formula(formula, allow_external_formulas)
        elif isinstance(value, str):
            _validate_formula(value, allow_external_formulas)
        ws[coord] = formula if formula is not None else value

    return _finalize(wb, path, recalc=recalculate, backup=backup)


def append_rows(
    path: str,
    sheet: str,
    rows: list[list[Any]],
    create_sheet_if_missing: bool = False,
    recalculate: bool = True,
    allow_external_formulas: bool = False,
    backup: bool = AUTO_BACKUP,
) -> dict[str, Any]:
    if not rows:
        return {"saved": True, "recalculated": False, "errors_found": [], "message": "Nothing to write; file unchanged."}
    wb = _load(path)
    if sheet not in wb.sheetnames and create_sheet_if_missing:
        wb.create_sheet(sheet)
    ws = _ws(wb, sheet)

    for row in rows:
        for cell_val in row:
            if isinstance(cell_val, str):
                _validate_formula(cell_val, allow_external_formulas)
        ws.append(row)

    return _finalize(wb, path, recalc=recalculate, backup=backup)


def create_sheet(path: str, sheet: str, index: int | None = None, backup: bool = AUTO_BACKUP) -> dict[str, Any]:
    wb = _load(path)
    if sheet in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet}' already exists")
    wb.create_sheet(sheet, index)
    return _finalize(wb, path, recalc=False, backup=backup)


def delete_sheet(path: str, sheet: str, backup: bool = AUTO_BACKUP) -> dict[str, Any]:
    wb = _load(path)
    _ws(wb, sheet)
    if len(wb.sheetnames) == 1:
        raise ValueError("Cannot delete the only sheet in a workbook")
    del wb[sheet]
    return _finalize(wb, path, recalc=False, backup=backup)


def insert_rows(
    path: str, sheet: str, start_row: int, count: int = 1, recalculate: bool = True, backup: bool = AUTO_BACKUP
) -> dict[str, Any]:
    if start_row < 1:
        raise ValueError(f"start_row must be >= 1, got {start_row}")
    if count < 1:
        raise ValueError(f"count must be >= 1, got {count}")
    wb = _load(path)
    _ws(wb, sheet).insert_rows(start_row, count)
    return _finalize(wb, path, recalc=recalculate, backup=backup)


def delete_rows(
    path: str, sheet: str, start_row: int, count: int = 1, recalculate: bool = True, backup: bool = AUTO_BACKUP
) -> dict[str, Any]:
    if start_row < 1:
        raise ValueError(f"start_row must be >= 1, got {start_row}")
    if count < 1:
        raise ValueError(f"count must be >= 1, got {count}")
    wb = _load(path)
    _ws(wb, sheet).delete_rows(start_row, count)
    return _finalize(wb, path, recalc=recalculate, backup=backup)


def insert_columns(
    path: str, sheet: str, start_column: int, count: int = 1, recalculate: bool = True, backup: bool = AUTO_BACKUP
) -> dict[str, Any]:
    if start_column < 1:
        raise ValueError(f"start_column must be >= 1, got {start_column}")
    if count < 1:
        raise ValueError(f"count must be >= 1, got {count}")
    wb = _load(path)
    _ws(wb, sheet).insert_cols(start_column, count)
    return _finalize(wb, path, recalc=recalculate, backup=backup)


def delete_columns(
    path: str, sheet: str, start_column: int, count: int = 1, recalculate: bool = True, backup: bool = AUTO_BACKUP
) -> dict[str, Any]:
    if start_column < 1:
        raise ValueError(f"start_column must be >= 1, got {start_column}")
    if count < 1:
        raise ValueError(f"count must be >= 1, got {count}")
    wb = _load(path)
    _ws(wb, sheet).delete_cols(start_column, count)
    return _finalize(wb, path, recalc=recalculate, backup=backup)


def merge_cells(path: str, sheet: str, cell_range: str, backup: bool = AUTO_BACKUP) -> dict[str, Any]:
    wb = _load(path)
    _ws(wb, sheet).merge_cells(cell_range)
    return _finalize(wb, path, recalc=False, backup=backup)


def unmerge_cells(path: str, sheet: str, cell_range: str, backup: bool = AUTO_BACKUP) -> dict[str, Any]:
    wb = _load(path)
    _ws(wb, sheet).unmerge_cells(cell_range)
    return _finalize(wb, path, recalc=False, backup=backup)


def set_cell_style(path: str, sheet: str, cell_range: str, style: dict[str, Any], backup: bool = AUTO_BACKUP) -> dict[str, Any]:
    wb = _load(path)
    ws = _ws(wb, sheet)
    min_col, min_row, max_col, max_row = range_boundaries(cell_range)
    min_col = min_col or 1
    min_row = min_row or 1
    if max_col is None:
        max_col = max(ws.max_column or 1, min_col)
    if max_row is None:
        max_row = max(ws.max_row or 1, min_row)

    font_kwargs = {k: style[k] for k in ("bold", "italic") if k in style}
    if "font_size" in style:
        font_kwargs["size"] = style["font_size"]
    if "font_color" in style:
        font_kwargs["color"] = style["font_color"]
    font = Font(**font_kwargs) if font_kwargs else None

    fill = None
    if "bg_color" in style:
        fill = PatternFill(start_color=style["bg_color"], end_color=style["bg_color"], fill_type="solid")

    alignment = None
    if "horizontal" in style or "vertical" in style:
        alignment = Alignment(horizontal=style.get("horizontal"), vertical=style.get("vertical"))

    border = None
    if "border" in style:
        side = Side(style=style["border"])
        border = Border(left=side, right=side, top=side, bottom=side)

    number_format = style.get("number_format")

    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            if font:
                cell.font = font
            if fill:
                cell.fill = fill
            if alignment:
                cell.alignment = alignment
            if border:
                cell.border = border
            if number_format:
                cell.number_format = number_format

    return _finalize(wb, path, recalc=False, backup=backup)


def _validate_sheet_name(name: str) -> None:
    if not name or len(name) > 31:
        raise ValueError("Sheet name must be 1-31 characters long")
    if any(c in name for c in r"\/?*[]:"):
        raise ValueError(r"Sheet name cannot contain \ / ? * : [ ]")


def rename_sheet(path: str, old_name: str, new_name: str, backup: bool = AUTO_BACKUP) -> dict[str, Any]:
    """Rename an existing sheet.

    Note: Renaming a sheet does not automatically rewrite formulas in other sheets
    referencing the old name (openpyxl limitation). Dependent formulas will break.
    """
    wb = _load(path)
    ws = _ws(wb, old_name)

    if old_name == new_name:
        return _finalize(wb, path, recalc=False, backup=backup)

    if new_name in wb.sheetnames:
        raise ValueError(f"Sheet '{new_name}' already exists")
    _validate_sheet_name(new_name)

    ws.title = new_name
    return _finalize(wb, path, recalc=False, backup=backup)


def copy_sheet(path: str, source_sheet: str, target_sheet: str, backup: bool = AUTO_BACKUP) -> dict[str, Any]:
    wb = _load(path)
    source_ws = _ws(wb, source_sheet)

    if target_sheet in wb.sheetnames:
        raise ValueError(f"Sheet '{target_sheet}' already exists")
    _validate_sheet_name(target_sheet)

    new_ws = wb.copy_worksheet(source_ws)
    new_ws.title = target_sheet
    return _finalize(wb, path, recalc=False, backup=backup)


def autofit_columns(
    path: str, sheet: str, min_width: int = 10, max_width: int = 50, padding: int = 3, backup: bool = AUTO_BACKUP
) -> dict[str, Any]:
    if min_width < 0:
        raise ValueError(f"min_width must be non-negative, got {min_width}")
    if max_width < 0:
        raise ValueError(f"max_width must be non-negative, got {max_width}")
    if min_width > max_width:
        raise ValueError(f"min_width ({min_width}) cannot be greater than max_width ({max_width})")

    wb = _load(path)
    ws = _ws(wb, sheet)

    for col in ws.columns:
        if not col:
            continue
        max_length = 0
        column_letter = col[0].column_letter

        for cell in col:
            try:
                if cell.value is not None:
                    max_length = max(max_length, len(str(cell.value)))
            except Exception:
                pass

        adjusted_width = min(max_width, max(min_width, max_length + padding))
        ws.column_dimensions[column_letter].width = adjusted_width

    return _finalize(wb, path, recalc=False, backup=backup)


def clear_range(
    path: str, sheet: str, cell_range: str, clear_values: bool = True, clear_styles: bool = False, backup: bool = AUTO_BACKUP
) -> dict[str, Any]:
    wb = _load(path)
    ws = _ws(wb, sheet)
    min_col, min_row, max_col, max_row = range_boundaries(cell_range)

    # Bound coordinates to worksheet used area to prevent DoS / OOM on full-column/row ranges like "A:A"
    if ws.max_row is None or ws.max_column is None:
        return _finalize(wb, path, recalc=False, backup=backup)

    min_col = min_col or 1
    min_row = min_row or 1
    max_col = min(max_col, ws.max_column) if max_col is not None else ws.max_column
    max_row = min(max_row, ws.max_row) if max_row is not None else ws.max_row

    if min_col > max_col or min_row > max_row:
        return _finalize(wb, path, recalc=False, backup=backup)

    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            if clear_values:
                cell.value = None
            if clear_styles:
                cell.style = "Normal"

    return _finalize(wb, path, recalc=clear_values, backup=backup)



_TABLE_NAME_PATTERN = re.compile(r"^[A-Za-z_][A-Za-z0-9_.]*$")
_CELL_COORD_PATTERN = re.compile(r"^[A-Za-z]+[1-9][0-9]*$")
_R1C1_PATTERN = re.compile(r"^R[1-9][0-9]*C[1-9][0-9]*$", re.IGNORECASE)


def _validate_table_name(name: str) -> None:
    if not name or len(name) > 255 or not _TABLE_NAME_PATTERN.match(name):
        raise ValueError(
            f"Invalid table name '{name}'. Table names must start with a letter or underscore, "
            "contain only alphanumeric characters, underscores, or periods, and cannot contain spaces."
        )
    if _CELL_COORD_PATTERN.match(name) or _R1C1_PATTERN.match(name):
        raise ValueError(f"Invalid table name '{name}'. Table names cannot be cell coordinates.")


def create_table(
    path: str,
    sheet: str,
    cell_range: str,
    table_name: str,
    style_name: str = "TableStyleMedium9",
    show_filter: bool = True,
    show_row_stripes: bool = True,
    backup: bool = AUTO_BACKUP
) -> dict[str, Any]:
    """Create a formal Excel Table on the specified range.

    Args:
        path: Path to the .xlsx file.
        sheet: Sheet name.
        cell_range: A1-style range (e.g. "A1:D10").
        table_name: Unique table identifier across the workbook.
        style_name: Table style name (e.g. "TableStyleMedium9").
        show_filter: Whether to display auto-filter dropdown arrows.
        show_row_stripes: Whether to apply alternating row shading.
        backup: Whether to create a backup before modifying the workbook.

    Returns:
        Standardized write result dict.
    """
    from openpyxl.worksheet.table import Table, TableStyleInfo
    from openpyxl.worksheet.filters import AutoFilter

    _validate_table_name(table_name)
    wb = _load(path)
    ws = _ws(wb, sheet)

    for other_ws in wb.worksheets:
        if table_name in other_ws.tables:
            raise ValueError(f"Table name '{table_name}' already exists in sheet '{other_ws.title}'")

    min_col, min_row, max_col, max_row = range_boundaries(cell_range)
    if min_col is None or min_row is None or max_col is None or max_row is None:
        raise ValueError(f"Invalid cell_range: {cell_range}")
    if min_row >= max_row:
        raise ValueError(f"Table range '{cell_range}' must span at least 2 rows (1 header row and at least 1 data row)")

    seen_headers: set[str] = set()
    for col_idx in range(min_col, max_col + 1):
        cell_val = ws.cell(row=min_row, column=col_idx).value
        if cell_val is None or str(cell_val).strip() == "":
            raise ValueError(
                f"Table header cell at row {min_row}, column {col_idx} is empty. "
                "Excel tables require all header cells in the top row to contain non-empty text."
            )
        header_str = str(cell_val).strip()
        if header_str in seen_headers:
            raise ValueError(
                f"Duplicate table column header '{header_str}' found at row {min_row}, column {col_idx}. "
                "Excel tables require unique column header names."
            )
        seen_headers.add(header_str)

    tab = Table(displayName=table_name, ref=cell_range)
    style = TableStyleInfo(
        name=style_name,
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=show_row_stripes,
        showColumnStripes=False
    )
    tab.tableStyleInfo = style

    if not show_filter:
        tab.autoFilter = None
    else:
        tab.autoFilter = AutoFilter(ref=cell_range)

    ws.add_table(tab)
    return _finalize(wb, path, recalc=False, backup=backup)


def create_chart(
    path: str,
    sheet: str,
    chart_type: str,
    data_range: str,
    categories_range: str | None = None,
    title: str = "Chart",
    target_cell: str = "E2",
    backup: bool = AUTO_BACKUP
) -> dict[str, Any]:
    """Create a native Excel chart (bar, line, pie, scatter).

    Args:
        path: Path to the .xlsx file.
        sheet: Sheet name.
        chart_type: One of "bar", "line", "pie", "scatter".
        data_range: A1-style range holding the chart series values (e.g. "B1:B10").
        categories_range: Optional A1-style range holding series category labels (e.g. "A2:A10").
        title: Title string displayed on the chart.
        target_cell: Cell coordinate where the top-left corner of the chart is placed.
        backup: Whether to create a backup before modifying the workbook.

    Returns:
        Standardized write result dict.
    """
    from openpyxl.chart import BarChart, LineChart, PieChart, ScatterChart, Reference

    if not isinstance(target_cell, str) or not _CELL_COORD_PATTERN.match(target_cell.strip()):
        raise ValueError(f"Invalid target_cell coordinate: '{target_cell}'. Must be an A1-style reference like 'E2'.")
    target_cell = target_cell.strip().upper()

    wb = _load(path)
    ws = _ws(wb, sheet)

    chart_map = {
        "bar": BarChart,
        "line": LineChart,
        "pie": PieChart,
        "scatter": ScatterChart
    }

    if chart_type.lower() not in chart_map:
        raise ValueError(f"Unsupported chart type: {chart_type}. Supported: bar, line, pie, scatter.")

    chart = chart_map[chart_type.lower()]()
    chart.title = title

    d_min_col, d_min_row, d_max_col, d_max_row = range_boundaries(data_range)
    if any(v is None for v in (d_min_col, d_min_row, d_max_col, d_max_row)):
        raise ValueError(f"Invalid data_range: {data_range}")

    data = Reference(ws, min_col=d_min_col, min_row=d_min_row, max_col=d_max_col, max_row=d_max_row)
    chart.add_data(data, titles_from_data=True)

    if categories_range:
        c_min_col, c_min_row, c_max_col, c_max_row = range_boundaries(categories_range)
        if any(v is None for v in (c_min_col, c_min_row, c_max_col, c_max_row)):
            raise ValueError(f"Invalid categories_range: {categories_range}")
        cats = Reference(ws, min_col=c_min_col, min_row=c_min_row, max_col=c_max_col, max_row=c_max_row)
        chart.set_categories(cats)

    ws.add_chart(chart, target_cell)

    return _finalize(wb, path, recalc=False, backup=backup)

