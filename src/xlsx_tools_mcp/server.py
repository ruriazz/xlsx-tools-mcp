"""MCP server exposing accurate, structure-preserving read/write tools for Excel (.xlsx) files.

Read path: python-calamine (fast, accurate types) with an openpyxl fallback for
formulas/styles/comments. Write path: openpyxl (preserves everything it doesn't
touch — styles, merges, other formulas) followed by a LibreOffice headless pass
whose computed formula values are injected back into the file (or, in
``roundtrip`` mode, whose export replaces it).

Each tool is a thin wrapper: resolve the path, take the per-file lock, call one
io-layer function. Parameter meanings live in the input schema (see ``params``),
tool-choice guidance in the server instructions, and every tool declares MCP
annotations (read-only / destructive / idempotent) so clients can gate calls.
"""

from __future__ import annotations

import inspect
from typing import Annotated, Any, Callable, TypeVar

from mcp.server import MCPServer
from mcp.types import ToolAnnotations
from pydantic import Field

from . import ledger
from . import params as P
from . import rules as rules_mod
from .errors import XlsxMcpError
from .io import copies, reader, review, sharing, transform, writer
from .io.writer import restore_backup as io_restore_backup
from .locking import file_lock, file_locks
from .recalc import recalculate
from .settings import CONFIGURED_FILES, resolve_literal_path, resolve_path

T = TypeVar("T")

_files_note = (
    f"\n\nPreloaded files (pass as `path`, or omit `path` when there is only one): "
    f"{', '.join(sorted(CONFIGURED_FILES))}."
    if CONFIGURED_FILES
    else ""
)

INSTRUCTIONS = """\
Read and write Excel .xlsx/.xlsm workbooks. Writes keep every part of the file a tool does not touch.

Pick a read tool:
- New workbook: list_sheets (names, sizes), then profile_sheet (column names, types, min/max).
- Need rows: read_sheet with cell_range or max_rows. For large sheets use query_sheet (rows matching a filter) or aggregate_sheet (totals per group) instead of reading everything.
- Find where a value is: search_workbook. One cell's formula, style or comment: get_cell. Hidden sheets, exact size, defined names: get_workbook_info.

Pick a write tool:
- Set cells at known addresses (values or formulas): write_cells. Add records below the existing data: append_rows.
- Make or remove space, shifting other cells: insert_rows, delete_rows, insert_columns, delete_columns. Empty cells but keep the layout: clear_range.

Write responses: {saved, recalculated, errors_found, message}, plus recalc_mode, rules, rolled_back and ripple when relevant. \
Check errors_found (formula cells showing #REF!, #DIV/0!, ...) before reporting success. \
rolled_back=true means a rule from set_rules failed and the file is unchanged. ripple lists other formula results the write changed. \
Read tools return formulas' last computed results; writes recompute them with LibreOffice unless recalculate=false.\
"""

mcp = MCPServer("xlsx-tools-mcp", instructions=INSTRUCTIONS + _files_note)


def _tool(title: str, *, read_only: bool = False, destructive: bool = False, idempotent: bool = False):
    """Register an MCP tool with a title, a dedented description and behavior annotations.

    `destructive` means the call can remove or overwrite existing content (cells,
    sheets, rules or whole files). All tools work on local files only (closed world).
    """
    # Hints equal to the MCP defaults (readOnly=false, destructive=true, idempotent=false)
    # are left out to keep tools/list small; openWorld defaults to true, so it is always set.
    hints: dict[str, bool] = {"open_world_hint": False}
    if read_only:
        hints.update(read_only_hint=True)
    elif not destructive:
        hints.update(destructive_hint=False)
    if idempotent and not read_only:
        hints.update(idempotent_hint=True)
    annotations = ToolAnnotations(**hints)

    def register(fn: Callable[..., T]) -> Callable[..., T]:
        description = inspect.cleandoc(fn.__doc__ or "")
        return mcp.tool(title=title, description=description, annotations=annotations)(fn)

    return register


def _run(fn: Callable[[], T]) -> T:
    """Translate internal domain errors into MCP tool errors with a clear message."""
    try:
        return fn()
    except XlsxMcpError as exc:
        raise ValueError(str(exc)) from exc
    except FileExistsError as exc:
        raise ValueError(str(exc)) from exc
    except FileNotFoundError as exc:
        raise ValueError(str(exc)) from exc
    except PermissionError as exc:
        winerror = getattr(exc, "winerror", None)
        if winerror in (32, 33):
            raise ValueError(f"File is currently locked by another application: {getattr(exc, 'filename', None) or 'target file'}") from exc
        raise ValueError(f"Permission denied: {exc}") from exc


def _backup(backup: bool | None) -> bool:
    return backup if backup is not None else writer.AUTO_BACKUP


# ── Inspect / read ────────────────────────────────────────────────────────────


@_tool("List preloaded files", read_only=True)
def list_configured_files() -> dict[str, str]:
    """List workbooks preloaded at startup (XLSX_MCP_FILES) as name -> absolute path.

    Call this first when unsure which file to use: any tool's `path` accepts these names.
    """
    return dict(CONFIGURED_FILES)


@_tool("List sheets", read_only=True)
def list_sheets(path: P.Path = None) -> list[dict[str, Any]]:
    """List every sheet with its approximate row and column count. Fast.

    For exact dimensions, hidden state and defined names use get_workbook_info.
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: reader.list_sheets(path))


@_tool("Get workbook info", read_only=True)
def get_workbook_info(path: P.Path = None) -> dict[str, Any]:
    """Get workbook metadata: exact dimensions and visibility of each sheet, active sheet, defined names.

    Slower than list_sheets on large files.
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: reader.workbook_info(path))


@_tool("Read sheet values", read_only=True)
def read_sheet(
    sheet: P.Sheet,
    cell_range: P.OptionalCellRange = None,
    max_rows: Annotated[int | None, Field(description="Return at most this many rows. Omit for no limit.")] = None,
    offset_row: Annotated[int, Field(description="Skip this many rows first (after the header row for records/markdown); use with max_rows to page.")] = 0,
    format: Annotated[P.ReadFormat, Field(description='"array": 2D list. "records": list of objects keyed by the first row. "markdown": table text.')] = "array",
    path: P.Path = None,
) -> dict[str, Any]:
    """Return the raw cell values of a sheet or range. Formula cells give their last computed result.

    Best for small sheets or a known range. On large sheets, filter with query_sheet,
    summarize with aggregate_sheet, or inspect columns with profile_sheet instead.
    For a cell's formula text or style use get_cell.
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: reader.read_sheet(path, sheet, cell_range, max_rows, offset_row, format))


@_tool("Get cell details", read_only=True)
def get_cell(
    sheet: P.Sheet,
    cell: Annotated[str, Field(description='Cell address, e.g. "C5".')],
    path: P.Path = None,
) -> dict[str, Any]:
    """Get everything about one cell: computed value, formula text, number format, font, fill, merge state, comment."""
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: reader.get_cell(path, sheet, cell))


@_tool("Search workbook", read_only=True)
def search_workbook(
    query: Annotated[str, Field(description="Text to look for inside cell values (substring match).")],
    sheet: Annotated[str | None, Field(description="Search only this sheet. Omit to search all sheets.")] = None,
    match_case: Annotated[bool, Field(description="Match upper/lower case exactly.")] = False,
    limit: Annotated[int | None, Field(description="Return at most this many matches. Omit for no limit.")] = None,
    path: P.Path = None,
) -> list[dict[str, Any]]:
    """Find the cells whose value contains some text. Returns {sheet, cell, value} per match.

    Use to locate data when you don't know its address; to filter table rows by
    conditions use query_sheet.
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: reader.search_workbook(path, query, sheet, match_case, limit))


@_tool("Profile sheet columns", read_only=True)
def profile_sheet(
    sheet: P.Sheet,
    sample_rows: Annotated[int, Field(description="Example values to show per column.")] = 3,
    path: P.Path = None,
) -> dict[str, Any]:
    """Summarize each column of a table-shaped sheet (first row = headers): type, empty count, min/max, sample values.

    Use first on an unfamiliar sheet to learn its column names before query_sheet or
    aggregate_sheet. Returns no cell data beyond the samples.
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: reader.profile_sheet(path, sheet, sample_rows))


@_tool("Query sheet rows", read_only=True)
def query_sheet(
    sheet: P.Sheet,
    filter_query: Annotated[str, Field(description="Condition on header names, e.g. \"Status == 'PAID' and Amount > 500\". Allowed: comparisons, and/or/not, in, arithmetic. Header names must be valid identifiers (no spaces).")],
    columns: Annotated[list[str] | None, Field(description="Only return these columns. Omit for all.")] = None,
    max_rows: Annotated[int, Field(description="Return at most this many matching rows.")] = 100,
    cell_range: P.OptionalCellRange = None,
    path: P.Path = None,
) -> dict[str, Any]:
    """Return only the rows of a table-shaped sheet (first row = headers) that match a condition.

    Returns {columns, records, row_count, limited}; row_count counts all matches.
    Use instead of read_sheet when you need a subset of a large table. For totals
    per group use aggregate_sheet.
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: transform.query_sheet(path, sheet, filter_query, columns, max_rows, cell_range))


@_tool("Aggregate sheet", read_only=True)
def aggregate_sheet(
    sheet: P.Sheet,
    group_by: Annotated[list[str], Field(description='Header names to group by, e.g. ["Region"].')],
    agg: Annotated[dict[str, str], Field(description='Header name -> function: sum, mean, min, max, count, median, nunique, first or last. E.g. {"Amount": "sum"}.')],
    cell_range: P.OptionalCellRange = None,
    has_header: Annotated[bool, Field(description="First row holds column names. When false, columns are named col_0, col_1, ...")] = True,
    path: P.Path = None,
) -> dict[str, Any]:
    """Group rows and compute totals per group (e.g. sum of Amount by Region). Returns one record per group.

    Use for summaries of a table; to list individual matching rows use query_sheet.
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: transform.aggregate_sheet(path, sheet, group_by, agg, cell_range, has_header))


# ── Write: workbook and cells ─────────────────────────────────────────────────


@_tool("Create workbook", destructive=True)
def create_workbook(
    path: Annotated[str, Field(description="File path for the new .xlsx or .xlsm file (preloaded names do not apply).")],
    sheets: Annotated[list[str] | None, Field(description='Sheet names to create. Default: ["Sheet1"].')] = None,
    overwrite: Annotated[bool, Field(description="Replace an existing file at path. Refused by default.")] = False,
) -> dict[str, Any]:
    """Create a new, empty workbook file."""
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.create_workbook(path, sheets, overwrite))


@_tool("Write cells", destructive=True, idempotent=True)
def write_cells(
    sheet: P.Sheet,
    cells: Annotated[list[dict[str, Any]], Field(description='Cells to set: [{"cell": "A1", "value": 100}, {"cell": "B1", "formula": "=A1*2"}]. Existing content is overwritten.')],
    create_sheet_if_missing: P.CreateSheetIfMissing = False,
    recalculate: P.Recalculate = True,
    allow_external_formulas: P.AllowExternalFormulas = False,
    backup: P.Backup = None,
    reason: P.Reason = None,
    path: P.Path = None,
) -> dict[str, Any]:
    """Set values or formulas in cells at given addresses, overwriting what is there.

    To add new records below existing data without computing addresses use
    append_rows; to make room between rows use insert_rows.
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.write_cells(path, sheet, cells, create_sheet_if_missing, recalculate, allow_external_formulas, _backup(backup), reason=reason))


@_tool("Append rows")
def append_rows(
    sheet: P.Sheet,
    rows: Annotated[list[list[Any]], Field(description='Rows to add, each a list of values from column A, e.g. [["INV-3", 7, "=B5*2"]]. Strings starting with "=" become formulas.')],
    create_sheet_if_missing: P.CreateSheetIfMissing = False,
    recalculate: P.Recalculate = True,
    allow_external_formulas: P.AllowExternalFormulas = False,
    backup: P.Backup = None,
    reason: P.Reason = None,
    path: P.Path = None,
) -> dict[str, Any]:
    """Add rows below the last used row of a sheet. Never overwrites existing cells.

    To change existing cells use write_cells.
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.append_rows(path, sheet, rows, create_sheet_if_missing, recalculate, allow_external_formulas, _backup(backup), reason=reason))


@_tool("Clear range", destructive=True, idempotent=True)
def clear_range(
    sheet: P.Sheet,
    cell_range: P.CellRange,
    clear_values: Annotated[bool, Field(description="Remove values and formulas.")] = True,
    clear_styles: Annotated[bool, Field(description="Reset formatting to the Normal style.")] = False,
    backup: P.Backup = None,
    reason: P.Reason = None,
    path: P.Path = None,
) -> dict[str, Any]:
    """Empty the cells in a range (values and/or formatting). Rows and columns stay in place.

    To remove rows or columns and shift the rest, use delete_rows or delete_columns.
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.clear_range(path, sheet, cell_range, clear_values, clear_styles, backup=_backup(backup), reason=reason))


@_tool("Recalculate workbook", destructive=True, idempotent=True)
def recalculate_workbook(
    mode: Annotated[P.RecalcMode | None, Field(description='"inject": write only computed values, keep every other part. "roundtrip": replace the file with LibreOffice\'s export (may drop pivots, charts, validation). Default: XLSX_MCP_RECALC_MODE (inject).')] = None,
    path: P.Path = None,
) -> dict[str, Any]:
    """Recompute every formula with LibreOffice and report formula errors.

    Writes already do this. Needed only after writes with recalculate=false or
    after the file was edited outside this server.
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        result = _run(lambda: recalculate(path, mode=mode))
        return {
            "saved": True,
            "path": path,
            "recalculated": result.success,
            "errors_found": result.errors_found,
            "message": result.message,
            "recalc_mode": result.mode,
        }


@_tool("Restore backup", destructive=True, idempotent=True)
def restore_backup(
    backup_path: Annotated[str, Field(description="The .bak file to restore from (created by writes with backup enabled).")],
    target_path: Annotated[str, Field(description="The .xlsx/.xlsm workbook to overwrite.")],
) -> dict[str, Any]:
    """Replace a workbook with one of its .bak backups, atomically."""
    backup_path = _run(lambda: resolve_path(backup_path))
    target_path = _run(lambda: resolve_path(target_path))
    with file_lock(target_path):
        return _run(lambda: io_restore_backup(backup_path, target_path))


# ── Write: sheets ─────────────────────────────────────────────────────────────


@_tool("Create sheet")
def create_sheet(
    sheet: Annotated[str, Field(description="Name for the new sheet; must not exist yet.")],
    index: Annotated[int | None, Field(description="Zero-based position. Omit to add at the end.")] = None,
    backup: P.Backup = None,
    reason: P.Reason = None,
    path: P.Path = None,
) -> dict[str, Any]:
    """Add a new empty sheet. To duplicate an existing sheet use copy_sheet."""
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.create_sheet(path, sheet, index, backup=_backup(backup), reason=reason))


@_tool("Copy sheet")
def copy_sheet(
    source_sheet: Annotated[str, Field(description="Sheet to duplicate.")],
    target_sheet: Annotated[str, Field(description="Name for the copy (1-31 chars, must not exist yet).")],
    backup: P.Backup = None,
    reason: P.Reason = None,
    path: P.Path = None,
) -> dict[str, Any]:
    """Duplicate a sheet with its values, formulas and styles, inside the same workbook."""
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.copy_sheet(path, source_sheet, target_sheet, backup=_backup(backup), reason=reason))


@_tool("Rename sheet", destructive=True)
def rename_sheet(
    old_name: Annotated[str, Field(description="Current sheet name.")],
    new_name: Annotated[str, Field(description="New sheet name (1-31 chars, none of \\ / ? * : [ ]).")],
    backup: P.Backup = None,
    reason: P.Reason = None,
    path: P.Path = None,
) -> dict[str, Any]:
    """Rename a sheet. Formulas in other sheets that use the old name are NOT updated and will break."""
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.rename_sheet(path, old_name, new_name, backup=_backup(backup), reason=reason))


@_tool("Delete sheet", destructive=True)
def delete_sheet(
    sheet: Annotated[str, Field(description="Sheet to delete.")],
    backup: P.Backup = None,
    reason: P.Reason = None,
    path: P.Path = None,
) -> dict[str, Any]:
    """Delete a sheet and everything on it. Refused when it is the only sheet."""
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.delete_sheet(path, sheet, backup=_backup(backup), reason=reason))


# ── Write: rows and columns (cells shift) ─────────────────────────────────────

_StartRow = Annotated[int, Field(description="1-based row number.")]
_StartColumn = Annotated[int, Field(description="1-based column number (A=1, B=2, ...).")]
_Count = Annotated[int, Field(description="How many to insert or delete.")]


@_tool("Insert rows")
def insert_rows(
    sheet: P.Sheet,
    start_row: _StartRow,
    count: _Count = 1,
    recalculate: P.Recalculate = True,
    backup: P.Backup = None,
    reason: P.Reason = None,
    path: P.Path = None,
) -> dict[str, Any]:
    """Insert blank rows before start_row; rows below move down."""
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.insert_rows(path, sheet, start_row, count, recalculate, backup=_backup(backup), reason=reason))


@_tool("Delete rows", destructive=True)
def delete_rows(
    sheet: P.Sheet,
    start_row: _StartRow,
    count: _Count = 1,
    recalculate: P.Recalculate = True,
    backup: P.Backup = None,
    reason: P.Reason = None,
    path: P.Path = None,
) -> dict[str, Any]:
    """Delete rows starting at start_row; rows below move up. To empty cells without shifting use clear_range."""
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.delete_rows(path, sheet, start_row, count, recalculate, backup=_backup(backup), reason=reason))


@_tool("Insert columns")
def insert_columns(
    sheet: P.Sheet,
    start_column: _StartColumn,
    count: _Count = 1,
    recalculate: P.Recalculate = True,
    backup: P.Backup = None,
    reason: P.Reason = None,
    path: P.Path = None,
) -> dict[str, Any]:
    """Insert blank columns before start_column; columns to the right move right."""
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.insert_columns(path, sheet, start_column, count, recalculate, backup=_backup(backup), reason=reason))


@_tool("Delete columns", destructive=True)
def delete_columns(
    sheet: P.Sheet,
    start_column: _StartColumn,
    count: _Count = 1,
    recalculate: P.Recalculate = True,
    backup: P.Backup = None,
    reason: P.Reason = None,
    path: P.Path = None,
) -> dict[str, Any]:
    """Delete columns starting at start_column; columns to the right move left. To empty cells without shifting use clear_range."""
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.delete_columns(path, sheet, start_column, count, recalculate, backup=_backup(backup), reason=reason))


# ── Write: formatting and objects ─────────────────────────────────────────────


@_tool("Merge cells", destructive=True)
def merge_cells(
    sheet: P.Sheet,
    cell_range: P.CellRange,
    backup: P.Backup = None,
    reason: P.Reason = None,
    path: P.Path = None,
) -> dict[str, Any]:
    """Merge a rectangular range into one cell. Only the top-left value is kept."""
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.merge_cells(path, sheet, cell_range, backup=_backup(backup), reason=reason))


@_tool("Unmerge cells")
def unmerge_cells(
    sheet: P.Sheet,
    cell_range: Annotated[str, Field(description='The merged range exactly as merged, e.g. "A1:C1".')],
    backup: P.Backup = None,
    reason: P.Reason = None,
    path: P.Path = None,
) -> dict[str, Any]:
    """Split a merged range back into separate cells."""
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.unmerge_cells(path, sheet, cell_range, backup=_backup(backup), reason=reason))


@_tool("Set cell style", idempotent=True)
def set_cell_style(
    sheet: P.Sheet,
    cell_range: P.CellRange,
    style: Annotated[dict[str, Any], Field(description='Any of: bold, italic (bool); font_size (number); font_color, bg_color (hex RGB, e.g. "FF0000"); horizontal, vertical (alignment, e.g. "center"); border ("thin", "medium", "thick", ...); number_format (e.g. "#,##0.00").')],
    backup: P.Backup = None,
    reason: P.Reason = None,
    path: P.Path = None,
) -> dict[str, Any]:
    """Apply font, fill, alignment, border or number format to every cell in a range. Values are not changed."""
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.set_cell_style(path, sheet, cell_range, style, backup=_backup(backup), reason=reason))


@_tool("Autofit columns", idempotent=True)
def autofit_columns(
    sheet: P.Sheet,
    min_width: Annotated[int, Field(description="Narrowest width, in characters.")] = 10,
    max_width: Annotated[int, Field(description="Widest width, in characters.")] = 50,
    padding: Annotated[int, Field(description="Characters added to the longest value.")] = 3,
    backup: P.Backup = None,
    reason: P.Reason = None,
    path: P.Path = None,
) -> dict[str, Any]:
    """Set each column's width to fit its longest value, within min_width..max_width."""
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.autofit_columns(path, sheet, min_width, max_width, padding, backup=_backup(backup), reason=reason))


@_tool("Create table")
def create_table(
    sheet: P.Sheet,
    cell_range: Annotated[str, Field(description='Range including the header row, e.g. "A1:D10". Headers must be non-empty and unique.')],
    table_name: Annotated[str, Field(description="Workbook-unique name: letters, digits, _ or ., starting with a letter or _.")],
    style_name: Annotated[str, Field(description='Excel table style, e.g. "TableStyleMedium9".')] = "TableStyleMedium9",
    show_filter: Annotated[bool, Field(description="Show filter buttons in the header row.")] = True,
    show_row_stripes: Annotated[bool, Field(description="Shade alternate rows.")] = True,
    backup: P.Backup = None,
    reason: P.Reason = None,
    path: P.Path = None,
) -> dict[str, Any]:
    """Turn a range into a formatted Excel Table (ListObject) with optional filter buttons."""
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.create_table(path, sheet, cell_range, table_name, style_name, show_filter, show_row_stripes, backup=_backup(backup), reason=reason))


@_tool("Create chart")
def create_chart(
    sheet: P.Sheet,
    chart_type: P.ChartType,
    data_range: Annotated[str, Field(description='Series values with the series name in the first row, e.g. "B1:B10".')],
    categories_range: Annotated[str | None, Field(description='Category labels, e.g. "A2:A10".')] = None,
    title: Annotated[str, Field(description="Chart title.")] = "Chart",
    target_cell: Annotated[str, Field(description="Cell where the chart's top-left corner is placed.")] = "E2",
    backup: P.Backup = None,
    reason: P.Reason = None,
    path: P.Path = None,
) -> dict[str, Any]:
    """Add a native Excel chart built from ranges on the same sheet."""
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.create_chart(path, sheet, chart_type, data_range, categories_range, title, target_cell, backup=_backup(backup), reason=reason))


# ── Rules, review, sharing, change log ───────────────────────────────────────


@_tool("Set rules", destructive=True, idempotent=True)
def set_rules(
    rules: Annotated[list[dict[str, Any]], Field(description="Rule objects, each with a unique id and a type (see tool description). An empty list removes all rules.")],
    on_violation: Annotated[P.OnViolation, Field(description='"rollback": undo a write that breaks a rule. "warn": keep it and report.')] = "rollback",
    on_unverifiable: Annotated[P.OnUnverifiable, Field(description='What to do when a rule cannot be checked (e.g. LibreOffice missing): "warn" or "rollback".')] = "warn",
    path: P.Path = None,
) -> dict[str, Any]:
    """Replace the workbook's rules and check them once; later value-changing writes are checked automatically.

    A write that makes a passing rule fail (or a failing one worse) is rolled back;
    rules already failing never block. Stored as "<file>.rules.json". Each rule has an
    "id" plus one of:
      {"type": "cells_equal", "left": "S!B40", "right": "S!C40", "tolerance": 0.01}
      {"type": "cell_compare", "cell": "S!D4", "op": "<=", "value": 1000}  (or "cell_ref")
      {"type": "column_check", "sheet": "Data", "expr": "Qty >= 0", "cell_range": "A1:F500"}
      {"type": "column_unique", "sheet": "Data", "column": "Invoice"}
      {"type": "not_empty", "range": "Data!A2:A500"}
      {"type": "no_new_errors"}
    """
    path = _run(lambda: resolve_path(path))

    def apply() -> dict[str, Any]:
        sheets = reader.list_sheets(path)
        ruleset, warnings = rules_mod.validate(rules, on_violation, on_unverifiable, [s["name"] for s in sheets])
        rules_mod.save_ruleset(path, ruleset)
        if not ruleset.rules:
            return {"rules": None, "evaluation": None, "warnings": warnings, "message": "Rules removed."}
        evaluation = rules_mod.evaluate(path, ruleset)
        return {
            "rules": ruleset.to_json(),
            "evaluation": {"status": rules_mod.status_of(evaluation.results), "results": evaluation.results},
            "warnings": warnings,
            "message": f"Saved {len(ruleset.rules)} rule(s) to {rules_mod.sidecar_path(path)}.",
        }

    with file_lock(path):
        return _run(apply)


@_tool("Check rules", read_only=True)
def check_rules(path: P.Path = None) -> dict[str, Any]:
    """Check the workbook's rules (from set_rules) against its current values without writing."""
    path = _run(lambda: resolve_path(path))

    def check() -> dict[str, Any]:
        ruleset = rules_mod.load_ruleset(path)
        if ruleset is None:
            return {"rules": None, "evaluation": None}
        evaluation = rules_mod.evaluate(path, ruleset)
        return {
            "rules": ruleset.to_json(),
            "evaluation": {"status": rules_mod.status_of(evaluation.results), "results": evaluation.results},
        }

    with file_lock(path):
        return _run(check)


@_tool("Export review copy", destructive=True)
def export_review_copy(
    baseline_path: Annotated[str | None, Field(description="Workbook to compare against. Default: this file's newest .bak backup.")] = None,
    output_path: Annotated[str | None, Field(description='Where to write the copy. Default: "<name>.review.xlsx" next to the file.')] = None,
    overwrite: P.Overwrite = False,
    highlight_color: Annotated[str, Field(description="Hex RGB fill for changed cells.")] = "FFF2CC",
    include_recalc_changes: Annotated[bool, Field(description="Also mark formulas whose text is the same but whose result changed.")] = False,
    max_annotated_cells: Annotated[int, Field(description="Highlight at most this many cells; all changes are still listed in _Changes.")] = 2000,
    path: P.Path = None,
) -> dict[str, Any]:
    """Write a copy for a human reviewer: cells changed since a baseline are highlighted with a Before/After comment.

    A `_Changes` sheet lists every change with a link. The workbook itself is not modified.
    """
    path = _run(lambda: resolve_path(path))
    baseline = _run(lambda: resolve_literal_path(baseline_path or review.find_latest_backup(path)))
    output = _run(lambda: resolve_literal_path(output_path or review.default_output(path)))
    copies.ensure_distinct(output, path, baseline)
    # The baseline is only read; backups are written once and replaced atomically.
    with file_locks(path, output):
        return _run(
            lambda: review.export_review_copy(
                path, baseline, output, overwrite, highlight_color, include_recalc_changes, max_annotated_cells
            )
        )


@_tool("Inspect for sharing", read_only=True)
def inspect_for_sharing(path: P.Path = None) -> dict[str, Any]:
    """Report hidden or sensitive content to review before sharing a workbook.

    Covers hidden sheets, rows and columns, comments, document properties, external
    links, pivot caches, invisible text, embedded objects, macros, custom XML and the
    change log. Each finding says whether sanitize_copy can remove it.
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: sharing.inspect_for_sharing(path))


@_tool("Sanitize copy", destructive=True)
def sanitize_copy(
    output_path: Annotated[str, Field(description="Path for the cleaned copy; must differ from the source.")],
    remove: Annotated[list[P.Sanitizable] | None, Field(description="What to strip. Default: comments, document_properties, change_log. Hidden sheets are removed only when nothing refers to them; external links become their cached values.")] = None,
    overwrite: P.Overwrite = False,
    path: P.Path = None,
) -> dict[str, Any]:
    """Write a cleaned copy of the workbook for sharing (run inspect_for_sharing first). The source is not modified."""
    path = _run(lambda: resolve_path(path))
    output = _run(lambda: resolve_literal_path(output_path))
    copies.ensure_distinct(output, path)
    with file_locks(path, output):
        return _run(lambda: sharing.sanitize_copy(path, output, remove, overwrite))


@_tool("Get change log", read_only=True)
def get_change_log(
    sheet: Annotated[str | None, Field(description="Only entries for this sheet.")] = None,
    cell: Annotated[str | None, Field(description="Only entries for this cell, including ranges that contain it.")] = None,
    limit: Annotated[int, Field(description="Maximum entries (1-1000).")] = 50,
    path: P.Path = None,
) -> dict[str, Any]:
    """Read the change log stored inside the workbook, newest first: time, tool, cell, before/after, reason.

    Entries exist only for writes made while XLSX_MCP_CHANGE_LOG=true.
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: ledger.read(path, sheet, cell, limit))


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
