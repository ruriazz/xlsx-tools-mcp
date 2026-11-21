"""MCP server exposing accurate, structure-preserving read/write tools for Excel (.xlsx) files.

Read path: python-calamine (fast, accurate types) with an openpyxl fallback for
formulas/styles/comments. Write path: openpyxl (preserves everything it doesn't
touch — styles, merges, other formulas) followed by a LibreOffice headless
recalculation pass, since openpyxl never evaluates formulas itself.

Note: the LibreOffice round-trip is a tradeoff, not a guarantee of bit-perfect
preservation. It recomputes formulas but re-exports the whole workbook, which is
NOT lossless for every feature openpyxl preserves (pivot tables, charts, data
validation, some formats/defined names). Callers can pass ``recalculate=False``
on the value/formula-writing tools to write with openpyxl only and skip it.
"""

from __future__ import annotations

from typing import Any, Callable, TypeVar

from mcp.server import MCPServer

from . import ledger
from . import rules as rules_mod
from .errors import XlsxMcpError
from .io import copies, reader, review, sharing, transform, writer
from .io.writer import restore_backup as io_restore_backup
from .locking import file_lock, file_locks
from .recalc import recalculate
from .settings import CONFIGURED_FILES, resolve_literal_path, resolve_path

T = TypeVar("T")

_files_note = (
    f" Files preloaded at startup (call with no `path`, or with these names, to use them "
    f"directly — no need to search the filesystem): {', '.join(sorted(CONFIGURED_FILES))}."
    if CONFIGURED_FILES
    else ""
)

mcp = MCPServer(
    "xlsx-tools-mcp",
    instructions=(
        "Tools for reading and writing Excel (.xlsx) workbooks with high accuracy. "
        "Writes are saved with openpyxl (preserving the file's existing structure, "
        "styles and formulas) and then recalculated with a LibreOffice headless pass "
        "so formula results are never stale. Every write response includes "
        "`errors_found`: any Excel error values (#REF!, #DIV/0!, ...) discovered in "
        "the recalculated file — check this before treating a write as fully successful. "
        "When the workbook has rules (see set_rules), write responses also carry `rules`; "
        "`rolled_back: true` means the write broke a rule and the file was restored. "
        "`ripple` lists other formula results that changed because of the write."
        + _files_note
    ),
)


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



@mcp.tool()
def restore_backup(backup_path: str, target_path: str) -> dict[str, Any]:
    """Restore a backup file over a target workbook atomically.

    Args:
        backup_path: Path to the .bak file to restore from.
        target_path: Path to the target .xlsx/.xlsm file to overwrite.
    """
    backup_path = _run(lambda: resolve_path(backup_path))
    target_path = _run(lambda: resolve_path(target_path))
    with file_lock(target_path):
        return _run(lambda: io_restore_backup(backup_path, target_path))

# ── Inspect / read ────────────────────────────────────────────────────────────


@mcp.tool()
def list_configured_files() -> dict[str, str]:
    """List files preloaded at server startup (via the XLSX_MCP_FILES env var).

    Call this first if unsure which files are available — every other tool's `path`
    argument accepts these names directly (or can be omitted if only one is configured),
    no filesystem search needed.
    """
    return dict(CONFIGURED_FILES)


@mcp.tool()
def list_sheets(path: str | None = None) -> list[dict[str, Any]]:
    """List every sheet in a workbook with its approximate row/column counts.

    Args:
        path: Path to the .xlsx file, or the name of a preloaded file. Omit if only
            one file is configured.
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: reader.list_sheets(path))


@mcp.tool()
def get_workbook_info(path: str | None = None) -> dict[str, Any]:
    """Get workbook-level metadata: sheets, exact dimensions, active sheet, defined names.

    Args:
        path: Path to the .xlsx file, or the name of a preloaded file. Omit if only
            one file is configured.
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: reader.workbook_info(path))


@mcp.tool()
def read_sheet(
    sheet: str,
    cell_range: str | None = None,
    max_rows: int | None = None,
    offset_row: int = 0,
    format: str = "array",
    path: str | None = None,
) -> dict[str, Any]:
    """Read cell values from a sheet.

    Args:
        path: Path to the .xlsx file, or the name of a preloaded file. Omit if only
            one file is configured.
        sheet: Sheet name.
        cell_range: Optional A1-style range (e.g. "B2:F20"). Omit to read the full used area.
        max_rows: Optional cap on the number of rows returned, to bound response size for large sheets.
        offset_row: Skip a number of data rows (0-indexed from first data row).
        format: "array" (default 2D list), "records" (list of dicts, first row as keys), or "markdown" (table string).
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: reader.read_sheet(path, sheet, cell_range, max_rows, offset_row, format))


@mcp.tool()
def get_cell(sheet: str, cell: str, path: str | None = None) -> dict[str, Any]:
    """Get full detail for a single cell: value, formula, number format, font, fill, merge, comment.

    Args:
        path: Path to the .xlsx file, or the name of a preloaded file. Omit if only
            one file is configured.
        sheet: Sheet name.
        cell: Cell reference, e.g. "C5".
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: reader.get_cell(path, sheet, cell))


@mcp.tool()
def search_workbook(
    query: str, sheet: str | None = None, match_case: bool = False, limit: int | None = None, path: str | None = None
) -> list[dict[str, Any]]:
    """Search cell values across one or all sheets for a substring match.

    Args:
        path: Path to the .xlsx file, or the name of a preloaded file. Omit if only
            one file is configured.
        query: Substring to search for.
        sheet: Restrict the search to one sheet; omit to search every sheet.
        match_case: Case-sensitive match when True.
        limit: Optional cap on the number of matches returned, to bound response size for large sheets.
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: reader.search_workbook(path, query, sheet, match_case, limit))


@mcp.tool()
def aggregate_sheet(
    sheet: str,
    group_by: list[str],
    agg: dict[str, str],
    cell_range: str | None = None,
    has_header: bool = True,
    path: str | None = None,
) -> dict[str, Any]:
    """Group and aggregate sheet data with pandas (e.g. sum sales by region).

    Args:
        path: Path to the .xlsx file, or the name of a preloaded file. Omit if only
            one file is configured.
        sheet: Sheet name.
        group_by: Column names to group by (taken from the header row).
        agg: Mapping of column name to aggregation function, e.g. {"amount": "sum"}.
        cell_range: Optional A1-style range to read before aggregating.
        has_header: Whether the first row of the range holds column names.
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: transform.aggregate_sheet(path, sheet, group_by, agg, cell_range, has_header))


# ── Write ──────────────────────────────────────────────────────────────────────

@mcp.tool()
def create_workbook(path: str, sheets: list[str] | None = None, overwrite: bool = False) -> dict[str, Any]:
    """Create a new .xlsx workbook.

    Args:
        path: Path for the new file. New files aren't preloaded, so this must be a real path.
        sheets: Sheet names to create (defaults to a single "Sheet1").
        overwrite: Replace an existing file at path when True.
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.create_workbook(path, sheets, overwrite))


@mcp.tool()
def write_cells(
    sheet: str,
    cells: list[dict[str, Any]],
    create_sheet_if_missing: bool = False,
    recalculate: bool = True,
    allow_external_formulas: bool = False,
    backup: bool | None = None,
    reason: str | None = None,
    path: str | None = None,
) -> dict[str, Any]:
    """Write values and/or formulas into specific cells, then recalculate the workbook.

    Set `recalculate=False` to skip the LibreOffice recompute (e.g. for structurally
    complex workbooks you don't want round-tripped).

    Args:
        path: Path to the .xlsx file, or the name of a preloaded file. Omit if only
            one file is configured.
        sheet: Sheet name.
        cells: List of {"cell": "A1", "value": ...} or {"cell": "B1", "formula": "=A1*2"}.
        create_sheet_if_missing: Create the sheet first if it doesn't exist yet.
        recalculate: Run the LibreOffice recompute pass after saving.
        reason: Optional note recorded in the in-workbook change log (XLSX_MCP_CHANGE_LOG=true).

    Returns:
        Write result including `errors_found` — any Excel error values produced by recalculation.
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.write_cells(path, sheet, cells, create_sheet_if_missing, recalculate, allow_external_formulas, backup if backup is not None else writer.AUTO_BACKUP, reason=reason))


@mcp.tool()
def append_rows(
    sheet: str,
    rows: list[list[Any]],
    create_sheet_if_missing: bool = False,
    recalculate: bool = True,
    allow_external_formulas: bool = False,
    backup: bool | None = None,
    reason: str | None = None,
    path: str | None = None,
) -> dict[str, Any]:
    """Append rows after the last used row of a sheet, then recalculate the workbook.

    Set `recalculate=False` to skip the LibreOffice recompute (e.g. for structurally
    complex workbooks you don't want round-tripped).

    Args:
        path: Path to the .xlsx file, or the name of a preloaded file. Omit if only
            one file is configured.
        sheet: Sheet name.
        rows: List of rows, each a list of cell values in column order.
        create_sheet_if_missing: Create the sheet first if it doesn't exist yet.
        recalculate: Run the LibreOffice recompute pass after saving.
        reason: Optional note recorded in the in-workbook change log (XLSX_MCP_CHANGE_LOG=true).
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.append_rows(path, sheet, rows, create_sheet_if_missing, recalculate, allow_external_formulas, backup if backup is not None else writer.AUTO_BACKUP, reason=reason))


@mcp.tool()
def create_sheet(sheet: str, index: int | None = None, backup: bool | None = None, reason: str | None = None, path: str | None = None) -> dict[str, Any]:
    """Add a new empty sheet to an existing workbook.

    Args:
        path: Path to the .xlsx file, or the name of a preloaded file. Omit if only
            one file is configured.
        sheet: Name for the new sheet.
        index: Zero-based position to insert at; omit to append at the end.
        reason: Optional note recorded in the in-workbook change log (XLSX_MCP_CHANGE_LOG=true).
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.create_sheet(path, sheet, index, backup=backup if backup is not None else writer.AUTO_BACKUP, reason=reason))


@mcp.tool()
def delete_sheet(sheet: str, backup: bool | None = None, reason: str | None = None, path: str | None = None) -> dict[str, Any]:
    """Delete a sheet from a workbook. Fails if it is the only sheet left.

    Args:
        path: Path to the .xlsx file, or the name of a preloaded file. Omit if only
            one file is configured.
        sheet: Sheet name to delete.
        reason: Optional note recorded in the in-workbook change log (XLSX_MCP_CHANGE_LOG=true).
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.delete_sheet(path, sheet, backup=backup if backup is not None else writer.AUTO_BACKUP, reason=reason))


@mcp.tool()
def insert_rows(
    sheet: str, start_row: int, count: int = 1, recalculate: bool = True, backup: bool | None = None, reason: str | None = None, path: str | None = None
) -> dict[str, Any]:
    """Insert blank rows, shifting existing rows down.

    Set `recalculate=False` to skip the LibreOffice recompute (e.g. for structurally
    complex workbooks you don't want round-tripped).

    Args:
        path: Path to the .xlsx file, or the name of a preloaded file. Omit if only
            one file is configured.
        sheet: Sheet name.
        start_row: 1-based row index to insert before.
        count: Number of rows to insert.
        recalculate: Run the LibreOffice recompute pass after saving.
        reason: Optional note recorded in the in-workbook change log (XLSX_MCP_CHANGE_LOG=true).
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.insert_rows(path, sheet, start_row, count, recalculate, backup=backup if backup is not None else writer.AUTO_BACKUP, reason=reason))


@mcp.tool()
def delete_rows(
    sheet: str, start_row: int, count: int = 1, recalculate: bool = True, backup: bool | None = None, reason: str | None = None, path: str | None = None
) -> dict[str, Any]:
    """Delete rows, shifting the rows below them upward.

    Set `recalculate=False` to skip the LibreOffice recompute (e.g. for structurally
    complex workbooks you don't want round-tripped).

    Args:
        path: Path to the .xlsx file, or the name of a preloaded file. Omit if only
            one file is configured.
        sheet: Sheet name.
        start_row: 1-based row index to start deleting from.
        count: Number of rows to delete.
        recalculate: Run the LibreOffice recompute pass after saving.
        reason: Optional note recorded in the in-workbook change log (XLSX_MCP_CHANGE_LOG=true).
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.delete_rows(path, sheet, start_row, count, recalculate, backup=backup if backup is not None else writer.AUTO_BACKUP, reason=reason))


@mcp.tool()
def insert_columns(
    sheet: str, start_column: int, count: int = 1, recalculate: bool = True, backup: bool | None = None, reason: str | None = None, path: str | None = None
) -> dict[str, Any]:
    """Insert blank columns, shifting existing columns right.

    Set `recalculate=False` to skip the LibreOffice recompute (e.g. for structurally
    complex workbooks you don't want round-tripped).

    Args:
        path: Path to the .xlsx file, or the name of a preloaded file. Omit if only
            one file is configured.
        sheet: Sheet name.
        start_column: 1-based column index to insert before.
        count: Number of columns to insert.
        recalculate: Run the LibreOffice recompute pass after saving.
        reason: Optional note recorded in the in-workbook change log (XLSX_MCP_CHANGE_LOG=true).
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.insert_columns(path, sheet, start_column, count, recalculate, backup=backup if backup is not None else writer.AUTO_BACKUP, reason=reason))


@mcp.tool()
def delete_columns(
    sheet: str, start_column: int, count: int = 1, recalculate: bool = True, backup: bool | None = None, reason: str | None = None, path: str | None = None
) -> dict[str, Any]:
    """Delete columns, shifting the columns to their right leftward.

    Set `recalculate=False` to skip the LibreOffice recompute (e.g. for structurally
    complex workbooks you don't want round-tripped).

    Args:
        path: Path to the .xlsx file, or the name of a preloaded file. Omit if only
            one file is configured.
        sheet: Sheet name.
        start_column: 1-based column index to start deleting from.
        count: Number of columns to delete.
        recalculate: Run the LibreOffice recompute pass after saving.
        reason: Optional note recorded in the in-workbook change log (XLSX_MCP_CHANGE_LOG=true).
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.delete_columns(path, sheet, start_column, count, recalculate, backup=backup if backup is not None else writer.AUTO_BACKUP, reason=reason))


@mcp.tool()
def merge_cells(sheet: str, cell_range: str, backup: bool | None = None, reason: str | None = None, path: str | None = None) -> dict[str, Any]:
    """Merge a rectangular range of cells into one.

    Args:
        path: Path to the .xlsx file, or the name of a preloaded file. Omit if only
            one file is configured.
        sheet: Sheet name.
        cell_range: A1-style range, e.g. "A1:C1".
        reason: Optional note recorded in the in-workbook change log (XLSX_MCP_CHANGE_LOG=true).
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.merge_cells(path, sheet, cell_range, backup=backup if backup is not None else writer.AUTO_BACKUP, reason=reason))


@mcp.tool()
def unmerge_cells(sheet: str, cell_range: str, backup: bool | None = None, reason: str | None = None, path: str | None = None) -> dict[str, Any]:
    """Undo a merge on a range of cells.

    Args:
        path: Path to the .xlsx file, or the name of a preloaded file. Omit if only
            one file is configured.
        sheet: Sheet name.
        cell_range: A1-style range that was previously merged.
        reason: Optional note recorded in the in-workbook change log (XLSX_MCP_CHANGE_LOG=true).
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.unmerge_cells(path, sheet, cell_range, backup=backup if backup is not None else writer.AUTO_BACKUP, reason=reason))


@mcp.tool()
def set_cell_style(sheet: str, cell_range: str, style: dict[str, Any], backup: bool | None = None, reason: str | None = None, path: str | None = None) -> dict[str, Any]:
    """Apply formatting to a range of cells.

    Args:
        path: Path to the .xlsx file, or the name of a preloaded file. Omit if only
            one file is configured.
        sheet: Sheet name.
        cell_range: A1-style range, e.g. "A1:D1".
        style: Keys among: bold, italic, font_size, font_color (hex RGB, e.g. "FF0000"),
            bg_color (hex RGB), horizontal, vertical (alignment), border ("thin", "medium",
            "thick", ...), number_format (e.g. "#,##0.00").
        reason: Optional note recorded in the in-workbook change log (XLSX_MCP_CHANGE_LOG=true).
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.set_cell_style(path, sheet, cell_range, style, backup=backup if backup is not None else writer.AUTO_BACKUP, reason=reason))


@mcp.tool()
def rename_sheet(old_name: str, new_name: str, backup: bool | None = None, reason: str | None = None, path: str | None = None) -> dict[str, Any]:
    """Rename an existing sheet.

    Warning: This tool does NOT automatically update formula references in other sheets.
    Formulas depending on the old sheet name will break.

    Args:
        path: Path to the .xlsx file, or the name of a preloaded file. Omit if only
            one file is configured.
        old_name: Current name of the sheet.
        new_name: New name for the sheet (1-31 chars, valid Excel sheet name).
        reason: Optional note recorded in the in-workbook change log (XLSX_MCP_CHANGE_LOG=true).
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.rename_sheet(path, old_name, new_name, backup=backup if backup is not None else writer.AUTO_BACKUP, reason=reason))


@mcp.tool()
def copy_sheet(source_sheet: str, target_sheet: str, backup: bool | None = None, reason: str | None = None, path: str | None = None) -> dict[str, Any]:
    """Duplicate a sheet including its contents, formulas, and styles.

    Args:
        path: Path to the .xlsx file, or the name of a preloaded file. Omit if only
            one file is configured.
        source_sheet: Name of the sheet to duplicate.
        target_sheet: Name of the new sheet (1-31 chars, valid Excel sheet name).
        reason: Optional note recorded in the in-workbook change log (XLSX_MCP_CHANGE_LOG=true).
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.copy_sheet(path, source_sheet, target_sheet, backup=backup if backup is not None else writer.AUTO_BACKUP, reason=reason))


@mcp.tool()
def autofit_columns(
    sheet: str,
    min_width: int = 10,
    max_width: int = 50,
    padding: int = 3,
    backup: bool | None = None,
    reason: str | None = None,
    path: str | None = None,
) -> dict[str, Any]:
    """Automatically adjust column widths based on maximum text length.

    Args:
        path: Path to the .xlsx file, or the name of a preloaded file. Omit if only
            one file is configured.
        sheet: Sheet name.
        min_width: Minimum allowed column width.
        max_width: Maximum allowed column width.
        padding: Extra padding characters to add to max length.
        reason: Optional note recorded in the in-workbook change log (XLSX_MCP_CHANGE_LOG=true).
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.autofit_columns(path, sheet, min_width, max_width, padding, backup=backup if backup is not None else writer.AUTO_BACKUP, reason=reason))


@mcp.tool()
def clear_range(
    sheet: str,
    cell_range: str,
    clear_values: bool = True,
    clear_styles: bool = False,
    backup: bool | None = None,
    reason: str | None = None,
    path: str | None = None,
) -> dict[str, Any]:
    """Clear cell values and/or styles within a specified range without deleting rows or columns.

    Args:
        path: Path to the .xlsx file, or the name of a preloaded file. Omit if only
            one file is configured.
        sheet: Sheet name.
        cell_range: A1-style range (e.g. "A1:D10") to clear.
        clear_values: Whether to clear cell values (True by default).
        clear_styles: Whether to clear cell styles/formats (False by default).
        reason: Optional note recorded in the in-workbook change log (XLSX_MCP_CHANGE_LOG=true).
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.clear_range(path, sheet, cell_range, clear_values, clear_styles, backup=backup if backup is not None else writer.AUTO_BACKUP, reason=reason))


@mcp.tool()
def recalculate_workbook(mode: str | None = None, path: str | None = None) -> dict[str, Any]:
    """Force a LibreOffice headless recalculation pass and report any formula errors found.

    Args:
        path: Path to the .xlsx file, or the name of a preloaded file. Omit if only
            one file is configured.
        mode: "inject" writes only the computed values into the file and keeps every
            other part as is; "roundtrip" replaces the file with LibreOffice's export.
            Defaults to XLSX_MCP_RECALC_MODE (inject).
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



@mcp.tool()
def profile_sheet(
    sheet: str, sample_rows: int = 3, path: str | None = None
) -> dict[str, Any]:
    """Profile sheet data: metadata, inferred types, null counts, min/max, sample values.

    Args:
        path: Path to the .xlsx file, or the name of a preloaded file. Omit if only
            one file is configured.
        sheet: Sheet name.
        sample_rows: Number of non-null sample values to return per column.
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: reader.profile_sheet(path, sheet, sample_rows))

@mcp.tool()
def query_sheet(
    sheet: str, filter_query: str, columns: list[str] | None = None, max_rows: int = 100, cell_range: str | None = None, path: str | None = None
) -> dict[str, Any]:
    """Query sheet data using pandas expressions without returning the entire sheet into the LLM context.

    Args:
        path: Path to the .xlsx file, or the name of a preloaded file. Omit if only
            one file is configured.
        sheet: Sheet name.
        filter_query: Pandas query expression (e.g. "status == 'PAID' and amount > 500").
        columns: Optional list of columns to return.
        max_rows: Maximum rows to return (default 100).
        cell_range: Optional range to restrict reading before querying.
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: transform.query_sheet(path, sheet, filter_query, columns, max_rows, cell_range))

@mcp.tool()
def create_table(
    sheet: str, cell_range: str, table_name: str, style_name: str = "TableStyleMedium9", show_filter: bool = True, show_row_stripes: bool = True, backup: bool | None = None, reason: str | None = None, path: str | None = None
) -> dict[str, Any]:
    """Create a formal Excel Table (ListObject) on the specified range.

    Args:
        path: Path to the .xlsx file, or the name of a preloaded file. Omit if only
            one file is configured.
        sheet: Sheet name.
        cell_range: A1-style range (e.g. "A1:D10").
        table_name: Unique table identifier across the workbook.
        style_name: Table style name (e.g. "TableStyleMedium9").
        show_filter: Whether to display auto-filter dropdown arrows.
        show_row_stripes: Whether to apply alternating row shading.
        backup: Whether to create a backup before modifying the workbook.
        reason: Optional note recorded in the in-workbook change log (XLSX_MCP_CHANGE_LOG=true).
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.create_table(path, sheet, cell_range, table_name, style_name, show_filter, show_row_stripes, backup=backup if backup is not None else writer.AUTO_BACKUP, reason=reason))

@mcp.tool()
def create_chart(
    sheet: str, chart_type: str, data_range: str, categories_range: str | None = None, title: str = "Chart", target_cell: str = "E2", backup: bool | None = None, reason: str | None = None, path: str | None = None
) -> dict[str, Any]:
    """Create a native Excel chart (bar, line, pie, scatter).

    Args:
        path: Path to the .xlsx file, or the name of a preloaded file. Omit if only
            one file is configured.
        sheet: Sheet name.
        chart_type: One of "bar", "line", "pie", "scatter".
        data_range: A1-style range holding the chart series values (e.g. "B1:B10").
        categories_range: Optional A1-style range holding series category labels (e.g. "A2:A10").
        title: Title string displayed on the chart.
        target_cell: Cell coordinate where the top-left corner of the chart is placed.
        backup: Whether to create a backup before modifying the workbook.
        reason: Optional note recorded in the in-workbook change log (XLSX_MCP_CHANGE_LOG=true).
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.create_chart(path, sheet, chart_type, data_range, categories_range, title, target_cell, backup=backup if backup is not None else writer.AUTO_BACKUP, reason=reason))


# ── Rules, review, sharing, change log ───────────────────────────────────────


@mcp.tool()
def set_rules(
    rules: list[dict[str, Any]],
    on_violation: str = "rollback",
    on_unverifiable: str = "warn",
    path: str | None = None,
) -> dict[str, Any]:
    """Replace the workbook's rules (an empty list removes them), then evaluate them once.

    After every later value-changing write, rules are checked against the recalculated
    values. A write that makes a passing rule fail (or a failing rule worse) is rolled
    back when on_violation="rollback". Rules that already fail and don't get worse
    never block. Stored next to the workbook as "<file>.rules.json".

    Rule types (each needs a unique "id"):
      {"type": "cells_equal", "left": "Sheet!B40", "right": "Sheet!C40", "tolerance": 0.01}
      {"type": "cell_compare", "cell": "Sheet!D4", "op": "<=", "value": 1000}   (or "cell_ref")
      {"type": "column_check", "sheet": "Data", "expr": "Qty >= 0", "cell_range": "A1:F500"}
      {"type": "column_unique", "sheet": "Data", "column": "Invoice"}
      {"type": "not_empty", "range": "Data!A2:A500"}
      {"type": "no_new_errors"}

    Args:
        path: Path to the .xlsx file, or the name of a preloaded file. Omit if only
            one file is configured.
        rules: List of rule objects as above.
        on_violation: "rollback" (default) or "warn".
        on_unverifiable: "warn" (default) or "rollback" — what to do when a rule cannot
            be checked (e.g. LibreOffice missing, sheet deleted).
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


@mcp.tool()
def check_rules(path: str | None = None) -> dict[str, Any]:
    """Evaluate the workbook's rules against its current values, without writing anything.

    Args:
        path: Path to the .xlsx file, or the name of a preloaded file. Omit if only
            one file is configured.
    """
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


@mcp.tool()
def export_review_copy(
    baseline_path: str | None = None,
    output_path: str | None = None,
    overwrite: bool = False,
    highlight_color: str = "FFF2CC",
    include_recalc_changes: bool = False,
    max_annotated_cells: int = 2000,
    path: str | None = None,
) -> dict[str, Any]:
    """Write a copy of the workbook that highlights every cell changed since a baseline.

    Changed cells get a fill and a "Before / After" comment; a `_Changes` sheet lists
    every change with a link to the cell. The workbook itself is never modified.

    Args:
        path: Path to the .xlsx file, or the name of a preloaded file. Omit if only
            one file is configured.
        baseline_path: Workbook to compare against. Defaults to this file's newest .bak backup.
        output_path: Where to write the copy. Defaults to "<name>.review.xlsx" next to the file.
        overwrite: Replace an existing output file.
        highlight_color: Hex RGB fill for changed cells.
        include_recalc_changes: Also list formulas whose text is unchanged but whose result changed.
        max_annotated_cells: Cap on highlighted cells (all changes are still listed in `_Changes`).
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


@mcp.tool()
def inspect_for_sharing(path: str | None = None) -> dict[str, Any]:
    """Report hidden or sensitive content to review before a workbook is shared. Read-only.

    Checks hidden/very hidden sheets, hidden rows/columns holding data, comments and
    their authors, document properties (author, company), external workbook links,
    defined names pointing elsewhere, pivot cache data, ';;;' or same-color-as-fill
    text, embedded objects, macros, custom XML and the change log sheet.

    Args:
        path: Path to the .xlsx file, or the name of a preloaded file. Omit if only
            one file is configured.
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: sharing.inspect_for_sharing(path))


@mcp.tool()
def sanitize_copy(
    output_path: str,
    remove: list[str] | None = None,
    overwrite: bool = False,
    path: str | None = None,
) -> dict[str, Any]:
    """Write a cleaned copy of the workbook for sharing. The source file is never modified.

    Args:
        path: Path to the .xlsx file, or the name of a preloaded file. Omit if only
            one file is configured.
        output_path: Path for the cleaned copy (must differ from the source).
        remove: Any of "comments", "document_properties", "hidden_sheets",
            "external_links", "change_log", "macros". Defaults to comments,
            document_properties and change_log. Hidden sheets are only removed when no
            formula, name, chart or validation still refers to them. External links are
            replaced by their last cached values.
        overwrite: Replace an existing output file.
    """
    path = _run(lambda: resolve_path(path))
    output = _run(lambda: resolve_literal_path(output_path))
    copies.ensure_distinct(output, path)
    with file_locks(path, output):
        return _run(lambda: sharing.sanitize_copy(path, output, remove, overwrite))


@mcp.tool()
def get_change_log(
    sheet: str | None = None, cell: str | None = None, limit: int = 50, path: str | None = None
) -> dict[str, Any]:
    """Read the change log stored inside the workbook, newest first.

    Entries are written when the server runs with XLSX_MCP_CHANGE_LOG=true. Each has
    ts_utc, op_id, tool, sheet, ref, before, after and reason.

    Args:
        path: Path to the .xlsx file, or the name of a preloaded file. Omit if only
            one file is configured.
        sheet: Only entries for this sheet.
        cell: Only entries for this cell, including range entries that contain it.
        limit: Maximum entries to return (1-1000).
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: ledger.read(path, sheet, cell, limit))


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
