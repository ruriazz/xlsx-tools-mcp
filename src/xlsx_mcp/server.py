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

from .errors import XlsxMcpError
from .io import reader, transform, writer
from .locking import file_lock
from .recalc import recalculate
from .settings import CONFIGURED_FILES, resolve_path

T = TypeVar("T")

_files_note = (
    f" Files preloaded at startup (call with no `path`, or with these names, to use them "
    f"directly — no need to search the filesystem): {', '.join(sorted(CONFIGURED_FILES))}."
    if CONFIGURED_FILES
    else ""
)

mcp = MCPServer(
    "xlsx-mcp",
    instructions=(
        "Tools for reading and writing Excel (.xlsx) workbooks with high accuracy. "
        "Writes are saved with openpyxl (preserving the file's existing structure, "
        "styles and formulas) and then recalculated with a LibreOffice headless pass "
        "so formula results are never stale. Every write response includes "
        "`errors_found`: any Excel error values (#REF!, #DIV/0!, ...) discovered in "
        "the recalculated file — check this before treating a write as fully successful."
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
    sheet: str, cell_range: str | None = None, max_rows: int | None = None, path: str | None = None
) -> dict[str, Any]:
    """Read cell values from a sheet as a 2D array, addressed absolutely from A1.

    Args:
        path: Path to the .xlsx file, or the name of a preloaded file. Omit if only
            one file is configured.
        sheet: Sheet name.
        cell_range: Optional A1-style range (e.g. "B2:F20"). Omit to read the full used area.
        max_rows: Optional cap on the number of rows returned, to bound response size for large sheets.
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: reader.read_sheet(path, sheet, cell_range, max_rows))


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
    with file_lock(path):
        return _run(lambda: writer.create_workbook(path, sheets, overwrite))


@mcp.tool()
def write_cells(
    sheet: str,
    cells: list[dict[str, Any]],
    create_sheet_if_missing: bool = False,
    recalculate: bool = True,
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

    Returns:
        Write result including `errors_found` — any Excel error values produced by recalculation.
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.write_cells(path, sheet, cells, create_sheet_if_missing, recalculate))


@mcp.tool()
def append_rows(
    sheet: str,
    rows: list[list[Any]],
    create_sheet_if_missing: bool = False,
    recalculate: bool = True,
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
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.append_rows(path, sheet, rows, create_sheet_if_missing, recalculate))


@mcp.tool()
def create_sheet(sheet: str, index: int | None = None, path: str | None = None) -> dict[str, Any]:
    """Add a new empty sheet to an existing workbook.

    Args:
        path: Path to the .xlsx file, or the name of a preloaded file. Omit if only
            one file is configured.
        sheet: Name for the new sheet.
        index: Zero-based position to insert at; omit to append at the end.
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.create_sheet(path, sheet, index))


@mcp.tool()
def delete_sheet(sheet: str, path: str | None = None) -> dict[str, Any]:
    """Delete a sheet from a workbook. Fails if it is the only sheet left.

    Args:
        path: Path to the .xlsx file, or the name of a preloaded file. Omit if only
            one file is configured.
        sheet: Sheet name to delete.
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.delete_sheet(path, sheet))


@mcp.tool()
def insert_rows(
    sheet: str, start_row: int, count: int = 1, recalculate: bool = True, path: str | None = None
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
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.insert_rows(path, sheet, start_row, count, recalculate))


@mcp.tool()
def delete_rows(
    sheet: str, start_row: int, count: int = 1, recalculate: bool = True, path: str | None = None
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
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.delete_rows(path, sheet, start_row, count, recalculate))


@mcp.tool()
def insert_columns(
    sheet: str, start_column: int, count: int = 1, recalculate: bool = True, path: str | None = None
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
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.insert_columns(path, sheet, start_column, count, recalculate))


@mcp.tool()
def delete_columns(
    sheet: str, start_column: int, count: int = 1, recalculate: bool = True, path: str | None = None
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
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.delete_columns(path, sheet, start_column, count, recalculate))


@mcp.tool()
def merge_cells(sheet: str, cell_range: str, path: str | None = None) -> dict[str, Any]:
    """Merge a rectangular range of cells into one.

    Args:
        path: Path to the .xlsx file, or the name of a preloaded file. Omit if only
            one file is configured.
        sheet: Sheet name.
        cell_range: A1-style range, e.g. "A1:C1".
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.merge_cells(path, sheet, cell_range))


@mcp.tool()
def unmerge_cells(sheet: str, cell_range: str, path: str | None = None) -> dict[str, Any]:
    """Undo a merge on a range of cells.

    Args:
        path: Path to the .xlsx file, or the name of a preloaded file. Omit if only
            one file is configured.
        sheet: Sheet name.
        cell_range: A1-style range that was previously merged.
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.unmerge_cells(path, sheet, cell_range))


@mcp.tool()
def set_cell_style(sheet: str, cell_range: str, style: dict[str, Any], path: str | None = None) -> dict[str, Any]:
    """Apply formatting to a range of cells.

    Args:
        path: Path to the .xlsx file, or the name of a preloaded file. Omit if only
            one file is configured.
        sheet: Sheet name.
        cell_range: A1-style range, e.g. "A1:D1".
        style: Keys among: bold, italic, font_size, font_color (hex RGB, e.g. "FF0000"),
            bg_color (hex RGB), horizontal, vertical (alignment), border ("thin", "medium",
            "thick", ...), number_format (e.g. "#,##0.00").
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        return _run(lambda: writer.set_cell_style(path, sheet, cell_range, style))


@mcp.tool()
def recalculate_workbook(path: str | None = None) -> dict[str, Any]:
    """Force a LibreOffice headless recalculation pass and report any formula errors found.

    Args:
        path: Path to the .xlsx file, or the name of a preloaded file. Omit if only
            one file is configured.
    """
    path = _run(lambda: resolve_path(path))
    with file_lock(path):
        result = recalculate(path)
        return {
            "saved": True,
            "path": path,
            "recalculated": result.success,
            "errors_found": result.errors_found,
            "message": result.message,
        }


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
