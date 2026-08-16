# Changelog

All notable changes to this project are documented in this file.

## Unreleased

- Renamed the project from `xlsx_mcp` / `xlsx-mcp` to module `xlsx_tools_mcp` and
  dist name `xlsx-tools-mcp` for public GitHub/PyPI publication.
- Added publication packaging: MIT license, project metadata/classifiers, console
  entry point, and GitHub Actions workflow for the new name.

## 0.1.0 — initial release

- Domain primitives (`src/xlsx_tools_mcp/errors.py`, `settings.py`, `locking.py`,
  `recalc.py`):
  - `XLSX_MCP_FILES` to preload workbooks at startup for path-free / aliased access.
  - Per-file `<path>.lock` via filelock, with `XLSX_MCP_LOCK_TIMEOUT`.
  - LibreOffice headless recalculation pass with error scanning (`errors_found`),
    capped by `XLSX_MCP_RECALC_TIMEOUT`, tolerating a missing LibreOffice.
  - Typed domain errors for missing sheets, lock timeouts, and missing/ambiguous paths.
- Workbook io layer (`src/xlsx_tools_mcp/io/`):
  - `reader.py` — list sheets, workbook info, `read_sheet` (calamine primary,
    openpyxl fallback), `get_cell`, `search_workbook`.
  - `transform.py` — pandas-based `aggregate_sheet` on top of the read path.
  - `writer.py` — create workbook, write cells/formulas, append rows, add/delete
    sheets, insert/delete rows & columns, merge/unmerge cells, set cell styles —
    with atomic saves and optional recalculation.
- MCP server (`src/xlsx_tools_mcp/server.py`) exposing all of the above as 20 tools
  over stdio.
- Test suite covering reader, writer, transform, recalc, locking, and settings.