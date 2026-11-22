# Changelog

All notable changes to this project are documented in this file.

## [0.4.0rc2] — 2026-11-22 (pre-release)

### Changed
- Tool surface for agents: every tool has a `title` and MCP annotations (read-only, destructive, idempotent, closed world). Parameter descriptions moved from the docstrings into the input schema, fixed choices (`format`, `mode`, `on_violation`, `on_unverifiable`, `remove`, `chart_type`) are enums, and descriptions say which neighboring tool to use instead. The server instructions carry a tool-choice guide and the shared write-response contract. Tool names, parameter names and defaults are unchanged.

## [0.4.0rc1] — 2026-11-22 (pre-release)

### Added
- **Value injection recalculation** (`XLSX_MCP_RECALC_MODE=inject`, default): LibreOffice computes formulas on a temporary copy and only the computed values are written into the workbook's formula cells. All other parts (pivot tables, charts, data validation, defined names, `.xlsm` macros) are copied unchanged. `recalculate_workbook` takes a `mode` argument; `roundtrip` keeps the v0.3 behavior.
- **Rule-guarded writes**: `set_rules` / `check_rules` with six rule types (`cells_equal`, `cell_compare`, `column_check`, `column_unique`, `not_empty`, `no_new_errors`) stored in `<file>.rules.json`. A value write that makes a passing rule fail, or a failing rule worse, is rolled back (`rolled_back: true`); `on_violation="warn"` only reports.
- **Ripple report**: value-write responses list other formula results changed by the write (`ripple`), controlled by `XLSX_MCP_RIPPLE` and `XLSX_MCP_RIPPLE_MAX_CELLS`.
- **`export_review_copy`**: copy with changed cells highlighted, Before/After comments and a `_Changes` sheet, compared against the newest `.bak` or a given baseline.
- **`inspect_for_sharing` / `sanitize_copy`**: report hidden sheets, hidden rows/columns with data, comments, document properties, external links, pivot cache data, invisible text, embedded objects, macros, custom XML and the change log; write a cleaned copy.
- **In-workbook change log** (opt-in, `XLSX_MCP_CHANGE_LOG=true`): writes are recorded on a very hidden `_xlsx_mcp_log` sheet; all write tools accept `reason`; `get_change_log` reads entries. The sheet name is reserved for all other tools.
- `XLSX_MCP_MAX_PART_BYTES` caps the XML parts the server parses itself.
- `soffice` pytest marker and a CI job running the integration tests against a real LibreOffice.

### Changed
- `lxml` is now a dependency (needed to rewrite sheet XML without dropping namespace declarations). openpyxl also uses it when present.
- Write responses may include `recalc_mode`, `rules`, `rolled_back` and `ripple`; existing keys are unchanged.
- In inject mode `errors_found` lists only formula cells whose result is an error. Text typed as `#N/A` is no longer reported.
- Sheet-removal checks count only visible sheets, so the change-log sheet never keeps the last user sheet alive.
- Formatting-only tools (`set_cell_style`, `autofit_columns`, `create_table`, `create_chart`, `create_sheet`, `copy_sheet`, `unmerge_cells`) now keep the workbook's cached formula results instead of leaving them empty after the save.
- `export_review_copy` and `sanitize_copy` resolve `output_path`/`baseline_path` literally (no preloaded-alias lookup) and check them against `XLSX_MCP_ALLOWED_DIRS`.
- The LibreOffice temporary folder is removed with `ignore_cleanup_errors`, so a lingering Windows handle no longer turns a finished recalculation into an error.

### Fixed
- `recalculate_workbook` now reports domain errors as clean tool errors (it was not wrapped like the other tools).
- Formula cached values are no longer left empty after a write when LibreOffice is available in inject mode, so calamine/pandas readers see current results.

## [0.3.2] — 2026-10-17

### Fixed
- **Windows In-Place Recalculation Collision**: In `recalculate()`, route LibreOffice `--outdir` into a dedicated `out/` subdirectory within the temporary working directory instead of the input directory. Prevents Windows file sharing / in-place overwrite errors (`0x4c0c`, `Code:12`) when evaluating formulas (#10).

## [0.3.1] — 2026-10-04

### Added
- **Explicit Domain Error for Missing Files**: Added `WorkbookNotFoundError(XlsxMcpError)` across read and write operations. Calling read tools on non-existent paths now yields a clean, immediate error without unnecessary parser fallbacks, and write operations explicitly advise callers to use `create_workbook` first.
- **Corrupted / Invalid Workbook Defense**: Added `InvalidWorkbookError(XlsxMcpError)` that cleanly intercepts unhandled `zipfile.BadZipFile`, `InvalidFileException`, and Calamine read errors when accessing 0-byte or corrupted files.
- **File Handle Leak Prevention**: Wrapped open workbook access in `writer.py` with `_open_wb` context manager, guaranteeing that openpyxl file handles are cleanly closed even when domain exceptions (e.g. `SheetNotFoundError`, invalid coordinates) occur prior to finalization.
- **PermissionError Preservation**: Ensured `PermissionError` (e.g. Windows file sharing locks winerror 32/33) is re-raised and never masked as `InvalidWorkbookError` in `reader.py` and `writer.py`, preserving clear lock diagnostics in `server._run()`.
- **Server Exception Translation**: Added `FileNotFoundError` safety net translation in `server._run()`, ensuring all file-related exceptions format cleanly as standard MCP tool errors instead of unhandled internal server failures.
- **Windows Case-Insensitive Preloaded Aliases**: Preloaded files in `XLSX_MCP_FILES` now support case-insensitive alias lookups on Windows (`os.path.normcase`).

### Documentation
- Updated `SECURITY.md` with complete enterprise protection details (path confinement, formula injection guard, query AST validator, safe atomic replacement, auto-backup).
- Updated `CONTRIBUTING.md` with multi-OS CI matrix details, 29 tools architecture guidelines, and pre-PR checklist.

## [0.3.0] — 2026-10-04

### Added
- **Multi-OS CI Matrix**: GitHub Actions workflow testing on `ubuntu-latest`, `macos-latest`, and `windows-latest` across Python 3.10, 3.11, and 3.12.
- **Cross-Platform Path Confinement**: Path confinement now supports both semicolon (`;`) and comma (`,`) delimiters in `XLSX_MCP_ALLOWED_DIRS` and `XLSX_MCP_FILES`.
- **Windows Path Resilience**: Enforces case-insensitive path confinement on Windows via `os.path.normcase` and safely handles comparisons across separate drive letters.
- **File Sharing Lock Resilience**: `safe_replace` retry loop with exponential backoff (up to 4 attempts) to mitigate transient Windows file locks (`ERROR_SHARING_VIOLATION` / `ERROR_LOCK_VIOLATION` winerror 32 & 33) caused by antivirus scanners or file indexers.
- **Typed Concurrency Exception**: Added `FileInUseError` domain exception with user-friendly MCP error messages when a file remains permanently locked by another application (e.g. desktop Excel).
- **Headless LibreOffice Windows Detection**: Added automatic discovery across standard Windows paths (`Program Files`, `Program Files (x86)`, `LocalAppData\Programs`) and `winget` installation guidance.
- **Isolated Recalc Profile**: Passes `-env:UserInstallation` URI to LibreOffice subprocess to prevent collisions with active desktop instances.
- **Console Window Suppression**: Hidden process creation (`CREATE_NO_WINDOW`) on Windows for headless recalculation.
- **Comprehensive E2E Suite**: Added `tests/test_e2e_all_tools.py` verifying all 29 MCP tools end-to-end across platforms.

## [0.2.0] — 2026-10-03

### Added
- **Sheet Operations**: Added `rename_sheet` and `copy_sheet` tools with Excel sheet name validation.
- **Formatting & Layout**: Added `autofit_columns` with customizable padding/bounds and `clear_range` for selective value/style clearing.
- **Flexible Reading Formats**: Added `offset_row` and `format` (`"array"`, `"records"`, `"markdown"`) parameters to `read_sheet`.
- **Data Profiling & Querying**: Added `profile_sheet` for column summaries and `query_sheet` with AST sandboxing (`SafeQueryValidator`) for safe pandas expressions.
- **Native Excel Objects**: Added `create_table` (ListObject with TableStyle) and `create_chart` (Bar, Line, Pie, Scatter).
- **Path Confinement**: Added `XLSX_MCP_ALLOWED_DIRS` environment variable to restrict filesystem traversal for untrusted agents.
- **Formula Injection Guard**: Validation against unsafe formula functions (`WEBSERVICE`, `HYPERLINK`, `INDIRECT`, etc.) and DDE attacks unless explicitly allowed via `allow_external_formulas=True`.
- **Auto-Backup & Rollback**: Added `backup=True` parameter, `XLSX_MCP_AUTO_BACKUP` environment variable, and `restore_backup` tool.

## [0.1.1] — 2026-10-02

### Changed
- Renamed project from `xlsx_mcp` / `xlsx-mcp` to module `xlsx_tools_mcp` and distribution `xlsx-tools-mcp` for public GitHub and PyPI publication.
- Added MIT license, classifiers, console scripts entrypoint, and automated CI publishing workflow.

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