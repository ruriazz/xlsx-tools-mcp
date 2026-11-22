# xlsx-tools-mcp

An MCP server for reading and writing Excel (.xlsx) files with high accuracy, while preserving the file's existing structure, styles, and formulas.

<!-- mcp-name: io.github.ruriazz/xlsx-tools-mcp -->

[![CI](https://github.com/ruriazz/xlsx-tools-mcp/actions/workflows/ci.yml/badge.svg)](https://github.com/ruriazz/xlsx-tools-mcp/actions/workflows/ci.yml) ![PyPI Version](https://img.shields.io/pypi/v/xlsx-tools-mcp) [![Downloads/month](https://img.shields.io/pypi/dm/xlsx-tools-mcp)](https://pypistats.org/packages/xlsx-tools-mcp) [![Listed on mcpservers.org](https://mcpservers.org/badge.svg)](https://mcpservers.org/servers/ruriazz/xlsx-tools-mcp)

---

## Overview

`xlsx-tools-mcp` exposes 35 Model Context Protocol (MCP) tools that give an LLM agent accurate, structure-preserving read **and** write access to Excel `.xlsx` files. It runs as a standard stdio MCP server: you install it and register it with an MCP client (Claude Code, OpenCode, etc.), and the client's agent can list sheets, read cell ranges, search values, aggregate data, write cells/formulas, manage sheets/rows/columns, apply styles, and force formula recalculation. It can also guard writes with business rules, produce a highlighted review copy of what changed, check a workbook for hidden data before it is shared, and keep a change log inside the file.

It is built around the principle that editing an existing workbook should **not** destroy what it doesn't touch.

### Features

- **Structure-preserving writes via openpyxl** — writes load the existing workbook and save it back, preserving styles, merged cells, comments, charts, data validation, conditional formatting and defined names. openpyxl does not keep shapes and other drawing objects it cannot read.
- **Never-stale formula results without the round-trip** — openpyxl writes formula *strings* and leaves every cached result empty. After each value/formula write, LibreOffice computes the formulas on a temporary copy and only the computed values are injected into the workbook (`XLSX_MCP_RECALC_MODE=inject`, the default). Every other part of the file is copied byte-for-byte. The response carries `errors_found` — formula cells whose result is an Excel error (`#REF!`, `#DIV/0!`, `#N/A`, …).
- **Rule-guarded writes** — declare checks once (`set_rules`: cells that must match, limits, column checks, uniqueness, required cells, "no new errors"). After every value-changing write the rules are re-evaluated on the recalculated values; a write that breaks a passing rule is rolled back and the file is left unchanged.
- **Ripple report** — each write response lists other formula results that changed because of the write (`ripple`), so an agent notices unintended side effects.
- **Review copy for humans** — `export_review_copy` writes a copy where every cell changed since the last backup is highlighted with a "Before / After" comment, plus a `_Changes` sheet linking to each change.
- **Pre-share inspection** — `inspect_for_sharing` reports hidden and very hidden sheets, hidden rows/columns with data, comments and their authors, author/company properties, external links, pivot cache data, invisible text, macros and more; `sanitize_copy` writes a cleaned copy.
- **In-file change log (opt-in)** — with `XLSX_MCP_CHANGE_LOG=true`, every write records time, tool, cell, before/after and an optional `reason` on a very hidden sheet; `get_change_log` reads it back.
- **Fast reads via python-calamine** — a Rust-backed parser for accurate, fast type inference, with an automatic openpyxl fallback when you need formulas/styles/comments or when calamine can't parse the file.
- **pandas-based grouping/aggregation & querying** — `aggregate_sheet` and `query_sheet` perform operations without destroying source structure, equipped with AST expression validation to block code injection.
- **Native Excel objects** — create formal Excel Tables (`create_table`) with autofilter and styles, and native charts (`create_chart` for bar, line, pie, scatter).
- **Sheet lifecycle & formatting** — rename sheets (`rename_sheet`), duplicate sheets with styles intact (`copy_sheet`), clear ranges (`clear_range`), and auto-fit column widths (`autofit_columns`).
- **Cross-platform reliability (Windows, macOS, Linux)** — case-insensitive path confinement, comma or semicolon `;` environment delimiters, safe atomic file replacement with transient lock retry (`WinError 32` / antivirus defense), and silent headless recalculation without window flicker.
- **Per-file locking** — concurrent tool calls (or other processes) touching the same workbook are serialized via a sibling `<path>.lock` file (filelock), so writes never interleave and corrupt the file.
- **Path confinement** — `XLSX_MCP_ALLOWED_DIRS` (comma or semicolon-separated) restricts all reads and writes to specific directories. If set, any attempt to access a file outside these boundaries raises `AccessDeniedError`. Case-insensitive path comparisons and multi-drive paths on Windows are supported.
- **Auto-backup & restore** — Write tools automatically create a `.bak` sibling file before modifying the workbook if `backup=True` (which defaults to `XLSX_MCP_AUTO_BACKUP=true`). Use `restore_backup` to atomically revert changes.
- **Formula injection guard** — `write_cells` and `append_rows` reject formulas containing `WEBSERVICE`, `HYPERLINK`, `INDIRECT`, `RTD`, `CALL`, `REGISTER`, or DDE commands (`|`) by default to prevent data exfiltration or malicious links. Pass `allow_external_formulas=True` to bypass this check.
- **XML-bomb protection** — the `defusedxml` package is an automatic dependency; openpyxl detects it and uses its hardened XML parser, so hostile `xlsx` XML can't expand into resource exhaustion.
- **Preload files at startup** — set `XLSX_MCP_FILES` to preload one or more workbooks (comma or semicolon-separated); tools can then be called with `path` omitted or with a short alias instead of a full filesystem path.

---

## Download Stats

[![Downloads/month](https://img.shields.io/pypi/dm/xlsx-tools-mcp)](https://pypistats.org/packages/xlsx-tools-mcp) [![Downloads/week](https://img.shields.io/pypi/dw/xlsx-tools-mcp)](https://pypistats.org/packages/xlsx-tools-mcp)

Live via [pypistats.org](https://pypistats.org/packages/xlsx-tools-mcp), non-mirror downloads. These count download events, not unique users or installs — one user can trigger many downloads (CI, reinstalls, Docker rebuilds, mirrors).

---

## Architecture

```
┌──────────────────────── Supervisor (MCP transport, stdio)
│  src/xlsx_tools_mcp/server.py     35 MCP tools + instructions
│  src/xlsx_tools_mcp/params.py     shared parameter types and descriptions
│  src/xlsx_tools_mcp/settings.py   env vars, preloaded files, path resolution
│  src/xlsx_tools_mcp/locking.py    per-file <path>.lock serialization
│  src/xlsx_tools_mcp/errors.py     domain error types
│  src/xlsx_tools_mcp/recalc.py     LibreOffice headless recalc (inject or roundtrip)
│  src/xlsx_tools_mcp/ooxml.py      hardened zip/XML access to the package (lxml)
│  src/xlsx_tools_mcp/inject.py     read/write cached formula results in the XML
│  src/xlsx_tools_mcp/guard.py      per-write rules check, rollback snapshot, ripple
│  src/xlsx_tools_mcp/rules.py      rule schema, sidecar file, evaluation
│  src/xlsx_tools_mcp/ripple.py     before/after diff of formula results
│  src/xlsx_tools_mcp/ledger.py     opt-in change log sheet
│
├─ Read path
│  src/xlsx_tools_mcp/io/reader.py      calamine primary → openpyxl fallback
│  src/xlsx_tools_mcp/io/transform.py   pandas aggregation on read results
│
├─ Write path
│  src/xlsx_tools_mcp/io/writer.py      openpyxl → LibreOffice values injected → rules/ripple
│
└─ Derived copies (never modify the source)
   src/xlsx_tools_mcp/io/review.py      highlighted review copy vs a baseline
   src/xlsx_tools_mcp/io/sharing.py     inspect_for_sharing / sanitize_copy
```

The **io layer** (`io/`) is deliberately decoupled from the MCP transport (`server.py`). Each MCP tool is a thin wrapper that resolves the target path, takes the per-file lock, and calls one io-layer function. This keeps the core logic independent of MCP, so it can be tested directly (see `tests/`).

### How recalculation works

After a write that touches cell values or formulas, the server copies the file to a temporary folder and runs `soffice --headless --convert-to xlsx` on the copy so every formula gets a computed value. What happens next depends on `XLSX_MCP_RECALC_MODE`:

- **`inject` (default)** — the computed values are read from LibreOffice's export and written into the `<v>` elements of the formula cells in *your* file. Nothing else in the file changes: pivot tables, slicers, charts, data validation, defined names, `.xlsm` macros and every other part are copied byte-for-byte. Excel still recalculates on open (openpyxl sets `fullCalcOnLoad`).
- **`roundtrip`** — the v0.3 behavior: the file is replaced by LibreOffice's export. This re-exports the whole workbook and may not keep pivot tables, charts, data validation, some formats and some defined names.

If injection cannot be completed (for example an unexpected XML layout), the file is left exactly as openpyxl saved it, `recalculated` is `false`, and `message` explains why. Formulas that use functions LibreOffice does not know report `#NAME?` in `errors_found` with a hint; Excel computes them when the file is opened.

You can still pass `recalculate=False` on the value/formula-writing tools (`write_cells`, `append_rows`, `insert_rows`, `delete_rows`, `insert_columns`, `delete_columns`) to skip LibreOffice entirely; formula results are then empty until the file is recalculated.

---

## Requirements

- **Python ≥ 3.10**
- **LibreOffice** — *optional but recommended*. Needed only for formula recalculation. Without it, writes still succeed (saved via openpyxl) but formulas are **not** recomputed and a warning is returned in the `message` field.

Install LibreOffice:

```bash
# macOS
brew install --cask libreoffice

# Debian / Ubuntu
sudo apt-get install -y libreoffice-calc

# Windows
winget install TheDocumentFoundation.LibreOffice
```

The server finds LibreOffice by checking `soffice` / `libreoffice` on `PATH`, the standard macOS install location (`/Applications/LibreOffice.app/Contents/MacOS/soffice`), and standard Windows paths (`%ProgramFiles%\LibreOffice\program\soffice.exe`, etc.). On Windows, headless recalculation runs silently without window flicker (`CREATE_NO_WINDOW`) and with isolated user profiles.

---

## Installation

The server speaks **stdio** transport (standard MCP): after installation it waits for an MCP client to connect and call tools. You don't usually run it yourself; you register it with a client.

### 1. From PyPI via `uvx` (recommended — no clone)

```bash
uvx xlsx-tools-mcp
```

`uvx` fetches and runs the published package without polluting your project. This is the simplest way to power up an MCP client (see configuration snippets below).

### 2. From source

```bash
git clone https://github.com/ruriazz/xlsx-tools-mcp.git
cd xlsx-tools-mcp
uv sync
# run the server (useful for local dev / debugging):
uv run xlsx-tools-mcp
```

### 3. Via `pip`

```bash
pip install xlsx-tools-mcp
```

This installs the console entry point, so you can run the server directly:

```bash
xlsx-tools-mcp
```

---

## Configuration for MCP clients

The simplest registration for every client uses `uvx xlsx-tools-mcp` (no clone, always the published version).

> **Tip for Windows:** In JSON client configuration files (`.mcp.json`, `opencode.json`), use forward slashes (e.g. `"C:/path/to/data.xlsx"`) or escaped backslashes (e.g. `"C:\\\\path\\\\to\\\\data.xlsx"`) to prevent JSON escape sequence errors.

### Claude Code

```bash
claude mcp add xlsx-tools-mcp -- uvx xlsx-tools-mcp
```

Or via `.mcp.json` in your project:

```json
{
  "mcpServers": {
    "xlsx-tools-mcp": { "command": "uvx", "args": ["xlsx-tools-mcp"] }
  }
}
```

### OpenCode

In `opencode.json` (project) or `~/.config/opencode/opencode.json` (global):

```json
{
  "mcp": {
    "xlsx-tools-mcp": { "type": "local", "command": ["uvx", "xlsx-tools-mcp"], "enabled": true }
  }
}
```

### When running from a source clone

If you cloned the repo instead of installing from PyPI, point the client at your local checkout by swapping `uvx xlsx-tools-mcp` for the dynamic `uv run` form (use the **absolute** path to the clone):

**Claude Code `.mcp.json`:**

```json
{
  "mcpServers": {
    "xlsx-tools-mcp": {
      "command": "uv",
      "args": ["--directory", "/absolute/path/to/xlsx-reader", "run", "xlsx-tools-mcp"]
    }
  }
}
```

**OpenCode:**

```json
{
  "mcp": {
    "xlsx-tools-mcp": {
      "type": "local",
      "command": ["uv", "--directory", "/absolute/path/to/xlsx-reader", "run", "xlsx-tools-mcp"],
      "enabled": true
    }
  }
}
```

Replace `/absolute/path/to/xlsx-reader` with the actual location of your clone.

---

## Preloading files (`XLSX_MCP_FILES`)

Set the `XLSX_MCP_FILES` environment variable in the **MCP server config `env`** section (not your interactive shell — the server is launched by the client) to preload workbooks at startup. Format: comma- or semicolon-separated `alias=absolute/path` entries, or bare absolute paths:

```
XLSX_MCP_FILES=name=/abs/path/to/name.xlsx,report=/data/report.xlsx
```

Bare paths get an alias defaulting to the filename:

```
XLSX_MCP_FILES=/abs/path/to/sales.xlsx
```

With `alias`/`filename` as the alias:

- **One file configured** → every tool can be called with `path` omitted entirely.
- **Multiple files configured** → pass the alias (or filename) as `path`.
- `list_configured_files()` returns the alias → absolute-path mapping.
- Raw absolute **and relative** paths still work for files you didn't preload.

**Claude Code — `.mcp.json` with preloading:**

```json
{
  "mcpServers": {
    "xlsx-tools-mcp": {
      "command": "uvx",
      "args": ["xlsx-tools-mcp"],
      "env": {
        "XLSX_MCP_FILES": "report=/data/report.xlsx,sales=/data/sales.xlsx"
      }
    }
  }
}
```

**OpenCode with preloading:**

```json
{
  "mcp": {
    "xlsx-tools-mcp": {
      "type": "local",
      "command": ["uvx", "xlsx-tools-mcp"],
      "env": { "XLSX_MCP_FILES": "report=/data/report.xlsx,sales=/data/sales.xlsx" },
      "enabled": true
    }
  }
}
```

---

## Tool reference

All 35 tools. Unless noted, `path` accepts a filesystem path, a preloaded alias/filename, or may be omitted when exactly one file is preloaded. `create_workbook` is the exception — its `path` is required because a new file is never preloaded.

> **Tool metadata:** every tool has a `title` and MCP annotations: `readOnlyHint` on the 12 read-only tools, `destructiveHint: false` on writes that only add or format, and `openWorldHint: false` everywhere (local files only). Every parameter has a description in the input schema, and fixed choices are enums. The server `instructions` include a short guide for choosing between overlapping tools (`read_sheet` / `query_sheet` / `aggregate_sheet` / `profile_sheet`, `write_cells` / `append_rows`, `clear_range` / `delete_rows`).

> **Response shape (all write tools):** every write tool returns `{"saved": bool, "recalculated": bool, "errors_found": list, "message": str}`. When non-empty, `errors_found` is a list of `{"sheet": "...", "cell": "B2", "error": "#DIV/0!"}`. Extra keys appear when relevant: `recalc_mode` (when LibreOffice was asked to run), `rules` and `rolled_back` (when the workbook has rules), and `ripple` (other formula results changed by a value write). Every write tool also accepts an optional `reason` string, recorded in the change log when it is enabled.

### Inspect / Read

| Tool | Description |
|------|-------------|
| `list_configured_files()` | List files preloaded at startup via `XLSX_MCP_FILES`, as an alias → absolute-path map. Call this first if unsure what's available. |
| `list_sheets(path?)` | List every sheet in the workbook with approximate row/column counts (calamine). |
| `get_workbook_info(path?)` | Workbook-level metadata: per-sheet exact dimensions, `max_row`/`max_column`, sheet state, the active sheet, and defined names. |
| `read_sheet(sheet, cell_range?, max_rows?, offset_row?, format?, path?)` | Read cell values. `cell_range` limits the reading area. `max_rows` caps rows. `offset_row` skips a number of rows for pagination. `format` can be "array" (2D list), "records" (list of dicts using first row as keys), or "markdown" (table string). |
| `get_cell(sheet, cell, path?)` | Full detail for a single cell: value (cached computed), formula, number format, font (bold/italic/size/color), fill color, merge state, comment. |
| `search_workbook(query, sheet?, match_case?, limit?, path?)` | Substring search across one or all sheets. `sheet` restricts to one sheet; `match_case=True` makes it case-sensitive; `limit` caps matches. Returns `{"sheet", "cell", "value"}`. |
| `aggregate_sheet(sheet, group_by, agg, cell_range?, has_header?, path?)` | Group and aggregate with pandas. `group_by` is a list of column names (taken from the header row); `agg` maps column name → aggregation function, e.g. `{"amount": "sum"}`. `has_header=True` (default) reads column names from the first row. Returns `{columns, records, row_count}`. |
| `profile_sheet(sheet, sample_rows?, path?)` | Profile sheet metadata: inferred types, null counts, min/max, sample values. |
| `query_sheet(sheet, filter_query, columns?, max_rows?, cell_range?, path?)` | Filter sheet data using pandas expressions with AST-hardened code injection defense. |

### Write

| Tool | Description |
|------|-------------|
| `create_workbook(path, sheets?, overwrite?)` | Create a new `.xlsx`/`.xlsm` workbook. `sheets` defaults to `["Sheet1"]`. `overwrite=True` replaces an existing file. `path` is **required** (new files are never preloaded). |
| `write_cells(sheet, cells, create_sheet_if_missing?, recalculate?, allow_external_formulas?, backup?, path?)` | Write values and/or formulas into specific cells. `cells` is a list of `{"cell": "A1", "value": ...}` or `{"cell": "B1", "formula": "=A1*2"}`. Optionally create the sheet first; `recalculate=True` (default) runs the LibreOffice recompute. |
| `append_rows(sheet, rows, create_sheet_if_missing?, recalculate?, allow_external_formulas?, backup?, path?)` | Append rows after the last used row. `rows` is a list of rows, each a list of cell values in column order. |
| `create_sheet(sheet, index?, path?)` | Add a new empty sheet. `index` is a zero-based insert position; omit to append at the end. |
| `delete_sheet(sheet, path?)` | Delete a sheet. Fails if it's the only sheet left. |
| `rename_sheet(old_name, new_name, path?)` | Rename an existing sheet. |
| `copy_sheet(source_sheet, target_sheet, path?)` | Duplicate a sheet including its contents, formulas, and styles. |
| `insert_rows(sheet, start_row, count?, recalculate?, path?)` | Insert blank rows before `start_row` (1-based), shifting existing rows down. `count` defaults to 1. |
| `delete_rows(sheet, start_row, count?, recalculate?, path?)` | Delete rows starting at `start_row` (1-based), shifting rows below upward. `count` defaults to 1. |
| `insert_columns(sheet, start_column, count?, recalculate?, path?)` | Insert blank columns before `start_column` (1-based), shifting existing columns right. `count` defaults to 1. |
| `delete_columns(sheet, start_column, count?, recalculate?, path?)` | Delete columns starting at `start_column` (1-based), shifting columns to the right leftward. `count` defaults to 1. |
| `merge_cells(sheet, cell_range, path?)` | Merge a rectangular range (e.g. `"A1:C1"`) into one cell. |
| `unmerge_cells(sheet, cell_range, path?)` | Undo a merge on a previously-merged range. |
| `set_cell_style(sheet, cell_range, style, path?)` | Apply formatting to a range (e.g. `"A1:D1"`). `style` keys: `bold`, `italic`, `font_size`, `font_color` (hex RGB, e.g. `"FF0000"`), `bg_color` (hex RGB), `horizontal`, `vertical` (alignment), `border` (`"thin"`, `"medium"`, `"thick"`, …), `number_format` (e.g. `"#,##0.00"`). |
| `autofit_columns(sheet, min_width?, max_width?, padding?, path?)` | Automatically adjust column widths based on maximum text length to prevent clipping. |
| `clear_range(sheet, cell_range, clear_values?, clear_styles?, path?)` | Clear cell values and/or styles within a specified range without deleting rows or columns. |
| `create_table(sheet, cell_range, table_name, style_name?, show_filter?, show_row_stripes?, path?)` | Create a formal Excel Table over a range. |
| `create_chart(sheet, chart_type, data_range, categories_range?, title?, target_cell?, path?)` | Add a native Excel chart (bar, line, pie, scatter). |
| `recalculate_workbook(mode?, path?)` | Force a LibreOffice headless recalculation pass and report any formula errors found. `mode` is `"inject"` or `"roundtrip"` (default from `XLSX_MCP_RECALC_MODE`). |
| `restore_backup(backup_path, target_path)` | Restore a `.bak` backup file over a target workbook atomically. |

### Rules, review, sharing & change log

| Tool | Description |
|------|-------------|
| `set_rules(rules, on_violation?, on_unverifiable?, path?)` | Replace the workbook's rules (empty list removes them) and evaluate them once. Stored as `<file>.rules.json` next to the workbook. `on_violation`: `"rollback"` (default) or `"warn"`. `on_unverifiable`: `"warn"` (default) or `"rollback"`. |
| `check_rules(path?)` | Evaluate the rules against the current values without writing. |
| `export_review_copy(baseline_path?, output_path?, overwrite?, highlight_color?, include_recalc_changes?, max_annotated_cells?, path?)` | Write a copy that highlights every cell changed since the baseline (default: the newest `.bak` of the file) with a Before/After comment, plus a `_Changes` sheet. Default output: `<name>.review.xlsx`. The workbook itself is not modified. |
| `inspect_for_sharing(path?)` | Read-only report of hidden or sensitive content, each finding with a severity and whether `sanitize_copy` can remove it. |
| `sanitize_copy(output_path, remove?, overwrite?, path?)` | Write a cleaned copy. `remove` is any of `comments`, `document_properties`, `hidden_sheets`, `external_links`, `change_log`, `macros` (default: the first two plus `change_log`). Hidden sheets are removed only when nothing references them; external links become their last cached values. |
| `get_change_log(sheet?, cell?, limit?, path?)` | Read the in-workbook change log, newest first; `cell` also matches range entries that contain it. |

#### Rule types

| `type` | Fields | Fails when |
|--------|--------|------------|
| `cells_equal` | `left`, `right` (`"Sheet!B40"`), `tolerance?` | the two values differ by more than `tolerance` |
| `cell_compare` | `cell`, `op` (`==`, `!=`, `<`, `<=`, `>`, `>=`), `value` **or** `cell_ref` | the comparison is false |
| `column_check` | `sheet`, `expr` (safe pandas expression on header names), `cell_range?` | any data row does not satisfy `expr` (blank values count as not satisfying it) |
| `column_unique` | `sheet`, `column`, `cell_range?` | a non-empty value appears twice |
| `not_empty` | `range` (`"Data!A2:A500"`) | a cell in the range is empty |
| `no_new_errors` | — | the write produced formula errors that were not there before |

A write is rolled back only when it makes a rule fail that passed before, or makes a failing rule worse (more bad rows, larger difference). Rules that already failed and did not get worse are reported under `pre_existing` and never block. When a rule's state before the write could not be determined (for example, formula results were empty because an earlier write skipped recalculation), a failure is reported under `unknown_before` and only blocks with `on_unverifiable="rollback"`; run `recalculate_workbook` to refresh the results. `no_new_errors` compares error counts instead of positions for row/column inserts and deletes. If saving or recalculating raises an error, the pre-write copy is restored, so rules never fail open.

Rules are evaluated for every tool that can change values: value writes, row/column inserts and deletes, `clear_range`, `merge_cells`, `rename_sheet` and `delete_sheet` (the last three do not recalculate, so rules on formula results are `unverifiable` there). Rule references are not shifted by `insert_rows`/`delete_rows`; update them with `set_rules`. Formatting-only tools (`set_cell_style`, `autofit_columns`, `create_table`, `create_chart`, `create_sheet`, `copy_sheet`, `unmerge_cells`) report `"rules": {"status": "skipped"}` and keep the workbook's cached formula results across the save.

### Example payload — `write_cells`

A call writing a formula and a value:

```json
{
  "sheet": "Sheet1",
  "cells": [
    { "cell": "A1", "value": 100 },
    { "cell": "B1", "formula": "=A1*2" }
  ],
  "recalculate": true,
  "path": "/data/budget.xlsx"
}
```

Matching response:

```json
{
  "saved": true,
  "recalculated": true,
  "errors_found": [],
  "message": "Recalculated with LibreOffice headless."
}
```

If a formula this touches produced an error, `errors_found` would look like:

```json
{
  "saved": true,
  "recalculated": true,
  "errors_found": [
    { "sheet": "Sheet1", "cell": "C5", "error": "#DIV/0!" }
  ],
  "message": "Recalculated with LibreOffice headless."
}
```

### Example payload — `set_rules` and a rolled-back write

```json
{
  "rules": [
    { "id": "balance", "type": "cells_equal", "left": "Balance!B40", "right": "Balance!C40", "tolerance": 0.01 },
    { "id": "qty", "type": "column_check", "sheet": "Orders", "expr": "Qty >= 0" },
    { "id": "errors", "type": "no_new_errors" }
  ]
}
```

A later `write_cells` that unbalances the sheet:

```json
{
  "saved": false,
  "recalculated": true,
  "errors_found": [],
  "message": "Write rolled back: rule(s) ['balance'] failed after the change. The file is unchanged.",
  "recalc_mode": "inject",
  "rules": {
    "status": "violated",
    "results": [{ "id": "balance", "type": "cells_equal", "status": "fail", "detail": "Balance!B40=1200, Balance!C40=1150, diff=50", "metric": 50.0 }],
    "regressions": ["balance"],
    "pre_existing": []
  },
  "rolled_back": true,
  "ripple": { "skipped": "write was rolled back" }
}
```

A write that is kept reports its side effects:

```json
"ripple": {
  "changed_count": 2,
  "unknown_before": 0,
  "sample": [
    { "sheet": "Summary", "cell": "D4", "before": 1200000, "after": 1500000, "delta": 300000 },
    { "sheet": "Summary", "cell": "D9", "before": 0.31, "after": 0.27, "delta": -0.04 }
  ],
  "truncated": false
}
```

### Example payload — `query_sheet`

```json
{
  "sheet": "Sales",
  "filter_query": "Revenue > 500 and Status == 'PAID'",
  "columns": ["Customer", "Revenue"],
  "max_rows": 10
}
```

### Example payload — `create_table` & `create_chart`

```json
{
  "sheet": "Summary",
  "cell_range": "A1:D10",
  "table_name": "SalesTable",
  "style_name": "TableStyleMedium9"
}
```

```json
{
  "sheet": "Summary",
  "chart_type": "bar",
  "data_range": "B1:B10",
  "categories_range": "A2:A10",
  "title": "Monthly Revenue",
  "target_cell": "F2"
}
```

---

## Environment variables

| Variable | Default | Description |
|----------|---------|-------------|
| `XLSX_MCP_FILES` | *empty* | Comma or semicolon-separated `alias=path` or `path` to preload at startup. |
| `XLSX_MCP_ALLOWED_DIRS` | *empty* | Comma or semicolon-separated directories to confine file access (paths outside raise `AccessDeniedError`). Supports case-insensitivity on Windows. |
| `XLSX_MCP_AUTO_BACKUP` | `true` | Automatically create a `.bak` copy before modifying a workbook. |
| `XLSX_MCP_BACKUP_DIR` | *empty* | Custom directory for backup files (defaults to beside the target file). |
| `XLSX_MCP_RECALC_TIMEOUT` | `60` | Timeout in seconds for headless LibreOffice formula recalculation. |
| `XLSX_MCP_LOCK_TIMEOUT` | `10` | Timeout in seconds to acquire per-file lock before failing. |
| `XLSX_MCP_RECALC_MODE` | `inject` | `inject` writes only computed values into the file; `roundtrip` replaces the file with LibreOffice's export. |
| `XLSX_MCP_MAX_PART_BYTES` | `268435456` | Largest uncompressed XML part (bytes) the server parses directly. |
| `XLSX_MCP_RIPPLE` | `true` | Include the `ripple` report in value-write responses. |
| `XLSX_MCP_RIPPLE_MAX_CELLS` | `200000` | Skip the ripple report above this many formula cells. |
| `XLSX_MCP_CHANGE_LOG` | `false` | Record every write on a very hidden `_xlsx_mcp_log` sheet inside the workbook. |
| `XLSX_MCP_CHANGE_LOG_MAX_ROWS` | `10000` | Oldest change-log rows are removed above this count. |
| `XLSX_MCP_CHANGE_LOG_MAX_CELLS_PER_OP` | `500` | Per-cell entries per tool call; the rest are summarized in one row. |

---

## Security & concurrency

- **Path confinement** — `XLSX_MCP_ALLOWED_DIRS` (comma or semicolon-separated) restricts all reads and writes to specific directories. If set, any attempt to access a file outside these boundaries raises `AccessDeniedError`. Case-insensitive path comparisons and multi-drive paths on Windows are supported.
- **Auto-backup & restore** — Write tools will automatically create a `.bak` sibling file before modifying the workbook if `backup=True` (which defaults to `XLSX_MCP_AUTO_BACKUP=true`). Use `XLSX_MCP_BACKUP_DIR` to save backups to a specific folder. Use `restore_backup` to atomically revert changes.
- **Formula injection guard** — `write_cells` and `append_rows` reject formulas containing `WEBSERVICE`, `HYPERLINK`, `INDIRECT`, `RTD`, `CALL`, `REGISTER`, or DDE commands (`|`) by default to prevent data exfiltration or command execution. Pass `allow_external_formulas=True` to bypass this check.
- **AST safe query validation** — `query_sheet` uses an AST parser to strictly allow only comparisons, boolean logic, and safe operators, blocking `@` variable lookups, code execution calls, and private dunder attributes.
- **XML-bomb protection** — `defusedxml` is an automatic dependency of this package. openpyxl auto-detects it and uses its hardened XML parser, so a malicious `.xlsx` (a zip of XML) can't trigger entity-expansion resource exhaustion. No configuration needed.
- **Per-file locking** — every read/write acquires a sibling `<path>.lock` file (via `filelock`). Concurrent tool calls or other processes touching the same workbook are serialized so writes never interleave and corrupt the file.
- **Transient lock resilience** — on Windows NTFS, atomic file swaps can momentarily encounter sharing violations (`WinError 32` / `33`) while antivirus scanners or search indexers inspect newly saved files. The server performs an exponential backoff retry to transparently succeed once released, or provides a clear guidance message if the file is permanently open in Microsoft Excel.
- **Hardened XML access** — the parts the server reads or rewrites itself are parsed with lxml with entity resolution, DTD loading and network access disabled; parts declaring a `DOCTYPE` are refused, and each part is size-capped (`XLSX_MCP_MAX_PART_BYTES`) before it is read.
- **Derived files never overwrite their source** — `export_review_copy` and `sanitize_copy` refuse an `output_path` equal to an input, apply path confinement to it, and lock all files involved in a fixed order.
- **Rules fail closed** — a malformed `<file>.rules.json` blocks writes with a clear error instead of silently skipping the checks. A rollback restores a snapshot taken under the same file lock.
- **Change-log sheet is reserved** — `_xlsx_mcp_log` cannot be read, written, renamed, copied or deleted through the regular tools; logged values are stored as text so they can never become live formulas. `inspect_for_sharing` reports it and `sanitize_copy` removes it by default.
- **Recalc timeout** — `XLSX_MCP_RECALC_TIMEOUT` (seconds, default `60`) caps how long the LibreOffice recalculation pass may run.
- **Lock timeout** — `XLSX_MCP_LOCK_TIMEOUT` (seconds, default `10`) caps how long a tool will wait to acquire the per-file lock before failing.

---

## Troubleshooting

- **`errors_found` is empty even though my formula is broken** — recalculation likely didn't run. Check the `message` field: if it says LibreOffice wasn't found, the file was saved via openpyxl as-is and formulas were **not** recomputed (cached values may be stale). Install LibreOffice (see [Requirements](#requirements)).
- **Recalculation is slow or times out** — raise `XLSX_MCP_RECALC_TIMEOUT` (default 60s). On timeout, the file is still saved, but `recalculated` will be `false` and `message` says the recalc timed out.
- **`LockTimeoutError` on concurrent access** — another operation holds the lock. Raise `XLSX_MCP_LOCK_TIMEOUT` (default 10s), or retry when the other operation finishes.
- **`Workbook is locked by another process` / `WinError 32`** — on Windows, Microsoft Excel and some applications lock open files exclusively. Close the workbook in Excel and retry. Transient locks from antivirus scanners or search indexers are automatically retried by the server before failing.
- **`recalculated` is `false` with "values could not be written back"** — injection met an unexpected file layout; the file was left as openpyxl saved it. Set `XLSX_MCP_RECALC_MODE=roundtrip` (or call `recalculate_workbook(mode="roundtrip")`) to use LibreOffice's export instead.
- **`saved: false, rolled_back: true`** — the write broke one of the workbook's rules; see `rules.regressions` and `rules.results`. Use `check_rules` to inspect, or `set_rules` to change the rules.
- **`Rules file ... could not be read`** — `<file>.rules.json` is malformed. Fix it with `set_rules` (or `set_rules([])` to remove the rules) or delete the file.
- **`rules.status` is `unverifiable`** — values could not be checked, usually because LibreOffice is missing and formula results are empty. Install LibreOffice, or choose `on_unverifiable="rollback"` to block such writes.
- **`Sheet '_xlsx_mcp_log' is reserved`** — that sheet holds the change log; read it with `get_change_log`.
- **"Sheet not found"** — the error message lists the available sheet names, so you can pick the correct one.
- **`path` required / no file configured** — you called a tool without `path` but no (or multiple) files are preloaded. Preload one file via `XLSX_MCP_FILES`, pass an explicit alias, or pass a raw path.

---

## Development / Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md). Run the test suite with:

```bash
uv run pytest
```