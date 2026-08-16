# xlsx-tools-mcp

An MCP server for reading and writing Excel (.xlsx) files with high accuracy, while preserving the file's existing structure, styles, and formulas.

[![CI](https://github.com/ruriazz/xlsx-tools-mcp/actions/workflows/ci.yml/badge.svg)](https://github.com/ruriazz/xlsx-tools-mcp/actions/workflows/ci.yml) ![PyPI Version](https://img.shields.io/pypi/v/xlsx-tools-mcp)

---

## Overview

`xlsx-tools-mcp` exposes 20 Model Context Protocol (MCP) tools that give an LLM agent accurate, structure-preserving read **and** write access to Excel `.xlsx` files. It runs as a standard stdio MCP server: you install it and register it with an MCP client (Claude Code, OpenCode, etc.), and the client's agent can list sheets, read cell ranges, search values, aggregate data, write cells/formulas, manage sheets/rows/columns, apply styles, and force formula recalculation.

It is built around the principle that editing an existing workbook should **not** destroy what it doesn't touch.

### Features

- **Structure-preserving writes via openpyxl** — writes load the existing workbook and save it back, preserving styles, merged cells, comments, and any aspect the edit doesn't touch.
- **Never-stale formula results via LibreOffice recalculation** — openpyxl writes formula *strings* but never evaluates them. After every value/formula write the server runs a headless LibreOffice pass to recompute real results, then returns `errors_found` — any Excel error values (`#REF!`, `#DIV/0!`, `#N/A`, …) produced by the recalculation.
- **Fast reads via python-calamine** — a Rust-backed parser for accurate, fast type inference, with an automatic openpyxl fallback when you need formulas/styles/comments or when calamine can't parse the file.
- **pandas-based grouping/aggregation** — `aggregate_sheet` groups and aggregates on top of the normal read path, so merged cells and styling in the source range are preserved before flattening.
- **Per-file locking** — concurrent tool calls (or other processes) touching the same workbook are serialized via a sibling `<path>.lock` file (filelock), so writes never interleave and corrupt the file.
- **XML-bomb protection** — the `defusedxml` package is an automatic dependency; openpyxl detects it and uses its hardened XML parser, so hostile `xlsx` XML can't expand into resource exhaustion.
- **Preload files at startup** — set `XLSX_MCP_FILES` to preload one or more workbooks; tools can then be called with `path` omitted or with a short alias instead of a full filesystem path.

---

## Architecture

```
┌──────────────────────── Supervisor (MCP transport, stdio)
│  src/xlsx_tools_mcp/server.py     20 MCP tools + instructions
│  src/xlsx_tools_mcp/settings.py   env vars, preloaded files, path resolution
│  src/xlsx_tools_mcp/locking.py    per-file <path>.lock serialization
│  src/xlsx_tools_mcp/errors.py     domain error types
│  src/xlsx_tools_mcp/recalc.py     LibreOffice headless recalc + error scanning
│
├─ Read path
│  src/xlsx_tools_mcp/io/reader.py      calamine primary → openpyxl fallback
│  src/xlsx_tools_mcp/io/transform.py   pandas aggregation on read results
│
└─ Write path
   src/xlsx_tools_mcp/io/writer.py      openpyxl → LibreOffice recalc → scan errors
```

The **io layer** (`io/`) is deliberately decoupled from the MCP transport (`server.py`). Each MCP tool is a thin wrapper that resolves the target path, takes the per-file lock, and calls one io-layer function. This keeps the core logic independent of MCP, so it can be tested directly (see `tests/`).

### The recalculation tradeoff

After a write that touches cell values or formulas, the server runs `soffice --headless --convert-to xlsx` on the file so every formula gets a real computed value. This round-trip recomputes formulas but **re-exports the whole workbook** — it is a tradeoff, **not** a guarantee of bit-perfect preservation. Features that openpyxl would otherwise preserve may not survive identically: pivot tables, charts, data validation, some formats, and some defined names.

If you're working on a structurally complex workbook where that risk matters, you can pass `recalculate=False` on the value/formula-writing tools (`write_cells`, `append_rows`, `insert_rows`, `delete_rows`, `insert_columns`, `delete_columns`) to save with openpyxl only and skip the round-trip entirely.

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
```

The server finds LibreOffice by checking `soffice` / `libreoffice` on `PATH` and the standard macOS install location (`/Applications/LibreOffice.app/Contents/MacOS/soffice`).

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

Set the `XLSX_MCP_FILES` environment variable in the **MCP server config `env`** section (not your interactive shell — the server is launched by the client) to preload workbooks at startup. Format: comma-separated `alias=absolute/path` entries, or bare absolute paths:

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

All 20 tools. Unless noted, `path` accepts a filesystem path, a preloaded alias/filename, or may be omitted when exactly one file is preloaded. `create_workbook` is the exception — its `path` is required because a new file is never preloaded.

> **Response shape (all write tools):** every write tool returns `{"saved": bool, "recalculated": bool, "errors_found": list, "message": str}`. When non-empty, `errors_found` is a list of `{"sheet": "...", "cell": "B2", "error": "#DIV/0!"}`.

### Inspect / Read

| Tool | Description |
|------|-------------|
| `list_configured_files()` | List files preloaded at startup via `XLSX_MCP_FILES`, as an alias → absolute-path map. Call this first if unsure what's available. |
| `list_sheets(path?)` | List every sheet in the workbook with approximate row/column counts (calamine). |
| `get_workbook_info(path?)` | Workbook-level metadata: per-sheet exact dimensions, `max_row`/`max_column`, sheet state, the active sheet, and defined names. |
| `read_sheet(sheet, cell_range?, max_rows?, path?)` | Read cell values as a 2D array addressed absolutely from A1. `cell_range` is an optional A1-style range (e.g. `"B2:F20"`); omit to read the full used area. `max_rows` optionally caps the number of rows returned. |
| `get_cell(sheet, cell, path?)` | Full detail for a single cell: value (cached computed), formula, number format, font (bold/italic/size/color), fill color, merge state, comment. |
| `search_workbook(query, sheet?, match_case?, limit?, path?)` | Substring search across one or all sheets. `sheet` restricts to one sheet; `match_case=True` makes it case-sensitive; `limit` caps matches. Returns `{"sheet", "cell", "value"}`. |
| `aggregate_sheet(sheet, group_by, agg, cell_range?, has_header?, path?)` | Group and aggregate with pandas. `group_by` is a list of column names (taken from the header row); `agg` maps column name → aggregation function, e.g. `{"amount": "sum"}`. `has_header=True` (default) reads column names from the first row. Returns `{columns, records, row_count}`. |

### Write

| Tool | Description |
|------|-------------|
| `create_workbook(path, sheets?, overwrite?)` | Create a new `.xlsx`/`.xlsm` workbook. `sheets` defaults to `["Sheet1"]`. `overwrite=True` replaces an existing file. `path` is **required** (new files are never preloaded). |
| `write_cells(sheet, cells, create_sheet_if_missing?, recalculate?, path?)` | Write values and/or formulas into specific cells. `cells` is a list of `{"cell": "A1", "value": ...}` or `{"cell": "B1", "formula": "=A1*2"}`. Optionally create the sheet first; `recalculate=True` (default) runs the LibreOffice recompute. |
| `append_rows(sheet, rows, create_sheet_if_missing?, recalculate?, path?)` | Append rows after the last used row. `rows` is a list of rows, each a list of cell values in column order. |
| `create_sheet(sheet, index?, path?)` | Add a new empty sheet. `index` is a zero-based insert position; omit to append at the end. |
| `delete_sheet(sheet, path?)` | Delete a sheet. Fails if it's the only sheet left. |
| `insert_rows(sheet, start_row, count?, recalculate?, path?)` | Insert blank rows before `start_row` (1-based), shifting existing rows down. `count` defaults to 1. |
| `delete_rows(sheet, start_row, count?, recalculate?, path?)` | Delete rows starting at `start_row` (1-based), shifting rows below upward. `count` defaults to 1. |
| `insert_columns(sheet, start_column, count?, recalculate?, path?)` | Insert blank columns before `start_column` (1-based), shifting existing columns right. `count` defaults to 1. |
| `delete_columns(sheet, start_column, count?, recalculate?, path?)` | Delete columns starting at `start_column` (1-based), shifting columns to the right leftward. `count` defaults to 1. |
| `merge_cells(sheet, cell_range, path?)` | Merge a rectangular range (e.g. `"A1:C1"`) into one cell. |
| `unmerge_cells(sheet, cell_range, path?)` | Undo a merge on a previously-merged range. |
| `set_cell_style(sheet, cell_range, style, path?)` | Apply formatting to a range (e.g. `"A1:D1"`). `style` keys: `bold`, `italic`, `font_size`, `font_color` (hex RGB, e.g. `"FF0000"`), `bg_color` (hex RGB), `horizontal`, `vertical` (alignment), `border` (`"thin"`, `"medium"`, `"thick"`, …), `number_format` (e.g. `"#,##0.00"`). |
| `recalculate_workbook(path?)` | Force a LibreOffice headless recalculation pass and report any formula errors found. |

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

---

## Security & concurrency

- **XML-bomb protection** — `defusedxml` is an automatic dependency of this package. openpyxl auto-detects it and uses its hardened XML parser, so a malicious `.xlsx` (a zip of XML) can't trigger entity-expansion resource exhaustion. No configuration needed.
- **Per-file locking** — every read/write acquires a sibling `<path>.lock` file (via `filelock`). Concurrent tool calls or other processes touching the same workbook are serialized so writes never interleave and corrupt the file.
- **Recalc timeout** — `XLSX_MCP_RECALC_TIMEOUT` (seconds, default `60`) caps how long the LibreOffice recalculation pass may run.
- **Lock timeout** — `XLSX_MCP_LOCK_TIMEOUT` (seconds, default `10`) caps how long a tool will wait to acquire the per-file lock before failing.

---

## Troubleshooting

- **`errors_found` is empty even though my formula is broken** — recalculation likely didn't run. Check the `message` field: if it says LibreOffice wasn't found, the file was saved via openpyxl as-is and formulas were **not** recomputed (cached values may be stale). Install LibreOffice (see [Requirements](#requirements)).
- **Recalculation is slow or times out** — raise `XLSX_MCP_RECALC_TIMEOUT` (default 60s). On timeout, the file is still saved, but `recalculated` will be `false` and `message` says the recalc timed out.
- **`LockTimeoutError` on concurrent access** — another operation holds the lock. Raise `XLSX_MCP_LOCK_TIMEOUT` (default 10s), or retry when the other operation finishes.
- **"Sheet not found"** — the error message lists the available sheet names, so you can pick the correct one.
- **`path` required / no file configured** — you called a tool without `path` but no (or multiple) files are preloaded. Preload one file via `XLSX_MCP_FILES`, pass an explicit alias, or pass a raw path.

---

## Development / Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md). Run the test suite with:

```bash
uv run pytest
```