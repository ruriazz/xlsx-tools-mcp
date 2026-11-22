"""Review copy: a highlighted, annotated copy showing what changed since a baseline.

The current workbook is loaded, every changed cell gets a fill and a
"Before / After" comment, and a `_Changes` sheet lists each change with a link to
the cell. The source and the baseline are never written.
"""

from __future__ import annotations

import datetime
import io
import os
import re
import sys
import zipfile
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.comments import Comment
from openpyxl.styles import Font, PatternFill
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.worksheet.hyperlink import Hyperlink

from .. import inject, ledger, settings
from ..errors import BaselineNotFoundError, InvalidWorkbookError, WorkbookNotFoundError
from ..refs import quote_sheet
from ..ripple import same_value
from . import copies, writer

REVIEW_AUTHOR = "xlsx-tools-mcp"
CHANGES_SHEET = "_Changes"
NEW_SHEET_TAB = "FFC000"
MAX_LISTED = 50_000
MAX_DISPLAY = 500
SHIFT_RATIO = 0.5
SHIFT_MIN_CHANGES = 10
_HEX_COLOR = re.compile(r"^(?:[0-9A-Fa-f]{6}|[0-9A-Fa-f]{8})$")
_CASE_INSENSITIVE_FS = os.name == "nt" or sys.platform == "darwin"


def default_output(path: str) -> str:
    p = Path(path)
    return str(p.with_name(f"{p.stem}.review{p.suffix}"))


def find_latest_backup(path: str) -> str:
    """Newest `<name>.<YYYYmmdd-HHMMSS>.bak` for this workbook (by the timestamp in its name)."""
    p = Path(path)
    folder = Path(writer.BACKUP_DIR or settings.BACKUP_DIR or p.parent)
    flags = re.IGNORECASE if _CASE_INSENSITIVE_FS else 0
    pattern = re.compile(rf"^{re.escape(p.name)}\.(\d{{8}}-\d{{6}})\.bak$", flags)
    best: tuple[str, str] | None = None
    try:
        names = os.listdir(folder)
    except OSError:
        names = []
    for name in names:
        m = pattern.match(name)
        if m and (best is None or m.group(1) > best[0]):
            best = (m.group(1), str(folder / name))
    if best is None:
        raise BaselineNotFoundError(
            f"No backup found for '{p.name}' in '{folder}'. Enable XLSX_MCP_AUTO_BACKUP (or pass "
            "backup=True on writes) so a .bak exists, or pass baseline_path explicitly."
        )
    return best[1]


def _plain(v: Any) -> Any:
    text = getattr(v, "text", None)  # ArrayFormula / DataTableFormula
    if isinstance(text, str):
        return text if text.startswith("=") else f"={text}"
    if isinstance(v, str) and v == "":
        return None
    return v


def _is_formula(v: Any) -> bool:
    return isinstance(v, str) and v.startswith("=")


def _display(v: Any) -> str:
    if v is None:
        return "(empty)"
    if isinstance(v, (datetime.datetime, datetime.date, datetime.time)):
        text = v.isoformat()
    else:
        text = str(v)
    return text if len(text) <= MAX_DISPLAY else text[: MAX_DISPLAY - 1] + "…"


def _sheet_cells(ws) -> dict[str, Any]:
    out: dict[str, Any] = {}
    for row in ws.iter_rows():
        for c in row:
            v = _plain(c.value)
            if v is not None:
                out[c.coordinate] = v
    return out


def _load_baseline(baseline: str) -> dict[str, dict[str, Any]]:
    if not Path(baseline).is_file():
        raise BaselineNotFoundError(f"Baseline file not found: '{baseline}'")
    if Path(baseline).suffix.lower() not in (".bak", ".xlsx", ".xlsm"):
        raise ValueError(f"baseline_path must be a .bak, .xlsx or .xlsm file, got '{baseline}'")
    try:
        # Loaded from bytes: openpyxl refuses paths whose extension is not .xlsx/.xlsm,
        # and backups end in .bak.
        wb = openpyxl.load_workbook(io.BytesIO(Path(baseline).read_bytes()), read_only=True, data_only=False)
    except PermissionError:
        raise
    except (zipfile.BadZipFile, InvalidFileException, KeyError) as exc:
        raise InvalidWorkbookError(f"Baseline '{baseline}' is not a valid Excel workbook: {exc}") from exc
    try:
        out = {}
        for ws in wb.worksheets:
            if ledger.is_reserved(ws.title):
                continue
            ws.reset_dimensions()
            out[ws.title] = _sheet_cells(ws)
        return out
    finally:
        wb.close()


def _classify(b: Any, c: Any, cb: Any, cc: Any, include_recalc: bool) -> tuple[str, Any, Any] | None:
    if b is None and c is None:
        return None
    if b is None:
        return ("added", None, c)
    if c is None:
        return ("removed", b, None)
    if _is_formula(b) and _is_formula(c):
        if b != c:
            return ("formula", b, c)
        if include_recalc and not same_value(cb, cc):
            return ("recalc_only", cb, cc)
        return None
    if _is_formula(b) != _is_formula(c) or not same_value(b, c):
        return ("value", b, c)
    return None


def _diff(
    base: dict[str, dict[str, Any]],
    cur: dict[str, dict[str, Any]],
    base_cache: dict[str, inject.SheetValues],
    cur_cache: dict[str, inject.SheetValues],
    include_recalc: bool,
) -> tuple[list[dict[str, Any]], list[str]]:
    changes: list[dict[str, Any]] = []
    warnings: list[str] = []
    for sheet, cells in cur.items():
        if sheet not in base:
            continue
        before = base[sheet]
        bc, cc = base_cache.get(sheet, {}), cur_cache.get(sheet, {})
        sheet_changes = 0
        for coord in sorted(set(before) | set(cells), key=_coord_key):
            kind = _classify(
                before.get(coord), cells.get(coord),
                inject.to_python(bc.get(coord)), inject.to_python(cc.get(coord)),
                include_recalc,
            )
            if kind:
                changes.append({"sheet": sheet, "cell": coord, "kind": kind[0], "before": kind[1], "after": kind[2]})
                sheet_changes += 1
        used = max(len(before), len(cells), 1)
        if sheet_changes >= SHIFT_MIN_CHANGES and sheet_changes / used > SHIFT_RATIO:
            warnings.append(
                f"Sheet '{sheet}': {round(100 * sheet_changes / used)}% of cells changed — "
                "possible row/column shift; the comparison is by cell position"
            )
    return changes, warnings


def _coord_key(coord: str) -> tuple[int, int]:
    m = re.match(r"([A-Z]+)(\d+)", coord)
    if not m:
        return (0, 0)
    col = 0
    for ch in m.group(1):
        col = col * 26 + ord(ch) - 64
    return (int(m.group(2)), col)


def _annotate(ws, coord: str, change: dict[str, Any], fill: PatternFill) -> bool:
    cell = ws[coord]
    if isinstance(cell, MergedCell):
        return False
    cell.fill = fill
    text = f"[{REVIEW_AUTHOR} review] {change['kind']}\nBefore: {_display(change['before'])}\nAfter: {_display(change['after'])}"
    if cell.comment is not None:
        cell.comment = Comment(f"{cell.comment.text}\n\n{text}", cell.comment.author or REVIEW_AUTHOR, width=260, height=120)
    else:
        cell.comment = Comment(text, REVIEW_AUTHOR, width=260, height=120)
    return True


def _set_text(cell, value: Any) -> None:
    cell.value = value
    if isinstance(value, str):
        cell.data_type = "s"  # keep "=..." before/after values as text


def _changes_sheet(wb, rows: list[dict[str, Any]]):
    name = CHANGES_SHEET
    n = 2
    taken = {s.casefold() for s in wb.sheetnames}  # Excel sheet names are case-insensitive
    while name.casefold() in taken:
        name = f"{CHANGES_SHEET}_{n}"
        n += 1
    ws = wb.create_sheet(name, 0)
    header = ["Sheet", "Cell", "Kind", "Before", "After"]
    for col, text in enumerate(header, start=1):
        c = ws.cell(row=1, column=col)
        _set_text(c, text)
        c.font = Font(bold=True)
    for r, ch in enumerate(rows, start=2):
        _set_text(ws.cell(row=r, column=1), ch["sheet"])
        cell_c = ws.cell(row=r, column=2)
        _set_text(cell_c, ch["cell"])
        if ch["cell"] and ch["sheet"] in wb.sheetnames:
            cell_c.hyperlink = Hyperlink(ref=cell_c.coordinate, location=f"{quote_sheet(ch['sheet'])}!{ch['cell']}")
        _set_text(ws.cell(row=r, column=3), ch["kind"])
        _set_text(ws.cell(row=r, column=4), _display(ch["before"]))
        _set_text(ws.cell(row=r, column=5), _display(ch["after"]))
    for letter, width in zip("ABCDE", (24, 10, 14, 40, 40)):
        ws.column_dimensions[letter].width = width
    ws.freeze_panes = "A2"
    # Exactly one selected tab: Excel groups every tabSelected sheet, and edits to a
    # group apply to all of its sheets.
    for other in wb.worksheets:
        other.sheet_view.tabSelected = False
    ws.sheet_view.tabSelected = True
    wb.active = ws
    return ws


def export_review_copy(
    path: str,
    baseline_path: str,
    output_path: str,
    overwrite: bool = False,
    highlight_color: str = "FFF2CC",
    include_recalc_changes: bool = False,
    max_annotated_cells: int = 2000,
) -> dict[str, Any]:
    if not _HEX_COLOR.match(highlight_color or ""):
        raise ValueError(f"highlight_color must be a 6 or 8 digit hex color, got {highlight_color!r}")
    if not 0 <= max_annotated_cells <= 100_000:
        raise ValueError(f"max_annotated_cells must be between 0 and 100000, got {max_annotated_cells}")
    if not Path(path).is_file():
        raise WorkbookNotFoundError(f"Workbook not found: '{path}'.")
    copies.check_output(path, output_path, overwrite, baseline_path)

    base = _load_baseline(baseline_path)
    base_cache = inject.read_cached_values(baseline_path)
    cur_cache = inject.read_cached_values(path)
    wb = copies.load_source(path, output_path)
    try:
        cur = {ws.title: _sheet_cells(ws) for ws in wb.worksheets if not ledger.is_reserved(ws.title)}
        changes, warnings = _diff(base, cur, base_cache, cur_cache, include_recalc_changes)
        sheets_added = [s for s in cur if s not in base]
        sheets_removed = [s for s in base if s not in cur]

        fill = PatternFill(fill_type="solid", start_color=highlight_color, end_color=highlight_color)
        annotated = 0
        capped = False
        for ch in changes:
            if annotated >= max_annotated_cells:
                capped = True
                break
            if _annotate(wb[ch["sheet"]], ch["cell"], ch, fill):
                annotated += 1
        for s in sheets_added:
            wb[s].sheet_properties.tabColor = NEW_SHEET_TAB

        listed_rows = [{"sheet": s, "cell": "", "kind": "sheet_added", "before": None, "after": s} for s in sheets_added]
        listed_rows += [{"sheet": s, "cell": "", "kind": "sheet_removed", "before": s, "after": None} for s in sheets_removed]
        listed_rows += changes
        listed = listed_rows[:MAX_LISTED]
        _changes_sheet(wb, listed)
    except Exception:
        wb.close()
        raise
    copies.save_copy(wb, path, output_path)

    summary: dict[str, Any] = {k: 0 for k in ("value", "formula", "added", "removed", "recalc_only")}
    for ch in changes:
        summary[ch["kind"]] += 1
    summary["sheets_added"] = sheets_added
    summary["sheets_removed"] = sheets_removed
    return {
        "output_path": output_path,
        "baseline_path": baseline_path,
        "summary": summary,
        "annotated": annotated,
        "listed": len(listed),
        "truncated": capped or len(listed_rows) > len(listed),
        "warnings": warnings,
    }
