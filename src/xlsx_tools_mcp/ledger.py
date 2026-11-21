"""Opt-in change log kept inside the workbook on a very hidden sheet.

Each write appends who/what/when/why rows, so the history travels with the file.
Values are stored as JSON text with an explicit string type, so a logged "=A1*2"
can never become a live formula.
"""

from __future__ import annotations

import datetime
import json
import uuid
from pathlib import Path
from typing import Any
import zipfile

import openpyxl
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from . import settings
from .context import LedgerEntry, WriteContext
from .errors import InvalidWorkbookError, ReservedSheetError, WorkbookNotFoundError
from .refs import cell_in_range, normalize_cell

LOG_SHEET = "_xlsx_mcp_log"
HEADER = ("ts_utc", "op_id", "tool", "sheet", "ref", "before", "after", "reason")
MAX_VALUE_CHARS = 2000
MAX_REASON_CHARS = 500
REF_NOTE = (
    "References are recorded at write time; later row/column inserts or deletes "
    "are not applied to older entries."
)


def is_reserved(name: Any) -> bool:
    return isinstance(name, str) and name.casefold() == LOG_SHEET.casefold()


def assert_not_reserved(*names: Any) -> None:
    for name in names:
        if is_reserved(name):
            raise ReservedSheetError(
                f"Sheet '{name}' is reserved for the xlsx-tools-mcp change log. "
                "Use get_change_log to read it."
            )


def visible_names(names: list[str]) -> list[str]:
    return [n for n in names if not is_reserved(n)]


def find_sheet_name(names: list[str]) -> str | None:
    return next((n for n in names if is_reserved(n)), None)


def _clean(text: str, limit: int) -> str:
    text = ILLEGAL_CHARACTERS_RE.sub("", text)
    return text if len(text) <= limit else text[: limit - 1] + "…"


def _plain(value: Any) -> Any:
    # openpyxl ArrayFormula / DataTableFormula expose their formula as `.text`.
    text = getattr(value, "text", None)
    return text if isinstance(text, str) else value


def encode(value: Any) -> str:
    return _clean(json.dumps(_plain(value), default=str, ensure_ascii=False), MAX_VALUE_CHARS)


def _decode(text: Any) -> Any:
    if not isinstance(text, str) or text == "":
        return None
    try:
        return json.loads(text)
    except ValueError:
        return text


def _log_sheet(wb: Workbook) -> Worksheet:
    existing = find_sheet_name(wb.sheetnames)
    if existing:
        return wb[existing]
    ws = wb.create_sheet(LOG_SHEET)
    ws.sheet_state = "veryHidden"
    _write_row(ws, 1, list(HEADER))
    return ws


def _write_row(ws: Worksheet, row: int, values: list[Any]) -> None:
    for col, value in enumerate(values, start=1):
        cell = ws.cell(row=row, column=col)
        cell.value = value
        if isinstance(value, str):
            cell.data_type = "s"  # never let a logged "=..." become a formula


def _rows_for(ctx: WriteContext) -> list[LedgerEntry]:
    entries = ctx.ledger or [LedgerEntry(None, "", has_values=False)]
    cap = settings.CHANGE_LOG_MAX_CELLS_PER_OP
    cells = [e for e in entries if e.has_values]
    notes = [e for e in entries if not e.has_values]
    if len(cells) > cap:
        extra = len(cells) - cap
        notes.append(LedgerEntry(cells[cap].sheet, f"(+{extra} more cells)", has_values=False))
        cells = cells[:cap]
    return notes + cells


def append(wb: Workbook, ctx: WriteContext, now: datetime.datetime | None = None) -> None:
    """Append this call's entries to the workbook's change log (no-op when disabled)."""
    if not settings.CHANGE_LOG_ENABLED:
        return
    ws = _log_sheet(wb)
    ts = (now or datetime.datetime.now(datetime.timezone.utc)).isoformat(timespec="seconds")
    op_id = uuid.uuid4().hex[:12]
    reason = _clean(ctx.reason, MAX_REASON_CHARS) if ctx.reason else ""
    row = ws.max_row + 1
    for entry in _rows_for(ctx):
        before = encode(entry.before) if entry.has_values else ""
        after = encode(entry.after) if entry.has_values else ""
        _write_row(
            ws,
            row,
            [ts, op_id, ctx.tool, _clean(entry.sheet or "", 255), _clean(entry.ref, 255), before, after, reason],
        )
        row += 1
    excess = (ws.max_row - 1) - settings.CHANGE_LOG_MAX_ROWS
    if excess > 0:
        ws.delete_rows(2, excess)


def _matches(entry: dict[str, Any], sheet: str | None, cell: str | None) -> bool:
    if sheet is not None and entry.get("sheet") != sheet:
        return False
    if cell is None:
        return True
    ref = str(entry.get("ref") or "")
    return ref.upper() == cell or cell_in_range(cell, ref)


def read(path: str, sheet: str | None = None, cell: str | None = None, limit: int = 50) -> dict[str, Any]:
    """Change-log entries, newest first, optionally filtered by sheet and/or cell."""
    if not 1 <= limit <= 1000:
        raise ValueError(f"limit must be between 1 and 1000, got {limit}")
    cell = normalize_cell(cell) if cell else None
    if not Path(path).is_file():
        raise WorkbookNotFoundError(f"Workbook not found: '{path}'.")
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except PermissionError:
        raise
    except (zipfile.BadZipFile, InvalidFileException) as exc:
        raise InvalidWorkbookError(f"File '{path}' is not a valid Excel workbook: {exc}") from exc
    try:
        name = find_sheet_name(wb.sheetnames)
        if name is None:
            return {"enabled_in_file": False, "entries": [], "truncated": False, "note": REF_NOTE}
        rows = list(wb[name].iter_rows(min_row=2, values_only=True))
    finally:
        wb.close()

    entries: list[dict[str, Any]] = []
    truncated = False
    for raw in reversed(rows):
        values = list(raw) + [None] * (len(HEADER) - len(raw))
        entry = dict(zip(HEADER, values[: len(HEADER)]))
        if not entry.get("ts_utc"):
            continue
        entry["before"] = _decode(entry["before"])
        entry["after"] = _decode(entry["after"])
        if not _matches(entry, sheet, cell):
            continue
        if len(entries) == limit:
            truncated = True
            break
        entries.append(entry)
    return {"enabled_in_file": True, "entries": entries, "truncated": truncated, "note": REF_NOTE}
