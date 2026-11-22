"""Parsing of A1-style references with an optional sheet prefix.

Accepts "B40", "Sheet1!B40", "'My Sheet'!B40", "'It''s'!A1:C5" and "$B$40".
"""

from __future__ import annotations

import re

from openpyxl.utils import column_index_from_string, get_column_letter, range_boundaries

MAX_ROW = 1_048_576
MAX_COL = 16_384

_CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})$")


def split_sheet(ref: str) -> tuple[str | None, str]:
    """Split "Sheet!A1" into ("Sheet", "A1"). Quoted names may contain "!" and "''"."""
    ref = ref.strip()
    if ref.startswith("'"):
        i = 1
        name: list[str] = []
        while i < len(ref):
            ch = ref[i]
            if ch == "'":
                if i + 1 < len(ref) and ref[i + 1] == "'":
                    name.append("'")
                    i += 2
                    continue
                break
            name.append(ch)
            i += 1
        else:
            raise ValueError(f"Unterminated sheet quote in reference '{ref}'")
        rest = ref[i + 1:]
        if not rest.startswith("!"):
            raise ValueError(f"Expected '!' after quoted sheet name in '{ref}'")
        return "".join(name), rest[1:]
    if "!" in ref:
        sheet, rest = ref.rsplit("!", 1)
        return sheet, rest
    return None, ref


def normalize_cell(cell: str) -> str:
    m = _CELL_RE.match(cell.strip())
    if not m:
        raise ValueError(f"Invalid cell coordinate '{cell}'. Use an A1-style reference like 'B5'.")
    col, row = m.group(1).upper(), int(m.group(2))
    if column_index_from_string(col) > MAX_COL or row > MAX_ROW:
        raise ValueError(f"Cell coordinate '{cell}' is outside Excel's grid")
    return f"{col}{row}"


def parse_cell_ref(ref: str, require_sheet: bool = True) -> tuple[str | None, str]:
    sheet, cell = split_sheet(ref)
    if require_sheet and not sheet:
        raise ValueError(f"Reference '{ref}' needs a sheet name, e.g. 'Sheet1!B5'")
    return sheet, normalize_cell(cell)


def bounds(cell_range: str) -> tuple[int, int, int, int]:
    """(min_col, min_row, max_col, max_row) of a bounded A1 range or single cell."""
    text = cell_range.replace("$", "").strip()
    try:
        min_col, min_row, max_col, max_row = range_boundaries(text)
    except (ValueError, TypeError) as exc:
        raise ValueError(f"Invalid range '{cell_range}'") from exc
    if None in (min_col, min_row, max_col, max_row):
        raise ValueError(f"Range '{cell_range}' must be bounded, e.g. 'A2:A500'")
    if max_col > MAX_COL or max_row > MAX_ROW:
        raise ValueError(f"Range '{cell_range}' is outside Excel's grid")
    return min_col, min_row, max_col, max_row


def parse_range_ref(ref: str, require_sheet: bool = True) -> tuple[str | None, str]:
    sheet, rng = split_sheet(ref)
    if require_sheet and not sheet:
        raise ValueError(f"Reference '{ref}' needs a sheet name, e.g. 'Sheet1!A2:A500'")
    min_col, min_row, max_col, max_row = bounds(rng)
    first = f"{get_column_letter(min_col)}{min_row}"
    last = f"{get_column_letter(max_col)}{max_row}"
    return sheet, first if first == last else f"{first}:{last}"


def cell_in_range(cell: str, cell_range: str) -> bool:
    """True when `cell` lies inside `cell_range` (row/column-only ranges allowed)."""
    try:
        min_col, min_row, max_col, max_row = range_boundaries(cell_range.replace("$", ""))
        col, row = range_boundaries(normalize_cell(cell))[:2]
    except (ValueError, TypeError):
        return False
    return (
        (min_col is None or col >= min_col)
        and (max_col is None or col <= max_col)
        and (min_row is None or row >= min_row)
        and (max_row is None or row <= max_row)
    )


def quote_sheet(name: str) -> str:
    return "'" + name.replace("'", "''") + "'"
