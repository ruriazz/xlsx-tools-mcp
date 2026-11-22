"""Copy cached formula results from one workbook package into another.

openpyxl never stores formula results: every save writes `<v></v>` for formula
cells. LibreOffice computes them, but saving its export over the user's file loses
features it doesn't round-trip (pivot tables, slicers, some charts and names).

This module reads the computed values out of a LibreOffice export (or any package
that has them) and writes only the `<v>` text and `t` type of formula cells into the
original package. Every other zip entry is copied unchanged.
"""

from __future__ import annotations

import contextlib
import os
import shutil
import zipfile
import zlib
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Union

from lxml import etree

from . import ooxml
from .errors import InjectionError, InvalidWorkbookError
from .locking import safe_replace
from .refs import bounds

# (type, text): type is "n" (number), "str", "b" (bool) or "e" (error).
CachedValue = tuple[str, str]
SheetValues = dict[str, Union[CachedValue, None]]

# Corrupt or unsupported zip content surfaces as any of these while reading entries.
_ZIP_ERRORS = (zipfile.BadZipFile, zlib.error, NotImplementedError, EOFError)

_C, _F, _V, _IS, _ROW, _SHEET_DATA = (
    ooxml.q("c"), ooxml.q("f"), ooxml.q("v"), ooxml.q("is"), ooxml.q("row"), ooxml.q("sheetData"),
)


@dataclass
class InjectStats:
    injected: int = 0
    missing: int = 0
    skipped_sheets: list[str] = field(default_factory=list)
    errors: list[dict[str, Any]] = field(default_factory=list)

    def summary(self) -> str:
        text = f"inject: {self.injected} values, {self.missing} missing"
        if self.skipped_sheets:
            text += f", skipped sheets: {', '.join(self.skipped_sheets)}"
        return text


def _array_bounds(f: etree._Element) -> tuple[int, int, int, int] | None:
    if f.get("t") != "array" or not f.get("ref"):
        return None
    try:
        return bounds(f.get("ref"))
    except ValueError:
        return None


def _coord_parts(coord: str) -> tuple[int, int]:
    min_col, min_row, _, _ = bounds(coord)
    return min_col, min_row


def _in_any(coord: str, ranges: list[tuple[int, int, int, int]]) -> bool:
    if not ranges:
        return False
    col, row = _coord_parts(coord)
    return any(r[0] <= col <= r[2] and r[1] <= row <= r[3] for r in ranges)


def _iter_cells(sheet_root: etree._Element):
    data = sheet_root.find(_SHEET_DATA)
    if data is None:
        return
    for row in data.iterchildren(_ROW):
        yield from row.iterchildren(_C)


def _cell_value(c: etree._Element, sst: list[str] | None) -> CachedValue | None:
    t = c.get("t") or "n"
    if t == "inlineStr":
        is_el = c.find(_IS)
        return None if is_el is None else ("str", ooxml.rich_text(is_el))
    v = c.find(_V)
    if v is None:
        return None
    text = v.text
    if t == "str":
        return ("str", text or "")
    if text is None or text.strip() == "":
        return None
    text = text.strip()
    if t == "n":
        return ("n", text)
    if t == "s":
        if sst is None:
            return None
        try:
            idx = int(text)
        except ValueError:
            return None
        return ("str", sst[idx]) if 0 <= idx < len(sst) else None
    if t == "b":
        return ("b", "1" if text.lower() in ("1", "true") else "0")
    if t == "e":
        return ("e", text)
    if t == "d":
        return ("str", text)
    return None


def read_cached_values(path: str | Path, include_empty: bool = False) -> dict[str, SheetValues]:
    """Cached results of formula cells (and array-formula spill cells), per sheet.

    With `include_empty`, formula cells lacking a cached value map to None so callers
    can tell "no value yet" apart from "not a formula".
    """
    out: dict[str, SheetValues] = {}
    try:
        with zipfile.ZipFile(path) as zf:
            sst_cache: list[list[str]] = []

            def sst() -> list[str]:
                if not sst_cache:
                    sst_cache.append(ooxml.shared_strings(zf))
                return sst_cache[0]

            for sp in ooxml.workbook_sheets(zf):
                root = ooxml.parse_part(zf, sp.part).getroot()
                values: SheetValues = {}
                arrays: list[tuple[int, int, int, int]] = []
                for c in _iter_cells(root):
                    r = c.get("r")
                    if not r:
                        continue
                    f = c.find(_F)
                    if f is not None:
                        ab = _array_bounds(f)
                        if ab:
                            arrays.append(ab)
                    elif not _in_any(r, arrays):
                        continue
                    needs_sst = (c.get("t") == "s")
                    val = _cell_value(c, sst() if needs_sst else None)
                    if val is not None or (include_empty and f is not None):
                        values[r.upper()] = val
                out[sp.name] = values
    except _ZIP_ERRORS as exc:
        raise InvalidWorkbookError(f"File '{path}' is not a valid Excel workbook: {exc}") from exc
    return out


def to_python(value: CachedValue | None) -> Any:
    """Convert a cached value to a plain Python value (errors stay as "#DIV/0!" strings)."""
    if value is None:
        return None
    t, text = value
    if t == "n":
        try:
            num = float(text)
        except ValueError:
            return text
        return int(num) if num.is_integer() and "." not in text and "e" not in text.lower() else num
    if t == "b":
        return text == "1"
    return text


def _error_entry(sheet: str, coord: str, error: str) -> dict[str, Any]:
    entry: dict[str, Any] = {"sheet": sheet, "cell": coord, "error": error}
    if error == "#NAME?":
        entry["hint"] = "may be a function LibreOffice does not support; Excel recalculates on open"
    return entry


def _apply(c: etree._Element, f: etree._Element | None, val: CachedValue) -> None:
    t, text = val
    is_el = c.find(_IS)
    if is_el is not None:
        c.remove(is_el)
    v = c.find(_V)
    if v is None:
        v = c.makeelement(_V, {})
        if f is not None:
            f.addnext(v)
        else:
            c.insert(0, v)
    v.text = text
    if t == "n":
        c.attrib.pop("t", None)
    else:
        c.set("t", t)


@dataclass
class _SheetResult:
    injected: int = 0
    missing: int = 0
    formula_count: int = 0
    errors: list[dict[str, Any]] = field(default_factory=list)
    missing_ref: bool = False


def _inject_sheet(root: etree._Element, sheet: str, sheet_vals: SheetValues | None) -> _SheetResult:
    res = _SheetResult()
    arrays: list[tuple[int, int, int, int]] = []
    for c in _iter_cells(root):
        r = c.get("r")
        if not r:
            # Positions would have to be inferred; the caller leaves this sheet untouched.
            res.missing_ref = True
            return res
        f = c.find(_F)
        if f is not None:
            res.formula_count += 1
            ab = _array_bounds(f)
            if ab:
                arrays.append(ab)
        elif not _in_any(r, arrays):
            continue
        val = sheet_vals.get(r.upper()) if sheet_vals is not None else None
        if val is None:
            if f is not None:
                res.missing += 1
            continue
        _apply(c, f, val)
        res.injected += 1
        if val[0] == "e":
            res.errors.append(_error_entry(sheet, r.upper(), val[1]))
    return res


def _same_path(a: str, b: str) -> bool:
    return os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))


def inject_cached_values(
    src: str | Path, values: dict[str, SheetValues], dst: str | Path
) -> InjectStats:
    """Write `values` into the formula cells of package `src`, saving the result to `dst`.

    Sheets are matched by name. Only cells that hold a formula in `src` (or sit inside
    an array formula's range) are touched. `dst` may equal `src`; the result is staged
    next to `dst` and swapped in atomically, so `dst` is never left half-written.
    """
    stats = InjectStats()
    replacements: dict[str, bytes] = {}
    try:
        with zipfile.ZipFile(src) as zf:
            for sp in ooxml.workbook_sheets(zf):
                tree = ooxml.parse_part(zf, sp.part)
                sheet_vals = values.get(sp.name)
                res = _inject_sheet(tree.getroot(), sp.name, sheet_vals)
                if res.missing_ref or (sheet_vals is None and res.formula_count):
                    stats.skipped_sheets.append(sp.name)
                if res.missing_ref:
                    continue
                stats.injected += res.injected
                stats.missing += res.missing
                stats.errors.extend(res.errors)
                if res.injected:
                    replacements[sp.part] = ooxml.serialize(tree)
    except _ZIP_ERRORS as exc:
        raise InvalidWorkbookError(f"File '{src}' is not a valid Excel workbook: {exc}") from exc

    src_s, dst_s = str(src), str(dst)
    if not replacements:
        if not _same_path(src_s, dst_s):
            shutil.copyfile(src_s, dst_s)
        return stats

    staged = f"{dst_s}.inj-{os.getpid()}"
    try:
        ooxml.rewrite_package(src_s, staged, replacements)
        safe_replace(staged, dst_s)
    except PermissionError:
        raise
    except (OSError, ValueError, *_ZIP_ERRORS) as exc:
        raise InjectionError(f"Could not write computed values into '{dst_s}': {exc}") from exc
    finally:
        with contextlib.suppress(OSError):
            os.remove(staged)
    return stats

