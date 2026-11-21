"""Test helpers: a fake LibreOffice that really computes simple formulas.

The fake `subprocess.run` reads the input workbook, evaluates each formula with a
tiny evaluator (cell refs, Sheet!refs, SUM(range), + - * /, numbers, strings),
and writes a LibreOffice-like export with cached values into `--outdir`.
Division by zero yields #DIV/0!, unknown names yield #NAME?.
"""

from __future__ import annotations

import hashlib
import re
import zipfile
from pathlib import Path
from types import SimpleNamespace
from typing import Any

import openpyxl
from openpyxl.utils import get_column_letter, range_boundaries

from xlsx_tools_mcp import inject, recalc

_SUM = re.compile(r"SUM\((?:(\w+)!)?([A-Z]+\d+):([A-Z]+\d+)\)")
_REF = re.compile(r"(?:(\w+)!)?\$?([A-Z]{1,3})\$?(\d+)")


def _to_cached(value: Any) -> inject.CachedValue:
    if isinstance(value, bool):
        return ("b", "1" if value else "0")
    if isinstance(value, (int, float)):
        if isinstance(value, float) and value.is_integer():
            return ("n", str(int(value)))
        return ("n", repr(value))
    return ("str", str(value))


def compute_values(path: str) -> dict[str, dict[str, inject.CachedValue]]:
    wb = openpyxl.load_workbook(path, data_only=False)
    consts: dict[tuple[str, str], Any] = {}
    formulas: dict[tuple[str, str], str] = {}
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and v.startswith("="):
                    formulas[(ws.title, c.coordinate)] = v[1:]
                elif v is not None:
                    consts[(ws.title, c.coordinate)] = v
    wb.close()
    results: dict[tuple[str, str], Any] = {}

    def get(sheet: str, coord: str) -> Any:
        key = (sheet, coord)
        if key in results:
            return results[key]
        return consts.get(key, 0)

    for _ in range(10):
        changed = False
        for (sheet, coord), text in formulas.items():
            val = _evaluate(text, sheet, get)
            if results.get((sheet, coord), object()) != val:
                results[(sheet, coord)] = val
                changed = True
        if not changed:
            break
    out: dict[str, dict[str, inject.CachedValue]] = {}
    for (sheet, coord), val in results.items():
        out.setdefault(sheet, {})[coord] = val if isinstance(val, tuple) else _to_cached(val)
    return out


def _evaluate(text: str, sheet: str, get) -> Any:
    def sum_repl(m: re.Match) -> str:
        s = m.group(1) or sheet
        min_col, min_row, max_col, max_row = range_boundaries(f"{m.group(2)}:{m.group(3)}")
        total = 0.0
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                v = get(s, f"{get_column_letter(c)}{r}")
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    total += v
        return repr(total)

    def ref_repl(m: re.Match) -> str:
        v = get(m.group(1) or sheet, f"{m.group(2)}{m.group(3)}")
        if isinstance(v, tuple):  # propagated error
            raise _ErrorValue(v[1])
        return repr(v)

    try:
        expr = _SUM.sub(sum_repl, text)
        expr = _REF.sub(ref_repl, expr)
        value = eval(expr, {"__builtins__": {}}, {})  # noqa: S307 — test-only evaluator
    except ZeroDivisionError:
        return ("e", "#DIV/0!")
    except _ErrorValue as exc:
        return ("e", str(exc))
    except Exception:
        return ("e", "#NAME?")
    return value


class _ErrorValue(Exception):
    pass


def fake_soffice_run(cmd, capture_output=False, timeout=None, **kwargs):
    outdir = Path(cmd[cmd.index("--outdir") + 1])
    src = Path(cmd[-1])
    out = outdir / f"{src.stem}.xlsx"
    inject.inject_cached_values(src, compute_values(str(src)), out)
    return SimpleNamespace(returncode=0, stderr=b"")


def use_fake_libreoffice(monkeypatch) -> None:
    monkeypatch.setattr(recalc, "find_soffice", lambda: "/fake/soffice")
    monkeypatch.setattr(recalc.subprocess, "run", fake_soffice_run)


def file_hash(path: str | Path) -> str:
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def sheet_xml(path: str | Path, part: str = "xl/worksheets/sheet1.xml") -> str:
    with zipfile.ZipFile(path) as zf:
        return zf.read(part).decode("utf-8")


def replace_part(path: str | Path, part: str, data: bytes | str) -> None:
    """Rewrite one zip entry in place (keeps entry order)."""
    path = Path(path)
    if isinstance(data, str):
        data = data.encode("utf-8")
    tmp = path.with_suffix(path.suffix + ".tmp")
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            zout.writestr(info, data if info.filename == part else zin.read(info))
    tmp.replace(path)


def add_part(path: str | Path, part: str, data: bytes | str) -> None:
    path = Path(path)
    if isinstance(data, str):
        data = data.encode("utf-8")
    with zipfile.ZipFile(path, "a", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(part, data)
