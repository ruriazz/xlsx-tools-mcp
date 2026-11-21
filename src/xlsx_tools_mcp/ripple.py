"""Ripple report: which formula results changed as a side effect of a write.

Only formula (and array spill) cells can change value on recalculation, so the
snapshot is just their cached results, read straight from the package XML. That
keeps error values (calamine reads them as "") and avoids reading constant cells.
"""

from __future__ import annotations

import math
from typing import Any

from . import inject, ledger, settings

SAMPLE_SIZE = 20

Snapshot = dict[tuple[str, str], "inject.CachedValue | None"]


def snapshot(path: str) -> Snapshot | None:
    """Formula results keyed by (sheet, cell), or None when over XLSX_MCP_RIPPLE_MAX_CELLS."""
    values = inject.read_cached_values(path, include_empty=True)
    total = sum(len(cells) for name, cells in values.items() if not ledger.is_reserved(name))
    if total > settings.RIPPLE_MAX_CELLS:
        return None
    return {
        (sheet, coord): val
        for sheet, cells in values.items()
        if not ledger.is_reserved(sheet)
        for coord, val in cells.items()
    }


def same_value(a: Any, b: Any) -> bool:
    if isinstance(a, bool) or isinstance(b, bool):
        return a == b
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return math.isclose(a, b, rel_tol=1e-9, abs_tol=1e-12)
    return a == b


def diff(before: Snapshot, after: Snapshot, touched: set[tuple[str, str]]) -> dict[str, Any]:
    changed: list[dict[str, Any]] = []
    unknown_before = 0
    for key, new in after.items():
        if key in touched or key not in before or new is None:
            continue
        old = before[key]
        if old is None:
            unknown_before += 1
            continue
        a, b = inject.to_python(old), inject.to_python(new)
        if same_value(a, b):
            continue
        item: dict[str, Any] = {"sheet": key[0], "cell": key[1], "before": a, "after": b}
        numeric = all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in (a, b))
        if numeric:
            item["delta"] = b - a
        changed.append(item)

    def order(item: dict[str, Any]) -> tuple[int, float]:
        return (0, -abs(item["delta"])) if "delta" in item else (1, 0.0)

    changed.sort(key=order)
    return {
        "changed_count": len(changed),
        "unknown_before": unknown_before,
        "sample": changed[:SAMPLE_SIZE],
        "truncated": len(changed) > SAMPLE_SIZE,
    }
