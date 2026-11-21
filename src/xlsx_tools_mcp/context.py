"""Per-call description of a write, shared by the writer, guard and change log."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class LedgerEntry:
    sheet: str | None
    ref: str
    before: Any = None
    after: Any = None
    has_values: bool = True  # False for summary rows (structural edits, formatting)


@dataclass
class WriteContext:
    tool: str
    # (sheet, "B2") cells written directly by this call — excluded from ripple reports.
    touched: set[tuple[str, str]] = field(default_factory=set)
    # True when cell coordinates shift (row/column insert or delete, sheet delete/rename).
    structural: bool = False
    # True when cell values or formulas may change, i.e. rules and ripple apply.
    changes_values: bool = False
    # True when the edit cannot change any cell value or formula (styles, tables,
    # charts, new sheets): cached formula results are carried over across the save.
    preserves_values: bool = False
    ledger: list[LedgerEntry] = field(default_factory=list)
    reason: str | None = None

    def record(self, sheet: str, coord: str, before: Any, after: Any) -> None:
        self.touched.add((sheet, coord))
        self.ledger.append(LedgerEntry(sheet, coord, before, after))

    def note(self, sheet: str | None, ref: str) -> None:
        self.ledger.append(LedgerEntry(sheet, ref, has_values=False))
