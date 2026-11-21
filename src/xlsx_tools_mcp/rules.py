"""Workbook rules: declarative checks evaluated after every value-changing write.

Rules live in a sidecar file `<workbook>.rules.json`. A write that turns a passing
rule into a failing one (or makes a failing rule worse) is a *regression*; with
`on_violation="rollback"` the write is undone. Rules that already failed before the
write and did not get worse are reported as `pre_existing` and never block, so one
broken rule cannot lock a workbook against the very write that would fix it.
"""

from __future__ import annotations

import contextlib
import json
import math
import os
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable

import openpyxl

from . import inject, ledger
from .errors import InvalidRulesError, SheetNotFoundError
from .io.transform import _validate_safe_query, frame_from_rows
from .locking import safe_replace
from .recalc import EXCEL_ERROR_VALUES
from .refs import bounds, parse_cell_ref, parse_range_ref

RULES_VERSION = 1
MAX_RULES = 100
MAX_RULES_FILE_BYTES = 1024 * 1024
MAX_DETAIL_ITEMS = 10
_ID_RE = re.compile(r"^[A-Za-z0-9_.-]{1,64}$")
_OPS: dict[str, Callable[[Any, Any], bool]] = {
    "==": lambda a, b: a == b,
    "!=": lambda a, b: a != b,
    "<": lambda a, b: a < b,
    "<=": lambda a, b: a <= b,
    ">": lambda a, b: a > b,
    ">=": lambda a, b: a >= b,
}
_COMMON_KEYS = {"id", "type", "description"}
_SCHEMA: dict[str, tuple[set[str], set[str]]] = {
    "cells_equal": ({"left", "right"}, {"tolerance"}),
    "cell_compare": ({"cell", "op"}, {"value", "cell_ref"}),
    "column_check": ({"sheet", "expr"}, {"cell_range"}),
    "column_unique": ({"sheet", "column"}, {"cell_range"}),
    "not_empty": ({"range"}, set()),
    "no_new_errors": (set(), set()),
}
POLICIES_VIOLATION = ("rollback", "warn")
POLICIES_UNVERIFIABLE = ("warn", "rollback")


@dataclass
class Ruleset:
    rules: list[dict[str, Any]]
    on_violation: str = "rollback"
    on_unverifiable: str = "warn"
    version: int = RULES_VERSION

    @property
    def needs_snapshot(self) -> bool:
        return self.on_violation == "rollback" or self.on_unverifiable == "rollback"

    def to_json(self) -> dict[str, Any]:
        return {
            "version": self.version,
            "on_violation": self.on_violation,
            "on_unverifiable": self.on_unverifiable,
            "rules": self.rules,
        }


def sidecar_path(path: str) -> str:
    return f"{path}.rules.json"


# ── Validation ────────────────────────────────────────────────────────────────


def _is_number(v: Any) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool) and math.isfinite(v)


def _check_ref(rule_id: str, key: str, value: Any, kind: str, warnings: list[str], sheets: list[str] | None) -> None:
    if not isinstance(value, str) or not value.strip():
        raise InvalidRulesError(f"Rule '{rule_id}': '{key}' must be a non-empty reference string")
    try:
        sheet, _ = parse_cell_ref(value) if kind == "cell" else parse_range_ref(value)
    except ValueError as exc:
        raise InvalidRulesError(f"Rule '{rule_id}': {exc}") from exc
    _check_sheet(rule_id, sheet, warnings, sheets)


def _check_sheet(rule_id: str, sheet: str | None, warnings: list[str], sheets: list[str] | None) -> None:
    if ledger.is_reserved(sheet):
        raise InvalidRulesError(f"Rule '{rule_id}': sheet '{sheet}' is reserved")
    if sheets is not None and sheet not in sheets:
        warnings.append(f"Rule '{rule_id}': sheet '{sheet}' does not exist yet")


def _check_cells_equal(rule_id: str, raw: dict[str, Any], warnings: list[str], sheets: list[str] | None) -> None:
    _check_ref(rule_id, "left", raw["left"], "cell", warnings, sheets)
    _check_ref(rule_id, "right", raw["right"], "cell", warnings, sheets)
    tol = raw.get("tolerance", 0)
    if not _is_number(tol) or tol < 0:
        raise InvalidRulesError(f"Rule '{rule_id}': tolerance must be a number >= 0")


def _check_cell_compare(rule_id: str, raw: dict[str, Any], warnings: list[str], sheets: list[str] | None) -> None:
    _check_ref(rule_id, "cell", raw["cell"], "cell", warnings, sheets)
    if raw["op"] not in _OPS:
        raise InvalidRulesError(f"Rule '{rule_id}': op must be one of {list(_OPS)}")
    if ("value" in raw) == ("cell_ref" in raw):
        raise InvalidRulesError(f"Rule '{rule_id}': give exactly one of 'value' or 'cell_ref'")
    if "cell_ref" in raw:
        _check_ref(rule_id, "cell_ref", raw["cell_ref"], "cell", warnings, sheets)
    elif not (_is_number(raw["value"]) or isinstance(raw["value"], (str, bool))):
        raise InvalidRulesError(f"Rule '{rule_id}': value must be a number, string or boolean")


def _check_column_rule(rule_id: str, raw: dict[str, Any], warnings: list[str], sheets: list[str] | None) -> None:
    if not isinstance(raw["sheet"], str) or not raw["sheet"]:
        raise InvalidRulesError(f"Rule '{rule_id}': sheet must be a non-empty string")
    _check_sheet(rule_id, raw["sheet"], warnings, sheets)
    if "cell_range" in raw:
        try:
            range_sheet, _ = parse_range_ref(str(raw["cell_range"]), require_sheet=False)
        except ValueError as exc:
            raise InvalidRulesError(f"Rule '{rule_id}': {exc}") from exc
        if range_sheet:
            raise InvalidRulesError(f"Rule '{rule_id}': cell_range must not include a sheet name")
    if raw["type"] == "column_unique":
        if not isinstance(raw["column"], str) or not raw["column"]:
            raise InvalidRulesError(f"Rule '{rule_id}': column must be a non-empty string")
        return
    expr = raw["expr"]
    if not isinstance(expr, str) or not expr.strip() or len(expr) > 1000:
        raise InvalidRulesError(f"Rule '{rule_id}': expr must be a non-empty string up to 1000 chars")
    try:
        _validate_safe_query(expr)
    except ValueError as exc:
        raise InvalidRulesError(f"Rule '{rule_id}': {exc}") from exc


def _check_not_empty(rule_id: str, raw: dict[str, Any], warnings: list[str], sheets: list[str] | None) -> None:
    _check_ref(rule_id, "range", raw["range"], "range", warnings, sheets)


def _check_nothing(rule_id: str, raw: dict[str, Any], warnings: list[str], sheets: list[str] | None) -> None:
    return None


_TYPE_CHECKS: dict[str, Callable[[str, dict[str, Any], list[str], list[str] | None], None]] = {
    "cells_equal": _check_cells_equal,
    "cell_compare": _check_cell_compare,
    "column_check": _check_column_rule,
    "column_unique": _check_column_rule,
    "not_empty": _check_not_empty,
    "no_new_errors": _check_nothing,
}


def _validate_rule(raw: Any, seen: set[str], warnings: list[str], sheets: list[str] | None) -> dict[str, Any]:
    if not isinstance(raw, dict):
        raise InvalidRulesError("Each rule must be an object")
    rule_id = raw.get("id")
    if not isinstance(rule_id, str) or not _ID_RE.match(rule_id):
        raise InvalidRulesError(f"Rule id {rule_id!r} must match {_ID_RE.pattern}")
    if rule_id in seen:
        raise InvalidRulesError(f"Duplicate rule id '{rule_id}'")
    seen.add(rule_id)
    rtype = raw.get("type")
    if rtype not in _SCHEMA:
        raise InvalidRulesError(f"Rule '{rule_id}': type must be one of {sorted(_SCHEMA)}")
    required, optional = _SCHEMA[rtype]
    missing = required - set(raw)
    if missing:
        raise InvalidRulesError(f"Rule '{rule_id}': missing keys {sorted(missing)}")
    unknown = set(raw) - required - optional - _COMMON_KEYS
    if unknown:
        raise InvalidRulesError(f"Rule '{rule_id}': unknown keys {sorted(unknown)}")
    if "description" in raw and not isinstance(raw["description"], str):
        raise InvalidRulesError(f"Rule '{rule_id}': description must be a string")
    _TYPE_CHECKS[rtype](rule_id, raw, warnings, sheets)
    return dict(raw)


def validate(
    rules: Any,
    on_violation: Any = "rollback",
    on_unverifiable: Any = "warn",
    sheets: list[str] | None = None,
) -> tuple[Ruleset, list[str]]:
    """Validate a rule list; returns the ruleset and non-fatal warnings."""
    if on_violation not in POLICIES_VIOLATION:
        raise InvalidRulesError(f"on_violation must be one of {POLICIES_VIOLATION}")
    if on_unverifiable not in POLICIES_UNVERIFIABLE:
        raise InvalidRulesError(f"on_unverifiable must be one of {POLICIES_UNVERIFIABLE}")
    if not isinstance(rules, list):
        raise InvalidRulesError("rules must be a list")
    if len(rules) > MAX_RULES:
        raise InvalidRulesError(f"At most {MAX_RULES} rules are allowed, got {len(rules)}")
    seen: set[str] = set()
    warnings: list[str] = []
    clean = [_validate_rule(r, seen, warnings, sheets) for r in rules]
    return Ruleset(clean, on_violation, on_unverifiable), warnings


# ── Persistence ───────────────────────────────────────────────────────────────


def load_ruleset(path: str) -> Ruleset | None:
    """The workbook's ruleset, or None. A corrupt sidecar fails closed."""
    side = Path(sidecar_path(path))
    if not side.is_file():
        return None
    hint = f"Fix it with set_rules or delete '{side}'."
    try:
        if side.stat().st_size > MAX_RULES_FILE_BYTES:
            raise InvalidRulesError(f"Rules file '{side}' is larger than {MAX_RULES_FILE_BYTES} bytes. {hint}")
        data = json.loads(side.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        raise InvalidRulesError(f"Rules file '{side}' could not be read: {exc}. {hint}") from exc
    if not isinstance(data, dict) or data.get("version") != RULES_VERSION:
        raise InvalidRulesError(f"Rules file '{side}' must be an object with version {RULES_VERSION}. {hint}")
    try:
        ruleset, _ = validate(
            data.get("rules"), data.get("on_violation", "rollback"), data.get("on_unverifiable", "warn")
        )
    except InvalidRulesError as exc:
        raise InvalidRulesError(f"Rules file '{side}' is invalid: {exc}. {hint}") from exc
    return ruleset


def save_ruleset(path: str, ruleset: Ruleset | None) -> None:
    """Atomically write (or, for None / no rules, remove) the workbook's sidecar."""
    side = sidecar_path(path)
    if ruleset is None or not ruleset.rules:
        with contextlib.suppress(FileNotFoundError):
            os.remove(side)
        return
    tmp = f"{side}.tmp-{os.getpid()}"
    try:
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(ruleset.to_json(), fh, ensure_ascii=False, indent=2)
        safe_replace(tmp, side)
    finally:
        with contextlib.suppress(OSError):
            os.remove(tmp)


# ── Evaluation ────────────────────────────────────────────────────────────────


class _Values:
    """Cached cell values of one workbook, loaded once per evaluation."""

    def __init__(self, path: str, unknown: set[tuple[str, str]] | None = None):
        self.wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        self._rows: dict[str, list[tuple[Any, ...]]] = {}
        # Formula cells with no cached result: an empty value there means "not computed".
        self.unknown = unknown or set()

    def close(self) -> None:
        self.wb.close()

    def rows(self, sheet: str) -> list[tuple[Any, ...]]:
        if sheet not in self._rows:
            if sheet not in self.wb.sheetnames or ledger.is_reserved(sheet):
                raise SheetNotFoundError(f"Sheet '{sheet}' not found")
            ws = self.wb[sheet]
            ws.reset_dimensions()  # don't trust a stale <dimension> tag
            self._rows[sheet] = list(ws.iter_rows(values_only=True))
        return self._rows[sheet]

    def cell(self, ref: str) -> Any:
        sheet, coord = parse_cell_ref(ref)
        col, row, _, _ = bounds(coord)
        rows = self.rows(sheet)
        if row > len(rows):
            return None
        values = rows[row - 1]
        return values[col - 1] if col - 1 < len(values) else None

    def block(self, sheet: str, cell_range: str | None) -> list[list[Any]]:
        rows = [list(r) for r in self.rows(sheet)]
        if not cell_range:
            return rows
        min_col, min_row, max_col, max_row = bounds(cell_range)
        out = []
        for r in range(min_row, max_row + 1):
            src = rows[r - 1] if r - 1 < len(rows) else []
            out.append([src[c - 1] if c - 1 < len(src) else None for c in range(min_col, max_col + 1)])
        return out


NO_CACHED_VALUE = "no_cached_value"


class _Unverifiable(Exception):
    def __init__(self, message: str, cause: str | None = None):
        super().__init__(message)
        self.cause = cause


def make_result(
    rule: dict[str, Any], status: str, detail: str, metric: float | None, cause: str | None = None
) -> dict[str, Any]:
    out = {"id": rule["id"], "type": rule["type"], "status": status, "detail": detail, "metric": metric}
    if cause:
        out["cause"] = cause
    return out


def _is_error(v: Any) -> bool:
    return isinstance(v, str) and v in EXCEL_ERROR_VALUES


def _num(v: Any) -> float | None:
    return float(v) if _is_number(v) else None


def _need(vals: _Values, ref: str) -> Any:
    v = vals.cell(ref)
    if v is None or (isinstance(v, str) and v == ""):
        sheet, coord = parse_cell_ref(ref)
        if (sheet, coord) in vals.unknown:
            raise _Unverifiable(f"{ref} is a formula with no computed result", NO_CACHED_VALUE)
        raise _Unverifiable(f"{ref} has no value")
    return v


def _eval_cells_equal(rule: dict[str, Any], vals: _Values) -> dict[str, Any]:
    left = _need(vals, rule["left"])
    right = _need(vals, rule["right"])
    if _is_error(left) or _is_error(right):
        return make_result(rule, "fail", f"{rule['left']}={left!r}, {rule['right']}={right!r}", 1.0)
    ln, rn = _num(left), _num(right)
    if ln is not None and rn is not None:
        diff = abs(ln - rn)
        ok = diff <= float(rule.get("tolerance", 0))
        return make_result(rule, "pass" if ok else "fail", f"{rule['left']}={left}, {rule['right']}={right}, diff={diff:g}", diff)
    ok = left == right
    return make_result(rule, "pass" if ok else "fail", f"{rule['left']}={left!r}, {rule['right']}={right!r}", 0.0 if ok else 1.0)


def _eval_cell_compare(rule: dict[str, Any], vals: _Values) -> dict[str, Any]:
    left = _need(vals, rule["cell"])
    if "cell_ref" in rule:
        right = _need(vals, rule["cell_ref"])
    else:
        right = rule["value"]
    if _is_error(left) or _is_error(right):
        return make_result(rule, "fail", f"{rule['cell']}={left!r} is an error", 1.0)
    ln, rn = _num(left), _num(right)
    if ln is not None and rn is not None:
        left, right = ln, rn
    elif rule["op"] not in ("==", "!=") and type(left) is not type(right):
        raise _Unverifiable(f"cannot compare {type(left).__name__} with {type(right).__name__}")
    try:
        ok = _OPS[rule["op"]](left, right)
    except TypeError as exc:
        raise _Unverifiable(str(exc)) from exc
    return make_result(rule, "pass" if ok else "fail", f"{rule['cell']}={left!r} {rule['op']} {right!r}", 0.0 if ok else 1.0)


def _require_computed(vals: _Values, sheet: str, cell_range: str | None) -> None:
    """Formula cells without a computed result make a range check meaningless."""
    cells = [c for s, c in vals.unknown if s == sheet]
    if not cells:
        return
    if cell_range:
        min_col, min_row, max_col, max_row = bounds(cell_range)
        cells = [c for c in cells if min_col <= bounds(c)[0] <= max_col and min_row <= bounds(c)[1] <= max_row]
    if cells:
        raise _Unverifiable(
            f"{len(cells)} formula cell(s) in '{sheet}' have no computed result", NO_CACHED_VALUE
        )


def _frame(rule: dict[str, Any], vals: _Values):
    _require_computed(vals, rule["sheet"], rule.get("cell_range"))
    block = vals.block(rule["sheet"], rule.get("cell_range"))
    if not block:
        raise _Unverifiable(f"sheet '{rule['sheet']}' has no data")
    df = frame_from_rows(block)
    first_data_row = (bounds(rule["cell_range"])[1] if rule.get("cell_range") else 1) + 1
    df.index = range(first_data_row, first_data_row + len(df))
    return df.dropna(how="all")


def _eval_column_check(rule: dict[str, Any], vals: _Values) -> dict[str, Any]:
    df = _frame(rule, vals)
    try:
        bad = df.query(f"not ({rule['expr']})")
    except Exception as exc:  # pandas raises many types for bad columns/types
        raise _Unverifiable(f"expression could not be evaluated: {exc}") from exc
    rows = [int(i) for i in bad.index[:MAX_DETAIL_ITEMS]]
    detail = f"{len(bad)} row(s) violate '{rule['expr']}'" + (f"; rows {rows}" if rows else "")
    return make_result(rule, "fail" if len(bad) else "pass", detail, float(len(bad)))


def _eval_column_unique(rule: dict[str, Any], vals: _Values) -> dict[str, Any]:
    df = _frame(rule, vals)
    if rule["column"] not in df.columns:
        raise _Unverifiable(f"column '{rule['column']}' not found; available: {list(df.columns)}")
    col = df[rule["column"]]
    col = col[col.notna() & (col.astype(str).str.strip() != "")]
    dupes = col[col.duplicated(keep="first")]
    sample = [str(v) for v in dupes.unique()[:MAX_DETAIL_ITEMS]]
    detail = f"{len(dupes)} duplicate value(s)" + (f": {sample}" if sample else "")
    return make_result(rule, "fail" if len(dupes) else "pass", detail, float(len(dupes)))


def _eval_not_empty(rule: dict[str, Any], vals: _Values) -> dict[str, Any]:
    sheet, rng = parse_range_ref(rule["range"])
    _require_computed(vals, sheet, rng)
    block = vals.block(sheet, rng)
    empty = sum(1 for row in block for v in row if v is None or (isinstance(v, str) and not v.strip()))
    return make_result(rule, "fail" if empty else "pass", f"{empty} empty cell(s) in {rule['range']}", float(empty))


_EVALUATORS: dict[str, Callable[[dict[str, Any], _Values], dict[str, Any]]] = {
    "cells_equal": _eval_cells_equal,
    "cell_compare": _eval_cell_compare,
    "column_check": _eval_column_check,
    "column_unique": _eval_column_unique,
    "not_empty": _eval_not_empty,
}


def formula_state(path: str) -> tuple[bool, set[tuple[str, str]], set[tuple[str, str]]]:
    """(has formulas, formula cells holding an error, formula cells with no cached value)."""
    values = inject.read_cached_values(path, include_empty=True)
    has_formulas = any(values.values())
    errors = {(s, c) for s, cells in values.items() for c, v in cells.items() if v is not None and v[0] == "e"}
    unknown = {(s, c) for s, cells in values.items() for c, v in cells.items() if v is None}
    return has_formulas, errors, unknown


@dataclass
class Evaluation:
    results: list[dict[str, Any]] = field(default_factory=list)
    errors: set[tuple[str, str]] = field(default_factory=set)
    # Formula cells with no cached result (e.g. saved by openpyxl without recalculation).
    unknown: set[tuple[str, str]] = field(default_factory=set)


def evaluate(
    path: str,
    ruleset: Ruleset,
    baseline: Evaluation | None = None,
    values_current: bool = True,
    structural: bool = False,
) -> Evaluation:
    """Evaluate every rule against the workbook's cached values.

    `baseline` is the evaluation taken before the write (used by `no_new_errors`).
    `values_current=False` (recalculation did not run on a workbook with formulas)
    marks every rule unverifiable: formula results are missing or stale.
    `structural=True` (rows/columns shifted) compares error counts, not positions.
    """
    has_formulas, errors_now, unknown_now = formula_state(path)
    out = Evaluation(errors=errors_now, unknown=unknown_now)
    if not values_current and has_formulas:
        for rule in ruleset.rules:
            out.results.append(make_result(
                rule, "unverifiable", "formula results are not current (recalculation did not run)", None, NO_CACHED_VALUE
            ))
        return out
    try:
        vals = _Values(path, unknown_now)
    except Exception as exc:  # any unreadable workbook makes every rule unverifiable
        for rule in ruleset.rules:
            out.results.append(make_result(rule, "unverifiable", f"workbook could not be read: {exc}", None))
        return out
    try:
        for rule in ruleset.rules:
            out.results.append(_evaluate_one(rule, vals, baseline, errors_now, structural))
    finally:
        vals.close()
    return out


def _new_errors(
    baseline: Evaluation | None, errors_now: set[tuple[str, str]], structural: bool
) -> tuple[int, list[tuple[str, str]]]:
    if baseline is None:  # evaluating the baseline itself: nothing is "new" yet
        return 0, []
    if structural:
        # Inserted/deleted rows or columns move existing errors to new coordinates.
        allowed = len(baseline.errors) + len(baseline.unknown)
        count = max(0, len(errors_now) - allowed)
        return count, sorted(errors_now - baseline.errors)[:count]
    # Cells whose result was unknown before the write cannot be called new errors.
    new = sorted(errors_now - baseline.errors - baseline.unknown)
    return len(new), new


def _evaluate_one(
    rule: dict[str, Any],
    vals: _Values,
    baseline: Evaluation | None,
    errors_now: set[tuple[str, str]],
    structural: bool,
) -> dict[str, Any]:
    if rule["type"] == "no_new_errors":
        count, new = _new_errors(baseline, errors_now, structural)
        sample = [f"{s}!{c}" for s, c in new[:MAX_DETAIL_ITEMS]]
        detail = f"{count} new formula error(s)" + (f": {sample}" if sample else "")
        return make_result(rule, "fail" if count else "pass", detail, float(count))
    try:
        return _EVALUATORS[rule["type"]](rule, vals)
    except _Unverifiable as exc:
        return make_result(rule, "unverifiable", str(exc), None, exc.cause)
    except Exception as exc:  # a rule that cannot be computed is unverifiable, never a crash
        return make_result(rule, "unverifiable", str(exc), None)


@dataclass
class Classification:
    regressions: list[str] = field(default_factory=list)
    pre_existing: list[str] = field(default_factory=list)
    unverifiable: list[str] = field(default_factory=list)
    # Failing now, but the state before the write could not be determined.
    unknown_before: list[str] = field(default_factory=list)


def classify(pre: list[dict[str, Any]], post: list[dict[str, Any]]) -> Classification:
    """Sort post-write results by what the write did to each rule.

    A rule is a regression when it passed (or could not be checked) before and fails
    now, or failed before and got worse. Only when the earlier check was impossible
    because formula results had never been computed (`cause == "no_cached_value"`) is
    a failure reported as `unknown_before` and treated like an unverifiable rule, so
    a workbook can never be locked into rolling back every write.
    """
    before = {r["id"]: r for r in pre}
    out = Classification()
    for r in post:
        if r["status"] == "unverifiable":
            out.unverifiable.append(r["id"])
            continue
        if r["status"] != "fail":
            continue
        p = before.get(r["id"])
        if p is not None and p["status"] == "unverifiable" and p.get("cause") == NO_CACHED_VALUE:
            out.unknown_before.append(r["id"])
        elif p is None or p["status"] != "fail":
            # Passing, or unverifiable for any other reason (empty input, missing
            # sheet): conservatively treated as passing before.
            out.regressions.append(r["id"])
        elif (r["metric"] or 0) > (p["metric"] or 0):
            out.regressions.append(r["id"])
        else:
            out.pre_existing.append(r["id"])
    return out


def status_of(results: list[dict[str, Any]]) -> str:
    if any(r["status"] == "fail" for r in results):
        return "violated"
    if any(r["status"] == "unverifiable" for r in results):
        return "unverifiable"
    return "passed"
