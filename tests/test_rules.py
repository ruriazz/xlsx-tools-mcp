import json
from pathlib import Path

import openpyxl
import pytest

from xlsx_tools_mcp import rules, server
from xlsx_tools_mcp.errors import InvalidRulesError
from xlsx_tools_mcp.io import writer

from .helpers import file_hash, use_fake_libreoffice

BALANCE = {"id": "balance", "type": "cells_equal", "left": "Neraca!B3", "right": "Neraca!C3"}


@pytest.fixture
def book(tmp_path, monkeypatch):
    use_fake_libreoffice(monkeypatch)
    path = str(tmp_path / "model.xlsx")
    writer.create_workbook(path, ["Neraca", "Data"])
    writer.write_cells(path, "Neraca", [
        {"cell": "B1", "value": 60}, {"cell": "B2", "value": 40}, {"cell": "B3", "formula": "=SUM(B1:B2)"},
        {"cell": "C1", "value": 100}, {"cell": "C3", "formula": "=SUM(C1:C2)"},
    ], backup=False)
    writer.append_rows(path, "Data", [["Invoice", "Qty"], ["INV-1", 5], ["INV-2", 3]], backup=False)
    return path


def _snaps(path: str) -> list[Path]:
    return list(Path(path).parent.glob("*.snap-*"))


def test_set_rules_saves_sidecar_and_evaluates(book):
    res = server.set_rules([BALANCE], path=book)
    assert res["evaluation"]["status"] == "passed"
    stored = json.loads(Path(rules.sidecar_path(book)).read_text())
    assert stored["rules"][0]["id"] == "balance"
    assert server.check_rules(path=book)["evaluation"]["results"][0]["status"] == "pass"


def test_regression_rolls_back_byte_identical(book):
    server.set_rules([BALANCE], path=book)
    before = file_hash(book)
    res = server.write_cells("Neraca", [{"cell": "B1", "value": 999}], path=book)
    assert res["rolled_back"] is True
    assert res["saved"] is False
    assert res["rules"]["regressions"] == ["balance"]
    assert file_hash(book) == before
    assert _snaps(book) == []


def test_passing_write_keeps_change(book):
    server.set_rules([BALANCE], path=book)
    res = server.write_cells("Neraca", [{"cell": "C2", "value": 0}], path=book)
    assert res["rolled_back"] is False
    assert res["rules"]["status"] == "passed"
    assert _snaps(book) == []


def test_pre_existing_failure_does_not_block(book):
    server.write_cells("Neraca", [{"cell": "B1", "value": 70}], path=book, backup=False)  # now unbalanced
    server.set_rules([BALANCE], path=book)
    res = server.write_cells("Data", [{"cell": "C1", "value": "note"}], path=book)
    assert res["rolled_back"] is False
    assert res["rules"]["pre_existing"] == ["balance"]


def test_worse_failure_counts_as_regression(book):
    server.write_cells("Data", [{"cell": "B2", "value": -1}], path=book, backup=False)  # 1 bad row, no rules yet
    server.set_rules([{"id": "qty", "type": "column_check", "sheet": "Data", "expr": "Qty >= 0"}], path=book)
    ok = server.write_cells("Data", [{"cell": "A4", "value": "INV-3"}, {"cell": "B4", "value": 1}], path=book)
    assert ok["rolled_back"] is False and ok["rules"]["pre_existing"] == ["qty"]
    res = server.write_cells("Data", [{"cell": "B3", "value": -2}], path=book)  # 2 bad rows: worse
    assert res["rolled_back"] is True
    assert res["rules"]["regressions"] == ["qty"]
    assert openpyxl.load_workbook(book)["Data"]["B3"].value == 3


def test_warn_policy_keeps_write(book):
    server.set_rules([BALANCE], on_violation="warn", path=book)
    res = server.write_cells("Neraca", [{"cell": "B1", "value": 999}], path=book)
    assert res["rolled_back"] is False
    assert res["saved"] is True
    assert res["rules"]["regressions"] == ["balance"]
    assert "Rule check" in res["message"]
    assert _snaps(book) == []


def test_corrupt_sidecar_fails_closed(book):
    Path(rules.sidecar_path(book)).write_text("{not json")
    before = file_hash(book)
    with pytest.raises(ValueError, match="could not be read"):
        server.write_cells("Neraca", [{"cell": "B1", "value": 1}], path=book)
    assert file_hash(book) == before


@pytest.mark.parametrize("expr", ["__import__('os')", "@x > 1", "Qty.sum() > 0", "abs(Qty) > 0"])
def test_unsafe_expression_rejected(book, expr):
    with pytest.raises(ValueError):
        server.set_rules([{"id": "x", "type": "column_check", "sheet": "Data", "expr": expr}], path=book)


@pytest.mark.parametrize("rule", [
    {"id": "bad id!", "type": "no_new_errors"},
    {"id": "a", "type": "unknown"},
    {"id": "a", "type": "cells_equal", "left": "B3", "right": "Neraca!C3"},
    {"id": "a", "type": "cell_compare", "cell": "Neraca!B3", "op": "<", "value": 1, "cell_ref": "Neraca!C3"},
    {"id": "a", "type": "cell_compare", "cell": "Neraca!B3", "op": "~", "value": 1},
    {"id": "a", "type": "not_empty", "range": "Data!A:A"},
    {"id": "a", "type": "no_new_errors", "extra": 1},
    {"id": "a", "type": "cells_equal", "left": "_xlsx_mcp_log!A1", "right": "Neraca!C3"},
])
def test_invalid_rules_rejected(rule):
    with pytest.raises(InvalidRulesError):
        rules.validate([rule])


def test_duplicate_ids_rejected():
    with pytest.raises(InvalidRulesError, match="Duplicate"):
        rules.validate([{"id": "a", "type": "no_new_errors"}, {"id": "a", "type": "no_new_errors"}])


def test_empty_rules_remove_sidecar(book):
    server.set_rules([BALANCE], path=book)
    server.set_rules([], path=book)
    assert not Path(rules.sidecar_path(book)).exists()
    assert server.check_rules(path=book) == {"rules": None, "evaluation": None}


def test_each_rule_type_evaluates(book):
    server.write_cells("Neraca", [{"cell": "D1", "formula": "=1/0"}], path=book, backup=False)
    res = server.set_rules([
        {"id": "cmp", "type": "cell_compare", "cell": "Neraca!B3", "op": "<=", "value": 100},
        {"id": "cmpref", "type": "cell_compare", "cell": "Neraca!B1", "op": ">", "cell_ref": "Neraca!B2"},
        {"id": "uniq", "type": "column_unique", "sheet": "Data", "column": "Invoice"},
        {"id": "filled", "type": "not_empty", "range": "Data!A2:A3"},
        {"id": "gap", "type": "not_empty", "range": "Data!A2:A5"},
        {"id": "qty", "type": "column_check", "sheet": "Data", "expr": "Qty > 4"},
        {"id": "err", "type": "cells_equal", "left": "Neraca!D1", "right": "Neraca!B3"},
        {"id": "missingcol", "type": "column_unique", "sheet": "Data", "column": "Nope"},
        {"id": "errors", "type": "no_new_errors"},
    ], path=book)
    status = {r["id"]: r["status"] for r in res["evaluation"]["results"]}
    assert status == {
        "cmp": "pass", "cmpref": "pass", "uniq": "pass", "filled": "pass", "gap": "fail",
        "qty": "fail", "err": "fail", "missingcol": "unverifiable", "errors": "pass",
    }


def test_no_new_errors_rolls_back(book):
    server.set_rules([{"id": "errors", "type": "no_new_errors"}], path=book)
    res = server.write_cells("Neraca", [{"cell": "E1", "formula": "=B1/0"}], path=book)
    assert res["rolled_back"] is True
    assert openpyxl.load_workbook(book)["Neraca"]["E1"].value is None


def test_unverifiable_when_libreoffice_missing(book, monkeypatch):
    server.set_rules([BALANCE], path=book)
    monkeypatch.setattr("xlsx_tools_mcp.recalc.find_soffice", lambda: None)
    res = server.write_cells("Neraca", [{"cell": "B1", "value": 999}], path=book)
    assert res["rules"]["status"] == "unverifiable"
    assert res["rolled_back"] is False  # on_unverifiable defaults to warn


def test_unverifiable_rollback_policy(book, monkeypatch):
    server.set_rules([BALANCE], on_unverifiable="rollback", path=book)
    before = file_hash(book)
    monkeypatch.setattr("xlsx_tools_mcp.recalc.find_soffice", lambda: None)
    res = server.write_cells("Neraca", [{"cell": "B1", "value": 999}], path=book)
    assert res["rolled_back"] is True
    assert file_hash(book) == before


def test_style_change_skips_rules(book):
    server.set_rules([BALANCE], path=book)
    res = server.set_cell_style("Neraca", "A1", {"bold": True}, path=book)
    assert res["rules"]["status"] == "skipped"
    assert "rolled_back" not in res


def test_exception_after_save_restores_file_and_removes_snapshot(book, monkeypatch):
    server.set_rules([BALANCE], path=book)
    before = file_hash(book)

    def boom(path, *a, **k):
        raise RuntimeError("simulated crash")

    monkeypatch.setattr(writer, "recalculate", boom)
    with pytest.raises(RuntimeError):
        writer.write_cells(book, "Neraca", [{"cell": "B1", "value": 1}])
    assert file_hash(book) == before  # rules never fail open
    assert _snaps(book) == []


def test_no_lockout_after_values_were_emptied(book):
    """Review finding #1: empty cached values must not make every later write roll back."""
    server.write_cells("Neraca", [{"cell": "D1", "formula": "=1/0"}], path=book, backup=False)
    server.set_rules([{"id": "errors", "type": "no_new_errors"}, BALANCE], path=book)
    server.write_cells("Neraca", [{"cell": "E1", "value": 1}], recalculate=False, path=book, backup=False)
    first = server.write_cells("Data", [{"cell": "C1", "value": "x"}], path=book)
    assert first["rolled_back"] is False
    second = server.write_cells("Data", [{"cell": "C2", "value": "y"}], path=book)
    assert second["rolled_back"] is False
    assert second["rules"]["status"] == "passed"  # the existing #DIV/0! is not a new error
    assert second["rules"]["regressions"] == []


def test_style_change_keeps_cached_values(book):
    import openpyxl as _o
    server.set_cell_style("Neraca", "A1", {"bold": True}, path=book)
    assert _o.load_workbook(book, data_only=True)["Neraca"]["B3"].value == 100


def test_structural_edit_does_not_count_moved_errors(book):
    server.write_cells("Neraca", [{"cell": "D5", "formula": "=1/0"}], path=book, backup=False)
    server.set_rules([{"id": "errors", "type": "no_new_errors"}], path=book)
    res = server.insert_rows("Neraca", 1, count=2, path=book)
    assert res["rolled_back"] is False
    assert res["rules"]["status"] == "passed"


def test_merge_is_checked_by_rules(book):
    server.set_rules([BALANCE], path=book)
    res = server.merge_cells("Neraca", "B1:B2", path=book)
    assert res["rules"]["status"] != "skipped"


def test_rollback_failure_keeps_snapshot(book, monkeypatch):
    server.set_rules([BALANCE], path=book)
    from xlsx_tools_mcp import guard
    from xlsx_tools_mcp.errors import FileInUseError

    def locked(src, dst):
        raise FileInUseError("locked")

    monkeypatch.setattr(guard, "safe_replace", locked)
    res = server.write_cells("Neraca", [{"cell": "B1", "value": 999}], path=book)
    assert res["rolled_back"] is False
    assert "rollback_error" in res
    assert Path(res["snapshot_path"]).exists()


def test_no_rules_response_unchanged(book):
    res = server.write_cells("Neraca", [{"cell": "C2", "value": 0}], path=book)
    assert "rules" not in res and "rolled_back" not in res


def test_partial_snapshot_copy_never_overwrites_workbook(book, monkeypatch):
    """Re-review N1: a failed snapshot copy must leave the workbook untouched."""
    server.set_rules([BALANCE], path=book)
    before = file_hash(book)
    from xlsx_tools_mcp import guard

    def partial_copy(src, dst):
        Path(dst).write_bytes(Path(src).read_bytes()[:100])
        raise OSError("disk full")

    monkeypatch.setattr(guard.shutil, "copy2", partial_copy)
    with pytest.raises(OSError, match="disk full"):
        server.write_cells("Neraca", [{"cell": "B1", "value": 999}], path=book)
    assert file_hash(book) == before
    assert _snaps(book) == []


def test_empty_input_cell_rule_still_rolls_back(book):
    """Re-review N2: unverifiable-before for a plain empty cell is not 'unknown_before'."""
    server.set_rules([{"id": "cap", "type": "cell_compare", "cell": "Data!F1", "op": "<=", "value": 1000}], path=book)
    res = server.write_cells("Data", [{"cell": "F1", "value": 5000}], path=book)
    assert res["rolled_back"] is True
    assert res["rules"]["regressions"] == ["cap"]


def test_unique_rule_on_empty_sheet_still_rolls_back(book):
    server.create_sheet("Inv", path=book)
    server.set_rules([{"id": "uniq", "type": "column_unique", "sheet": "Inv", "column": "No"}], path=book)
    res = server.append_rows("Inv", [["No"], ["A"], ["A"]], path=book)
    assert res["rolled_back"] is True


def test_failed_restore_after_crash_reports_snapshot(book, monkeypatch):
    """Re-review N3: if undoing a crashed write fails, the kept copy is named."""
    server.set_rules([BALANCE], path=book)
    from xlsx_tools_mcp import guard
    from xlsx_tools_mcp.errors import FileInUseError

    def boom(path, *a, **k):
        raise RuntimeError("simulated crash")

    def locked(src, dst):
        raise FileInUseError("locked")

    monkeypatch.setattr(writer, "recalculate", boom)
    monkeypatch.setattr(guard, "safe_replace", locked)
    with pytest.raises(Exception, match="kept at") as info:
        writer.write_cells(book, "Neraca", [{"cell": "B1", "value": 1}])
    kept = str(info.value).split("kept at '")[1].rstrip("'")
    assert Path(kept).exists()


def test_column_rule_on_uncomputed_formulas_does_not_lock_out(book):
    """Re-review L1: duplicate keys computed by formulas, cached results wiped."""
    server.write_cells("Data", [{"cell": "A4", "formula": "=A2"}], path=book, backup=False)  # duplicate INV-1
    server.set_rules([{"id": "uniq", "type": "column_unique", "sheet": "Data", "column": "Invoice"}], path=book)
    server.write_cells("Data", [{"cell": "D1", "value": 1}], recalculate=False, path=book, backup=False)
    res = server.write_cells("Data", [{"cell": "D2", "value": 2}], path=book)
    assert res["rolled_back"] is False
    assert res["rules"]["unknown_before"] == ["uniq"]
