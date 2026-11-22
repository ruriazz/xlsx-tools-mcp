"""End-to-end run of the v0.4 tools through the MCP server, including tool registration."""

import asyncio

import openpyxl

from xlsx_tools_mcp import server, settings

from .helpers import use_fake_libreoffice

NEW_TOOLS = {"set_rules", "check_rules", "export_review_copy", "inspect_for_sharing", "sanitize_copy", "get_change_log"}


def test_server_registers_35_tools():
    names = {t.name for t in asyncio.run(server.mcp.list_tools())}
    assert len(names) == 35
    assert NEW_TOOLS <= names


def test_new_tools_lifecycle(tmp_path, monkeypatch):
    use_fake_libreoffice(monkeypatch)
    monkeypatch.setattr(settings, "CHANGE_LOG_ENABLED", True)
    data = tmp_path / "Data Dir"
    data.mkdir()
    monkeypatch.setattr(settings, "ALLOWED_DIRS", [str(data.resolve())])
    path = str(data / "Budget.xlsx")

    server.create_workbook(path, ["Plan"])
    server.write_cells("Plan", [
        {"cell": "A1", "value": 100}, {"cell": "A2", "value": 50}, {"cell": "A3", "formula": "=SUM(A1:A2)"},
        {"cell": "B3", "value": 150},
    ], path=path)

    res = server.set_rules([{"id": "total", "type": "cells_equal", "left": "Plan!A3", "right": "Plan!B3"}], path=path)
    assert res["evaluation"]["status"] == "passed"

    blocked = server.write_cells("Plan", [{"cell": "A1", "value": 1}], reason="typo", path=path)
    assert blocked["rolled_back"] is True
    ok = server.write_cells("Plan", [{"cell": "A1", "value": 90}, {"cell": "A2", "value": 60}], reason="rebalance", path=path)
    assert ok["rolled_back"] is False
    assert ok["ripple"]["changed_count"] == 0  # A3 stays 150
    assert server.check_rules(path=path)["evaluation"]["status"] == "passed"

    log = server.get_change_log(cell="A1", path=path)["entries"]
    assert log[0]["reason"] == "rebalance"
    assert all(e["reason"] != "typo" for e in log)  # rolled-back write left no entry

    review = server.export_review_copy(path=path)
    assert review["summary"]["value"] == 2
    assert openpyxl.load_workbook(review["output_path"], data_only=True)["Plan"]["A3"].value == 150

    findings = {f["category"] for f in server.inspect_for_sharing(path=path)["findings"]}
    assert "change_log" in findings
    clean = server.sanitize_copy(str(data / "Budget.shared.xlsx"), path=path)
    assert not {f["category"] for f in clean["remaining_findings"]} & {"change_log", "comments"}

    rec = server.recalculate_workbook(path=path)
    assert rec["recalc_mode"] == "inject" and rec["recalculated"] is True
