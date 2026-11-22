"""The MCP tool surface: names, annotations, schema descriptions and argument validation.

Other tests call the tool functions directly; these go through the MCP layer
(list_tools / call_tool), which is what an agent actually sees and sends.
"""

from __future__ import annotations

import asyncio
import re

import openpyxl
import pytest
from mcp.server.mcpserver.exceptions import ToolError

from xlsx_tools_mcp import server

# Renaming or dropping a tool breaks clients' saved permissions and prompts.
EXPECTED_TOOLS = {
    "list_configured_files", "list_sheets", "get_workbook_info", "read_sheet", "get_cell",
    "search_workbook", "profile_sheet", "query_sheet", "aggregate_sheet",
    "create_workbook", "write_cells", "append_rows", "clear_range", "recalculate_workbook",
    "restore_backup", "create_sheet", "copy_sheet", "rename_sheet", "delete_sheet",
    "insert_rows", "delete_rows", "insert_columns", "delete_columns",
    "merge_cells", "unmerge_cells", "set_cell_style", "autofit_columns", "create_table", "create_chart",
    "set_rules", "check_rules", "export_review_copy", "inspect_for_sharing", "sanitize_copy", "get_change_log",
}
READ_ONLY = {
    "list_configured_files", "list_sheets", "get_workbook_info", "read_sheet", "get_cell",
    "search_workbook", "profile_sheet", "query_sheet", "aggregate_sheet",
    "check_rules", "inspect_for_sharing", "get_change_log",
}
NON_DESTRUCTIVE_WRITES = {
    "append_rows", "create_sheet", "copy_sheet", "insert_rows", "insert_columns",
    "unmerge_cells", "set_cell_style", "autofit_columns", "create_table", "create_chart",
}


def _tools():
    return {t.name: t for t in asyncio.run(server.mcp.list_tools())}


def _call(name: str, args: dict):
    return asyncio.run(server.mcp.call_tool(name, args))


def test_tool_names_are_unchanged():
    assert set(_tools()) == EXPECTED_TOOLS


def test_every_tool_has_title_and_annotations():
    for name, tool in _tools().items():
        assert tool.title, name
        ann = tool.annotations
        assert ann is not None and ann.open_world_hint is False, name
        if name in READ_ONLY:
            assert ann.read_only_hint is True, name
        else:
            assert not ann.read_only_hint, name
            # destructive_hint left unset means "may be destructive" (MCP default).
            expected = False if name in NON_DESTRUCTIVE_WRITES else None
            assert ann.destructive_hint is expected, name


def test_descriptions_are_compact_and_params_documented():
    for name, tool in _tools().items():
        assert tool.description and "Args:" not in tool.description, name
        assert not re.search(r"\n {4}\S", tool.description), f"{name}: description not dedented"
        for param, schema in tool.input_schema["properties"].items():
            assert schema.get("description"), f"{name}.{param} has no description"


@pytest.mark.parametrize(
    "tool, param, values",
    [
        ("read_sheet", "format", ["array", "records", "markdown"]),
        ("create_chart", "chart_type", ["bar", "line", "pie", "scatter"]),
        ("set_rules", "on_violation", ["rollback", "warn"]),
    ],
)
def test_fixed_choices_are_enums(tool, param, values):
    assert _tools()[tool].input_schema["properties"][param]["enum"] == values


def test_instructions_name_only_real_tools():
    named = set(re.findall(r"\b[a-z]+(?:_[a-z]+)+\b", server.INSTRUCTIONS))
    candidates = {n for n in named if n.split("_")[0] in {
        "list", "get", "read", "search", "profile", "query", "aggregate", "write",
        "append", "insert", "delete", "clear", "set", "create",
    }}
    assert candidates and candidates <= EXPECTED_TOOLS, candidates - EXPECTED_TOOLS


@pytest.fixture
def book(tmp_path):
    path = tmp_path / "sales.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sales"
    for row in (["Region", "Amount"], ["East", 10], ["West", 5], ["East", 7]):
        ws.append(row)
    wb.save(path)
    return str(path)


def test_call_through_mcp_layer(book):
    res = _call("read_sheet", {"sheet": "Sales", "format": "records", "path": book})
    assert not res.is_error
    assert res.structured_content["rows"][0] == {"Region": "East", "Amount": 10}

    res = _call("aggregate_sheet", {"sheet": "Sales", "group_by": ["Region"], "agg": {"Amount": "sum"}, "path": book})
    assert not res.is_error
    assert {"Region": "East", "Amount": 17} in res.structured_content["records"]


def test_invalid_enum_rejected_by_schema(book):
    with pytest.raises(ToolError, match="format"):
        _call("read_sheet", {"sheet": "Sales", "format": "csv", "path": book})


def test_chart_type_stays_case_insensitive(book):
    res = _call("create_chart", {"sheet": "Sales", "chart_type": "Bar", "data_range": "B1:B4", "backup": False, "path": book})
    assert not res.is_error, res.content
    assert res.structured_content["saved"] is True
