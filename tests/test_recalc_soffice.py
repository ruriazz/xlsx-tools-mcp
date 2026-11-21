"""Integration tests against a real LibreOffice. Run with `pytest -m soffice`.

Skipped when LibreOffice is not installed. These verify what the unit tests can
only simulate: the shape of LibreOffice's export (string/bool/error encodings) and
that inject mode keeps parts LibreOffice would not round-trip.
"""

import zipfile

import openpyxl
import pytest
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

from xlsx_tools_mcp import recalc
from xlsx_tools_mcp.io import writer

pytestmark = [
    pytest.mark.soffice,
    pytest.mark.skipif(recalc._orig_find_soffice() is None, reason="LibreOffice not installed"),
]


@pytest.fixture
def rich_book(tmp_path):
    path = str(tmp_path / "rich.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data Sheet"
    ws.append(["k", "v"])
    for i in range(1, 5):
        ws.append([f"r{i}", i * 10])
    ws["C2"] = "=B2*2"
    ws["C3"] = "=SUM(B2:B5)"
    ws["C4"] = '="x"&A2'
    ws["C5"] = "=1/0"
    ws["D2"] = "=B2>5"
    ws["E2"] = "#N/A"
    chart = BarChart()
    chart.add_data(Reference(ws, min_col=2, min_row=1, max_row=5), titles_from_data=True)
    ws.add_chart(chart, "G2")
    dv = DataValidation(type="whole", operator="between", formula1="0", formula2="100")
    ws.add_data_validation(dv)
    dv.add("B2:B5")
    ws.conditional_formatting.add(
        "B2:B5", CellIsRule(operator="greaterThan", formula=["20"], fill=PatternFill("solid", start_color="FFFF00"))
    )
    wb.defined_names["Total"] = DefinedName("Total", attr_text="'Data Sheet'!$C$3")
    hidden = wb.create_sheet("Hidden")
    hidden.sheet_state = "veryHidden"
    hidden["A1"] = "=1+1"
    wb.save(path)
    return path


def _parts(path):
    with zipfile.ZipFile(path) as zf:
        return {n: zf.read(n) for n in zf.namelist()}


def test_inject_computes_all_value_types(rich_book):
    result = writer.write_cells(rich_book, "Data Sheet", [{"cell": "B2", "value": 11}], backup=False)
    assert result["recalculated"] is True, result["message"]
    assert result["recalc_mode"] == "inject"
    ws = openpyxl.load_workbook(rich_book, data_only=True)["Data Sheet"]
    assert ws["C2"].value == 22
    assert ws["C3"].value == 101
    assert ws["C4"].value == "xr1"
    assert ws["C5"].value == "#DIV/0!"
    assert ws["D2"].value is True
    assert result["errors_found"] == [{"sheet": "Data Sheet", "cell": "C5", "error": "#DIV/0!"}]
    assert openpyxl.load_workbook(rich_book, data_only=True)["Hidden"]["A1"].value == 2


def test_inject_keeps_structure(rich_book):
    writer.write_cells(rich_book, "Data Sheet", [{"cell": "B2", "value": 11}], recalculate=False, backup=False)
    before = _parts(rich_book)
    assert recalc.recalculate(rich_book, mode="inject").success
    after = _parts(rich_book)
    assert set(before) == set(after)
    changed = {n for n in before if before[n] != after[n]}
    assert changed <= {"xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"}
    wb = openpyxl.load_workbook(rich_book)
    ws = wb["Data Sheet"]
    assert ws._charts and ws.data_validations.dataValidation and ws.conditional_formatting
    assert "Total" in wb.defined_names
    assert wb["Hidden"].sheet_state == "veryHidden"
