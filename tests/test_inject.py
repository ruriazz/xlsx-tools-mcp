import re
import zipfile
from pathlib import Path
from types import SimpleNamespace

import openpyxl
import pytest
from openpyxl.worksheet.formula import ArrayFormula

from xlsx_tools_mcp import inject, ooxml, recalc
from xlsx_tools_mcp.errors import InvalidWorkbookError
from xlsx_tools_mcp.io import writer

from .helpers import add_part, file_hash, replace_part, sheet_xml, use_fake_libreoffice

SHEET = "xl/worksheets/sheet1.xml"


def _with_value(xml: str, coord: str, t: str | None, v: str) -> str:
    t_attr = f' t="{t}"' if t else ""
    pattern = rf'<c r="{coord}"( t="[^"]*")?><f([^>]*)>(.*?)</f><v\s*/?>(?:</v>)?</c>'
    new, n = re.subn(pattern, lambda m: f'<c r="{coord}"{t_attr}><f{m.group(2)}>{m.group(3)}</f><v>{v}</v></c>', xml)
    assert n == 1, coord
    return new


@pytest.fixture
def target(tmp_path):
    path = tmp_path / "target.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws["A1"] = 5
    ws["B1"] = "=A1*2"
    ws["C1"] = "=1/0"
    ws["D1"] = '="x"'
    ws["E1"] = "=TRUE()"
    ws["F1"] = "=A1"
    ws["G1"] = "#N/A"  # literal text, not an error
    ws["H1"] = "=A1&A1"
    ws["J1"] = ArrayFormula("J1:J2", "=A1:A2*2")
    ws["J2"] = 0
    other = wb.create_sheet("Other")
    other["A1"] = "=1+1"
    wb.save(path)
    return path


@pytest.fixture
def source(target, tmp_path):
    """A LibreOffice-like export of `target` with cached values for some formulas."""
    path = tmp_path / "source.xlsx"
    path.write_bytes(target.read_bytes())
    xml = sheet_xml(path)
    xml = _with_value(xml, "B1", None, "10")
    xml = _with_value(xml, "C1", "e", "#DIV/0!")
    xml = _with_value(xml, "D1", "str", "x")
    xml = _with_value(xml, "E1", "b", "1")
    xml = _with_value(xml, "H1", "s", "0")  # shared string index 0
    xml = _with_value(xml, "J1", None, "10")
    xml = xml.replace('<c r="J2" t="n"><v>0</v></c>', '<c r="J2" t="n"><v>7</v></c>')
    replace_part(path, SHEET, xml)
    ns = ooxml.NS_MAIN
    add_part(path, "xl/sharedStrings.xml", f'<sst xmlns="{ns}"><si><r><t>5</t></r><r><t>5</t></r></si></sst>')
    with zipfile.ZipFile(path) as zf:
        rels = zf.read("xl/_rels/workbook.xml.rels").decode()
    rel = ('<Relationship Id="rIdSST" Target="sharedStrings.xml" Type="http://schemas.openxmlformats.org/'
           'officeDocument/2006/relationships/sharedStrings"/>')
    replace_part(path, "xl/_rels/workbook.xml.rels", rels.replace("</Relationships>", rel + "</Relationships>"))
    return path


def test_read_cached_values_types(source):
    values = inject.read_cached_values(source)["Data"]
    assert values["B1"] == ("n", "10")
    assert values["C1"] == ("e", "#DIV/0!")
    assert values["D1"] == ("str", "x")
    assert values["E1"] == ("b", "1")
    assert values["H1"] == ("str", "55")
    assert values["J2"] == ("n", "7")  # array spill cell
    assert "G1" not in values and "A1" not in values  # constants are not read
    assert "F1" not in values  # formula without a cached value


def test_inject_writes_values_readable_by_openpyxl(target, source, tmp_path):
    out = tmp_path / "out.xlsx"
    stats = inject.inject_cached_values(target, inject.read_cached_values(source), out)
    wb = openpyxl.load_workbook(out, data_only=True)
    ws = wb["Data"]
    assert [ws[c].value for c in ("B1", "C1", "D1", "E1", "H1", "J2")] == [10, "#DIV/0!", "x", True, "55", 7]
    assert ws["F1"].value is None
    assert stats.missing == 2  # Data!F1 and Other!A1 have no cached value in the source
    assert stats.skipped_sheets == []
    assert stats.errors == [{"sheet": "Data", "cell": "C1", "error": "#DIV/0!"}]
    formulas = openpyxl.load_workbook(out)["Data"]
    assert formulas["B1"].value == "=A1*2"  # formulas untouched


def test_inject_only_changes_sheet_parts(target, source, tmp_path):
    out = tmp_path / "out.xlsx"
    inject.inject_cached_values(target, inject.read_cached_values(source), out)
    with zipfile.ZipFile(target) as a, zipfile.ZipFile(out) as b:
        assert a.namelist() == b.namelist()
        for name in a.namelist():
            if name != SHEET:
                assert a.read(name) == b.read(name), name


def test_sheet_absent_from_source_is_reported_skipped(target, source, tmp_path):
    values = inject.read_cached_values(source)
    del values["Other"]
    stats = inject.inject_cached_values(target, values, tmp_path / "o.xlsx")
    assert stats.skipped_sheets == ["Other"]


def test_literal_error_text_is_not_reported(target, source, tmp_path):
    stats = inject.inject_cached_values(target, inject.read_cached_values(source), tmp_path / "o.xlsx")
    assert all(e["cell"] != "G1" for e in stats.errors)


def test_inject_preserves_mc_ignorable_namespaces(target, source, tmp_path):
    xml = sheet_xml(target)
    xml = re.sub(
        r"<worksheet ([^>]*)>",
        lambda m: (
            '<worksheet xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
            'xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" '
            'mc:Ignorable="x14ac" ' + m.group(1) + ">"
        ),
        xml,
        count=1,
    )
    replace_part(target, SHEET, xml)
    out = tmp_path / "out.xlsx"
    inject.inject_cached_values(target, inject.read_cached_values(source), out)
    root = ooxml.parse_xml(sheet_xml(out).encode()).getroot()
    assert root.nsmap["x14ac"] == "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
    assert root.get("{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable") == "x14ac"
    assert "<c r=\"B1\"><f>A1*2</f><v>10</v></c>" in sheet_xml(out)


def test_inject_in_place_is_atomic_and_leaves_no_temp(target, source):
    inject.inject_cached_values(target, inject.read_cached_values(source), target)
    assert openpyxl.load_workbook(target, data_only=True)["Data"]["B1"].value == 10
    assert not list(Path(target).parent.glob("*.inj-*"))


def test_sheet_with_cell_missing_ref_is_left_untouched(target, source, tmp_path):
    xml = sheet_xml(target).replace('<c r="A1" t="n">', '<c t="n">', 1)
    replace_part(target, SHEET, xml)
    out = tmp_path / "out.xlsx"
    stats = inject.inject_cached_values(target, inject.read_cached_values(source), out)
    assert "Data" in stats.skipped_sheets
    assert sheet_xml(out) == xml


def test_doctype_in_sheet_is_rejected(target, source):
    xml = sheet_xml(target)
    replace_part(target, SHEET, xml.replace("<worksheet", '<!DOCTYPE x [<!ENTITY a "b">]><worksheet', 1))
    with pytest.raises(InvalidWorkbookError, match="DOCTYPE"):
        inject.inject_cached_values(target, inject.read_cached_values(source), target)


def test_to_python():
    assert inject.to_python(("n", "10")) == 10
    assert inject.to_python(("n", "2.5")) == 2.5
    assert inject.to_python(("n", "1E+3")) == 1000.0
    assert inject.to_python(("b", "0")) is False
    assert inject.to_python(("e", "#REF!")) == "#REF!"
    assert inject.to_python(None) is None


# ── recalculate() in inject mode ─────────────────────────────────────────────


def _book(tmp_path) -> str:
    path = str(tmp_path / "calc.xlsx")
    writer.create_workbook(path, ["Data"])
    wb = openpyxl.load_workbook(path)
    ws = wb["Data"]
    ws["A1"] = 4
    ws["B1"] = "=A1*3"
    ws["C1"] = "=A1/0"
    ws["D1"] = "#N/A"
    wb.save(path)
    return path


def test_recalculate_inject_mode(tmp_path, monkeypatch):
    path = _book(tmp_path)
    use_fake_libreoffice(monkeypatch)
    result = recalc.recalculate(path, mode="inject")
    assert result.success is True
    assert result.mode == "inject"
    assert result.errors_found == [{"sheet": "Data", "cell": "C1", "error": "#DIV/0!"}]
    assert "inject: 2 values, 0 missing" in result.message
    assert openpyxl.load_workbook(path, data_only=True)["Data"]["B1"].value == 12


def test_recalculate_inject_failure_leaves_file_unchanged(tmp_path, monkeypatch):
    path = _book(tmp_path)
    before = file_hash(path)

    def bad_run(cmd, capture_output=False, timeout=None, **kwargs):
        outdir = Path(cmd[cmd.index("--outdir") + 1])
        (outdir / f"{Path(cmd[-1]).stem}.xlsx").write_bytes(b"not a zip")
        return SimpleNamespace(returncode=0, stderr=b"")

    monkeypatch.setattr(recalc, "find_soffice", lambda: "/fake/soffice")
    monkeypatch.setattr(recalc.subprocess, "run", bad_run)
    result = recalc.recalculate(path, mode="inject")
    assert result.success is False
    assert "roundtrip" in result.message
    assert file_hash(path) == before


def test_recalculate_rejects_unknown_mode(tmp_path):
    with pytest.raises(ValueError):
        recalc.recalculate(_book(tmp_path), mode="nope")


def test_write_cells_uses_inject_mode_by_default(tmp_path, monkeypatch):
    path = _book(tmp_path)
    use_fake_libreoffice(monkeypatch)
    result = writer.write_cells(path, "Data", [{"cell": "A1", "value": 10}])
    assert result["recalc_mode"] == "inject"
    assert result["recalculated"] is True
    assert openpyxl.load_workbook(path, data_only=True)["Data"]["B1"].value == 30
