import zipfile

import openpyxl
import pytest
from openpyxl.comments import Comment
from openpyxl.packaging.relationship import Relationship
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.external_link.external import ExternalBook, ExternalLink, ExternalSheetNames

from xlsx_tools_mcp import server
from xlsx_tools_mcp.io import sharing

from .helpers import add_part, file_hash, replace_part


def _external(wb):
    book = ExternalBook(sheetNames=ExternalSheetNames(sheetName=["Other"]), id="rId1")
    link = ExternalLink(externalBook=book)
    link.file_link = Relationship(type="externalLinkPath", Target="other.xlsx", TargetMode="External", Id="rId1")
    wb._external_links.append(link)


@pytest.fixture
def leaky(tmp_path):
    path = tmp_path / "leaky.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Report"
    ws["A1"] = "total"
    ws["B1"] = 100
    ws["A1"].comment = Comment("internal only", "Alice")
    ws["C1"] = "secret"
    ws["C1"].number_format = ";;;"
    ws["D1"] = "white"
    ws["D1"].font = Font(color="FFFFFFFF")
    ws["E1"] = "same"
    ws["E1"].font = Font(color="FF112233")
    ws["E1"].fill = PatternFill("solid", start_color="FF112233")
    ws["A5"] = "hidden row data"
    ws.row_dimensions[5].hidden = True
    ws["F1"] = "=[1]Other!A1"
    hidden = wb.create_sheet("Payroll")
    hidden["A1"] = 5000
    hidden.sheet_state = "veryHidden"
    _external(wb)
    wb.properties.creator = "Alice"
    wb.save(path)
    with zipfile.ZipFile(path) as zf:
        app = zf.read("docProps/app.xml").decode()
    replace_part(path, "docProps/app.xml", app.replace("</Properties>", "<Company>ACME Corp</Company></Properties>"))
    add_part(path, "xl/pivotCache/pivotCacheRecords1.xml", "<pivotCacheRecords/>")
    return path


def _categories(result):
    return {f["category"]: f for f in result["findings"]}


def test_inspect_finds_every_category(leaky):
    before = file_hash(leaky)
    res = server.inspect_for_sharing(path=str(leaky))
    cats = _categories(res)
    assert file_hash(leaky) == before
    assert set(cats) >= {
        "hidden_sheets", "comments", "document_properties", "external_links",
        "pivot_cache", "hidden_formatting", "hidden_rows_columns",
    }
    assert cats["hidden_sheets"]["severity"] == "high"
    assert "Payroll" in cats["hidden_sheets"]["detail"]
    assert "Alice" in cats["comments"]["detail"]
    assert "ACME Corp" in cats["document_properties"]["detail"]
    assert "other.xlsx" in cats["external_links"]["detail"]
    assert cats["hidden_formatting"]["count"] == 3
    assert "row 5" in cats["hidden_rows_columns"]["detail"]
    assert res["findings"][0]["severity"] == "high"  # sorted by severity


def test_clean_workbook_has_no_findings(tmp_path):
    path = tmp_path / "clean.xlsx"
    wb = openpyxl.Workbook()
    wb.active["A1"] = 1
    wb.properties.creator = None
    wb.save(path)
    assert server.inspect_for_sharing(path=str(path))["findings"] == []


def test_sanitize_default_removes_comments_and_properties(leaky, tmp_path):
    before = file_hash(leaky)
    out = tmp_path / "clean.xlsx"
    res = server.sanitize_copy(str(out), path=str(leaky))
    assert file_hash(leaky) == before
    remaining = {f["category"] for f in res["remaining_findings"]}
    assert "comments" not in remaining
    assert "document_properties" not in remaining
    assert "hidden_sheets" in remaining  # not requested
    wb = openpyxl.load_workbook(out)
    assert wb["Report"]["A1"].comment is None
    with zipfile.ZipFile(out) as zf:
        core = zf.read("docProps/core.xml").decode()
    assert "Alice" not in core and "creator" not in core


def test_sanitize_hidden_sheet_and_external_links(leaky, tmp_path):
    out = tmp_path / "clean.xlsx"
    res = server.sanitize_copy(str(out), remove=["hidden_sheets", "external_links"], path=str(leaky))
    remaining = {f["category"] for f in res["remaining_findings"]}
    assert "hidden_sheets" not in remaining
    assert "external_links" not in remaining
    wb = openpyxl.load_workbook(out)
    assert "Payroll" not in wb.sheetnames
    assert wb["Report"]["F1"].value is None  # no cached value existed → cleared
    assert any("no cached value" in w for w in res["warnings"])
    with zipfile.ZipFile(out) as zf:
        assert not [n for n in zf.namelist() if n.startswith("xl/externalLinks/")]


def test_referenced_hidden_sheet_is_not_removed(leaky, tmp_path):
    wb = openpyxl.load_workbook(leaky)
    wb["Report"]["G1"] = "=Payroll!A1*2"
    wb.save(leaky)
    with pytest.raises(ValueError, match="still referenced by Report!G1"):
        server.sanitize_copy(str(tmp_path / "o.xlsx"), remove=["hidden_sheets"], path=str(leaky))
    assert not (tmp_path / "o.xlsx").exists()


def test_quoted_reference_detected():
    assert sharing._references_sheet("='My Sheet'!A1", "My Sheet")
    assert sharing._references_sheet("=SUM(payroll!A1:A3)", "Payroll")
    assert not sharing._references_sheet("=OtherPayroll!A1", "Payroll")


@pytest.mark.parametrize("remove", [["everything"], ["pivot_cache"]])
def test_unknown_remove_option_rejected(leaky, tmp_path, remove):
    with pytest.raises(ValueError, match="Unsupported"):
        server.sanitize_copy(str(tmp_path / "o.xlsx"), remove=remove, path=str(leaky))


def test_sanitize_output_cannot_be_source(leaky):
    with pytest.raises(ValueError, match="must differ"):
        server.sanitize_copy(str(leaky), path=str(leaky))


def test_xlsm_output_requires_xlsm_source(leaky, tmp_path):
    with pytest.raises(ValueError, match=".xlsm"):
        server.sanitize_copy(str(tmp_path / "o.xlsm"), path=str(leaky))


def test_external_link_cached_formula_text_stays_text(leaky, tmp_path):
    from .helpers import sheet_xml
    xml = sheet_xml(leaky)
    xml = xml.replace("<f>[1]Other!A1</f><v></v>", "<f>[1]Other!A1</f><v>=HYPERLINK(\"http://x\")</v>")
    xml = xml.replace('<c r="F1">', '<c r="F1" t="str">')
    replace_part(leaky, "xl/worksheets/sheet1.xml", xml)
    out = tmp_path / "clean.xlsx"
    server.sanitize_copy(str(out), remove=["external_links"], path=str(leaky))
    cell = openpyxl.load_workbook(out)["Report"]["F1"]
    assert cell.value == '=HYPERLINK("http://x")'
    assert cell.data_type == "s"


def test_output_named_like_preloaded_alias_is_not_redirected(leaky, tmp_path, monkeypatch):
    from xlsx_tools_mcp import settings
    other = tmp_path / "configured.xlsx"
    other.write_bytes(leaky.read_bytes())
    before = file_hash(other)
    monkeypatch.setattr(settings, "CONFIGURED_FILES", {"configured.xlsx": str(other)})
    target_dir = tmp_path / "share"
    target_dir.mkdir()
    out = target_dir / "configured.xlsx"
    server.sanitize_copy(str(out), path=str(leaky))
    assert out.exists()
    assert file_hash(other) == before


def test_custom_properties_reported_and_cleared(tmp_path):
    from openpyxl.packaging.custom import StringProperty
    path = tmp_path / "c.xlsx"
    wb = openpyxl.Workbook()
    wb.custom_doc_props.append(StringProperty(name="Client", value="ACME"))
    wb.save(path)
    cats = _categories(server.inspect_for_sharing(path=str(path)))
    assert "custom:Client" in cats["document_properties"]["detail"]
    res = server.sanitize_copy(str(tmp_path / "o.xlsx"), path=str(path))
    assert "document_properties" not in {f["category"] for f in res["remaining_findings"]}
