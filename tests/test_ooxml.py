import zipfile

import openpyxl
import pytest

from xlsx_tools_mcp import ooxml, settings
from xlsx_tools_mcp.errors import InvalidWorkbookError

from .helpers import replace_part


@pytest.fixture
def book(tmp_path):
    path = tmp_path / "book.xlsx"
    wb = openpyxl.Workbook()
    wb.active.title = "Data"
    wb["Data"]["A1"] = "hello"
    hidden = wb.create_sheet("Secret")
    hidden.sheet_state = "veryHidden"
    wb.create_sheet("My Sheet")
    wb.save(path)
    return path


def test_workbook_sheets_maps_names_states_and_parts(book):
    with zipfile.ZipFile(book) as zf:
        sheets = ooxml.workbook_sheets(zf)
    assert [(s.name, s.state) for s in sheets] == [("Data", "visible"), ("Secret", "veryHidden"), ("My Sheet", "visible")]
    assert all(s.part.startswith("xl/worksheets/") for s in sheets)


def test_absolute_relationship_targets_resolve(book):
    with zipfile.ZipFile(book) as zf:
        rels = zf.read("xl/_rels/workbook.xml.rels").decode()
    replace_part(book, "xl/_rels/workbook.xml.rels", rels.replace('Target="worksheets/', 'Target="/xl/worksheets/'))
    with zipfile.ZipFile(book) as zf:
        assert [s.name for s in ooxml.workbook_sheets(zf)] == ["Data", "Secret", "My Sheet"]


def test_parse_rejects_doctype():
    data = b'<?xml version="1.0"?><!DOCTYPE x [<!ENTITY a "aaaa">]><x>&a;</x>'
    with pytest.raises(InvalidWorkbookError, match="DOCTYPE"):
        ooxml.parse_xml(data)


def test_parse_does_not_expand_entities():
    data = b'<?xml version="1.0"?><!DOCTYPE x [<!ENTITY a "boom">]><x>&a;</x>'
    with pytest.raises(InvalidWorkbookError):
        ooxml.parse_xml(data)


def test_read_part_respects_size_limit(book, monkeypatch):
    monkeypatch.setattr(settings, "MAX_PART_BYTES", 50)
    with zipfile.ZipFile(book) as zf, pytest.raises(InvalidWorkbookError, match="MAX_PART_BYTES"):
        ooxml.read_part(zf, "xl/workbook.xml")


def test_rewrite_package_keeps_order_and_unreplaced_bytes(book, tmp_path):
    out = tmp_path / "out.xlsx"
    ooxml.rewrite_package(book, out, {"xl/worksheets/sheet1.xml": b"<replaced/>"})
    with zipfile.ZipFile(book) as a, zipfile.ZipFile(out) as b:
        assert a.namelist() == b.namelist()
        for info in a.infolist():
            if info.filename == "xl/worksheets/sheet1.xml":
                assert b.read(info.filename) == b"<replaced/>"
            else:
                assert a.read(info.filename) == b.read(info.filename)
                assert b.getinfo(info.filename).compress_type == info.compress_type


def test_rewrite_package_rejects_unknown_part_and_same_path(book, tmp_path):
    with pytest.raises(InvalidWorkbookError):
        ooxml.rewrite_package(book, tmp_path / "o.xlsx", {"xl/nope.xml": b"x"})
    with pytest.raises(ValueError):
        ooxml.rewrite_package(book, book, {})


def test_rich_text_ignores_phonetic_runs():
    ns = ooxml.NS_MAIN
    el = ooxml.parse_xml(
        f'<si xmlns="{ns}"><r><t>Hel</t></r><r><t>lo</t></r><rPh><t>IGNORED</t></rPh></si>'.encode()
    ).getroot()
    assert ooxml.rich_text(el) == "Hello"
