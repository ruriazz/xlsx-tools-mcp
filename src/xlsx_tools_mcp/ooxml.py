"""Low-level, read-safe access to the zip/XML package behind an .xlsx file.

Used where openpyxl is the wrong tool: injecting cached formula values without a
full openpyxl load/save (which drops shapes and cached values), and inspecting
parts openpyxl discards (docProps/app.xml, pivot cache records, external links).

Every XML part is parsed with a hardened lxml parser (no entity resolution, no
network, no DTD loading) and size-capped before it is read into memory. lxml is
required over xml.etree: etree drops namespace declarations that are only named in
`mc:Ignorable`, and Excel then refuses to open the file.
"""

from __future__ import annotations

import os
import posixpath
import shutil
import zipfile
from dataclasses import dataclass
from pathlib import Path

from lxml import etree

from . import settings
from .errors import InvalidWorkbookError

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

REL_OFFICE_DOCUMENT = "/officeDocument"
REL_SHARED_STRINGS = "/sharedStrings"
REL_WORKSHEET = "/worksheet"


def q(tag: str) -> str:
    """Qualify a SpreadsheetML main-namespace tag name."""
    return f"{{{NS_MAIN}}}{tag}"


@dataclass(frozen=True)
class SheetPart:
    name: str
    state: str  # "visible" | "hidden" | "veryHidden"
    part: str   # zip entry name, e.g. "xl/worksheets/sheet1.xml"


def _parser() -> etree.XMLParser:
    # A fresh parser per call: lxml parser objects are not thread-safe.
    return etree.XMLParser(
        resolve_entities=False,
        no_network=True,
        load_dtd=False,
        huge_tree=False,
        remove_blank_text=False,
    )


def read_part(zf: zipfile.ZipFile, name: str) -> bytes:
    """Read one zip entry, refusing entries larger than XLSX_MCP_MAX_PART_BYTES.

    The declared size is checked first, then the actual decompressed stream is capped
    too, since a hostile archive can lie about `file_size`.
    """
    limit = settings.MAX_PART_BYTES
    try:
        info = zf.getinfo(name)
    except KeyError as exc:
        raise InvalidWorkbookError(f"Workbook part '{name}' is missing") from exc
    if info.file_size > limit:
        raise InvalidWorkbookError(
            f"Workbook part '{name}' is {info.file_size} bytes, above XLSX_MCP_MAX_PART_BYTES={limit}"
        )
    with zf.open(info) as fh:
        data = fh.read(limit + 1)
    if len(data) > limit:
        raise InvalidWorkbookError(f"Workbook part '{name}' exceeds XLSX_MCP_MAX_PART_BYTES={limit}")
    return data


def parse_xml(data: bytes, name: str = "<xml>") -> etree._ElementTree:
    try:
        root = etree.fromstring(data, _parser())
    except etree.XMLSyntaxError as exc:
        raise InvalidWorkbookError(f"Workbook part '{name}' is not well-formed XML: {exc}") from exc
    tree = root.getroottree()
    if tree.docinfo.doctype:
        raise InvalidWorkbookError(f"Workbook part '{name}' declares a DOCTYPE, which is not allowed")
    return tree


def parse_part(zf: zipfile.ZipFile, name: str) -> etree._ElementTree:
    return parse_xml(read_part(zf, name), name)


def serialize(tree: etree._ElementTree) -> bytes:
    return etree.tostring(
        tree, xml_declaration=True, encoding="UTF-8", standalone=tree.docinfo.standalone
    )


def has_part(zf: zipfile.ZipFile, name: str) -> bool:
    try:
        zf.getinfo(name)
        return True
    except KeyError:
        return False


def _rels_name(part: str) -> str:
    folder, base = posixpath.split(part)
    return posixpath.join(folder, "_rels", f"{base}.rels")


def _resolve_target(source_part: str, target: str) -> str:
    if target.startswith("/"):
        return posixpath.normpath(target).lstrip("/")
    base = posixpath.dirname(source_part)
    joined = posixpath.join(base, target) if base else target
    return posixpath.normpath(joined).lstrip("/")


def relationships(zf: zipfile.ZipFile, part: str) -> list[dict[str, str]]:
    """Relationships declared by `part` (package root when part == "")."""
    rels_name = "_rels/.rels" if part == "" else _rels_name(part)
    if not has_part(zf, rels_name):
        return []
    root = parse_part(zf, rels_name).getroot()
    out: list[dict[str, str]] = []
    for rel in root.iter(f"{{{NS_PKG_REL}}}Relationship"):
        target = rel.get("Target") or ""
        external = rel.get("TargetMode") == "External"
        out.append(
            {
                "id": rel.get("Id") or "",
                "type": rel.get("Type") or "",
                "target": target if external else _resolve_target(part, target),
                "external": "1" if external else "",
            }
        )
    return out


def workbook_part(zf: zipfile.ZipFile) -> str:
    for rel in relationships(zf, ""):
        if rel["type"].endswith(REL_OFFICE_DOCUMENT) and not rel["external"]:
            return rel["target"]
    if has_part(zf, "xl/workbook.xml"):
        return "xl/workbook.xml"
    raise InvalidWorkbookError("Workbook package has no workbook part")


def workbook_sheets(zf: zipfile.ZipFile) -> list[SheetPart]:
    """Worksheets in workbook order, mapped to their zip entries.

    Chartsheets and dialog sheets are skipped: they hold no cells.
    """
    wb_part = workbook_part(zf)
    rel_targets = {
        r["id"]: r["target"]
        for r in relationships(zf, wb_part)
        if not r["external"] and r["type"].endswith(REL_WORKSHEET)
    }
    root = parse_part(zf, wb_part).getroot()
    sheets_el = root.find(q("sheets"))
    if sheets_el is None:
        return []
    out: list[SheetPart] = []
    for sheet in sheets_el.iterchildren(q("sheet")):
        target = rel_targets.get(sheet.get(f"{{{NS_REL}}}id") or "")
        if not target or not has_part(zf, target):
            continue
        out.append(SheetPart(sheet.get("name") or "", sheet.get("state") or "visible", target))
    return out


def shared_strings(zf: zipfile.ZipFile) -> list[str]:
    wb_part = workbook_part(zf)
    target = next(
        (r["target"] for r in relationships(zf, wb_part) if r["type"].endswith(REL_SHARED_STRINGS)),
        None,
    )
    if target is None or not has_part(zf, target):
        return []
    root = parse_part(zf, target).getroot()
    return [rich_text(si) for si in root.iterchildren(q("si"))]


def rich_text(el: etree._Element) -> str:
    """Text of a <si>/<is> element: direct <t> plus <r><t> runs, ignoring phonetic <rPh>."""
    parts: list[str] = []
    for child in el.iterchildren():
        if child.tag == q("t"):
            parts.append(child.text or "")
        elif child.tag == q("r"):
            parts.extend(t.text or "" for t in child.iterchildren(q("t")))
    return "".join(parts)


def rewrite_package(src: str | Path, dst: str | Path, replacements: dict[str, bytes]) -> None:
    """Copy the zip at `src` to `dst`, swapping in `replacements` for the named entries.

    Entry order, per-entry compression and timestamps are kept, so `[Content_Types].xml`
    stays first. Unreplaced entries are streamed, never fully buffered.
    """
    if os.path.normcase(os.path.abspath(src)) == os.path.normcase(os.path.abspath(dst)):
        raise ValueError("rewrite_package needs distinct src and dst paths")
    with zipfile.ZipFile(src) as zin:
        names = {i.filename for i in zin.infolist()}
        unknown = set(replacements) - names
        if unknown:
            raise InvalidWorkbookError(f"Cannot replace missing workbook parts: {sorted(unknown)}")
        with zipfile.ZipFile(dst, "w", allowZip64=True) as zout:
            for info in zin.infolist():
                new_info = zipfile.ZipInfo(info.filename, date_time=info.date_time)
                new_info.compress_type = info.compress_type
                new_info.external_attr = info.external_attr
                new_info.comment = info.comment
                if info.filename in replacements:
                    zout.writestr(new_info, replacements[info.filename])
                    continue
                new_info.file_size = info.file_size
                with zin.open(info) as fin, zout.open(new_info, "w", force_zip64=info.file_size > 0x7FFFFFFF) as fout:
                    shutil.copyfileobj(fin, fout, 1024 * 1024)
