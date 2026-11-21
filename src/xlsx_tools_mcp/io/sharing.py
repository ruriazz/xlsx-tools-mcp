"""Find hidden or sensitive content before a workbook is shared, and write a cleaned copy.

Inspection reads the original zip package (openpyxl drops parts such as
docProps/app.xml company fields and pivot cache records on load) plus an openpyxl
load for row/column visibility and cell formatting. Sanitizing always writes a new
file; the source is never modified.
"""

from __future__ import annotations

import re
import zipfile
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.packaging.custom import CustomPropertyList
from openpyxl.utils import column_index_from_string
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.workbook.workbook import Workbook
from openpyxl.xml.functions import tostring

from .. import inject, ledger, ooxml
from ..errors import InvalidWorkbookError, WorkbookNotFoundError
from ..refs import quote_sheet
from . import copies

SANITIZABLE = ("comments", "document_properties", "hidden_sheets", "external_links", "change_log", "macros")
DEFAULT_REMOVE = ("comments", "document_properties", "change_log")
SEVERITY = {
    "hidden_sheets": "high",
    "external_links": "high",
    "pivot_cache": "high",
    "macros": "high",
    "hidden_rows_columns": "medium",
    "comments": "medium",
    "document_properties": "medium",
    "hidden_formatting": "medium",
    "embedded_objects": "medium",
    "change_log": "medium",
    "defined_names": "low",
    "custom_xml": "low",
}
MAX_SAMPLES = 10
MAX_HIDDEN_LINES_SCANNED = 2000
_EXT_REF = re.compile(r"\[\d+\]")
_NS_DC = "http://purl.org/dc/elements/1.1/"
_NS_CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
_NS_APP = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
_CORE_FIELDS = (
    (_NS_DC, "creator"), (_NS_CP, "lastModifiedBy"), (_NS_DC, "title"), (_NS_DC, "subject"),
    (_NS_CP, "keywords"), (_NS_DC, "description"), (_NS_CP, "category"),
)
_APP_FIELDS = ("Company", "Manager")


def _finding(category: str, location: str, detail: str, count: int) -> dict[str, Any]:
    return {
        "category": category,
        "severity": SEVERITY[category],
        "location": location,
        "detail": detail,
        "count": count,
        "sanitizable": category in SANITIZABLE,
    }


def _sample(items: list[str]) -> str:
    shown = ", ".join(items[:MAX_SAMPLES])
    return shown + (f" (+{len(items) - MAX_SAMPLES} more)" if len(items) > MAX_SAMPLES else "")


# ── Package-level checks ──────────────────────────────────────────────────────


def _count_value_cells(zf: zipfile.ZipFile, part: str) -> int:
    root = ooxml.parse_part(zf, part).getroot()
    return sum(
        1 for c in root.iter(ooxml.q("c"))
        if c.find(ooxml.q("v")) is not None or c.find(ooxml.q("is")) is not None
    )


def _package_findings(zf: zipfile.ZipFile) -> list[dict[str, Any]]:
    names = zf.namelist()
    sheets = ooxml.workbook_sheets(zf)
    out: list[dict[str, Any]] = []

    for sp in sheets:
        if ledger.is_reserved(sp.name):
            out.append(_finding("change_log", sp.name, "xlsx-tools-mcp change log sheet (very hidden)", 1))
        elif sp.state in ("hidden", "veryHidden"):
            cells = _count_value_cells(zf, sp.part)
            out.append(_finding("hidden_sheets", sp.name, f"Sheet '{sp.name}' is {sp.state} and has {cells} non-empty cells", 1))

    out.extend(_comment_findings(zf, names))
    out.extend(_property_findings(zf))
    out.extend(_external_findings(zf, names, sheets))
    out.extend(_defined_name_findings(zf))

    simple = (
        ("pivot_cache", "xl/pivotCache/pivotCacheRecords", "pivot cache record part(s) store a copy of the source data"),
        ("embedded_objects", ("xl/embeddings/", "xl/activeX/"), "embedded object part(s)"),
        ("macros", "xl/vbaProject.bin", "VBA macro project"),
        ("custom_xml", "customXml/", "custom XML part(s)"),
    )
    for category, prefixes, text in simple:
        hits = [n for n in names if n.startswith(prefixes)]
        if hits:
            out.append(_finding(category, "package", f"{len(hits)} {text}: {_sample(hits)}", len(hits)))
    return out


def _comment_findings(zf: zipfile.ZipFile, names: list[str]) -> list[dict[str, Any]]:
    legacy = [n for n in names if re.match(r"^xl/comments[^/]*\.xml$|^xl/comments/[^/]+\.xml$", n)]
    threaded = [n for n in names if n.startswith("xl/threadedComments/") and n.endswith(".xml")]
    persons = [n for n in names if n.startswith("xl/persons/") and n.endswith(".xml")]
    count = 0
    authors: set[str] = set()
    for part in legacy:
        root = ooxml.parse_part(zf, part).getroot()
        count += sum(1 for _ in root.iter(ooxml.q("comment")))
        authors.update(a.text or "" for a in root.iter(ooxml.q("author")) if a.text)
    for part in threaded:
        root = ooxml.parse_part(zf, part).getroot()
        count += sum(1 for el in root.iter() if isinstance(el.tag, str) and el.tag.endswith("}threadedComment"))
    for part in persons:
        root = ooxml.parse_part(zf, part).getroot()
        authors.update(
            el.get("displayName") for el in root.iter()
            if isinstance(el.tag, str) and el.tag.endswith("}person") and el.get("displayName")
        )
    if not count:
        return []
    who = f"; authors: {_sample(sorted(authors))}" if authors else ""
    return [_finding("comments", "workbook", f"{count} comment(s){who}", count)]


def _property_findings(zf: zipfile.ZipFile) -> list[dict[str, Any]]:
    found: list[str] = []
    if ooxml.has_part(zf, "docProps/core.xml"):
        root = ooxml.parse_part(zf, "docProps/core.xml").getroot()
        for ns, local in _CORE_FIELDS:
            el = root.find(f"{{{ns}}}{local}")
            if el is not None and (el.text or "").strip():
                found.append(f"{local}={el.text.strip()!r}")
    if ooxml.has_part(zf, "docProps/app.xml"):
        root = ooxml.parse_part(zf, "docProps/app.xml").getroot()
        for local in _APP_FIELDS:
            el = root.find(f"{{{_NS_APP}}}{local}")
            if el is not None and (el.text or "").strip():
                found.append(f"{local}={el.text.strip()!r}")
    if ooxml.has_part(zf, "docProps/custom.xml"):
        root = ooxml.parse_part(zf, "docProps/custom.xml").getroot()
        found.extend(
            f"custom:{el.get('name')}" for el in root.iter()
            if isinstance(el.tag, str) and el.tag.endswith("}property") and el.get("name")
        )
    if not found:
        return []
    return [_finding("document_properties", "docProps", _sample(found), len(found))]


def _external_findings(zf: zipfile.ZipFile, names: list[str], sheets: list[ooxml.SheetPart]) -> list[dict[str, Any]]:
    parts = [n for n in names if n.startswith("xl/externalLinks/") and n.endswith(".xml") and "/_rels/" not in n]
    targets = sorted({r["target"] for p in parts for r in ooxml.relationships(zf, p) if r["external"]})
    cells: list[str] = []
    if parts:
        for sp in sheets:
            root = ooxml.parse_part(zf, sp.part).getroot()
            for c in root.iter(ooxml.q("c")):
                f = c.find(ooxml.q("f"))
                if f is not None and f.text and _EXT_REF.search(f.text):
                    cells.append(f"{sp.name}!{c.get('r')}")
    if not parts:
        return []
    detail = f"{len(parts)} external workbook link(s)"
    if targets:
        detail += f" to {_sample(targets)}"
    if cells:
        detail += f"; formulas using them: {_sample(cells)}"
    return [_finding("external_links", "workbook", detail, len(parts))]


def _defined_name_findings(zf: zipfile.ZipFile) -> list[dict[str, Any]]:
    root = ooxml.parse_part(zf, ooxml.workbook_part(zf)).getroot()
    bad = [
        f"{dn.get('name')}={dn.text}" for dn in root.iter(ooxml.q("definedName"))
        if dn.text and (_EXT_REF.search(dn.text) or "#REF!" in dn.text)
    ]
    if not bad:
        return []
    return [_finding("defined_names", "workbook", f"names pointing to other workbooks or #REF!: {_sample(bad)}", len(bad))]


# ── Cell-level checks (openpyxl) ──────────────────────────────────────────────


def _rgb(color: Any) -> str | None:
    rgb = getattr(color, "rgb", None)
    return rgb.upper() if isinstance(rgb, str) and len(rgb) in (6, 8) else None


def _is_hidden_format(cell: Any) -> bool:
    if isinstance(cell.number_format, str) and cell.number_format.strip() == ";;;":
        return True
    font = _rgb(cell.font.color) if cell.font is not None else None
    if not font:
        return False
    fill = cell.fill
    bg = _rgb(fill.fgColor) if fill is not None and fill.fill_type == "solid" else None
    if bg is None:
        return font[-6:] == "FFFFFF"
    return font[-6:] == bg[-6:]


def _hidden_lines(ws) -> list[str]:
    hits: list[str] = []
    rows = [r for r, d in ws.row_dimensions.items() if d.hidden][:MAX_HIDDEN_LINES_SCANNED]
    for r in rows:
        if any(v is not None for row in ws.iter_rows(min_row=r, max_row=r, values_only=True) for v in row):
            hits.append(f"row {r}")
    scanned = 0
    for key, dim in ws.column_dimensions.items():
        if not dim.hidden:
            continue
        start = dim.min or column_index_from_string(key)
        end = dim.max or start
        for c in range(start, end + 1):
            scanned += 1
            if scanned > MAX_HIDDEN_LINES_SCANNED:
                return hits
            if any(v is not None for col in ws.iter_cols(min_col=c, max_col=c, values_only=True) for v in col):
                hits.append(f"column {c}")
    return hits


def _cell_findings(wb: Workbook) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    lines: list[str] = []
    formatted: list[str] = []
    for ws in wb.worksheets:
        if ledger.is_reserved(ws.title):
            continue
        lines.extend(f"{ws.title}: {h}" for h in _hidden_lines(ws))
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None and not isinstance(cell, MergedCell) and _is_hidden_format(cell):
                    formatted.append(f"{ws.title}!{cell.coordinate}")
    if lines:
        out.append(_finding("hidden_rows_columns", "worksheets", f"hidden lines with data: {_sample(lines)}", len(lines)))
    if formatted:
        out.append(_finding(
            "hidden_formatting", "worksheets",
            f"cells hidden by ';;;' format or text colored like its background: {_sample(formatted)}", len(formatted),
        ))
    return out


def _open_zip(path: str) -> zipfile.ZipFile:
    if not Path(path).is_file():
        raise WorkbookNotFoundError(f"Workbook not found: '{path}'.")
    try:
        return zipfile.ZipFile(path)
    except zipfile.BadZipFile as exc:
        raise InvalidWorkbookError(f"File '{path}' is not a valid Excel workbook: {exc}") from exc


def inspect_for_sharing(path: str) -> dict[str, Any]:
    with _open_zip(path) as zf:
        findings = _package_findings(zf)
    try:
        wb = openpyxl.load_workbook(path, data_only=False)
    except PermissionError:
        raise
    except (zipfile.BadZipFile, InvalidFileException) as exc:
        raise InvalidWorkbookError(f"File '{path}' is not a valid Excel workbook: {exc}") from exc
    try:
        findings.extend(_cell_findings(wb))
    finally:
        wb.close()
    order = {"high": 0, "medium": 1, "low": 2}
    findings.sort(key=lambda f: order[f["severity"]])
    summary = {level: sum(1 for f in findings if f["severity"] == level) for level in order}
    return {"findings": findings, "summary": summary}


# ── Sanitizing ────────────────────────────────────────────────────────────────


def _formula_text(value: Any) -> str | None:
    text = getattr(value, "text", value)
    return text if isinstance(text, str) and (text.startswith("=") or value is not text) else None


def _references_sheet(text: str, name: str) -> bool:
    quoted = quote_sheet(name) + "!"
    if quoted.lower() in text.lower():
        return True
    return re.search(rf"(?<![\w.'\]]){re.escape(name)}!", text, re.IGNORECASE) is not None


def _sheet_references(wb: Workbook, name: str) -> list[str]:
    refs: list[str] = []
    for ws in wb.worksheets:
        if ws.title == name:
            continue
        for row in ws.iter_rows():
            for cell in row:
                text = _formula_text(cell.value)
                if text and _references_sheet(text, name):
                    refs.append(f"{ws.title}!{cell.coordinate}")
        for dv in ws.data_validations.dataValidation:
            if any(f and _references_sheet(str(f), name) for f in (dv.formula1, dv.formula2)):
                refs.append(f"{ws.title} data validation {dv.sqref}")
        for cf in ws.conditional_formatting:
            if any(_references_sheet(str(f), name) for rule in cf.rules for f in (rule.formula or [])):
                refs.append(f"{ws.title} conditional format {cf.sqref}")
        for chart in getattr(ws, "_charts", []):
            try:
                xml = tostring(chart._write()).decode("utf-8", "replace")
            except Exception:  # chart serialization is best-effort for reference checks
                continue
            if _references_sheet(xml, name):
                refs.append(f"{ws.title} chart")
        for pivot in getattr(ws, "_pivots", []):
            source = getattr(getattr(getattr(pivot, "cache", None), "cacheSource", None), "worksheetSource", None)
            if source is not None and getattr(source, "sheet", None) == name:
                refs.append(f"{ws.title} pivot table")
    for dn_name, dn in list(wb.defined_names.items()):
        if dn.attr_text and _references_sheet(dn.attr_text, name):
            refs.append(f"defined name {dn_name}")
    return refs


def _remove_hidden_sheets(wb: Workbook, applied: list[str]) -> None:
    targets = [ws.title for ws in wb.worksheets if ws.sheet_state in ("hidden", "veryHidden") and not ledger.is_reserved(ws.title)]
    for name in targets:
        refs = _sheet_references(wb, name)
        if refs:
            raise ValueError(
                f"Cannot remove hidden sheet '{name}': it is still referenced by {_sample(refs)}. "
                "Replace those references first, or leave 'hidden_sheets' out of `remove`."
            )
    for name in targets:
        del wb[name]
    if targets:
        visible = [i for i, ws in enumerate(wb.worksheets) if ws.sheet_state == "visible"]
        if visible:
            wb.active = visible[0]
        applied.append(f"hidden_sheets: removed {_sample(targets)}")


def _remove_comments(wb: Workbook, applied: list[str]) -> None:
    removed = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if not isinstance(cell, MergedCell) and cell.comment is not None:
                    cell.comment = None
                    removed += 1
    applied.append(f"comments: removed {removed}")


def _clear_properties(wb: Workbook, applied: list[str]) -> None:
    props = wb.properties
    for attr in ("creator", "lastModifiedBy", "title", "subject", "keywords", "description",
                 "category", "identifier", "contentStatus", "version", "language", "lastPrinted"):
        if hasattr(props, attr):
            setattr(props, attr, None)
    wb.custom_doc_props = CustomPropertyList()
    applied.append("document_properties: cleared author, title, subject, keywords, description, category and custom properties")


def _remove_external_links(wb: Workbook, source: str, applied: list[str], warnings: list[str]) -> None:
    links = getattr(wb, "_external_links", [])
    if not links:
        applied.append("external_links: none present")
        return
    cached = inject.read_cached_values(source)
    converted: list[str] = []
    for ws in wb.worksheets:
        sheet_cache = cached.get(ws.title, {})
        for row in ws.iter_rows():
            for cell in row:
                text = _formula_text(cell.value)
                if text and _EXT_REF.search(text):
                    cell_cache = sheet_cache.get(cell.coordinate)
                    value = inject.to_python(cell_cache)
                    if value is None:
                        warnings.append(f"{ws.title}!{cell.coordinate}: no cached value; cleared")
                    cell.value = value
                    if isinstance(value, str):
                        # A cached "=..." string must stay text, never become a live formula.
                        cell.data_type = "e" if cell_cache is not None and cell_cache[0] == "e" else "s"
                    converted.append(f"{ws.title}!{cell.coordinate}")
    for dn_name, dn in list(wb.defined_names.items()):
        if dn.attr_text and _EXT_REF.search(dn.attr_text):
            del wb.defined_names[dn_name]
            warnings.append(f"defined name '{dn_name}' pointed to an external workbook and was removed")
    for ws in wb.worksheets:
        for dn_name, dn in list(ws.defined_names.items()):
            if dn.attr_text and _EXT_REF.search(dn.attr_text):
                del ws.defined_names[dn_name]
                warnings.append(f"defined name '{dn_name}' on '{ws.title}' pointed to an external workbook and was removed")
    wb._external_links = []
    applied.append(f"external_links: removed {len(links)} link(s), converted {len(converted)} formula(s) to values")


def _remove_change_log(wb: Workbook, applied: list[str]) -> None:
    name = ledger.find_sheet_name(wb.sheetnames)
    if name:
        del wb[name]
        applied.append("change_log: removed")
    else:
        applied.append("change_log: none present")


def sanitize_copy(path: str, output_path: str, remove: list[str] | None = None, overwrite: bool = False) -> dict[str, Any]:
    actions = list(dict.fromkeys(remove if remove is not None else DEFAULT_REMOVE))
    invalid = [a for a in actions if a not in SANITIZABLE]
    if invalid:
        raise ValueError(f"Unsupported remove option(s) {invalid}. Choose from {list(SANITIZABLE)}.")
    if not Path(path).is_file():
        raise WorkbookNotFoundError(f"Workbook not found: '{path}'.")
    copies.check_output(path, output_path, overwrite)
    out_is_xlsm = Path(output_path).suffix.lower() == ".xlsm"
    if "macros" in actions and out_is_xlsm:
        raise ValueError("Removing macros needs a .xlsx output_path")
    with _open_zip(path) as zf:
        had_macros = ooxml.has_part(zf, "xl/vbaProject.bin")

    applied: list[str] = []
    warnings: list[str] = []
    wb = copies.load_source(path, output_path)
    try:
        if "hidden_sheets" in actions:
            _remove_hidden_sheets(wb, applied)
        if "comments" in actions:
            _remove_comments(wb, applied)
        if "document_properties" in actions:
            _clear_properties(wb, applied)
        if "external_links" in actions:
            _remove_external_links(wb, path, applied, warnings)
        if "change_log" in actions:
            _remove_change_log(wb, applied)
    except Exception:
        wb.close()
        raise
    if had_macros and not out_is_xlsm:
        if "macros" in actions:
            applied.append("macros: removed (a .xlsx file cannot hold macros)")
        else:
            warnings.append("macros were dropped because the output is .xlsx")
    copies.save_copy(wb, path, output_path)
    remaining = inspect_for_sharing(output_path)
    return {
        "output_path": output_path,
        "applied": applied,
        "warnings": warnings,
        "remaining_findings": remaining["findings"],
    }
