from pathlib import Path

import openpyxl
import pytest
from openpyxl.comments import Comment

from xlsx_tools_mcp import server, settings
from xlsx_tools_mcp.errors import BaselineNotFoundError
from xlsx_tools_mcp.io import review, writer

from .helpers import file_hash


@pytest.fixture
def book(tmp_path):
    path = str(tmp_path / "report.xlsx")
    writer.create_workbook(path, ["Data"])
    wb = openpyxl.load_workbook(path)
    ws = wb["Data"]
    ws["A1"] = "item"
    ws["B1"] = 10
    ws["C1"] = "=B1*2"
    ws["D1"] = "keep"
    ws["D1"].comment = Comment("original note", "Ann")
    wb.save(path)
    return path


def _change(path: str) -> None:
    writer.write_cells(path, "Data", [
        {"cell": "B1", "value": 99},
        {"cell": "C1", "formula": "=B1*3"},
        {"cell": "E1", "value": "new"},
        {"cell": "A1", "value": None},
        {"cell": "D1", "value": "changed"},
    ], recalculate=False, backup=True)


def test_review_copy_marks_changes_and_keeps_inputs(book):
    _change(book)
    baseline = review.find_latest_backup(book)
    src_hash, base_hash = file_hash(book), file_hash(baseline)
    res = server.export_review_copy(path=book)
    assert file_hash(book) == src_hash
    assert file_hash(baseline) == base_hash
    assert res["baseline_path"] == baseline
    assert res["summary"]["value"] == 2 and res["summary"]["formula"] == 1
    assert res["summary"]["added"] == 1 and res["summary"]["removed"] == 1

    out = openpyxl.load_workbook(res["output_path"])
    ws = out["Data"]
    assert ws["B1"].fill.fgColor.rgb.endswith("FFF2CC")
    assert "Before: 10" in ws["B1"].comment.text and "After: 99" in ws["B1"].comment.text
    assert ws["D1"].comment.text.startswith("original note")  # appended, not replaced
    assert ws["D1"].comment.author == "Ann"
    changes = out.worksheets[0]
    assert changes.title == "_Changes"
    rows = list(changes.iter_rows(min_row=2, values_only=True))
    formula_row = next(r for r in rows if r[1] == "C1")
    assert formula_row[3] == "=B1*2" and formula_row[4] == "=B1*3"
    assert changes["D2"].data_type == "s"  # stored as text, never evaluated
    assert sum(1 for ws_ in out.worksheets if ws_.sheet_view.tabSelected) == 1


def test_review_copy_without_backup_fails(book):
    with pytest.raises(BaselineNotFoundError):
        review.find_latest_backup(book)
    with pytest.raises(ValueError, match="No backup"):
        server.export_review_copy(path=book)


def test_output_must_differ_from_source(book):
    _change(book)
    with pytest.raises(ValueError, match="must differ"):
        server.export_review_copy(output_path=book, path=book)


def test_existing_output_needs_overwrite(book):
    _change(book)
    server.export_review_copy(path=book)
    with pytest.raises(ValueError, match="already exists"):
        server.export_review_copy(path=book)
    assert server.export_review_copy(overwrite=True, path=book)["listed"] > 0


def test_output_outside_allowed_dirs_rejected(book, tmp_path, monkeypatch):
    _change(book)
    monkeypatch.setattr(settings, "ALLOWED_DIRS", [str(Path(book).parent.resolve())])
    outside = tmp_path.parent / "elsewhere.xlsx"
    with pytest.raises(ValueError, match="Access denied"):
        server.export_review_copy(output_path=str(outside), path=book)


def test_latest_backup_chosen_by_name_timestamp(book, tmp_path, monkeypatch):
    backup_dir = tmp_path / "bak"
    backup_dir.mkdir()
    monkeypatch.setattr(writer, "BACKUP_DIR", str(backup_dir))
    name = Path(book).name
    for stamp in ("20260101-000000", "20261231-235959", "20260601-120000"):
        (backup_dir / f"{name}.{stamp}.bak").write_bytes(Path(book).read_bytes())
    (backup_dir / "other.xlsx.20991231-000000.bak").write_bytes(b"x")
    assert review.find_latest_backup(book).endswith("20261231-235959.bak")


def test_explicit_baseline_and_new_sheet(book, tmp_path):
    baseline = str(tmp_path / "base.xlsx")
    Path(baseline).write_bytes(Path(book).read_bytes())
    writer.create_sheet(book, "Added", backup=False)
    res = server.export_review_copy(baseline_path=baseline, output_path=str(tmp_path / "r.xlsx"), path=book)
    assert res["summary"]["sheets_added"] == ["Added"]
    assert openpyxl.load_workbook(res["output_path"])["Added"].sheet_properties.tabColor.rgb.endswith("FFC000")


def test_invalid_color_rejected(book, tmp_path):
    with pytest.raises(ValueError, match="hex"):
        server.export_review_copy(baseline_path=book, output_path=str(tmp_path / "x.xlsx"), highlight_color="red", path=book)


def test_annotation_cap_sets_truncated(book):
    _change(book)
    res = server.export_review_copy(max_annotated_cells=1, path=book)
    assert res["annotated"] == 1
    assert res["truncated"] is True
