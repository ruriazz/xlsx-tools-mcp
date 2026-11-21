"""Shared helpers for tools that write a *new* workbook derived from a source one.

Review copies and sanitized copies never touch their source: the output path must
differ from every input, and saving reuses the writer's atomic save. Because an
openpyxl save drops every cached formula result, the source's cached values are
injected back into the copy afterwards — no LibreOffice needed.
"""

from __future__ import annotations

import os
import zipfile
from pathlib import Path

import openpyxl
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.workbook.workbook import Workbook

from .. import inject
from ..errors import InvalidWorkbookError, WorkbookNotFoundError
from .writer import _EXCEL_EXTENSIONS, _atomic_save


def same_file(a: str, b: str) -> bool:
    """True for the same file, including case variants, hard links and mapped drives."""
    if os.path.exists(a) and os.path.exists(b):
        try:
            return os.path.samefile(a, b)
        except OSError:
            pass
    return os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))


def ensure_distinct(output: str, *inputs: str) -> None:
    """Refuse an output that is one of the inputs (call before taking file locks:
    two lock paths naming the same file would wait on each other)."""
    for other in inputs:
        if same_file(output, other):
            raise ValueError(f"output_path must differ from the input file '{other}'")


def check_output(source: str, output: str, overwrite: bool, *inputs: str) -> None:
    ensure_distinct(output, source, *inputs)
    out_suffix = Path(output).suffix.lower()
    if out_suffix not in _EXCEL_EXTENSIONS:
        raise ValueError(f"output_path must end in .xlsx or .xlsm, got '{output}'")
    if out_suffix == ".xlsm" and Path(source).suffix.lower() != ".xlsm":
        raise ValueError("output_path may only be .xlsm when the source workbook is .xlsm")
    if Path(output).exists() and not overwrite:
        raise FileExistsError(f"File already exists: {output}. Pass overwrite=True to replace it.")


def load_source(source: str, output: str, data_only: bool = False) -> Workbook:
    if not Path(source).is_file():
        raise WorkbookNotFoundError(f"Workbook not found: '{source}'.")
    keep_vba = Path(output).suffix.lower() == ".xlsm"
    try:
        return openpyxl.load_workbook(source, data_only=data_only, keep_vba=keep_vba)
    except PermissionError:
        raise
    except (zipfile.BadZipFile, InvalidFileException) as exc:
        raise InvalidWorkbookError(f"File '{source}' is not a valid Excel workbook: {exc}") from exc


def save_copy(wb: Workbook, source: str, output: str) -> inject.InjectStats:
    """Atomically save `wb` to `output`, then restore the source's cached formula results."""
    Path(output).parent.mkdir(parents=True, exist_ok=True)
    try:
        _atomic_save(wb, output, backup=False)
    finally:
        wb.close()
    return inject.inject_cached_values(output, inject.read_cached_values(source), output)
