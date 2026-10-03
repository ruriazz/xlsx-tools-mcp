from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import openpyxl

from .errors import FileInUseError
from .locking import safe_replace
from .settings import SOFFICE_TIMEOUT_SECONDS

# All standard Excel formula error values, including the newer dynamic-array ones.
EXCEL_ERROR_VALUES = {
    "#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#NULL!", "#NUM!", "#N/A", "#SPILL!", "#CALC!",
}


@dataclass
class RecalcResult:
    success: bool
    errors_found: list[dict[str, Any]] = field(default_factory=list)
    message: str = ""


def find_soffice() -> str | None:
    """Locate the LibreOffice binary, checking PATH and common install locations."""
    # 1. PATH lookup (Linux, macOS, Windows if added to PATH)
    for bin_name in ("soffice", "libreoffice", "soffice.exe"):
        found = shutil.which(bin_name)
        if found:
            return found

    # 2. Standard macOS path
    mac_path = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    if sys.platform == "darwin" and os.path.isfile(mac_path) and os.access(mac_path, os.X_OK):
        return mac_path

    # 3. Standard Windows paths
    if sys.platform == "win32":
        win_candidates = [
            os.path.join(os.environ.get("ProgramFiles", r"C:\Program Files"), "LibreOffice", "program", "soffice.exe"),
            os.path.join(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"), "LibreOffice", "program", "soffice.exe"),
        ]
        local_app_data = os.environ.get("LocalAppData")
        if local_app_data:
            win_candidates.append(os.path.join(local_app_data, "Programs", "LibreOffice", "program", "soffice.exe"))
        for candidate in win_candidates:
            if os.path.isfile(candidate):
                return candidate

    return None


_orig_find_soffice = find_soffice


def scan_formula_errors(path: str) -> list[dict[str, Any]]:
    """Scan every cell's cached value for Excel error strings (#DIV/0!, #REF!, ...)."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        found: list[dict[str, Any]] = []
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and cell.value in EXCEL_ERROR_VALUES:
                        found.append({"sheet": ws.title, "cell": cell.coordinate, "error": cell.value})
        return found
    finally:
        wb.close()


def recalculate(path: str, timeout: int = SOFFICE_TIMEOUT_SECONDS) -> RecalcResult:
    """Recalculate every formula in an xlsx file via a LibreOffice headless round-trip.

    openpyxl never evaluates formulas — it only writes formula strings with the old
    cached value (or none). This opens the file in LibreOffice Calc, lets it compute
    every formula, and writes the result back in place. Never raises for expected
    failure modes (missing LibreOffice, timeout, bad file) — those are reported in
    the returned RecalcResult so a failed recalculation never masks a successful save.
    """
    soffice = find_soffice()
    if not soffice:
        return RecalcResult(
            success=False,
            message=(
                "LibreOffice (soffice) not found on PATH. The file was saved but formulas "
                "were not recalculated — cached values may be stale until it is installed. "
                "macOS: brew install --cask libreoffice · Linux: apt-get install libreoffice-calc · Windows: winget install TheDocumentFoundation.LibreOffice"
            ),
        )

    with tempfile.TemporaryDirectory(prefix="xlsx_tools_mcp_recalc_") as tmpdir:
        src = Path(path)
        tmp_input = Path(tmpdir) / src.name
        shutil.copy(src, tmp_input)

        profile_dir = Path(tmpdir) / "lo_profile"
        cmd = [
            soffice,
            "--headless",
            "--norestore",
            "--nofirststartwizard",
            f"-env:UserInstallation={profile_dir.as_uri()}",
            "--infilter=Calc MS Excel 2007 XML",
            "--convert-to", "xlsx",
            "--outdir", tmpdir,
            str(tmp_input),
        ]

        extra_kwargs: dict[str, Any] = {}
        if sys.platform == "win32":
            extra_kwargs["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)

        try:
            proc = subprocess.run(cmd, capture_output=True, timeout=timeout, **extra_kwargs)
        except subprocess.TimeoutExpired:
            return RecalcResult(
                success=False,
                message=f"LibreOffice timed out after {timeout}s. File was saved but not recalculated.",
            )

        if proc.returncode != 0:
            stderr = proc.stderr.decode(errors="replace").strip()
            return RecalcResult(
                success=False,
                message=f"LibreOffice exited with code {proc.returncode}: {stderr}",
            )

        tmp_output = Path(tmpdir) / f"{tmp_input.stem}.xlsx"
        if not tmp_output.is_file():
            return RecalcResult(
                success=False,
                message="LibreOffice reported success but produced no output file.",
            )

        errors_found = scan_formula_errors(str(tmp_output))

        # Stage in src's own directory (os.replace needs same filesystem as tmpdir
        # may not guarantee) then atomically swap — never leaves `src` half-written.
        staged = src.parent / f"{src.name}.tmp-{os.getpid()}"
        try:
            shutil.copy(tmp_output, staged)
            try:
                safe_replace(staged, src)
            except FileInUseError:
                return RecalcResult(
                    success=False,
                    message=f"LibreOffice recalculated formulas, but could not overwrite '{src}' because it is locked by another application.",
                )
            except PermissionError as exc:
                return RecalcResult(
                    success=False,
                    message=f"LibreOffice recalculated formulas, but permission was denied when overwriting '{src}': {exc}",
                )
        finally:
            staged.unlink(missing_ok=True)

        return RecalcResult(success=True, errors_found=errors_found, message="Recalculated with LibreOffice headless.")
