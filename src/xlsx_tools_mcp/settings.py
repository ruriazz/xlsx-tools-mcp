import os
from pathlib import Path
import re
import sys

from .errors import FileNotConfiguredError, AccessDeniedError

IS_WINDOWS = sys.platform == "win32"

# Max seconds to wait for the LibreOffice headless recalculation pass.
SOFFICE_TIMEOUT_SECONDS = int(os.environ.get("XLSX_MCP_RECALC_TIMEOUT", "60"))

# Max seconds to wait to acquire the per-file lock before failing a read/write.
LOCK_TIMEOUT_SECONDS = float(os.environ.get("XLSX_MCP_LOCK_TIMEOUT", "10"))

AUTO_BACKUP = os.environ.get("XLSX_MCP_AUTO_BACKUP", "true").lower() in ("1", "true", "yes")
BACKUP_DIR = str(Path(os.environ["XLSX_MCP_BACKUP_DIR"]).expanduser().resolve()) if os.environ.get("XLSX_MCP_BACKUP_DIR") else ""


def _env_bool(name: str, default: bool) -> bool:
    raw = os.environ.get(name, "").strip().lower()
    if not raw:
        return default
    if raw in ("1", "true", "yes", "on"):
        return True
    if raw in ("0", "false", "no", "off"):
        return False
    raise ValueError(f"{name} must be true or false, got {raw!r}")


def _env_int(name: str, default: int, minimum: int = 0) -> int:
    raw = os.environ.get(name, "").strip()
    if not raw:
        return default
    try:
        value = int(raw)
    except ValueError as exc:
        raise ValueError(f"{name} must be an integer, got {raw!r}") from exc
    if value < minimum:
        raise ValueError(f"{name} must be >= {minimum}, got {value}")
    return value


RECALC_MODES = ("inject", "roundtrip")


def parse_recalc_mode(raw: str) -> str:
    mode = raw.strip().lower() or "inject"
    if mode not in RECALC_MODES:
        raise ValueError(f"XLSX_MCP_RECALC_MODE must be one of {RECALC_MODES}, got {raw!r}")
    return mode


# How formula results are written back after the LibreOffice pass:
# "inject" copies only computed values into the original file; "roundtrip" replaces
# the file with LibreOffice's own export (v0.3.x behavior).
RECALC_MODE = parse_recalc_mode(os.environ.get("XLSX_MCP_RECALC_MODE", "inject"))

# Upper bound on the uncompressed size of a single XML part parsed by the server.
MAX_PART_BYTES = _env_int("XLSX_MCP_MAX_PART_BYTES", 256 * 1024 * 1024, minimum=1)

# Report which formula results changed as a side effect of a write.
RIPPLE_ENABLED = _env_bool("XLSX_MCP_RIPPLE", True)
RIPPLE_MAX_CELLS = _env_int("XLSX_MCP_RIPPLE_MAX_CELLS", 200_000, minimum=0)

# Opt-in change log stored inside the workbook on a very hidden sheet.
CHANGE_LOG_ENABLED = _env_bool("XLSX_MCP_CHANGE_LOG", False)
CHANGE_LOG_MAX_ROWS = _env_int("XLSX_MCP_CHANGE_LOG_MAX_ROWS", 10_000, minimum=1)
CHANGE_LOG_MAX_CELLS_PER_OP = _env_int("XLSX_MCP_CHANGE_LOG_MAX_CELLS_PER_OP", 500, minimum=1)


def _split_env_list(raw: str) -> list[str]:
    return [entry.strip() for entry in re.split(r"[,;]", raw) if entry.strip()]

def _parse_allowed_dirs(raw: str) -> list[str]:
    dirs = []
    for entry in _split_env_list(raw):
        dirs.append(str(Path(entry).expanduser().resolve()))
    return dirs

ALLOWED_DIRS = _parse_allowed_dirs(os.environ.get("XLSX_MCP_ALLOWED_DIRS", ""))
if BACKUP_DIR and ALLOWED_DIRS and BACKUP_DIR not in ALLOWED_DIRS:
    ALLOWED_DIRS.append(BACKUP_DIR)

def _parse_configured_files(raw: str) -> dict[str, str]:
    """Parse `name=path` or bare `path` entries (comma- or semicolon-separated) into alias -> absolute path."""
    files: dict[str, str] = {}
    for entry in _split_env_list(raw):
        name, sep, path = entry.partition("=")
        path = path or name
        resolved = str(Path(path.strip()).expanduser().resolve())
        alias = name.strip() if sep else Path(resolved).name
        files[alias] = resolved
    return files


# Files preloaded at startup so tools can be called without hunting for a path.
# Format: "e2e.xlsx=/abs/path/to/e2e.xlsx,report=/abs/path/report.xlsx" or semicolon-separated
# bare absolute paths (alias defaults to the filename).
CONFIGURED_FILES: dict[str, str] = _parse_configured_files(os.environ.get("XLSX_MCP_FILES", ""))


def _check_allowed(resolved_path: str) -> str:
    if not ALLOWED_DIRS:
        return resolved_path
    
    rp = str(Path(resolved_path).resolve())
    for allowed in ALLOWED_DIRS:
        try:
            common = os.path.commonpath([rp, allowed])
            if os.path.normcase(common) == os.path.normcase(allowed):
                return resolved_path
        except ValueError:
            # Raised when paths are on different drives on Windows (e.g. C: vs D:)
            pass
    raise AccessDeniedError(f"Access denied: Path {resolved_path} is outside allowed directories.")

def resolve_literal_path(path: str) -> str:
    """Resolve a path exactly as given (no preloaded-alias lookup), then confine it.

    Used for files a tool *creates* or compares against, so an output name that
    happens to match a preloaded file's alias can never redirect to that file.
    """
    return _check_allowed(str(Path(path).expanduser().resolve()))


def resolve_path(path: str | None) -> str:
    """Resolve a tool's `path` argument against the preconfigured file set.

    - `path` omitted: use the single configured file, or fail listing choices if there
      are zero or multiple.
    - `path` matches a configured alias or filename: use its preloaded absolute path.
    - otherwise: use `path` as given (absolute or relative to cwd).
    """
    resolved = None
    if path is None:
        if len(CONFIGURED_FILES) == 1:
            resolved = next(iter(CONFIGURED_FILES.values()))
        elif not CONFIGURED_FILES:
            raise FileNotConfiguredError(
                "No path given and no files configured via XLSX_MCP_FILES."
            )
        else:
            raise FileNotConfiguredError(
                f"path required — multiple files configured: {', '.join(sorted(CONFIGURED_FILES))}"
            )
    elif path in CONFIGURED_FILES:
        resolved = CONFIGURED_FILES[path]
    else:
        name = Path(path).name
        if name in CONFIGURED_FILES:
            resolved = CONFIGURED_FILES[name]
        elif IS_WINDOWS:
            norm_path = os.path.normcase(path)
            norm_name = os.path.normcase(name)
            match = None
            for k, v in CONFIGURED_FILES.items():
                if os.path.normcase(k) in (norm_path, norm_name):
                    match = v
                    break
            resolved = match if match is not None else str(Path(path).expanduser().resolve())
        else:
            resolved = str(Path(path).expanduser().resolve())
    
    return _check_allowed(resolved)
