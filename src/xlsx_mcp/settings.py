import os
from pathlib import Path

from .errors import FileNotConfiguredError

# Max seconds to wait for the LibreOffice headless recalculation pass.
SOFFICE_TIMEOUT_SECONDS = int(os.environ.get("XLSX_MCP_RECALC_TIMEOUT", "60"))

# Max seconds to wait to acquire the per-file lock before failing a read/write.
LOCK_TIMEOUT_SECONDS = float(os.environ.get("XLSX_MCP_LOCK_TIMEOUT", "10"))


def _parse_configured_files(raw: str) -> dict[str, str]:
    """Parse `name=path` or bare `path` entries (comma-separated) into alias -> absolute path."""
    files: dict[str, str] = {}
    for entry in raw.split(","):
        entry = entry.strip()
        if not entry:
            continue
        name, sep, path = entry.partition("=")
        path = path or name
        resolved = str(Path(path.strip()).expanduser().resolve())
        alias = name.strip() if sep else Path(resolved).name
        files[alias] = resolved
    return files


# Files preloaded at startup so tools can be called without hunting for a path.
# Format: "e2e.xlsx=/abs/path/to/e2e.xlsx,report=/abs/path/report.xlsx" or just
# bare absolute paths (alias defaults to the filename).
CONFIGURED_FILES: dict[str, str] = _parse_configured_files(os.environ.get("XLSX_MCP_FILES", ""))


def resolve_path(path: str | None) -> str:
    """Resolve a tool's `path` argument against the preconfigured file set.

    - `path` omitted: use the single configured file, or fail listing choices if there
      are zero or multiple.
    - `path` matches a configured alias or filename: use its preloaded absolute path.
    - otherwise: use `path` as given (absolute or relative to cwd).
    """
    if path is None:
        if len(CONFIGURED_FILES) == 1:
            return next(iter(CONFIGURED_FILES.values()))
        if not CONFIGURED_FILES:
            raise FileNotConfiguredError(
                "No path given and no files configured via XLSX_MCP_FILES."
            )
        raise FileNotConfiguredError(
            f"path required — multiple files configured: {', '.join(sorted(CONFIGURED_FILES))}"
        )
    if path in CONFIGURED_FILES:
        return CONFIGURED_FILES[path]
    name = Path(path).name
    if name in CONFIGURED_FILES:
        return CONFIGURED_FILES[name]
    return path
