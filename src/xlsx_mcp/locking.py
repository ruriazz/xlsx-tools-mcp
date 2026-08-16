from __future__ import annotations

from contextlib import contextmanager
from pathlib import Path
from typing import Iterator

from filelock import FileLock, Timeout

from .errors import LockTimeoutError
from .settings import LOCK_TIMEOUT_SECONDS


@contextmanager
def file_lock(path: str | Path, timeout: float = LOCK_TIMEOUT_SECONDS) -> Iterator[None]:
    """Serialize concurrent read/write access to a single xlsx file.

    Uses a sibling `<path>.lock` file so concurrent MCP tool calls (or other
    processes) touching the same workbook never interleave writes and corrupt it.
    """
    lock = FileLock(f"{path}.lock", timeout=timeout)
    try:
        with lock:
            yield
    except Timeout as exc:
        raise LockTimeoutError(
            f"Timed out after {timeout}s waiting for a lock on {path}. "
            "Another operation may be in progress."
        ) from exc
