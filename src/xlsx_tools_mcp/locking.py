import os
import time
from contextlib import contextmanager
from pathlib import Path
from typing import Iterator

from filelock import FileLock, Timeout

from .errors import FileInUseError, LockTimeoutError
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


def safe_replace(src: str | Path, dst: str | Path, max_attempts: int = 4, delay: float = 0.05) -> None:
    """Atomically replace dst with src, retrying transient Windows locks.

    On Windows NTFS, antivirus scanners (e.g. Windows Defender) or search indexers
    briefly open newly written files, causing MoveFileExW to fail with
    ERROR_SHARING_VIOLATION (winerror 32) or ERROR_LOCK_VIOLATION (winerror 33).
    A short exponential backoff succeeds once the transient lock releases.
    If persistently locked by another process (winerror 32/33), raises FileInUseError.
    Other permission errors (e.g. read-only file) are raised immediately.
    """
    for attempt in range(max_attempts):
        try:
            os.replace(src, dst)
            return
        except PermissionError as exc:
            winerror = getattr(exc, "winerror", None)
            if winerror in (32, 33):
                if attempt < max_attempts - 1:
                    time.sleep(delay * (2 ** attempt))
                    continue
                raise FileInUseError(
                    f"Workbook '{dst}' is locked by another process (e.g. Microsoft Excel or Antivirus). "
                    "Please close the file and retry."
                ) from exc
            raise
