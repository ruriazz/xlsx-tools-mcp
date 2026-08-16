import pytest
from filelock import FileLock

from xlsx_tools_mcp.errors import LockTimeoutError
from xlsx_tools_mcp.locking import file_lock


def test_file_lock_times_out_when_already_held(tmp_path):
    path = str(tmp_path / "book.xlsx")
    lock = FileLock(f"{path}.lock")
    with lock:
        with pytest.raises(LockTimeoutError):
            with file_lock(path, timeout=0.01):
                pass
    with file_lock(path, timeout=0.01):
        pass