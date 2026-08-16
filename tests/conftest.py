import pytest


@pytest.fixture(autouse=True)
def _skip_live_soffice(monkeypatch):
    """All write tests must not shell out to real LibreOffice.
    recalculate() already degrades gracefully (returns success=False) when
    find_soffice() is None; this guarantees that path regardless of the host."""
    monkeypatch.setattr("xlsx_mcp.recalc.find_soffice", lambda: None)