import pytest


@pytest.fixture(autouse=True)
def _skip_live_soffice(request, monkeypatch):
    """All write tests must not shell out to real LibreOffice.
    recalculate() already degrades gracefully (returns success=False) when
    find_soffice() is None; this guarantees that path regardless of the host.

    Tests marked `@pytest.mark.soffice` opt out and use the real binary."""
    if request.node.get_closest_marker("soffice"):
        return
    monkeypatch.setattr("xlsx_tools_mcp.recalc.find_soffice", lambda: None)
