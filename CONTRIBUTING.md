# Contributing

Thanks for contributing to `xlsx-tools-mcp`. This project is small and deliberately
kept that way — readable, tested, and decoupled. Please read the notes below before
opening a PR.

## Setup

```bash
uv sync          # installs project + dev (pytest) dependencies
```

## Running tests

```bash
uv run pytest
```

All tests must pass before a PR is merged. The test suite (100+ tests) covers:
- The IO layer (`tests/test_reader.py`, `test_writer.py`, `test_transform.py`)
- Sheet profiling and safe pandas querying (`test_reader_profile.py`, `test_transform_query.py`)
- Tables and native charts (`test_writer_table_chart.py`)
- Cross-platform path confinement and multi-delimiter parsing (`test_settings.py`)
- LibreOffice headless detection, isolated profiles, and recalc (`test_recalc.py`)
- Safe file locking, retry backoff, and concurrency (`test_locking.py`, `test_writer.py`)
- Full cross-platform end-to-end integration (`test_e2e.py`, `test_e2e_all_tools.py`, `test_e2e_new_tools.py`)
- Package/XML access and value injection (`test_ooxml.py`, `test_inject.py`)
- Rules, rollback and ripple (`test_rules.py`, `test_ripple.py`)
- Review copies, sharing inspection/sanitizing and the change log (`test_review.py`, `test_sharing.py`, `test_ledger.py`)

`conftest.py` monkeypatches `find_soffice` so tests never require a real LibreOffice installation.
`tests/helpers.py` provides a fake LibreOffice that evaluates simple formulas, for tests that
need real recalculated values.

### LibreOffice integration tests

Tests marked `@pytest.mark.soffice` (`tests/test_recalc_soffice.py`) run against a real
LibreOffice and are skipped when it is not installed:

```bash
uv run pytest -m soffice -rs
```

CI runs them in a separate `soffice-integration` job on Ubuntu.

## Multi-OS CI Matrix

GitHub Actions runs the test suite across a matrix of 9 configurations:
- **Operating Systems**: `ubuntu-latest`, `macos-latest`, `windows-latest`
- **Python Versions**: `3.10`, `3.11`, `3.12`

When writing new features or tests:
- Always use `pathlib.Path` or `os.path` functions rather than hardcoded `/` or `\\` path strings.
- Support both comma (`,`) and semicolon (`;`) delimiters when dealing with environment list variables.
- Remember Windows path case-insensitivity (`os.path.normcase`) and multi-drive topologies.

## Adding tools

- Add the core logic in `src/xlsx_tools_mcp/io/` (e.g. `reader.py`, `writer.py`, or `transform.py`) — keep the IO layer **decoupled from MCP**.
- For write operations:
  - Route saves through `_finalize()` with a `WriteContext` so writes are atomic, recalculation (`errors_found`) stays consistent, rules/ripple apply, and the change log records the call. Set `changes_values=True` when the tool can change cell values or formulas, `structural=True` when it shifts coordinates, and `ctx.record(...)` each cell written.
  - Accept an optional `reason: str | None = None` and pass it into the `WriteContext`.
  - Call `ledger.assert_not_reserved()` on every sheet name the tool targets or creates.
  - Support `backup: bool = False` (or respect `XLSX_MCP_AUTO_BACKUP`) before overwriting existing workbooks.
  - Validate formula inputs against `_validate_formula()` to block unsafe external functions and DDE injections.
- Add a thin wrapper in `src/xlsx_tools_mcp/server.py` registered with `@_tool(title, read_only=..., destructive=..., idempotent=...)`. It resolves the path via `resolve_path()`, takes the lock via `file_lock`, and translates domain errors via `_run()`. Set `destructive=True` when the call can remove or overwrite cells, sheets, rules or files.
- Describe every parameter in the schema: reuse the shared types in `src/xlsx_tools_mcp/params.py` (`P.Path`, `P.Sheet`, `P.Backup`, ...) or use `Annotated[type, Field(description=...)]`, and `Literal[...]` for fixed choices. Keep the docstring to what the tool does and which neighboring tool to use instead; no `Args:` block. `tests/test_tool_surface.py` checks this.
- If the tool overlaps an existing one, update the tool-choice guide in `INSTRUCTIONS` (`server.py`).
- Add unit tests in `tests/` exercising the IO function directly.
- Add integration assertions to `tests/test_e2e_all_tools.py` ensuring the MCP tool entrypoint executes seamlessly.
- Update the tool table and documentation in `README.md`.

## Guidelines

- Match the existing code style (type hints everywhere, short docstrings).
- Writes must use the `_finalize` path so saves are atomic and recalculation (`errors_found`) stays consistent.
- Raise the module's own domain errors (`src/xlsx_tools_mcp/errors.py`) at API boundaries; let `server.py` translate them into MCP errors.
- Never introduce dependencies without evaluating security and cross-platform implications.

## Commit conventions

Concise, factual commit messages following Conventional Commits prefixes (e.g. `feat:`, `fix:`, `docs:`, `test:`, `chore:`, `refactor:`). Use imperative mood (e.g. `add`, not `added`). One logical change per commit.

## PR checklist

- [ ] Tests pass locally (`uv run pytest`)
- [ ] Cross-platform compatibility maintained (Linux, macOS, Windows)
- [ ] Changes to IO/server behavior are covered by unit and E2E tests
- [ ] Write operations adhere to formula validation, auto-backup, and atomic save conventions
- [ ] Tool reference and feature claims in `README.md` match the code
- [ ] Relevant documentation (`CHANGELOG.md`, `SECURITY.md`) updated if security or behavior changed