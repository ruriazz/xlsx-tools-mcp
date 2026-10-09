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
- Full 29-tool cross-platform end-to-end integration (`test_e2e.py`, `test_e2e_all_tools.py`)

`conftest.py` monkeypatches `find_soffice` so tests never require a real LibreOffice installation.

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
  - Route saves through `_finalize()` so writes are atomic and recalculation (`errors_found`) remains consistent.
  - Support `backup: bool = False` (or respect `XLSX_MCP_AUTO_BACKUP`) before overwriting existing workbooks.
  - Validate formula inputs against `_validate_formula()` to block unsafe external functions and DDE injections.
- Add a thin `@mcp.tool()` wrapper in `src/xlsx_tools_mcp/server.py` that resolves the path via `resolve_path()`, takes the lock via `file_lock`, and translates domain errors via `_run()`.
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