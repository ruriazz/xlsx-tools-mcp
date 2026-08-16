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

All tests must pass before a PR is merged. The suite covers the io layer
(`tests/test_reader.py`, `test_writer.py`, `test_transform.py`), recalculation,
locking, settings, and domain errors. `conftest.py` monkeypatches `find_soffice`
so tests never need a real LibreOffice install.

## Adding tools

- Add the core logic in `src/xlsx_tools_mcp/io/` (e.g. extend `reader.py` or
  `writer.py`) — keep the io layer **decoupled from MCP**.
- Add a thin `@mcp.tool()` wrapper in `src/xlsx_tools_mcp/server.py` that resolves
  the path via `resolve_path`, takes the lock via `file_lock`, and translates
  domain errors via `_run`.
- Add tests in `tests/` exercising the io function directly.
- Update the tool table in `README.md`.

## Guidelines

- Match the existing code style (type hints everywhere, short docstrings).
- Writes must use the `_finalize` path so saves are atomic and recalculation
  (`errors_found`) stays consistent for every write tool.
- Raise the module's own domain errors (`src/xlsx_tools_mcp/errors.py`) at API
  boundaries; let `server.py` translate them into MCP errors.

## Commit conventions

Concise, factual commit messages, conventional-ish prefixes (e.g. `feat:`,
`fix:`, `docs:`, `test:`, `refactor:`). Use present tense. One logical change
per commit.

## PR checklist

- [ ] Tests pass (`uv run pytest`)
- [ ] Changes to io/server behavior are covered by tests
- [ ] Tool reference and feature claims in `README.md` still match the code
- [ ] No undocumented behavior / feature drift (docs stay in sync)