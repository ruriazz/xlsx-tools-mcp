# Security Policy

## Reporting a vulnerability

Please disclose privately rather than opening a public issue. Open a GitHub issue
marked `security` on the repository, or contact the maintainers directly so details
are not exposed before a fix is released.

We aim to acknowledge reports promptly and keep you updated as a fix is prepared
and published. Do not create an exploit that impacts other users; a minimal
proof-of-concept is sufficient.

## Scope

By default in local mode, this server can read and write files accessible to the host
process. For autonomous agent setups or production deployments, built-in security
controls allow strict containment:

- **Path Confinement** — set `XLSX_MCP_ALLOWED_DIRS` to restrict read and write
  operations to specific directories. Any access outside these boundaries (or via
  symlink traversal) is rejected with `AccessDeniedError`.
- **Preloaded Workbooks** — set `XLSX_MCP_FILES` to alias specific workbooks for
  path-free access, preventing client agents from manipulating file paths.
- **Run with least-privilege** — execute the server under dedicated user accounts
  with restricted filesystem permissions.

## Built-in protections

- **Path confinement & traversal defense** — `resolve_path()` resolves real symlinks
  and verifies paths against `XLSX_MCP_ALLOWED_DIRS` using platform-appropriate
  comparisons (`os.path.normcase` for Windows case-insensitivity, cross-drive checks).
- **Formula injection & DDE guard** — write operations (`write_cells`, formulas)
  strictly inspect content starting with `=`, `+`, `-`, `@`, or containing DDE pipe syntax (`|`).
  Potentially dangerous functions (`WEBSERVICE`, `HYPERLINK`, `INDIRECT`, `RTD`,
  `CALL`, `REGISTER`) are rejected with `UnsafeFormulaError` unless explicitly permitted
  with `allow_external_formulas=True`.
- **Query expression sandboxing** — `query_sheet` uses an AST validator (`SafeQueryValidator`)
  that restricts pandas queries to safe comparison and logical expressions, blocking
  arbitrary Python code execution, built-in access, and method invocations.
- **Auto-backup & atomic rollback** — write mutations can create `.bak` snapshots
  prior to mutation (`backup=True` or `XLSX_MCP_AUTO_BACKUP`), with atomic rollback
  available via `restore_backup`.
- **Safe atomic replacement & lock resilience** — saves are staged in temporary files
  and swapped atomically via `safe_replace()` with exponential backoff retry for
  transient locks (e.g. Windows Defender scans or file indexers), and clean
  `FileInUseError` signaling if a file is permanently locked by another application.
- **XML-bomb protection** — `defusedxml` is integrated with openpyxl's XML parser
  to prevent entity-expansion (billion-laughs) denial-of-service from hostile spreadsheets.
- **Concurrency serialization** — per-file locking via `filelock` (sibling `<path>.lock`)
  serializes concurrent read/write operations to prevent file corruption.