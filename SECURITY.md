# Security Policy

## Reporting a vulnerability

Please disclose privately rather than opening a public issue. Open a GitHub issue
marked `security` on the repository, or contact the maintainers directly so details
are not exposed before a fix is released.

We aim to acknowledge reports promptly and keep you updated as a fix is prepared
and published. Do not create an exploit that impacts other users; a minimal
proof-of-concept is sufficient.

## Scope

This tool can **read and write arbitrary files on the machine it runs on**, at the
paths the calling agent/user points it at. It is a capability of the tool, not a
vulnerability in it — but treat it accordingly:

- Run the server **only against files you trust** and with least-privilege accounts.
- Do not expose it to untrusted agents with unrestricted filesystem access.
- A set of workbook files can be whitelisted/aliased via `XLSX_MCP_FILES` for
  controlled, path-free access.

## Built-in protections

- **XML-bomb protection** — the `defusedxml` dependency is auto-detected by
  openpyxl, which uses its hardened XML parser to prevent entity-expansion
  (billion-laughs-style) attacks from hostile `.xlsx` files.
- **Concurrency safety** — per-file locking via `filelock` (a sibling `<path>.lock`)
  serializes concurrent read/write access so writes cannot interleave and corrupt
  a workbook.
- **Atomic writes** — saves go to a temp file and are swapped in with `os.replace`,
  so an interrupted write never leaves a file half-written.