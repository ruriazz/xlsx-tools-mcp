class XlsxMcpError(Exception):
    """Base exception for all xlsx-tools-mcp domain errors."""


class SheetNotFoundError(XlsxMcpError):
    """Raised when a requested sheet name does not exist in the workbook."""


class LockTimeoutError(XlsxMcpError):
    """Raised when a file lock could not be acquired in time."""


class FileNotConfiguredError(XlsxMcpError):
    """Raised when `path` is omitted but zero or multiple files are preconfigured."""
