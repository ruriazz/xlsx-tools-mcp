class XlsxMcpError(Exception):
    """Base exception for all xlsx-tools-mcp domain errors."""


class SheetNotFoundError(XlsxMcpError):
    """Raised when a requested sheet name does not exist in the workbook."""


class LockTimeoutError(XlsxMcpError):
    """Raised when a file lock could not be acquired in time."""


class FileNotConfiguredError(XlsxMcpError):
    """Raised when `path` is omitted but zero or multiple files are preconfigured."""


class AccessDeniedError(XlsxMcpError):
    """Raised when path is outside the allowed directories boundaries."""


class UnsafeFormulaError(XlsxMcpError):
    """Raised when an unsafe formula function is detected."""


class FileInUseError(XlsxMcpError):
    """Raised when a file cannot be saved/overwritten because it is locked by another process."""


class WorkbookNotFoundError(XlsxMcpError):
    """Raised when a requested workbook file does not exist on disk."""


class InvalidWorkbookError(XlsxMcpError):
    """Raised when a file exists but is not a valid or readable Excel workbook."""


class InvalidRulesError(XlsxMcpError):
    """Raised when a workbook's rule set is malformed or cannot be read."""


class ReservedSheetError(XlsxMcpError):
    """Raised when a tool targets a sheet name reserved for internal use."""


class BaselineNotFoundError(XlsxMcpError):
    """Raised when no baseline workbook is available to compare against."""


class InjectionError(XlsxMcpError):
    """Raised when computed formula values cannot be written into a workbook package."""

