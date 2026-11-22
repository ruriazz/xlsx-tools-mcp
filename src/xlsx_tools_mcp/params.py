"""Shared parameter types for the MCP tool signatures.

Each alias carries its description into the tool's JSON input schema, so an agent
sees what a parameter means right next to its type instead of in a prose block.
Defining them once keeps the wording identical across all tools.
"""

from __future__ import annotations

from typing import Annotated, Literal

from pydantic import Field

Path = Annotated[
    str | None,
    Field(
        description="File path or preloaded name (list_configured_files). Omit if only one file is preloaded."
    ),
]
Sheet = Annotated[str, Field(description="Sheet name (case-sensitive).")]
CellRange = Annotated[str, Field(description='A1-style range, e.g. "A1:D10".')]
OptionalCellRange = Annotated[
    str | None,
    Field(description='Range to read, e.g. "A1:F500". Omit for the whole sheet.'),
]
Backup = Annotated[
    bool | None,
    Field(description="Save a .bak copy first. Default: XLSX_MCP_AUTO_BACKUP (true)."),
]
Reason = Annotated[
    str | None,
    Field(description="Why; kept in the change log (XLSX_MCP_CHANGE_LOG=true)."),
]
Recalculate = Annotated[
    bool,
    Field(
        description="Recompute formulas with LibreOffice after saving. False: faster, formula results stay empty."
    ),
]
AllowExternalFormulas = Annotated[
    bool,
    Field(
        description="Allow WEBSERVICE, HYPERLINK, INDIRECT, RTD, CALL, REGISTER and DDE ('|') in formulas."
    ),
]
CreateSheetIfMissing = Annotated[bool, Field(description="Create the sheet if missing.")]
Overwrite = Annotated[bool, Field(description="Replace an existing output file.")]

ReadFormat = Literal["array", "records", "markdown"]
RecalcMode = Literal["inject", "roundtrip"]
# Kept as str: the writer matches it case-insensitively since v0.2, so the enum is
# advertised in the schema without rejecting "Bar" at validation time.
ChartType = Annotated[
    str,
    Field(
        description='Chart kind: "bar", "line", "pie" or "scatter".',
        json_schema_extra={"enum": ["bar", "line", "pie", "scatter"]},
    ),
]
OnViolation = Literal["rollback", "warn"]
OnUnverifiable = Literal["warn", "rollback"]
Sanitizable = Literal["comments", "document_properties", "hidden_sheets", "external_links", "change_log", "macros"]
