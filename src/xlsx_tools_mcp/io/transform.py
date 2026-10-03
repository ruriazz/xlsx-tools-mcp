from __future__ import annotations

from typing import Any

import pandas as pd

from . import reader


def aggregate_sheet(
    path: str,
    sheet: str,
    group_by: list[str],
    agg: dict[str, str],
    cell_range: str | None = None,
    has_header: bool = True,
) -> dict[str, Any]:
    """Group and aggregate sheet data with pandas, built on top of `reader.read_sheet`.

    Never reads the file directly with pandas — merged cells and styling in the
    source range would otherwise get silently flattened/lost before pandas sees it.
    """
    rows = reader.read_sheet(path, sheet, cell_range=cell_range)["rows"]
    if not rows:
        return {"columns": [], "records": [], "row_count": 0}

    if has_header:
        header, *body = rows
        columns = [str(h) if h is not None else f"col_{i}" for i, h in enumerate(header)]
    else:
        columns = [f"col_{i}" for i in range(len(rows[0]))]
        body = rows

    df = pd.DataFrame(body, columns=columns)

    missing = [c for c in (*group_by, *agg.keys()) if c not in df.columns]
    if missing:
        raise ValueError(f"Unknown column(s): {missing}. Available columns: {list(df.columns)}")

    grouped = df.groupby(group_by, dropna=False).agg(agg).reset_index()
    return {
        "columns": [str(c) for c in grouped.columns],
        "records": grouped.to_dict(orient="records"),
        "row_count": len(grouped),
    }


def query_sheet(
    path: str,
    sheet: str,
    filter_query: str,
    columns: list[str] | None = None,
    max_rows: int = 100,
    cell_range: str | None = None,
) -> dict[str, Any]:
    """Query sheet data using pandas expressions.

    Args:
        path: Path to the .xlsx file.
        sheet: Sheet name.
        filter_query: Pandas query expression (e.g. "qty > 10 and price < 50").
        columns: Optional subset of column names to return.
        max_rows: Maximum number of rows to return (default 100).
        cell_range: Optional A1-style range to read before querying.

    Returns:
        Dict containing columns, records, total row_count, and whether results were limited.
    """
    rows = reader.read_sheet(path, sheet, cell_range=cell_range)["rows"]
    if not rows:
        return {"columns": [], "records": [], "row_count": 0, "limited": False}

    header, *body = rows
    df_columns = [str(h) if h is not None else f"col_{i}" for i, h in enumerate(header)]
    df = pd.DataFrame(body, columns=df_columns)
    
    try:
        filtered = df.query(filter_query)
    except Exception as e:
        raise ValueError(f"Invalid query: {e}")
        
    if columns:
        missing = [c for c in columns if c not in filtered.columns]
        if missing:
            raise ValueError(f"Unknown column(s): {missing}. Available columns: {list(filtered.columns)}")
        filtered = filtered[columns]
        
    result = filtered.head(max_rows)
    
    return {
        "columns": [str(c) for c in result.columns],
        "records": result.to_dict(orient="records"),
        "row_count": len(filtered),
        "limited": len(filtered) > max_rows
    }

