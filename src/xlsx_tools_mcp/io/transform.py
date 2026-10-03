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


import ast

_ALLOWED_AST_NODES = (
    ast.Expression,
    ast.BoolOp,
    ast.BinOp,
    ast.UnaryOp,
    ast.Compare,
    ast.Name,
    ast.Constant,
    ast.Load,
    ast.And,
    ast.Or,
    ast.Not,
    ast.Eq,
    ast.NotEq,
    ast.Lt,
    ast.LtE,
    ast.Gt,
    ast.GtE,
    ast.In,
    ast.NotIn,
    ast.List,
    ast.Tuple,
    ast.Add,
    ast.Sub,
    ast.Mult,
    ast.Div,
    ast.FloorDiv,
    ast.Mod,
    ast.Pow,
    ast.BitAnd,
    ast.BitOr,
    ast.BitXor,
    ast.Invert,
)


def _validate_safe_query(filter_query: str) -> None:
    if "@" in filter_query:
        raise ValueError("Invalid query: environment variable references ('@') are not permitted.")
    try:
        tree = ast.parse(filter_query, mode="eval")
    except SyntaxError as e:
        raise ValueError(f"Invalid query syntax: {e}") from e

    for node in ast.walk(tree):
        if not isinstance(node, _ALLOWED_AST_NODES):
            raise ValueError(f"Invalid query: operation '{type(node).__name__}' is not permitted.")
        if isinstance(node, ast.Name) and node.id.startswith("__"):
            raise ValueError(f"Invalid query: private/dunder names ('{node.id}') are not permitted.")


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
    if max_rows < 0:
        raise ValueError(f"max_rows must be non-negative, got {max_rows}")

    _validate_safe_query(filter_query)

    rows = reader.read_sheet(path, sheet, cell_range=cell_range)["rows"]
    if not rows:
        return {"columns": [], "records": [], "row_count": 0, "limited": False}

    header, *body = rows
    seen_cols: dict[str, int] = {}
    df_columns: list[str] = []
    for i, h in enumerate(header):
        name = str(h).strip() if (h is not None and str(h).strip() != "") else f"col_{i}"
        if name in seen_cols:
            seen_cols[name] += 1
            df_columns.append(f"{name}_{seen_cols[name]}")
        else:
            seen_cols[name] = 0
            df_columns.append(name)
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

