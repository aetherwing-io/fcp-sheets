"""Response formatting utilities for sheets queries."""

from __future__ import annotations

from fcp_sheets.model.refs import index_to_col


def format_cell_addr(col: int, row: int) -> str:
    """Format a (col, row) pair as an A1-style address."""
    return f"{index_to_col(col)}{row}"


def format_range(
    min_row: int, min_col: int, max_row: int, max_col: int
) -> str:
    """Format a bounding rectangle as A1:Z99."""
    return f"{index_to_col(min_col)}{min_row}:{index_to_col(max_col)}{max_row}"


def truncate_list(items: list[str], max_items: int = 8) -> str:
    """Join items, showing 'and N more' if truncated."""
    if len(items) <= max_items:
        return ", ".join(items)
    shown = ", ".join(items[:max_items])
    return f"{shown} ... +{len(items) - max_items} more"
