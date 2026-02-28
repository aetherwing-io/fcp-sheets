"""Cell operation handlers — set, data, fill, clear."""

from __future__ import annotations

from fcp_core import OpResult, ParsedOp

from fcp_sheets.lib.number_formats import resolve_format
from fcp_sheets.model.refs import index_to_col, parse_cell_ref
from fcp_sheets.server.resolvers import SheetsOpContext, resolve_cell_ref


def op_set(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    """Set a single cell value.

    Syntax: set CELL VALUE [fmt:FORMAT]

    Values starting with = are formulas.
    Numeric strings are converted to numbers.
    Quoted strings stay as text.
    """
    if len(op.positionals) < 2:
        return OpResult(success=False, message="Usage: set CELL VALUE [fmt:FORMAT]")

    ref_str = op.positionals[0]
    value_str = op.positionals[1]

    # Resolve cell reference (supports A1 and C6 anchors)
    resolved = resolve_cell_ref(ref_str, ctx)
    if resolved is None:
        return OpResult(success=False, message=f"Invalid cell reference: {ref_str!r}")

    col, row = resolved
    ws = ctx.active_sheet

    # Parse value
    value = _parse_cell_value(value_str)

    # Set the cell
    cell = ws.cell(row=row, column=col, value=value)

    # Apply number format if specified
    fmt = op.params.get("fmt")
    if fmt:
        cell.number_format = resolve_format(fmt)

    # Update index bounds
    ctx.index.expand_bounds(ws.title, row, col)
    addr = f"{index_to_col(col)}{row}"
    ctx.index.record_modified(ws.title, addr)

    # Format response
    display = repr(value) if isinstance(value, str) and not value.startswith("=") else str(value)
    return OpResult(success=True, message=f"{addr} = {display}", prefix="+")


def op_data(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    """Data block mode — stub for Wave 2."""
    raise NotImplementedError("data verb not yet implemented")


def op_fill(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    """Fill/drag formula — stub for Wave 2."""
    raise NotImplementedError("fill verb not yet implemented")


def op_clear(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    """Clear cell contents — stub for Wave 2."""
    raise NotImplementedError("clear verb not yet implemented")


def _parse_cell_value(s: str) -> str | int | float:
    """Parse a raw value string into the appropriate Python type.

    Rules:
    - Starts with '=' → formula (string)
    - Wrapped in quotes → text (strip quotes)
    - Leading zero with length > 1 → text (preserve leading zeros, C1)
    - Valid int → int
    - Valid float → float
    - Everything else → text
    """
    # Formula
    if s.startswith("="):
        return s

    # Quoted string — strip outer quotes
    if (s.startswith('"') and s.endswith('"')) or (s.startswith("'") and s.endswith("'")):
        return s[1:-1]

    # Leading zero protection (C1)
    if len(s) > 1 and s[0] == "0" and s[1:].isdigit():
        return s

    # Try int
    try:
        return int(s)
    except ValueError:
        pass

    # Try float
    try:
        return float(s)
    except ValueError:
        pass

    return s


HANDLERS: dict[str, callable] = {
    "set": op_set,
    "fill": op_fill,
    "clear": op_clear,
    # "data" is handled by adapter block mode, not dispatched here
}
