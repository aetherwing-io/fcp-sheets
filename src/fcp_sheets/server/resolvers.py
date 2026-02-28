"""Operation context and reference resolution for sheets verbs.

Provides SheetsOpContext and helpers for resolving cell references,
anchors (C6), and selectors.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Iterator

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from fcp_sheets.model.index import SheetIndex
from fcp_sheets.model.refs import (
    AnchorRef,
    CellRef,
    ColRef,
    RangeRef,
    RowRef,
    col_to_index,
    index_to_col,
    parse_anchor,
    parse_cell_ref,
    parse_range_ref,
)


@dataclass
class SheetsOpContext:
    """Context passed to every verb handler."""

    wb: Workbook
    index: SheetIndex
    named_styles: dict[str, dict]

    @property
    def active_sheet(self) -> Worksheet:
        """The currently active worksheet."""
        return self.wb.active  # type: ignore[return-value]

    @property
    def active_sheet_name(self) -> str:
        """Name of the active worksheet."""
        ws = self.wb.active
        return ws.title if ws else ""


def resolve_cell_ref(ref_str: str, ctx: SheetsOpContext) -> tuple[int, int] | None:
    """Resolve a cell reference string to (col, row).

    Handles standard A1 refs AND C6 spatial anchors.
    Returns None if ref cannot be parsed.
    """
    ref_str = ref_str.strip()

    # Try anchor first (C6)
    anchor = parse_anchor(ref_str)
    if anchor:
        return resolve_anchor(anchor, ctx)

    # Standard A1 ref
    cell = parse_cell_ref(ref_str)
    if cell:
        return (cell.col, cell.row)

    return None


def resolve_anchor(anchor: AnchorRef, ctx: SheetsOpContext) -> tuple[int, int] | None:
    """Resolve a spatial anchor to (col, row) based on current data bounds (C6).

    Anchors:
      @bottom_left   → (min_col, max_row + 1 + offset)
      @bottom_right  → (max_col, max_row + 1 + offset)
      @right_top     → (max_col + 1 + offset, min_row)
    """
    sheet_name = ctx.active_sheet_name
    bounds = ctx.index.get_bounds(sheet_name)

    if bounds is None:
        # No data — treat as (1, 1) + offset
        if anchor.anchor in ("bottom_left", "bottom_right"):
            return (1, 1 + anchor.offset)
        else:  # right_top
            return (1 + anchor.offset, 1)

    min_row, min_col, max_row, max_col = bounds

    if anchor.anchor == "bottom_left":
        return (min_col, max_row + 1 + anchor.offset)
    elif anchor.anchor == "bottom_right":
        return (max_col, max_row + 1 + anchor.offset)
    elif anchor.anchor == "right_top":
        return (max_col + 1 + anchor.offset, min_row)

    return None


def get_target_sheet(ref_str: str, ctx: SheetsOpContext) -> tuple[Worksheet, str]:
    """Extract target sheet from a cross-sheet reference.

    Returns (worksheet, ref_without_sheet_prefix).
    """
    if "!" in ref_str:
        sheet_name, _, rest = ref_str.partition("!")
        sheet_name = sheet_name.strip("'\"")
        if sheet_name in ctx.wb.sheetnames:
            return ctx.wb[sheet_name], rest
    return ctx.active_sheet, ref_str


def resolve_range_to_cells(
    range_str: str, ctx: SheetsOpContext
) -> Iterator[tuple[Worksheet, int, int]]:
    """Resolve a range string to an iterator of (worksheet, row, col) tuples."""
    ws, ref = get_target_sheet(range_str, ctx)

    # Single cell
    cell = parse_cell_ref(ref)
    if cell:
        yield (ws, cell.row, cell.col)
        return

    # Range
    range_ref = parse_range_ref(ref)
    if isinstance(range_ref, RangeRef):
        for row in range(range_ref.start.row, range_ref.end.row + 1):
            for col in range(range_ref.start.col, range_ref.end.col + 1):
                yield (ws, row, col)
        return

    if isinstance(range_ref, ColRef):
        # Use data bounds to limit iteration
        sheet_name = ws.title
        bounds = ctx.index.get_bounds(sheet_name)
        if bounds:
            min_row, _, max_row, _ = bounds
            for row in range(min_row, max_row + 1):
                for col in range(range_ref.start_col, range_ref.end_col + 1):
                    yield (ws, row, col)
        return

    if isinstance(range_ref, RowRef):
        bounds = ctx.index.get_bounds(ws.title)
        if bounds:
            _, min_col, _, max_col = bounds
            for row in range(range_ref.start_row, range_ref.end_row + 1):
                for col in range(min_col, max_col + 1):
                    yield (ws, row, col)
        return
