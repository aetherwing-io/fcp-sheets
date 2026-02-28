"""SheetIndex — lightweight metadata index for selector resolution.

Per C4 (Lazy Selector Evaluation):
- Tracks data bounds per sheet, active sheet, last modified ranges
- NO per-mutation cell scanning
- Bounds expand on writes, full rebuild only on undo/redo/open
"""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from fcp_sheets.model.snapshot import SheetsModel


class SheetIndex:
    """Lazy index tracking sheet metadata for efficient selector resolution."""

    def __init__(self) -> None:
        self._bounds: dict[str, tuple[int, int, int, int]] = {}
        self._active_sheet: str = ""
        self._last_modified: list[tuple[str, str]] = []  # [(sheet, range_str), ...]
        self._named_styles: set[str] = set()

    @property
    def active_sheet(self) -> str:
        return self._active_sheet

    @active_sheet.setter
    def active_sheet(self, name: str) -> None:
        self._active_sheet = name

    def expand_bounds(self, sheet_name: str, row: int, col: int) -> None:
        """Expand data bounds for a sheet after a write operation."""
        if sheet_name in self._bounds:
            mr, mc, xr, xc = self._bounds[sheet_name]
            self._bounds[sheet_name] = (
                min(mr, row), min(mc, col),
                max(xr, row), max(xc, col),
            )
        else:
            self._bounds[sheet_name] = (row, col, row, col)

    def get_bounds(self, sheet_name: str) -> tuple[int, int, int, int] | None:
        """Get data bounds for a sheet: (min_row, min_col, max_row, max_col)."""
        return self._bounds.get(sheet_name)

    def record_modified(self, sheet_name: str, range_str: str) -> None:
        """Record the last modified range (for @recent selector)."""
        self._last_modified.append((sheet_name, range_str))
        # Keep only last 20
        if len(self._last_modified) > 20:
            self._last_modified = self._last_modified[-20:]

    def get_recent(self, count: int = 1) -> list[tuple[str, str]]:
        """Get the last N modified (sheet, range) pairs."""
        return self._last_modified[-count:]

    def remove_sheet(self, sheet_name: str) -> None:
        """Remove a sheet from the index."""
        self._bounds.pop(sheet_name, None)
        self._last_modified = [
            (s, r) for s, r in self._last_modified if s != sheet_name
        ]

    def rename_sheet(self, old_name: str, new_name: str) -> None:
        """Rename a sheet in the index."""
        if old_name in self._bounds:
            self._bounds[new_name] = self._bounds.pop(old_name)
        self._last_modified = [
            (new_name if s == old_name else s, r) for s, r in self._last_modified
        ]
        if self._active_sheet == old_name:
            self._active_sheet = new_name

    def rebuild(self, model: SheetsModel) -> None:
        """Full rebuild from workbook state (after undo/redo/open).

        Uses openpyxl's min/max row/col properties (O(sheets), not O(cells)).
        """
        self._bounds.clear()
        wb = model.wb
        for ws in wb.worksheets:
            if ws.max_row is not None and ws.max_column is not None:
                min_row = ws.min_row or 1
                min_col = ws.min_column or 1
                self._bounds[ws.title] = (min_row, min_col, ws.max_row, ws.max_column)
        self._active_sheet = wb.active.title if wb.active else ""

    def clear(self) -> None:
        """Reset index state."""
        self._bounds.clear()
        self._active_sheet = ""
        self._last_modified.clear()
        self._named_styles.clear()
