"""Cell reference parser — A1 notation, ranges, cross-sheet, and spatial anchors.

Handles:
  A1          → CellRef(col=1, row=1)
  A1:D10      → RangeRef(start, end)
  B:B         → ColRef(start_col=2, end_col=2)
  3:3         → RowRef(start_row=3, end_row=3)
  A:E         → ColRef(start_col=1, end_col=5)
  1:5         → RowRef(start_row=1, end_row=5)
  Sheet2!A1   → CellRef with sheet="Sheet2"
  @bottom_left+2 → Anchor (resolved at runtime via SheetIndex)
"""

from __future__ import annotations

import re
from dataclasses import dataclass


@dataclass
class CellRef:
    """Single cell reference."""
    col: int  # 1-based
    row: int  # 1-based
    sheet: str | None = None


@dataclass
class RangeRef:
    """Rectangular range reference."""
    start: CellRef
    end: CellRef
    sheet: str | None = None


@dataclass
class ColRef:
    """Entire column or column range."""
    start_col: int
    end_col: int
    sheet: str | None = None


@dataclass
class RowRef:
    """Entire row or row range."""
    start_row: int
    end_row: int
    sheet: str | None = None


@dataclass
class AnchorRef:
    """Spatial anchor reference (C6)."""
    anchor: str  # "bottom_left", "bottom_right", "right_top"
    offset: int = 0  # row/col offset


# --- Column conversion ---

def col_to_index(col_str: str) -> int:
    """Convert column letter(s) to 1-based index. A→1, Z→26, AA→27."""
    col_str = col_str.upper()
    result = 0
    for ch in col_str:
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result


def index_to_col(index: int) -> str:
    """Convert 1-based index to column letter(s). 1→A, 26→Z, 27→AA."""
    result = []
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        result.append(chr(remainder + ord("A")))
    return "".join(reversed(result))


# --- Regex patterns ---

_CELL_RE = re.compile(r"^([A-Za-z]+)(\d+)$")
_COL_RE = re.compile(r"^([A-Za-z]+)$")
_ROW_RE = re.compile(r"^(\d+)$")
_ANCHOR_RE = re.compile(r"^@(bottom_left|bottom_right|right_top)(?:\+(\d+))?$")


# --- Parsers ---

def _split_sheet(ref_str: str) -> tuple[str | None, str]:
    """Split optional sheet prefix from reference."""
    if "!" in ref_str:
        sheet, _, rest = ref_str.partition("!")
        # Remove quotes if present
        sheet = sheet.strip("'\"")
        return sheet, rest
    return None, ref_str


def parse_cell_ref(s: str) -> CellRef | None:
    """Parse a single cell reference like 'A1' or 'Sheet2!B3'."""
    sheet, ref = _split_sheet(s.strip())
    m = _CELL_RE.match(ref)
    if not m:
        return None
    col = col_to_index(m.group(1))
    row = int(m.group(2))
    return CellRef(col=col, row=row, sheet=sheet)


def parse_range_ref(s: str) -> RangeRef | ColRef | RowRef | None:
    """Parse a range reference like 'A1:D10', 'B:B', '3:3', 'Sheet2!A1:B10'."""
    sheet, ref = _split_sheet(s.strip())

    if ":" not in ref:
        return None

    left, right = ref.split(":", 1)

    # Cell range: A1:D10
    m_left = _CELL_RE.match(left)
    m_right = _CELL_RE.match(right)
    if m_left and m_right:
        start = CellRef(col=col_to_index(m_left.group(1)), row=int(m_left.group(2)))
        end = CellRef(col=col_to_index(m_right.group(1)), row=int(m_right.group(2)))
        return RangeRef(start=start, end=end, sheet=sheet)

    # Column range: A:E or B:B
    cm_left = _COL_RE.match(left)
    cm_right = _COL_RE.match(right)
    if cm_left and cm_right:
        return ColRef(
            start_col=col_to_index(cm_left.group(1)),
            end_col=col_to_index(cm_right.group(1)),
            sheet=sheet,
        )

    # Row range: 1:5 or 3:3
    rm_left = _ROW_RE.match(left)
    rm_right = _ROW_RE.match(right)
    if rm_left and rm_right:
        return RowRef(
            start_row=int(rm_left.group(1)),
            end_row=int(rm_right.group(1)),
            sheet=sheet,
        )

    return None


def parse_anchor(s: str) -> AnchorRef | None:
    """Parse a spatial anchor reference (C6).

    Examples:
      @bottom_left      → AnchorRef("bottom_left", 0)
      @bottom_left+2    → AnchorRef("bottom_left", 2)
      @right_top        → AnchorRef("right_top", 0)
    """
    m = _ANCHOR_RE.match(s.strip())
    if not m:
        return None
    anchor = m.group(1)
    offset = int(m.group(2)) if m.group(2) else 0
    return AnchorRef(anchor=anchor, offset=offset)


def parse_ref(s: str) -> CellRef | RangeRef | ColRef | RowRef | AnchorRef | None:
    """Parse any cell/range/anchor reference."""
    s = s.strip()
    # Try anchor first
    anchor = parse_anchor(s)
    if anchor:
        return anchor
    # Try single cell
    cell = parse_cell_ref(s)
    if cell:
        return cell
    # Try range/col/row
    return parse_range_ref(s)
