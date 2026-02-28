"""Named color palette and hex parsing for spreadsheet formatting."""

from __future__ import annotations

# Standard Excel/Office color palette
NAMED_COLORS: dict[str, str] = {
    "blue": "4472C4",
    "orange": "ED7D31",
    "gray": "A5A5A5",
    "gold": "FFC000",
    "lt-blue": "5B9BD5",
    "green": "70AD47",
    "red": "FF0000",
    "dk-green": "00B050",
    "white": "FFFFFF",
    "black": "000000",
    "yellow": "FFFF00",
    "purple": "7030A0",
    # Conditional formatting fills
    "good-fill": "C6EFCE",
    "bad-fill": "FFC7CE",
    "neutral-fill": "FFEB9C",
}


def parse_color(color_str: str) -> str:
    """Parse a color string to a 6-char hex value (no #).

    Accepts:
      - Named colors: "blue", "red"
      - Hex with #: "#4472C4"
      - Hex without #: "4472C4"
    """
    # Check named colors
    name = color_str.lower().strip()
    if name in NAMED_COLORS:
        return NAMED_COLORS[name]

    # Strip # prefix
    hex_str = color_str.lstrip("#").strip()

    # Validate hex
    if len(hex_str) == 6 and all(c in "0123456789ABCDEFabcdef" for c in hex_str):
        return hex_str.upper()

    # 3-char shorthand
    if len(hex_str) == 3 and all(c in "0123456789ABCDEFabcdef" for c in hex_str):
        return "".join(c + c for c in hex_str).upper()

    raise ValueError(f"Invalid color: {color_str!r}")
