"""Table style catalog for Excel tables."""

from __future__ import annotations

# All built-in Excel table styles
TABLE_STYLES: list[str] = (
    [f"TableStyleLight{i}" for i in range(1, 22)]
    + [f"TableStyleMedium{i}" for i in range(1, 29)]
    + [f"TableStyleDark{i}" for i in range(1, 12)]
)

# Set for O(1) lookup
TABLE_STYLE_SET: set[str] = set(TABLE_STYLES)

# Case-insensitive lookup
_TABLE_STYLE_MAP: dict[str, str] = {s.lower(): s for s in TABLE_STYLES}


def resolve_table_style(name: str) -> str:
    """Resolve a table style name (case-insensitive).

    Raises ValueError if style not found.
    """
    key = name.lower()
    if key in _TABLE_STYLE_MAP:
        return _TABLE_STYLE_MAP[key]

    # Try with common shorthand: "medium9" → "TableStyleMedium9"
    for prefix in ("TableStyleLight", "TableStyleMedium", "TableStyleDark"):
        full = f"{prefix}{name}".lower()
        if full in _TABLE_STYLE_MAP:
            return _TABLE_STYLE_MAP[full]

    # Try: "light1" → "TableStyleLight1"
    for category in ("light", "medium", "dark"):
        if key.startswith(category):
            num = key[len(category):]
            full = f"TableStyle{category.capitalize()}{num}".lower()
            if full in _TABLE_STYLE_MAP:
                return _TABLE_STYLE_MAP[full]

    raise ValueError(f"Unknown table style: {name!r}")
