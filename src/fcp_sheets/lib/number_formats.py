"""Number format aliases for the fmt: parameter."""

from __future__ import annotations

# Common format string aliases
FORMAT_ALIASES: dict[str, str] = {
    "general": "General",
    "number": "0",
    "number2": "0.00",
    "comma": "#,##0",
    "comma2": "#,##0.00",
    "currency": "$#,##0",
    "currency2": "$#,##0.00",
    "accounting": '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)',
    "percent": "0%",
    "percent2": "0.00%",
    "date": "yyyy-mm-dd",
    "date-us": "mm/dd/yyyy",
    "date-eu": "dd/mm/yyyy",
    "time": "hh:mm:ss",
    "datetime": "yyyy-mm-dd hh:mm:ss",
    "text": "@",
    "fraction": "# ?/?",
    "scientific": "0.00E+00",
    "zip": "00000",
    "phone": "[<=9999999]###-####;(###) ###-####",
    "ssn": "000-00-0000",
}


def resolve_format(fmt_str: str) -> str:
    """Resolve a format alias or return the raw format string."""
    return FORMAT_ALIASES.get(fmt_str.lower(), fmt_str)
