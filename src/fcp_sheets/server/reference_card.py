"""Extra reference card sections for the sheets tool description."""

from __future__ import annotations

EXTRA_SECTIONS: dict[str, str] = {
    "Cell References": "A1 single | A1:D10 range | B:B column | 3:3 row | Sheet2!A1 cross-sheet\n  @bottom_left @bottom_right @right_top (spatial anchors, +N offset)",
    "Number Formats": (
        "General | 0 | 0.00 | #,##0 | $#,##0 | $#,##0.00\n"
        "  0% | 0.00% | yyyy-mm-dd | mm/dd/yyyy | hh:mm:ss | @"
    ),
    "Colors": (
        "#4472C4 blue  #ED7D31 orange  #A5A5A5 gray  #FFC000 gold\n"
        "  #5B9BD5 lt-blue  #70AD47 green  #FF0000 red  #00B050 dk-green\n"
        "  #C6EFCE good-fill  #FFC7CE bad-fill  #FFEB9C neutral-fill"
    ),
    "Chart Types": (
        "bar, column, line, pie, scatter, area, doughnut, radar, bubble\n"
        "  stacked-bar, stacked-column, stacked-area\n"
        "  100-bar, 100-column, 100-area\n"
        "  bar-3d, column-3d, line-3d, pie-3d, area-3d"
    ),
    "Selectors": (
        "@sheet:NAME  @range:A1:Z99  @row:N  @col:A  @type:formula|number|text|date|empty\n"
        "  @table:NAME  @name:NAME  @all  @recent  @recent:N  @not:TYPE:VALUE\n"
        "  Combine to intersect: @sheet:Revenue @col:E @type:formula"
    ),
    "Border Styles": "thin | medium | thick | dashed | dotted | double | hair\n  Sides: all | outline | top | bottom | left | right | inner | h | v",
    "Cond-Fmt Operators": "gt | lt | gte | lte | eq | neq | between | not-between",
    "Table Styles": "TableStyleLight1-21 | TableStyleMedium1-28 | TableStyleDark1-11",
    "Response Prefixes": (
        "+  cell/data created    ~  chart/table created\n"
        "  *  style/format modified  -  cell/range removed\n"
        "  !  error or meta         @  bulk/selector operation"
    ),
    "Conventions": (
        "- Cell references use A1 notation (case-insensitive)\n"
        "  - Values beginning with = are formulas\n"
        "  - Quoted strings are text; bare numbers are numeric\n"
        "  - Use data blocks for bulk entry (avoids cell-by-cell counting)\n"
        "  - Active sheet is implicit target; use sheet:NAME for cross-sheet\n"
        "  - Call sheets_help after context truncation for full reference"
    ),
}
