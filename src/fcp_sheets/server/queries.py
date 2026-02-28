"""Query handlers — read-only inspection of workbook state.

MVP queries: plan/map, stats, status, history.
Extended queries (Wave 3): describe, peek, list, find.
"""

from __future__ import annotations

from fcp_sheets.model.index import SheetIndex
from fcp_sheets.model.refs import index_to_col
from fcp_sheets.model.snapshot import SheetsModel
from fcp_sheets.server.formatter import format_range, truncate_list


def dispatch_query(query: str, model: SheetsModel, index: SheetIndex) -> str:
    """Route a query string to the appropriate handler."""
    query = query.strip()
    parts = query.split(None, 1)
    command = parts[0].lower() if parts else ""
    args = parts[1] if len(parts) > 1 else ""

    handlers = {
        "plan": _query_plan,
        "map": _query_plan,
        "stats": _query_stats,
        "status": _query_status,
        "history": _query_history,
        "describe": _query_describe,
        "peek": _query_peek,
        "list": _query_list,
        "find": _query_find,
    }

    handler = handlers.get(command)
    if handler is None:
        return (
            f"! Unknown query: {command!r}\n"
            "  try: plan, stats, status, history, describe, peek, list, find"
        )

    return handler(args, model, index)


# ---------------------------------------------------------------------------
# plan / map — primary overview
# ---------------------------------------------------------------------------

def _query_plan(args: str, model: SheetsModel, index: SheetIndex) -> str:
    """Show workbook topology — the most important query."""
    wb = model.wb
    lines: list[str] = []

    # Header
    saved = f", saved: {model.file_path}" if model.file_path else ", unsaved"
    lines.append(f'Workbook: "{model.title}" ({len(wb.sheetnames)} sheets{saved})')

    for ws in wb.worksheets:
        sheet_name = ws.title
        active = " [active]" if ws == wb.active else ""
        hidden = " [hidden]" if ws.sheet_state == "hidden" else ""
        lines.append(f"\n  Sheet: {sheet_name}{active}{hidden}")

        bounds = index.get_bounds(sheet_name)
        if bounds:
            min_row, min_col, max_row, max_col = bounds
            range_str = format_range(min_row, min_col, max_row, max_col)

            # Data bounds + freeze + filter
            meta_parts = [f"data: {range_str}"]
            if ws.freeze_panes:
                meta_parts.append(f"frozen: {ws.freeze_panes}")
            if ws.auto_filter and ws.auto_filter.ref:
                meta_parts.append(f"filter: {ws.auto_filter.ref}")
            lines.append(f"    {' | '.join(meta_parts)}")

            # Column headers (first row)
            cols = []
            for col in range(min_col, min(max_col + 1, min_col + 8)):
                cell = ws.cell(row=min_row, column=col)
                val = cell.value
                if val is not None:
                    cols.append(f"{index_to_col(col)}:{val}")
            if cols:
                lines.append(f"    cols: {truncate_list(cols)}")

            # Formula patterns
            formula_patterns = _detect_formula_patterns(ws, bounds)
            if formula_patterns:
                lines.append(f"    formulas: {' | '.join(formula_patterns)}")

            # Tables
            if ws.tables:
                for table_name in ws.tables:
                    table = ws.tables[table_name]
                    lines.append(f"    table: {table_name} {table.ref}")

            # Charts
            if ws._charts:
                for chart in ws._charts:
                    chart_title = chart.title or "Untitled"
                    chart_type = type(chart).__name__
                    lines.append(f"    chart: {chart_type} \"{chart_title}\"")

            # Conditional formatting
            if ws.conditional_formatting:
                cf_count = len(list(ws.conditional_formatting))
                if cf_count > 0:
                    lines.append(f"    cond-fmt: {cf_count} rule(s)")

            # Next empty row
            lines.append(f"    next-empty: row:{max_row + 1}")
        else:
            lines.append("    (empty)")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# stats
# ---------------------------------------------------------------------------

def _query_stats(args: str, model: SheetsModel, index: SheetIndex) -> str:
    """Quick summary of workbook contents."""
    wb = model.wb
    total_data = 0
    total_formula = 0
    total_charts = 0
    total_tables = 0
    total_merged = 0
    total_cond_fmt = 0

    for ws in wb.worksheets:
        bounds = index.get_bounds(ws.title)
        if bounds:
            min_row, min_col, max_row, max_col = bounds
            for row in ws.iter_rows(
                min_row=min_row, max_row=max_row,
                min_col=min_col, max_col=max_col,
            ):
                for cell in row:
                    if cell.value is not None:
                        if isinstance(cell.value, str) and cell.value.startswith("="):
                            total_formula += 1
                        else:
                            total_data += 1

        total_charts += len(ws._charts)
        total_tables += len(ws.tables)
        total_merged += len(ws.merged_cells.ranges)
        total_cond_fmt += len(list(ws.conditional_formatting))

    named_ranges = len(wb.defined_names.definedName) if wb.defined_names else 0

    lines = [
        f'Workbook: "{model.title}"',
        f"  Sheets: {len(wb.sheetnames)} ({', '.join(wb.sheetnames)})",
        f"  Data cells: {total_data:,} | Formula cells: {total_formula:,}",
        f"  Tables: {total_tables} | Charts: {total_charts} | Named ranges: {named_ranges}",
        f"  Merged regions: {total_merged} | Conditional formats: {total_cond_fmt}",
    ]
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# status
# ---------------------------------------------------------------------------

def _query_status(args: str, model: SheetsModel, index: SheetIndex) -> str:
    """Session status."""
    saved = model.file_path or "unsaved"
    active = model.wb.active.title if model.wb.active else "none"
    return (
        f'Session: "{model.title}"\n'
        f"  File: {saved}\n"
        f"  Sheets: {len(model.wb.sheetnames)}\n"
        f"  Active: {active}"
    )


# ---------------------------------------------------------------------------
# history
# ---------------------------------------------------------------------------

def _query_history(args: str, model: SheetsModel, index: SheetIndex) -> str:
    """Show recent operations — managed by session layer."""
    return "History managed by session layer. Use sheets_session for undo/redo."


# ---------------------------------------------------------------------------
# Stubs for Wave 3
# ---------------------------------------------------------------------------

def _query_describe(args: str, model: SheetsModel, index: SheetIndex) -> str:
    return "! describe query not yet implemented"

def _query_peek(args: str, model: SheetsModel, index: SheetIndex) -> str:
    return "! peek query not yet implemented"

def _query_list(args: str, model: SheetsModel, index: SheetIndex) -> str:
    return "! list query not yet implemented"

def _query_find(args: str, model: SheetsModel, index: SheetIndex) -> str:
    return "! find query not yet implemented"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _detect_formula_patterns(ws, bounds: tuple[int, int, int, int]) -> list[str]:
    """Detect formula patterns in a sheet for the plan query."""
    min_row, min_col, max_row, max_col = bounds
    patterns: dict[str, list[str]] = {}  # pattern → [cell_addrs]

    for row in ws.iter_rows(
        min_row=min_row, max_row=max_row,
        min_col=min_col, max_col=max_col,
    ):
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith("="):
                # Normalize formula pattern (strip row numbers for grouping)
                import re
                pattern = re.sub(r"\d+", "N", cell.value)
                addr = f"{index_to_col(cell.column)}{cell.row}"
                if pattern not in patterns:
                    patterns[pattern] = []
                patterns[pattern].append(addr)

    result = []
    for pattern, addrs in patterns.items():
        if len(addrs) == 1:
            result.append(f"{addrs[0]} {pattern.replace('N', 'N')}")
        else:
            # Show range
            first, last = addrs[0], addrs[-1]
            result.append(f"{first}:{last} pattern:{pattern}")

    return result[:5]  # Cap to avoid token explosion
