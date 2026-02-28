"""Structure operation handlers — merge, freeze, filter, width, height, etc.

Stubs for Wave 2 implementation.
"""

from __future__ import annotations

from fcp_core import OpResult, ParsedOp

from fcp_sheets.server.resolvers import SheetsOpContext


def op_merge(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    raise NotImplementedError("merge verb not yet implemented")

def op_unmerge(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    raise NotImplementedError("unmerge verb not yet implemented")

def op_freeze(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    raise NotImplementedError("freeze verb not yet implemented")

def op_unfreeze(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    raise NotImplementedError("unfreeze verb not yet implemented")

def op_filter(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    raise NotImplementedError("filter verb not yet implemented")

def op_width(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    raise NotImplementedError("width verb not yet implemented")

def op_height(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    raise NotImplementedError("height verb not yet implemented")

def op_hide_col(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    raise NotImplementedError("hide-col verb not yet implemented")

def op_hide_row(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    raise NotImplementedError("hide-row verb not yet implemented")

def op_unhide_col(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    raise NotImplementedError("unhide-col verb not yet implemented")

def op_unhide_row(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    raise NotImplementedError("unhide-row verb not yet implemented")

def op_group_rows(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    raise NotImplementedError("group-rows verb not yet implemented")

def op_group_cols(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    raise NotImplementedError("group-cols verb not yet implemented")

def op_ungroup_rows(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    raise NotImplementedError("ungroup-rows verb not yet implemented")

def op_ungroup_cols(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    raise NotImplementedError("ungroup-cols verb not yet implemented")


HANDLERS: dict[str, callable] = {
    "merge": op_merge,
    "unmerge": op_unmerge,
    "freeze": op_freeze,
    "unfreeze": op_unfreeze,
    "filter": op_filter,
    "width": op_width,
    "height": op_height,
    "hide-col": op_hide_col,
    "hide-row": op_hide_row,
    "unhide-col": op_unhide_col,
    "unhide-row": op_unhide_row,
    "group-rows": op_group_rows,
    "group-cols": op_group_cols,
    "ungroup-rows": op_ungroup_rows,
    "ungroup-cols": op_ungroup_cols,
}
