"""Editing operation handlers — copy, move, sort, insert/delete, remove.

Stubs for Wave 2 implementation.
"""

from __future__ import annotations

from fcp_core import OpResult, ParsedOp

from fcp_sheets.server.resolvers import SheetsOpContext


def op_remove(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    raise NotImplementedError("remove verb not yet implemented")

def op_copy(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    raise NotImplementedError("copy verb not yet implemented")

def op_move(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    raise NotImplementedError("move verb not yet implemented")

def op_sort(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    raise NotImplementedError("sort verb not yet implemented")

def op_insert_row(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    raise NotImplementedError("insert-row verb not yet implemented")

def op_insert_col(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    raise NotImplementedError("insert-col verb not yet implemented")

def op_delete_row(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    raise NotImplementedError("delete-row verb not yet implemented")

def op_delete_col(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    raise NotImplementedError("delete-col verb not yet implemented")


HANDLERS: dict[str, callable] = {
    "remove": op_remove,
    "copy": op_copy,
    "move": op_move,
    "sort": op_sort,
    "insert-row": op_insert_row,
    "insert-col": op_insert_col,
    "delete-row": op_delete_row,
    "delete-col": op_delete_col,
}
