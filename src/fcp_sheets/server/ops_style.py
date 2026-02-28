"""Style operation handlers — style, border, define-style, apply-style.

Stubs for Wave 2 implementation.
"""

from __future__ import annotations

from fcp_core import OpResult, ParsedOp

from fcp_sheets.server.resolvers import SheetsOpContext


def op_style(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    raise NotImplementedError("style verb not yet implemented")


def op_border(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    raise NotImplementedError("border verb not yet implemented")


def op_define_style(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    raise NotImplementedError("define-style verb not yet implemented")


def op_apply_style(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    raise NotImplementedError("apply-style verb not yet implemented")


HANDLERS: dict[str, callable] = {
    "style": op_style,
    "border": op_border,
    "define-style": op_define_style,
    "apply-style": op_apply_style,
}
