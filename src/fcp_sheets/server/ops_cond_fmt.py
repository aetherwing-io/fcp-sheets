"""Conditional formatting handlers — cond-fmt variants.

Stubs for Wave 2 implementation.
"""

from __future__ import annotations

from fcp_core import OpResult, ParsedOp

from fcp_sheets.server.resolvers import SheetsOpContext


def op_cond_fmt(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    raise NotImplementedError("cond-fmt verb not yet implemented")


HANDLERS: dict[str, callable] = {
    "cond-fmt": op_cond_fmt,
}
