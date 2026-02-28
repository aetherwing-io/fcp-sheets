"""Table operation handlers — table add/remove.

Stubs for Wave 2 implementation.
"""

from __future__ import annotations

from fcp_core import OpResult, ParsedOp

from fcp_sheets.server.resolvers import SheetsOpContext


def op_table(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    raise NotImplementedError("table verb not yet implemented")


HANDLERS: dict[str, callable] = {
    "table": op_table,
}
