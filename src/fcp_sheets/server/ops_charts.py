"""Chart operation handlers — chart add/series/axis/remove.

Stubs for Wave 2 implementation.
"""

from __future__ import annotations

from fcp_core import OpResult, ParsedOp

from fcp_sheets.server.resolvers import SheetsOpContext


def op_chart(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    raise NotImplementedError("chart verb not yet implemented")


HANDLERS: dict[str, callable] = {
    "chart": op_chart,
}
