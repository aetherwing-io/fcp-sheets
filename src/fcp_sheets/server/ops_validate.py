"""Data validation handlers — validate variants.

Stubs for Wave 2 implementation.
"""

from __future__ import annotations

from fcp_core import OpResult, ParsedOp

from fcp_sheets.server.resolvers import SheetsOpContext


def op_validate(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    raise NotImplementedError("validate verb not yet implemented")


HANDLERS: dict[str, callable] = {
    "validate": op_validate,
}
