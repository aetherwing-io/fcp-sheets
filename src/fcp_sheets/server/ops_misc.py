"""Miscellaneous operation handlers — name, image, link, comment, protect, page-setup.

Stubs for Wave 2 implementation.
"""

from __future__ import annotations

from fcp_core import OpResult, ParsedOp

from fcp_sheets.server.resolvers import SheetsOpContext


def op_name(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    raise NotImplementedError("name verb not yet implemented")

def op_image(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    raise NotImplementedError("image verb not yet implemented")

def op_link(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    raise NotImplementedError("link verb not yet implemented")

def op_comment(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    raise NotImplementedError("comment verb not yet implemented")

def op_protect(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    raise NotImplementedError("protect verb not yet implemented")

def op_unprotect(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    raise NotImplementedError("unprotect verb not yet implemented")

def op_lock(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    raise NotImplementedError("lock verb not yet implemented")

def op_unlock(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    raise NotImplementedError("unlock verb not yet implemented")

def op_page_setup(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    raise NotImplementedError("page-setup verb not yet implemented")


HANDLERS: dict[str, callable] = {
    "name": op_name,
    "image": op_image,
    "link": op_link,
    "comment": op_comment,
    "protect": op_protect,
    "unprotect": op_unprotect,
    "lock": op_lock,
    "unlock": op_unlock,
    "page-setup": op_page_setup,
}
