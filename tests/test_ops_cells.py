"""Tests for cell operations — set verb."""

from __future__ import annotations

import pytest
from fcp_core import ParsedOp

from fcp_sheets.adapter import SheetsAdapter
from fcp_sheets.model.snapshot import SheetsModel
from fcp_sheets.server.ops_cells import op_set, _parse_cell_value
from fcp_sheets.server.resolvers import SheetsOpContext


class TestParseValue:
    def test_integer(self):
        assert _parse_cell_value("42") == 42

    def test_float(self):
        assert _parse_cell_value("3.14") == 3.14

    def test_negative_int(self):
        assert _parse_cell_value("-5") == -5

    def test_negative_float(self):
        assert _parse_cell_value("-2.5") == -2.5

    def test_formula(self):
        assert _parse_cell_value("=SUM(A1:A10)") == "=SUM(A1:A10)"

    def test_quoted_string(self):
        assert _parse_cell_value('"Hello"') == "Hello"

    def test_single_quoted(self):
        assert _parse_cell_value("'World'") == "World"

    def test_plain_text(self):
        assert _parse_cell_value("Hello") == "Hello"

    def test_leading_zero_preserved(self):
        """C1: Leading zeros preserved as text."""
        assert _parse_cell_value("01234") == "01234"
        assert _parse_cell_value("007") == "007"

    def test_single_zero(self):
        """Single zero is a number, not leading-zero text."""
        assert _parse_cell_value("0") == 0

    def test_zero_point(self):
        assert _parse_cell_value("0.5") == 0.5


class TestSetVerb:
    def test_set_number(self, ctx: SheetsOpContext, adapter: SheetsAdapter):
        op = ParsedOp(verb="set", positionals=["A1", "42"], raw="set A1 42")
        result = op_set(op, ctx)
        assert result.success
        assert ctx.wb.active.cell(row=1, column=1).value == 42

    def test_set_text(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="set", positionals=["B2", "Hello"], raw="set B2 Hello")
        result = op_set(op, ctx)
        assert result.success
        assert ctx.wb.active.cell(row=2, column=2).value == "Hello"

    def test_set_formula(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="set", positionals=["C3", "=SUM(A1:B2)"], raw="set C3 =SUM(A1:B2)")
        result = op_set(op, ctx)
        assert result.success
        assert ctx.wb.active.cell(row=3, column=3).value == "=SUM(A1:B2)"

    def test_set_with_format(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="set", positionals=["D4", "0.156"], params={"fmt": "0.00%"}, raw="set D4 0.156 fmt:0.00%")
        result = op_set(op, ctx)
        assert result.success
        cell = ctx.wb.active.cell(row=4, column=4)
        assert cell.value == 0.156
        assert cell.number_format == "0.00%"

    def test_set_with_format_alias(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="set", positionals=["A1", "50000"], params={"fmt": "currency"}, raw="set A1 50000 fmt:currency")
        result = op_set(op, ctx)
        assert result.success
        cell = ctx.wb.active.cell(row=1, column=1)
        assert cell.value == 50000
        assert cell.number_format == "$#,##0"

    def test_set_missing_args(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="set", positionals=["A1"], raw="set A1")
        result = op_set(op, ctx)
        assert not result.success

    def test_set_invalid_ref(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="set", positionals=["INVALID", "42"], raw="set INVALID 42")
        result = op_set(op, ctx)
        assert not result.success

    def test_set_updates_index(self, ctx: SheetsOpContext, adapter: SheetsAdapter):
        op = ParsedOp(verb="set", positionals=["E5", "100"], raw="set E5 100")
        op_set(op, ctx)
        bounds = ctx.index.get_bounds("Sheet1")
        assert bounds is not None
        assert bounds[2] >= 5  # max_row includes row 5
        assert bounds[3] >= 5  # max_col includes col 5


class TestSetWithAnchor:
    """C6: Spatial anchor tests for set verb."""

    def test_set_at_bottom_left(self, ctx: SheetsOpContext, adapter: SheetsAdapter):
        # First put some data
        ctx.wb.active.cell(row=1, column=1, value="Header")
        ctx.wb.active.cell(row=2, column=1, value="Data")
        ctx.index.expand_bounds("Sheet1", 1, 1)
        ctx.index.expand_bounds("Sheet1", 2, 1)

        op = ParsedOp(verb="set", positionals=["@bottom_left", "Total"], raw="set @bottom_left Total")
        result = op_set(op, ctx)
        assert result.success
        # @bottom_left should be row 3 (max_row+1), col 1 (min_col)
        assert ctx.wb.active.cell(row=3, column=1).value == "Total"

    def test_set_at_bottom_left_offset(self, ctx: SheetsOpContext, adapter: SheetsAdapter):
        ctx.wb.active.cell(row=1, column=1, value="A")
        ctx.wb.active.cell(row=3, column=2, value="B")
        ctx.index.expand_bounds("Sheet1", 1, 1)
        ctx.index.expand_bounds("Sheet1", 3, 2)

        op = ParsedOp(verb="set", positionals=["@bottom_left+2", "Far"], raw="set @bottom_left+2 Far")
        result = op_set(op, ctx)
        assert result.success
        # @bottom_left+2 = row (3+1+2)=6, col 1
        assert ctx.wb.active.cell(row=6, column=1).value == "Far"
