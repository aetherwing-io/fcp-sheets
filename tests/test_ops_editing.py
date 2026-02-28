"""Tests for editing operations — remove, copy, move, sort, insert/delete row/col."""

from __future__ import annotations

import pytest
from fcp_core import ParsedOp

from fcp_sheets.server.ops_editing import (
    op_copy,
    op_delete_col,
    op_delete_row,
    op_insert_col,
    op_insert_row,
    op_move,
    op_remove,
    op_sort,
)
from fcp_sheets.server.resolvers import SheetsOpContext


# ── remove ──────────────────────────────────────────────────────────────

class TestRemove:
    def test_remove_single_cell(self, ctx: SheetsOpContext):
        ws = ctx.active_sheet
        ws.cell(row=1, column=1, value="Hello")
        op = ParsedOp(verb="remove", positionals=["A1"], raw="remove A1")
        result = op_remove(op, ctx)
        assert result.success
        assert ws.cell(row=1, column=1).value is None

    def test_remove_range(self, ctx: SheetsOpContext):
        ws = ctx.active_sheet
        for r in range(1, 4):
            for c in range(1, 4):
                ws.cell(row=r, column=c, value=r * 10 + c)
        op = ParsedOp(verb="remove", positionals=["A1:C3"], raw="remove A1:C3")
        result = op_remove(op, ctx)
        assert result.success
        for r in range(1, 4):
            for c in range(1, 4):
                assert ws.cell(row=r, column=c).value is None

    def test_remove_partial_range(self, ctx: SheetsOpContext):
        ws = ctx.active_sheet
        ws.cell(row=1, column=1, value="A")
        ws.cell(row=1, column=2, value="B")
        ws.cell(row=2, column=1, value="C")
        ws.cell(row=2, column=2, value="D")
        op = ParsedOp(verb="remove", positionals=["A1:A2"], raw="remove A1:A2")
        result = op_remove(op, ctx)
        assert result.success
        assert ws.cell(row=1, column=1).value is None
        assert ws.cell(row=2, column=1).value is None
        # Column B untouched
        assert ws.cell(row=1, column=2).value == "B"
        assert ws.cell(row=2, column=2).value == "D"

    def test_remove_no_args(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="remove", positionals=[], raw="remove")
        result = op_remove(op, ctx)
        assert not result.success

    def test_remove_invalid_range(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="remove", positionals=["INVALID"], raw="remove INVALID")
        result = op_remove(op, ctx)
        assert not result.success


# ── copy ────────────────────────────────────────────────────────────────

class TestCopy:
    def test_copy_range(self, ctx: SheetsOpContext):
        ws = ctx.active_sheet
        ws.cell(row=1, column=1, value=10)
        ws.cell(row=1, column=2, value=20)
        ws.cell(row=2, column=1, value=30)
        ws.cell(row=2, column=2, value=40)
        op = ParsedOp(
            verb="copy", positionals=["A1:B2"],
            params={"to": "D1"}, raw="copy A1:B2 to:D1",
        )
        result = op_copy(op, ctx)
        assert result.success
        # Source unchanged
        assert ws.cell(row=1, column=1).value == 10
        # Destination populated
        assert ws.cell(row=1, column=4).value == 10
        assert ws.cell(row=1, column=5).value == 20
        assert ws.cell(row=2, column=4).value == 30
        assert ws.cell(row=2, column=5).value == 40

    def test_copy_single_cell(self, ctx: SheetsOpContext):
        ws = ctx.active_sheet
        ws.cell(row=1, column=1, value="Hello")
        op = ParsedOp(
            verb="copy", positionals=["A1"],
            params={"to": "C3"}, raw="copy A1 to:C3",
        )
        result = op_copy(op, ctx)
        assert result.success
        assert ws.cell(row=3, column=3).value == "Hello"
        assert ws.cell(row=1, column=1).value == "Hello"

    def test_copy_cross_sheet(self, ctx: SheetsOpContext):
        ws1 = ctx.active_sheet
        ws1.cell(row=1, column=1, value="Data")
        ctx.wb.create_sheet("Sheet2")
        op = ParsedOp(
            verb="copy", positionals=["A1"],
            params={"to": "A1", "sheet": "Sheet2"}, raw="copy A1 to:A1 sheet:Sheet2",
        )
        result = op_copy(op, ctx)
        assert result.success
        assert ctx.wb["Sheet2"].cell(row=1, column=1).value == "Data"

    def test_copy_missing_to(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="copy", positionals=["A1:B2"], raw="copy A1:B2")
        result = op_copy(op, ctx)
        assert not result.success
        assert "to" in result.message.lower()

    def test_copy_no_args(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="copy", positionals=[], raw="copy")
        result = op_copy(op, ctx)
        assert not result.success

    def test_copy_invalid_source(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="copy", positionals=["INVALID"],
            params={"to": "A1"}, raw="copy INVALID to:A1",
        )
        result = op_copy(op, ctx)
        assert not result.success

    def test_copy_invalid_dest(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="copy", positionals=["A1"],
            params={"to": "INVALID"}, raw="copy A1 to:INVALID",
        )
        result = op_copy(op, ctx)
        assert not result.success

    def test_copy_nonexistent_sheet(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="copy", positionals=["A1"],
            params={"to": "A1", "sheet": "NoSuchSheet"}, raw="copy A1 to:A1 sheet:NoSuchSheet",
        )
        result = op_copy(op, ctx)
        assert not result.success

    def test_copy_preserves_formula(self, ctx: SheetsOpContext):
        ws = ctx.active_sheet
        ws.cell(row=1, column=1, value="=SUM(B1:B5)")
        op = ParsedOp(
            verb="copy", positionals=["A1"],
            params={"to": "C1"}, raw="copy A1 to:C1",
        )
        result = op_copy(op, ctx)
        assert result.success
        assert ws.cell(row=1, column=3).value == "=SUM(B1:B5)"


# ── move ────────────────────────────────────────────────────────────────

class TestMove:
    def test_move_range(self, ctx: SheetsOpContext):
        ws = ctx.active_sheet
        ws.cell(row=1, column=1, value=10)
        ws.cell(row=1, column=2, value=20)
        op = ParsedOp(
            verb="move", positionals=["A1:B1"],
            params={"to": "D1"}, raw="move A1:B1 to:D1",
        )
        result = op_move(op, ctx)
        assert result.success
        # Source cleared
        assert ws.cell(row=1, column=1).value is None
        assert ws.cell(row=1, column=2).value is None
        # Destination populated
        assert ws.cell(row=1, column=4).value == 10
        assert ws.cell(row=1, column=5).value == 20

    def test_move_cross_sheet(self, ctx: SheetsOpContext):
        ws1 = ctx.active_sheet
        ws1.cell(row=1, column=1, value="Move me")
        ctx.wb.create_sheet("Dest")
        op = ParsedOp(
            verb="move", positionals=["A1"],
            params={"to": "B2", "sheet": "Dest"}, raw="move A1 to:B2 sheet:Dest",
        )
        result = op_move(op, ctx)
        assert result.success
        assert ws1.cell(row=1, column=1).value is None
        assert ctx.wb["Dest"].cell(row=2, column=2).value == "Move me"

    def test_move_missing_to(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="move", positionals=["A1"], raw="move A1")
        result = op_move(op, ctx)
        assert not result.success

    def test_move_no_args(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="move", positionals=[], raw="move")
        result = op_move(op, ctx)
        assert not result.success


# ── sort ────────────────────────────────────────────────────────────────

class TestSort:
    def test_sort_ascending(self, ctx: SheetsOpContext):
        ws = ctx.active_sheet
        ws.cell(row=1, column=1, value=30)
        ws.cell(row=2, column=1, value=10)
        ws.cell(row=3, column=1, value=20)
        op = ParsedOp(
            verb="sort", positionals=["A1:A3"],
            params={"by": "A"}, raw="sort A1:A3 by:A",
        )
        result = op_sort(op, ctx)
        assert result.success
        assert ws.cell(row=1, column=1).value == 10
        assert ws.cell(row=2, column=1).value == 20
        assert ws.cell(row=3, column=1).value == 30

    def test_sort_descending(self, ctx: SheetsOpContext):
        ws = ctx.active_sheet
        ws.cell(row=1, column=1, value=10)
        ws.cell(row=2, column=1, value=30)
        ws.cell(row=3, column=1, value=20)
        op = ParsedOp(
            verb="sort", positionals=["A1:A3"],
            params={"by": "A", "dir": "desc"}, raw="sort A1:A3 by:A dir:desc",
        )
        result = op_sort(op, ctx)
        assert result.success
        assert ws.cell(row=1, column=1).value == 30
        assert ws.cell(row=2, column=1).value == 20
        assert ws.cell(row=3, column=1).value == 10

    def test_sort_with_multiple_columns(self, ctx: SheetsOpContext):
        ws = ctx.active_sheet
        # Name, Score
        ws.cell(row=1, column=1, value="Alice")
        ws.cell(row=1, column=2, value=90)
        ws.cell(row=2, column=1, value="Charlie")
        ws.cell(row=2, column=2, value=80)
        ws.cell(row=3, column=1, value="Bob")
        ws.cell(row=3, column=2, value=70)
        op = ParsedOp(
            verb="sort", positionals=["A1:B3"],
            params={"by": "B"}, raw="sort A1:B3 by:B",
        )
        result = op_sort(op, ctx)
        assert result.success
        # Sorted by column B (Score) ascending
        assert ws.cell(row=1, column=2).value == 70
        assert ws.cell(row=1, column=1).value == "Bob"
        assert ws.cell(row=2, column=2).value == 80
        assert ws.cell(row=3, column=2).value == 90

    def test_sort_secondary_key(self, ctx: SheetsOpContext):
        ws = ctx.active_sheet
        ws.cell(row=1, column=1, value="B")
        ws.cell(row=1, column=2, value=2)
        ws.cell(row=2, column=1, value="A")
        ws.cell(row=2, column=2, value=1)
        ws.cell(row=3, column=1, value="A")
        ws.cell(row=3, column=2, value=3)
        op = ParsedOp(
            verb="sort", positionals=["A1:B3"],
            params={"by": "A", "by2": "B"}, raw="sort A1:B3 by:A by2:B",
        )
        result = op_sort(op, ctx)
        assert result.success
        # Primary: col A asc, Secondary: col B asc
        assert ws.cell(row=1, column=1).value == "A"
        assert ws.cell(row=1, column=2).value == 1
        assert ws.cell(row=2, column=1).value == "A"
        assert ws.cell(row=2, column=2).value == 3
        assert ws.cell(row=3, column=1).value == "B"

    def test_sort_missing_by(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="sort", positionals=["A1:A3"], raw="sort A1:A3")
        result = op_sort(op, ctx)
        assert not result.success
        assert "by" in result.message.lower()

    def test_sort_no_args(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="sort", positionals=[], raw="sort")
        result = op_sort(op, ctx)
        assert not result.success

    def test_sort_invalid_direction(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="sort", positionals=["A1:A3"],
            params={"by": "A", "dir": "sideways"}, raw="sort A1:A3 by:A dir:sideways",
        )
        result = op_sort(op, ctx)
        assert not result.success

    def test_sort_column_outside_range(self, ctx: SheetsOpContext):
        ws = ctx.active_sheet
        ws.cell(row=1, column=1, value=1)
        op = ParsedOp(
            verb="sort", positionals=["A1:A3"],
            params={"by": "C"}, raw="sort A1:A3 by:C",
        )
        result = op_sort(op, ctx)
        assert not result.success
        assert "outside" in result.message.lower()

    def test_sort_with_none_values(self, ctx: SheetsOpContext):
        ws = ctx.active_sheet
        ws.cell(row=1, column=1, value=30)
        ws.cell(row=2, column=1, value=None)
        ws.cell(row=3, column=1, value=10)
        op = ParsedOp(
            verb="sort", positionals=["A1:A3"],
            params={"by": "A"}, raw="sort A1:A3 by:A",
        )
        result = op_sort(op, ctx)
        assert result.success
        # None sorts to end
        assert ws.cell(row=1, column=1).value == 10
        assert ws.cell(row=2, column=1).value == 30
        assert ws.cell(row=3, column=1).value is None


# ── insert-row ──────────────────────────────────────────────────────────

class TestInsertRow:
    def test_insert_single_row(self, ctx: SheetsOpContext):
        ws = ctx.active_sheet
        ws.cell(row=1, column=1, value="A")
        ws.cell(row=2, column=1, value="B")
        op = ParsedOp(verb="insert-row", positionals=["2"], raw="insert-row 2")
        result = op_insert_row(op, ctx)
        assert result.success
        assert ws.cell(row=1, column=1).value == "A"
        # Row 2 is now blank (inserted)
        assert ws.cell(row=2, column=1).value is None
        # Old row 2 shifted to row 3
        assert ws.cell(row=3, column=1).value == "B"

    def test_insert_multiple_rows(self, ctx: SheetsOpContext):
        ws = ctx.active_sheet
        ws.cell(row=1, column=1, value="A")
        ws.cell(row=2, column=1, value="B")
        op = ParsedOp(
            verb="insert-row", positionals=["2"],
            params={"count": "3"}, raw="insert-row 2 count:3",
        )
        result = op_insert_row(op, ctx)
        assert result.success
        assert ws.cell(row=1, column=1).value == "A"
        # 3 blank rows inserted
        assert ws.cell(row=2, column=1).value is None
        assert ws.cell(row=3, column=1).value is None
        assert ws.cell(row=4, column=1).value is None
        # Old row 2 moved to row 5
        assert ws.cell(row=5, column=1).value == "B"

    def test_insert_row_no_args(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="insert-row", positionals=[], raw="insert-row")
        result = op_insert_row(op, ctx)
        assert not result.success

    def test_insert_row_invalid_number(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="insert-row", positionals=["abc"], raw="insert-row abc")
        result = op_insert_row(op, ctx)
        assert not result.success


# ── insert-col ──────────────────────────────────────────────────────────

class TestInsertCol:
    def test_insert_single_col(self, ctx: SheetsOpContext):
        ws = ctx.active_sheet
        ws.cell(row=1, column=1, value="A")
        ws.cell(row=1, column=2, value="B")
        op = ParsedOp(verb="insert-col", positionals=["B"], raw="insert-col B")
        result = op_insert_col(op, ctx)
        assert result.success
        assert ws.cell(row=1, column=1).value == "A"
        assert ws.cell(row=1, column=2).value is None  # New column
        assert ws.cell(row=1, column=3).value == "B"  # Shifted right

    def test_insert_col_by_number(self, ctx: SheetsOpContext):
        ws = ctx.active_sheet
        ws.cell(row=1, column=1, value="X")
        op = ParsedOp(verb="insert-col", positionals=["1"], raw="insert-col 1")
        result = op_insert_col(op, ctx)
        assert result.success
        assert ws.cell(row=1, column=1).value is None
        assert ws.cell(row=1, column=2).value == "X"

    def test_insert_multiple_cols(self, ctx: SheetsOpContext):
        ws = ctx.active_sheet
        ws.cell(row=1, column=1, value="A")
        ws.cell(row=1, column=2, value="B")
        op = ParsedOp(
            verb="insert-col", positionals=["B"],
            params={"count": "2"}, raw="insert-col B count:2",
        )
        result = op_insert_col(op, ctx)
        assert result.success
        assert ws.cell(row=1, column=1).value == "A"
        assert ws.cell(row=1, column=4).value == "B"  # Shifted right by 2

    def test_insert_col_no_args(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="insert-col", positionals=[], raw="insert-col")
        result = op_insert_col(op, ctx)
        assert not result.success


# ── delete-row ──────────────────────────────────────────────────────────

class TestDeleteRow:
    def test_delete_single_row(self, ctx: SheetsOpContext):
        ws = ctx.active_sheet
        ws.cell(row=1, column=1, value="A")
        ws.cell(row=2, column=1, value="B")
        ws.cell(row=3, column=1, value="C")
        op = ParsedOp(verb="delete-row", positionals=["2"], raw="delete-row 2")
        result = op_delete_row(op, ctx)
        assert result.success
        assert ws.cell(row=1, column=1).value == "A"
        assert ws.cell(row=2, column=1).value == "C"

    def test_delete_multiple_rows(self, ctx: SheetsOpContext):
        ws = ctx.active_sheet
        ws.cell(row=1, column=1, value="A")
        ws.cell(row=2, column=1, value="B")
        ws.cell(row=3, column=1, value="C")
        ws.cell(row=4, column=1, value="D")
        op = ParsedOp(
            verb="delete-row", positionals=["2"],
            params={"count": "2"}, raw="delete-row 2 count:2",
        )
        result = op_delete_row(op, ctx)
        assert result.success
        assert ws.cell(row=1, column=1).value == "A"
        assert ws.cell(row=2, column=1).value == "D"

    def test_delete_row_no_args(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="delete-row", positionals=[], raw="delete-row")
        result = op_delete_row(op, ctx)
        assert not result.success


# ── delete-col ──────────────────────────────────────────────────────────

class TestDeleteCol:
    def test_delete_single_col(self, ctx: SheetsOpContext):
        ws = ctx.active_sheet
        ws.cell(row=1, column=1, value="A")
        ws.cell(row=1, column=2, value="B")
        ws.cell(row=1, column=3, value="C")
        op = ParsedOp(verb="delete-col", positionals=["B"], raw="delete-col B")
        result = op_delete_col(op, ctx)
        assert result.success
        assert ws.cell(row=1, column=1).value == "A"
        assert ws.cell(row=1, column=2).value == "C"

    def test_delete_multiple_cols(self, ctx: SheetsOpContext):
        ws = ctx.active_sheet
        ws.cell(row=1, column=1, value="A")
        ws.cell(row=1, column=2, value="B")
        ws.cell(row=1, column=3, value="C")
        ws.cell(row=1, column=4, value="D")
        op = ParsedOp(
            verb="delete-col", positionals=["B"],
            params={"count": "2"}, raw="delete-col B count:2",
        )
        result = op_delete_col(op, ctx)
        assert result.success
        assert ws.cell(row=1, column=1).value == "A"
        assert ws.cell(row=1, column=2).value == "D"

    def test_delete_col_no_args(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="delete-col", positionals=[], raw="delete-col")
        result = op_delete_col(op, ctx)
        assert not result.success

    def test_delete_col_invalid(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="delete-col", positionals=["!!!"], raw="delete-col !!!")
        result = op_delete_col(op, ctx)
        assert not result.success
