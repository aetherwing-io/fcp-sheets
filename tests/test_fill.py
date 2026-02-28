"""Tests for fill verb — drag/copy formulas and values."""

from __future__ import annotations

import pytest
from fcp_core import EventLog, ParsedOp

from fcp_sheets.adapter import SheetsAdapter
from fcp_sheets.model.snapshot import SheetsModel
from fcp_sheets.server.ops_cells import op_fill
from fcp_sheets.server.resolvers import SheetsOpContext


def _cell(model: SheetsModel, row: int, col: int):
    """Get cell value from active sheet."""
    return model.wb.active.cell(row=row, column=col).value


def _dispatch_fill(adapter, model, log, raw: str) -> "OpResult":
    """Helper to dispatch a fill op through the adapter (for snapshot/undo)."""
    parts = raw.split()
    verb = parts[0]
    positionals = []
    params = {}
    for part in parts[1:]:
        if ":" in part:
            k, v = part.split(":", 1)
            params[k] = v
        else:
            positionals.append(part)
    op = ParsedOp(verb=verb, positionals=positionals, params=params, raw=raw)
    return adapter.dispatch_op(op, model, log)


# -- Fill Numbers Down --


class TestFillNumberDown:
    def test_fill_number_down_by_count(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        model.wb.active.cell(row=1, column=1, value=42)
        adapter.index.expand_bounds("Sheet1", 1, 1)

        result = _dispatch_fill(adapter, model, log, "fill A1 dir:down count:3")
        assert result.success
        assert _cell(model, 2, 1) == 42
        assert _cell(model, 3, 1) == 42
        assert _cell(model, 4, 1) == 42
        assert _cell(model, 5, 1) is None  # not filled beyond count

    def test_fill_text_down(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        model.wb.active.cell(row=1, column=1, value="Hello")
        adapter.index.expand_bounds("Sheet1", 1, 1)

        result = _dispatch_fill(adapter, model, log, "fill A1 dir:down count:2")
        assert result.success
        assert _cell(model, 2, 1) == "Hello"
        assert _cell(model, 3, 1) == "Hello"

    def test_fill_float_down(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        model.wb.active.cell(row=1, column=1, value=3.14)
        adapter.index.expand_bounds("Sheet1", 1, 1)

        result = _dispatch_fill(adapter, model, log, "fill A1 dir:down count:1")
        assert result.success
        assert _cell(model, 2, 1) == 3.14


# -- Fill Formulas Down (C3) --


class TestFillFormulaDown:
    def test_simple_formula_shifts(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        """=A1+B1 at C1, filled down should become =A2+B2, =A3+B3."""
        model.wb.active.cell(row=1, column=3, value="=A1+B1")
        adapter.index.expand_bounds("Sheet1", 1, 3)

        result = _dispatch_fill(adapter, model, log, "fill C1 dir:down count:2")
        assert result.success
        assert _cell(model, 2, 3) == "=A2+B2"
        assert _cell(model, 3, 3) == "=A3+B3"

    def test_absolute_refs_stay_locked(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        """C3: =SUM($A$1:$A$10) filled down: absolute refs stay locked."""
        model.wb.active.cell(row=1, column=2, value="=SUM($A$1:$A$10)")
        adapter.index.expand_bounds("Sheet1", 1, 2)

        result = _dispatch_fill(adapter, model, log, "fill B1 dir:down count:2")
        assert result.success
        # $A$1:$A$10 should not shift
        assert _cell(model, 2, 2) == "=SUM($A$1:$A$10)"
        assert _cell(model, 3, 2) == "=SUM($A$1:$A$10)"

    def test_cross_sheet_ref(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        """C3: =Sheet2!$B$2*A1 filled down: $B$2 stays, A1 shifts."""
        model.wb.active.cell(row=1, column=1, value="=Sheet2!$B$2*A1")
        adapter.index.expand_bounds("Sheet1", 1, 1)

        result = _dispatch_fill(adapter, model, log, "fill A1 dir:down count:2")
        assert result.success
        # Sheet2!$B$2 stays, A1 becomes A2, A3
        val2 = _cell(model, 2, 1)
        val3 = _cell(model, 3, 1)
        assert "$B$2" in val2
        assert "A2" in val2
        assert "$B$2" in val3
        assert "A3" in val3

    def test_mixed_refs(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        """C3: =$A1+B$1 filled down: $A stays col-locked, $1 stays row-locked."""
        model.wb.active.cell(row=1, column=3, value="=$A1+B$1")
        adapter.index.expand_bounds("Sheet1", 1, 3)

        result = _dispatch_fill(adapter, model, log, "fill C1 dir:down count:2")
        assert result.success
        # $A (col locked) + row shifts: $A2, $A3
        # B$1 (row locked): stays B$1
        val2 = _cell(model, 2, 3)
        val3 = _cell(model, 3, 3)
        assert "$A2" in val2
        assert "B$1" in val2
        assert "$A3" in val3
        assert "B$1" in val3

    def test_vlookup_fill(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        """C3: =VLOOKUP(A2,Sheet2!$A:$B,2,FALSE) filled down."""
        model.wb.active.cell(row=2, column=2, value="=VLOOKUP(A2,Sheet2!$A:$B,2,FALSE)")
        adapter.index.expand_bounds("Sheet1", 2, 2)

        result = _dispatch_fill(adapter, model, log, "fill B2 dir:down count:2")
        assert result.success
        val3 = _cell(model, 3, 2)
        val4 = _cell(model, 4, 2)
        # A2 should shift to A3, A4; $A:$B should stay
        assert "A3" in val3
        assert "$A:$B" in val3
        assert "A4" in val4
        assert "$A:$B" in val4

    def test_sum_range_shifts(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        """=SUM(A1:A5) at B1 filled down: becomes =SUM(A2:A6), =SUM(A3:A7)."""
        model.wb.active.cell(row=1, column=2, value="=SUM(A1:A5)")
        adapter.index.expand_bounds("Sheet1", 1, 2)

        result = _dispatch_fill(adapter, model, log, "fill B1 dir:down count:2")
        assert result.success
        assert _cell(model, 2, 2) == "=SUM(A2:A6)"
        assert _cell(model, 3, 2) == "=SUM(A3:A7)"


# -- Fill Right --


class TestFillRight:
    def test_fill_number_right(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        model.wb.active.cell(row=1, column=1, value=100)
        adapter.index.expand_bounds("Sheet1", 1, 1)

        result = _dispatch_fill(adapter, model, log, "fill A1 dir:right count:3")
        assert result.success
        assert _cell(model, 1, 2) == 100
        assert _cell(model, 1, 3) == 100
        assert _cell(model, 1, 4) == 100

    def test_fill_formula_right(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        """=A1*2 at A2, filled right: =B1*2, =C1*2."""
        model.wb.active.cell(row=2, column=1, value="=A1*2")
        adapter.index.expand_bounds("Sheet1", 2, 1)

        result = _dispatch_fill(adapter, model, log, "fill A2 dir:right count:2")
        assert result.success
        assert _cell(model, 2, 2) == "=B1*2"
        assert _cell(model, 2, 3) == "=C1*2"


# -- Fill with to:CELL --


class TestFillTo:
    def test_fill_down_to_cell(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        model.wb.active.cell(row=1, column=1, value=10)
        adapter.index.expand_bounds("Sheet1", 1, 1)

        result = _dispatch_fill(adapter, model, log, "fill A1 dir:down to:A5")
        assert result.success
        assert "4 cells" in result.message
        for r in range(2, 6):
            assert _cell(model, r, 1) == 10
        assert _cell(model, 6, 1) is None

    def test_fill_right_to_cell(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        model.wb.active.cell(row=1, column=1, value="X")
        adapter.index.expand_bounds("Sheet1", 1, 1)

        result = _dispatch_fill(adapter, model, log, "fill A1 dir:right to:D1")
        assert result.success
        assert _cell(model, 1, 2) == "X"
        assert _cell(model, 1, 3) == "X"
        assert _cell(model, 1, 4) == "X"

    def test_fill_down_to_invalid_direction(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        """to: cell above source should fail for dir:down."""
        model.wb.active.cell(row=5, column=1, value=10)
        adapter.index.expand_bounds("Sheet1", 5, 1)

        result = _dispatch_fill(adapter, model, log, "fill A5 dir:down to:A3")
        assert not result.success
        assert "below" in result.message


# -- Fill with until:COL --


class TestFillUntil:
    def test_fill_until_column_empty(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        """Fill down until column A is empty."""
        ws = model.wb.active
        # Column A has data in rows 1-4
        ws.cell(row=1, column=1, value="a")
        ws.cell(row=2, column=1, value="b")
        ws.cell(row=3, column=1, value="c")
        ws.cell(row=4, column=1, value="d")
        for r in range(1, 5):
            adapter.index.expand_bounds("Sheet1", r, 1)

        # B1 has formula, fill down until A is empty
        ws.cell(row=1, column=2, value=100)
        adapter.index.expand_bounds("Sheet1", 1, 2)

        result = _dispatch_fill(adapter, model, log, "fill B1 dir:down until:A")
        assert result.success
        assert "3 cells" in result.message  # rows 2,3,4
        assert _cell(model, 2, 2) == 100
        assert _cell(model, 3, 2) == 100
        assert _cell(model, 4, 2) == 100
        assert _cell(model, 5, 2) is None


# -- Error Cases --


class TestFillErrors:
    def test_fill_no_args(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        result = _dispatch_fill(adapter, model, log, "fill")
        assert not result.success

    def test_fill_no_direction(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        model.wb.active.cell(row=1, column=1, value=1)
        result = _dispatch_fill(adapter, model, log, "fill A1 count:3")
        assert not result.success
        assert "dir:down" in result.message or "dir:" in result.message

    def test_fill_empty_source(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        result = _dispatch_fill(adapter, model, log, "fill A1 dir:down count:3")
        assert not result.success
        assert "empty" in result.message.lower()

    def test_fill_no_count_or_to(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        model.wb.active.cell(row=1, column=1, value=1)
        adapter.index.expand_bounds("Sheet1", 1, 1)
        result = _dispatch_fill(adapter, model, log, "fill A1 dir:down")
        assert not result.success
        assert "count" in result.message.lower() or "to" in result.message.lower()

    def test_fill_invalid_count(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        model.wb.active.cell(row=1, column=1, value=1)
        adapter.index.expand_bounds("Sheet1", 1, 1)
        result = _dispatch_fill(adapter, model, log, "fill A1 dir:down count:abc")
        assert not result.success

    def test_fill_zero_count(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        model.wb.active.cell(row=1, column=1, value=1)
        adapter.index.expand_bounds("Sheet1", 1, 1)
        result = _dispatch_fill(adapter, model, log, "fill A1 dir:down count:0")
        assert not result.success
