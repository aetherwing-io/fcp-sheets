"""Tests for data block mode — CSV input, C1 type inference, C2 markdown, C9 collisions."""

from __future__ import annotations

import pytest
from fcp_core import EventLog, ParsedOp

from fcp_sheets.adapter import SheetsAdapter
from fcp_sheets.model.snapshot import SheetsModel


# -- Helpers --


def _run_data_block(
    adapter: SheetsAdapter, model: SheetsModel, log: EventLog,
    anchor: str, lines: list[str],
) -> "from fcp_core import OpResult":
    """Send a data block through the adapter: data ANCHOR, lines..., data end."""
    # Start data block
    start_op = ParsedOp(verb="data", positionals=[anchor], raw=f"data {anchor}")
    result = adapter.dispatch_op(start_op, model, log)
    assert result.success, f"data start failed: {result.message}"

    # Feed lines (the adapter intercepts all ops while in data block mode)
    for line in lines:
        # Lines in data block mode are captured raw via op.raw
        line_op = ParsedOp(verb=line.split()[0] if line.split() else "", positionals=line.split()[1:] if line.split() else [], raw=line)
        result = adapter.dispatch_op(line_op, model, log)
        assert result.success, f"data line failed: {result.message}"

    # End data block
    end_op = ParsedOp(verb="data", positionals=["end"], raw="data end")
    return adapter.dispatch_op(end_op, model, log)


def _cell(model: SheetsModel, row: int, col: int):
    """Get cell value from active sheet."""
    return model.wb.active.cell(row=row, column=col).value


# -- Basic CSV Tests --


class TestDataBlockBasicCSV:
    """Basic CSV data entry tests."""

    def test_single_row_numbers(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        result = _run_data_block(adapter, model, log, "A1", ["1,2,3"])
        assert result.success
        assert _cell(model, 1, 1) == 1
        assert _cell(model, 1, 2) == 2
        assert _cell(model, 1, 3) == 3

    def test_single_row_text(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        result = _run_data_block(adapter, model, log, "A1", ["Hello,World,Foo"])
        assert result.success
        assert _cell(model, 1, 1) == "Hello"
        assert _cell(model, 1, 2) == "World"
        assert _cell(model, 1, 3) == "Foo"

    def test_multiple_rows(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        result = _run_data_block(adapter, model, log, "A1", [
            "Name,Age,City",
            "Alice,30,NYC",
            "Bob,25,LA",
        ])
        assert result.success
        assert "3 rows" in result.message
        assert _cell(model, 1, 1) == "Name"
        assert _cell(model, 2, 1) == "Alice"
        assert _cell(model, 2, 2) == 30
        assert _cell(model, 3, 1) == "Bob"
        assert _cell(model, 3, 3) == "LA"

    def test_mixed_types(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        result = _run_data_block(adapter, model, log, "A1", ["42,3.14,Hello,=SUM(A1:A3)"])
        assert result.success
        assert _cell(model, 1, 1) == 42
        assert _cell(model, 1, 2) == 3.14
        assert _cell(model, 1, 3) == "Hello"
        assert _cell(model, 1, 4) == "=SUM(A1:A3)"

    def test_anchor_offset(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        result = _run_data_block(adapter, model, log, "C3", ["10,20,30"])
        assert result.success
        assert _cell(model, 3, 3) == 10  # C3
        assert _cell(model, 3, 4) == 20  # D3
        assert _cell(model, 3, 5) == 30  # E3

    def test_empty_data_block_error(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        result = _run_data_block(adapter, model, log, "A1", [])
        assert not result.success
        assert "Empty data block" in result.message

    def test_csv_with_spaces(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        result = _run_data_block(adapter, model, log, "A1", [" hello , world , 42 "])
        assert result.success
        assert _cell(model, 1, 1) == "hello"
        assert _cell(model, 1, 2) == "world"
        assert _cell(model, 1, 3) == 42

    def test_csv_quoted_values(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        """CSV with quoted values containing commas."""
        result = _run_data_block(adapter, model, log, "A1", ['"one,two",three,four'])
        assert result.success
        assert _cell(model, 1, 1) == "one,two"
        assert _cell(model, 1, 2) == "three"

    def test_multicolumn_data(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        result = _run_data_block(adapter, model, log, "A1", [
            "A,B,C,D,E",
            "1,2,3,4,5",
        ])
        assert result.success
        for i in range(1, 6):
            assert _cell(model, 2, i) == i

    def test_data_end_without_start(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        """data end without data ANCHOR should fail."""
        end_op = ParsedOp(verb="data", positionals=["end"], raw="data end")
        result = adapter.dispatch_op(end_op, model, log)
        assert not result.success
        assert "without prior" in result.message

    def test_data_no_anchor(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        """data with no positionals should fail."""
        op = ParsedOp(verb="data", positionals=[], raw="data")
        result = adapter.dispatch_op(op, model, log)
        assert not result.success
        assert "Usage" in result.message


# -- C1: Robust Type Inference --


class TestDataBlockTypeInference:
    """C1: Leading zeros, quoted strings, formulas in data blocks."""

    def test_leading_zeros_preserved(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        """C1: Leading zeros like 01234 should be text, not number 1234."""
        result = _run_data_block(adapter, model, log, "A1", ["01234,007,00100"])
        assert result.success
        assert _cell(model, 1, 1) == "01234"
        assert _cell(model, 1, 2) == "007"
        assert _cell(model, 1, 3) == "00100"

    def test_single_zero_is_number(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        result = _run_data_block(adapter, model, log, "A1", ["0"])
        assert result.success
        assert _cell(model, 1, 1) == 0

    def test_zero_point_five_is_float(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        result = _run_data_block(adapter, model, log, "A1", ["0.5"])
        assert result.success
        assert _cell(model, 1, 1) == 0.5

    def test_quoted_string_strips_quotes(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        result = _run_data_block(adapter, model, log, "A1", ['"Hello World"'])
        assert result.success
        # CSV reader will strip the quotes for us
        assert _cell(model, 1, 1) == "Hello World"

    def test_formula_preserved(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        result = _run_data_block(adapter, model, log, "A1", ["=SUM(A2:A10),=B1*2"])
        assert result.success
        assert _cell(model, 1, 1) == "=SUM(A2:A10)"
        assert _cell(model, 1, 2) == "=B1*2"

    def test_negative_numbers(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        result = _run_data_block(adapter, model, log, "A1", ["-5,-3.14"])
        assert result.success
        assert _cell(model, 1, 1) == -5
        assert _cell(model, 1, 2) == -3.14

    def test_empty_cells_in_csv(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        result = _run_data_block(adapter, model, log, "A1", ["a,,c"])
        assert result.success
        assert _cell(model, 1, 1) == "a"
        assert _cell(model, 1, 2) == ""
        assert _cell(model, 1, 3) == "c"

    def test_integer_parsing(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        result = _run_data_block(adapter, model, log, "A1", ["100,200,300"])
        assert result.success
        assert isinstance(_cell(model, 1, 1), int)
        assert _cell(model, 1, 1) == 100

    def test_float_parsing(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        result = _run_data_block(adapter, model, log, "A1", ["1.5,2.7,3.9"])
        assert result.success
        assert isinstance(_cell(model, 1, 1), float)
        assert _cell(model, 1, 1) == 1.5

    def test_text_fallback(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        """Non-numeric, non-formula values are text."""
        result = _run_data_block(adapter, model, log, "A1", ["hello,world,foo-bar"])
        assert result.success
        assert isinstance(_cell(model, 1, 1), str)
        assert _cell(model, 1, 1) == "hello"
        assert _cell(model, 1, 3) == "foo-bar"


# -- C2: Markdown Table Detection --


class TestDataBlockMarkdown:
    """C2: Markdown table auto-detection and conversion."""

    def test_markdown_basic(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        result = _run_data_block(adapter, model, log, "A1", [
            "| Name | Age |",
            "| Alice | 30 |",
            "| Bob | 25 |",
        ])
        assert result.success
        assert "Markdown table detected" in result.message
        assert _cell(model, 1, 1) == "Name"
        assert _cell(model, 1, 2) == "Age"
        assert _cell(model, 2, 1) == "Alice"
        assert _cell(model, 2, 2) == 30
        assert _cell(model, 3, 1) == "Bob"
        assert _cell(model, 3, 2) == 25

    def test_markdown_separator_skipped(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        """Separator lines like |---|---| should be skipped."""
        result = _run_data_block(adapter, model, log, "A1", [
            "| Name | Age |",
            "|------|-----|",
            "| Alice | 30 |",
        ])
        assert result.success
        assert "2 rows" in result.message  # header + 1 data row
        assert _cell(model, 1, 1) == "Name"
        assert _cell(model, 2, 1) == "Alice"
        # Row 3 should be empty (separator was skipped)
        assert _cell(model, 3, 1) is None

    def test_markdown_separator_with_colons(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        """Separator with alignment colons should be skipped too."""
        result = _run_data_block(adapter, model, log, "A1", [
            "| Left | Center | Right |",
            "|:-----|:------:|------:|",
            "| A | B | C |",
        ])
        assert result.success
        assert "2 rows" in result.message

    def test_markdown_numbers_parsed(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        result = _run_data_block(adapter, model, log, "A1", [
            "| Val |",
            "| 42 |",
            "| 3.14 |",
        ])
        assert result.success
        assert _cell(model, 2, 1) == 42
        assert _cell(model, 3, 1) == 3.14

    def test_markdown_with_formulas(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        result = _run_data_block(adapter, model, log, "A1", [
            "| A | B |",
            "| 10 | =A1*2 |",
        ])
        assert result.success
        assert _cell(model, 2, 1) == 10
        assert _cell(model, 2, 2) == "=A1*2"

    def test_markdown_leading_zeros(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        """C1 type inference applies inside markdown tables too."""
        result = _run_data_block(adapter, model, log, "A1", [
            "| Code |",
            "| 007 |",
            "| 01234 |",
        ])
        assert result.success
        assert _cell(model, 2, 1) == "007"
        assert _cell(model, 3, 1) == "01234"

    def test_csv_not_markdown(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        """Lines not starting with | should be parsed as CSV."""
        result = _run_data_block(adapter, model, log, "A1", ["A,B,C", "1,2,3"])
        assert result.success
        assert "Markdown" not in result.message

    def test_markdown_empty_cells(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        result = _run_data_block(adapter, model, log, "A1", [
            "| A |  | C |",
        ])
        assert result.success
        assert _cell(model, 1, 1) == "A"
        assert _cell(model, 1, 2) == ""
        assert _cell(model, 1, 3) == "C"

    def test_markdown_three_columns(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        result = _run_data_block(adapter, model, log, "A1", [
            "| X | Y | Z |",
            "|---|---|---|",
            "| 1 | 2 | 3 |",
            "| 4 | 5 | 6 |",
        ])
        assert result.success
        assert "3 rows" in result.message
        assert _cell(model, 1, 1) == "X"
        assert _cell(model, 3, 3) == 6


# -- C6: Anchor Resolution --


class TestDataBlockAnchors:
    """C6 anchor resolution for data blocks."""

    def test_anchor_bottom_left(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        """Data at @bottom_left should append below existing data."""
        # First populate some data
        ws = model.wb.active
        ws.cell(row=1, column=1, value="Header")
        ws.cell(row=2, column=1, value="Data1")
        adapter.index.expand_bounds("Sheet1", 1, 1)
        adapter.index.expand_bounds("Sheet1", 2, 1)

        result = _run_data_block(adapter, model, log, "@bottom_left", ["NewData,Extra"])
        assert result.success
        # @bottom_left = (min_col=1, max_row+1=3)
        assert _cell(model, 3, 1) == "NewData"
        assert _cell(model, 3, 2) == "Extra"

    def test_anchor_with_offset(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        ws = model.wb.active
        ws.cell(row=1, column=1, value="X")
        adapter.index.expand_bounds("Sheet1", 1, 1)

        result = _run_data_block(adapter, model, log, "@bottom_left+1", ["Value"])
        assert result.success
        # @bottom_left+1 = (1, 1+1+1=3) with offset
        assert _cell(model, 3, 1) == "Value"

    def test_invalid_anchor(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        result = _run_data_block(adapter, model, log, "@invalid_anchor", ["1,2,3"])
        assert not result.success
        assert "Invalid anchor" in result.message


# -- C9: Collision Detection --


class TestDataBlockCollisions:
    """C9: Collision detection when overwriting existing cells."""

    def test_overwrite_warns(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        """Overwriting existing non-empty cells should produce a warning."""
        # Pre-populate cells
        ws = model.wb.active
        ws.cell(row=1, column=1, value="Existing1")
        ws.cell(row=1, column=2, value="Existing2")
        adapter.index.expand_bounds("Sheet1", 1, 1)
        adapter.index.expand_bounds("Sheet1", 1, 2)

        result = _run_data_block(adapter, model, log, "A1", ["New1,New2"])
        assert result.success
        assert "Overwrote" in result.message
        assert "2 non-empty" in result.message
        # Values should be overwritten
        assert _cell(model, 1, 1) == "New1"
        assert _cell(model, 1, 2) == "New2"

    def test_no_collision_empty_cells(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        """Writing to empty cells should not produce collision warning."""
        result = _run_data_block(adapter, model, log, "A1", ["1,2,3"])
        assert result.success
        assert "Overwrote" not in result.message

    def test_partial_collision(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        """Only some cells have existing data — warning lists them."""
        ws = model.wb.active
        ws.cell(row=1, column=1, value="Existing")
        adapter.index.expand_bounds("Sheet1", 1, 1)

        result = _run_data_block(adapter, model, log, "A1", ["New,Empty"])
        assert result.success
        assert "1 non-empty" in result.message
        assert "A1" in result.message

    def test_collision_many_cells(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        """Many collisions show preview with +N more."""
        ws = model.wb.active
        for i in range(1, 9):
            ws.cell(row=1, column=i, value=f"Old{i}")
            adapter.index.expand_bounds("Sheet1", 1, i)

        result = _run_data_block(adapter, model, log, "A1", [
            "1,2,3,4,5,6,7,8",
        ])
        assert result.success
        assert "8 non-empty" in result.message
        assert "+3 more" in result.message  # 8 - 5 preview = 3 more


# -- Snapshot / Undo Integration --


class TestDataBlockSnapshot:
    """Data block should create undo snapshot."""

    def test_data_block_creates_snapshot(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        result = _run_data_block(adapter, model, log, "A1", ["1,2,3"])
        assert result.success
        # EventLog should have one event
        assert len(log) == 1

    def test_data_block_records_modified(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        _run_data_block(adapter, model, log, "B2", ["x,y", "a,b"])
        recent = adapter.index.get_recent(1)
        assert len(recent) == 1
        assert "Sheet1" in recent[0][0]


# -- Edge Cases --


class TestDataBlockEdgeCases:
    """Edge cases for data block mode."""

    def test_blank_lines_ignored(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        result = _run_data_block(adapter, model, log, "A1", [
            "",
            "1,2",
            "",
            "3,4",
            "",
        ])
        assert result.success
        assert "2 rows" in result.message
        assert _cell(model, 1, 1) == 1
        assert _cell(model, 2, 1) == 3

    def test_single_cell_data_block(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        result = _run_data_block(adapter, model, log, "A1", ["42"])
        assert result.success
        assert _cell(model, 1, 1) == 42

    def test_all_blank_lines_error(self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog):
        """Data block with only blank lines should fail."""
        result = _run_data_block(adapter, model, log, "A1", ["", "  ", ""])
        assert not result.success
        assert "No data rows" in result.message
