"""Tests for cell reference parsing — all ref types + C6 anchors."""

from __future__ import annotations

import pytest

from fcp_sheets.model.refs import (
    AnchorRef,
    CellRef,
    ColRef,
    RangeRef,
    RowRef,
    col_to_index,
    index_to_col,
    parse_anchor,
    parse_cell_ref,
    parse_range_ref,
    parse_ref,
)


# -- Column conversion --

class TestColConversion:
    def test_a_to_1(self):
        assert col_to_index("A") == 1

    def test_z_to_26(self):
        assert col_to_index("Z") == 26

    def test_aa_to_27(self):
        assert col_to_index("AA") == 27

    def test_az_to_52(self):
        assert col_to_index("AZ") == 52

    def test_ba_to_53(self):
        assert col_to_index("BA") == 53

    def test_lowercase(self):
        assert col_to_index("a") == 1
        assert col_to_index("aa") == 27

    def test_index_to_col_1(self):
        assert index_to_col(1) == "A"

    def test_index_to_col_26(self):
        assert index_to_col(26) == "Z"

    def test_index_to_col_27(self):
        assert index_to_col(27) == "AA"

    def test_roundtrip(self):
        for i in range(1, 100):
            assert col_to_index(index_to_col(i)) == i


# -- Cell reference parsing --

class TestParseCellRef:
    def test_a1(self):
        ref = parse_cell_ref("A1")
        assert ref == CellRef(col=1, row=1)

    def test_b10(self):
        ref = parse_cell_ref("B10")
        assert ref == CellRef(col=2, row=10)

    def test_aa100(self):
        ref = parse_cell_ref("AA100")
        assert ref == CellRef(col=27, row=100)

    def test_lowercase(self):
        ref = parse_cell_ref("b2")
        assert ref == CellRef(col=2, row=2)

    def test_with_sheet(self):
        ref = parse_cell_ref("Sheet2!A1")
        assert ref == CellRef(col=1, row=1, sheet="Sheet2")

    def test_with_quoted_sheet(self):
        ref = parse_cell_ref("'My Sheet'!B3")
        assert ref == CellRef(col=2, row=3, sheet="My Sheet")

    def test_invalid(self):
        assert parse_cell_ref("123") is None
        assert parse_cell_ref("ABC") is None
        assert parse_cell_ref("") is None
        assert parse_cell_ref("A1:B2") is None


# -- Range reference parsing --

class TestParseRangeRef:
    def test_cell_range(self):
        ref = parse_range_ref("A1:D10")
        assert isinstance(ref, RangeRef)
        assert ref.start == CellRef(col=1, row=1)
        assert ref.end == CellRef(col=4, row=10)

    def test_col_range(self):
        ref = parse_range_ref("B:B")
        assert isinstance(ref, ColRef)
        assert ref.start_col == 2
        assert ref.end_col == 2

    def test_multi_col_range(self):
        ref = parse_range_ref("A:E")
        assert isinstance(ref, ColRef)
        assert ref.start_col == 1
        assert ref.end_col == 5

    def test_row_range(self):
        ref = parse_range_ref("3:3")
        assert isinstance(ref, RowRef)
        assert ref.start_row == 3
        assert ref.end_row == 3

    def test_multi_row_range(self):
        ref = parse_range_ref("1:5")
        assert isinstance(ref, RowRef)
        assert ref.start_row == 1
        assert ref.end_row == 5

    def test_with_sheet(self):
        ref = parse_range_ref("Sheet2!A1:B10")
        assert isinstance(ref, RangeRef)
        assert ref.sheet == "Sheet2"

    def test_no_colon(self):
        assert parse_range_ref("A1") is None

    def test_invalid(self):
        assert parse_range_ref("") is None


# -- Anchor parsing (C6) --

class TestParseAnchor:
    def test_bottom_left(self):
        ref = parse_anchor("@bottom_left")
        assert ref == AnchorRef(anchor="bottom_left", offset=0)

    def test_bottom_right(self):
        ref = parse_anchor("@bottom_right")
        assert ref == AnchorRef(anchor="bottom_right", offset=0)

    def test_right_top(self):
        ref = parse_anchor("@right_top")
        assert ref == AnchorRef(anchor="right_top", offset=0)

    def test_with_offset(self):
        ref = parse_anchor("@bottom_left+2")
        assert ref == AnchorRef(anchor="bottom_left", offset=2)

    def test_right_top_with_offset(self):
        ref = parse_anchor("@right_top+5")
        assert ref == AnchorRef(anchor="right_top", offset=5)

    def test_invalid_anchor(self):
        assert parse_anchor("@invalid") is None
        assert parse_anchor("bottom_left") is None  # Missing @
        assert parse_anchor("A1") is None


# -- parse_ref (unified) --

class TestParseRef:
    def test_cell(self):
        ref = parse_ref("A1")
        assert isinstance(ref, CellRef)

    def test_range(self):
        ref = parse_ref("A1:D10")
        assert isinstance(ref, RangeRef)

    def test_anchor(self):
        ref = parse_ref("@bottom_left+2")
        assert isinstance(ref, AnchorRef)

    def test_col(self):
        ref = parse_ref("B:B")
        assert isinstance(ref, ColRef)

    def test_row(self):
        ref = parse_ref("3:3")
        assert isinstance(ref, RowRef)
