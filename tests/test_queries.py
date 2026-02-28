"""Tests for MVP queries — plan/map, stats, status."""

from __future__ import annotations

import pytest

from fcp_sheets.adapter import SheetsAdapter
from fcp_sheets.model.snapshot import SheetsModel
from fcp_sheets.server.queries import dispatch_query
from fcp_core import EventLog, parse_op


class TestPlanQuery:
    def test_empty_workbook(self, adapter: SheetsAdapter, model: SheetsModel):
        result = dispatch_query("plan", model, adapter.index)
        assert "Test Workbook" in result
        assert "1 sheets" in result
        assert "Sheet1" in result
        assert "[active]" in result

    def test_with_data(self, adapter: SheetsAdapter, model: SheetsModel):
        log = EventLog()
        adapter.dispatch_op(parse_op("set A1 Name"), model, log)
        adapter.dispatch_op(parse_op("set B1 Score"), model, log)
        adapter.dispatch_op(parse_op("set A2 Alice"), model, log)
        adapter.dispatch_op(parse_op("set B2 95"), model, log)

        result = dispatch_query("plan", model, adapter.index)
        assert "data:" in result
        assert "A1:" in result  # Data bounds
        assert "next-empty:" in result

    def test_map_alias(self, adapter: SheetsAdapter, model: SheetsModel):
        result = dispatch_query("map", model, adapter.index)
        assert "Test Workbook" in result

    def test_multiple_sheets(self, adapter: SheetsAdapter):
        model = adapter.create_empty("Multi", {"sheets": "3"})
        result = dispatch_query("plan", model, adapter.index)
        assert "3 sheets" in result
        assert "Sheet1" in result
        assert "Sheet2" in result
        assert "Sheet3" in result


class TestStatsQuery:
    def test_empty(self, adapter: SheetsAdapter, model: SheetsModel):
        result = dispatch_query("stats", model, adapter.index)
        assert "Test Workbook" in result
        assert "Data cells: 0" in result

    def test_with_data(self, adapter: SheetsAdapter, model: SheetsModel):
        log = EventLog()
        adapter.dispatch_op(parse_op("set A1 42"), model, log)
        adapter.dispatch_op(parse_op("set A2 =A1*2"), model, log)

        result = dispatch_query("stats", model, adapter.index)
        assert "Data cells: 1" in result
        assert "Formula cells: 1" in result


class TestStatusQuery:
    def test_status(self, adapter: SheetsAdapter, model: SheetsModel):
        result = dispatch_query("status", model, adapter.index)
        assert "Test Workbook" in result
        assert "unsaved" in result
        assert "Sheet1" in result


class TestUnknownQuery:
    def test_unknown(self, adapter: SheetsAdapter, model: SheetsModel):
        result = dispatch_query("foobar", model, adapter.index)
        assert "Unknown query" in result
        assert "try:" in result
