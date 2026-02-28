"""Shared test fixtures for fcp-sheets."""

from __future__ import annotations

import pytest
from openpyxl import Workbook

from fcp_core import EventLog

from fcp_sheets.adapter import SheetsAdapter
from fcp_sheets.model.index import SheetIndex
from fcp_sheets.model.snapshot import SheetsModel
from fcp_sheets.server.resolvers import SheetsOpContext


@pytest.fixture
def adapter() -> SheetsAdapter:
    """Fresh SheetsAdapter instance."""
    return SheetsAdapter()


@pytest.fixture
def model(adapter: SheetsAdapter) -> SheetsModel:
    """SheetsModel with a new empty workbook."""
    return adapter.create_empty("Test Workbook", {})


@pytest.fixture
def log() -> EventLog:
    """Fresh EventLog instance."""
    return EventLog()


@pytest.fixture
def ctx(model: SheetsModel, adapter: SheetsAdapter) -> SheetsOpContext:
    """SheetsOpContext pointing at the model's active sheet."""
    return SheetsOpContext(
        wb=model.wb,
        index=adapter.index,
        named_styles={},
    )
