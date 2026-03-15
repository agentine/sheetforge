"""Shared test fixtures for sheetcraft."""

from __future__ import annotations

import os
import tempfile
from typing import Generator

import pytest

from sheetcraft import Workbook


@pytest.fixture
def tmp_xlsx() -> Generator[str, None, None]:
    """Provide a temporary .xlsx file path, cleaned up after test."""
    fd, path = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    yield path
    if os.path.exists(path):
        os.unlink(path)


@pytest.fixture
def sample_workbook() -> Workbook:
    """Create a sample workbook with various data types."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None

    # Headers
    ws["A1"] = "Name"
    ws["B1"] = "Age"
    ws["C1"] = "Score"
    ws["D1"] = "Active"

    # Data rows
    ws["A2"] = "Alice"
    ws["B2"] = 30
    ws["C2"] = 95.5
    ws["D2"] = True

    ws["A3"] = "Bob"
    ws["B3"] = 25
    ws["C3"] = 88.0
    ws["D3"] = False

    # Formula
    ws["C4"] = "=AVERAGE(C2:C3)"

    return wb


@pytest.fixture
def multi_sheet_workbook() -> Workbook:
    """Create a workbook with multiple sheets."""
    wb = Workbook()
    ws1 = wb.active
    assert ws1 is not None
    ws1["A1"] = "Sheet1 Data"

    ws2 = wb.create_sheet("Summary")
    ws2["A1"] = "Summary Data"
    ws2["A2"] = 100

    ws3 = wb.create_sheet("Config")
    ws3["A1"] = "Setting"
    ws3["B1"] = "Value"

    return wb
