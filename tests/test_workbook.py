"""Tests for sheetforge.workbook."""

from __future__ import annotations

import tempfile
import os

import pytest

from sheetforge import Workbook, load_workbook, Worksheet
from sheetforge.exceptions import SheetTitleException, InvalidFileException, WorkbookAlreadySaved


class TestWorkbookCreate:
    def test_default(self) -> None:
        wb = Workbook()
        assert len(wb) == 1
        assert wb.sheetnames == ["Sheet"]
        assert wb.active is not None

    def test_active(self) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        assert ws.title == "Sheet"


class TestWorkbookSheets:
    def test_create_sheet(self) -> None:
        wb = Workbook()
        ws = wb.create_sheet("Data")
        assert ws.title == "Data"
        assert "Data" in wb.sheetnames
        assert len(wb) == 2

    def test_create_sheet_auto_name(self) -> None:
        wb = Workbook()
        ws = wb.create_sheet()
        assert ws.title == "Sheet1"

    def test_create_sheet_at_index(self) -> None:
        wb = Workbook()
        ws = wb.create_sheet("First", index=0)
        assert wb.sheetnames[0] == "First"

    def test_create_sheet_duplicate_name(self) -> None:
        wb = Workbook()
        with pytest.raises(SheetTitleException):
            wb.create_sheet("Sheet")

    def test_remove_sheet(self) -> None:
        wb = Workbook()
        wb.create_sheet("Data")
        wb.remove_sheet("Data")
        assert "Data" not in wb.sheetnames

    def test_remove_by_object(self) -> None:
        wb = Workbook()
        ws = wb.create_sheet("Data")
        wb.remove(ws)
        assert len(wb) == 1

    def test_getitem(self) -> None:
        wb = Workbook()
        ws = wb["Sheet"]
        assert ws.title == "Sheet"

    def test_getitem_missing(self) -> None:
        wb = Workbook()
        with pytest.raises(KeyError):
            wb["Missing"]

    def test_contains(self) -> None:
        wb = Workbook()
        assert "Sheet" in wb
        assert "Missing" not in wb

    def test_iter(self) -> None:
        wb = Workbook()
        wb.create_sheet("Data")
        names = [ws.title for ws in wb]
        assert names == ["Sheet", "Data"]

    def test_set_active(self) -> None:
        wb = Workbook()
        ws2 = wb.create_sheet("Data")
        wb.active = ws2
        assert wb.active is ws2


class TestWorkbookSave:
    def test_save_and_load(self) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws["A1"] = "Hello"
        ws["B1"] = 42
        ws["C1"] = 3.14
        ws["D1"] = True
        ws["A2"] = "=SUM(B1:C1)"
        ws.append([10, 20, 30])

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            wb.save(f.name)
            path = f.name

        try:
            wb2 = load_workbook(path)
            ws2 = wb2.active
            assert ws2 is not None
            assert ws2["A1"].value == "Hello"
            assert ws2["B1"].value == 42
            assert ws2["C1"].value == 3.14
            assert ws2["D1"].value is True
            assert ws2["A2"].value == "=SUM(B1:C1)"
            assert ws2["A3"].value == 10
        finally:
            os.unlink(path)

    def test_save_multiple_sheets(self) -> None:
        wb = Workbook()
        ws1 = wb.active
        assert ws1 is not None
        ws1["A1"] = "Sheet1"
        ws2 = wb.create_sheet("Data")
        ws2["A1"] = "Sheet2"

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            wb.save(f.name)
            path = f.name

        try:
            wb2 = load_workbook(path)
            assert wb2.sheetnames == ["Sheet", "Data"]
            assert wb2["Sheet"]["A1"].value == "Sheet1"
            assert wb2["Data"]["A1"].value == "Sheet2"
        finally:
            os.unlink(path)

    def test_save_with_styles(self) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws["A1"] = "Styled"
        ws.column_dimensions[1] = 20.0
        ws.row_dimensions[1] = 30.0
        ws.merge_cells("B1:C1")
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = "A1:C10"

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            wb.save(f.name)
            path = f.name

        try:
            wb2 = load_workbook(path)
            ws2 = wb2.active
            assert ws2 is not None
            assert ws2["A1"].value == "Styled"
            assert len(ws2.merged_cells) == 1
            assert ws2.freeze_panes == "A2"
            assert ws2.auto_filter.ref == "A1:C10"
        finally:
            os.unlink(path)


class TestWorkbookClose:
    def test_close_then_save(self) -> None:
        wb = Workbook()
        wb.close()
        with pytest.raises(WorkbookAlreadySaved):
            wb.save("/tmp/test.xlsx")


class TestLoadWorkbook:
    def test_invalid_file(self) -> None:
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            f.write(b"not a zip file")
            path = f.name
        try:
            with pytest.raises(InvalidFileException):
                load_workbook(path)
        finally:
            os.unlink(path)
