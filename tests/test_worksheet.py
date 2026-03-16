"""Tests for sheetforge.worksheet."""

from __future__ import annotations

import pytest

from sheetforge.worksheet import Worksheet, AutoFilter, MergedCellRange
from sheetforge.exceptions import SheetTitleException


class TestWorksheetTitle:
    def test_default_title(self) -> None:
        ws = Worksheet()
        assert ws.title == "Sheet"

    def test_custom_title(self) -> None:
        ws = Worksheet(title="Data")
        assert ws.title == "Data"

    def test_set_title(self) -> None:
        ws = Worksheet()
        ws.title = "NewName"
        assert ws.title == "NewName"

    def test_invalid_characters(self) -> None:
        with pytest.raises(SheetTitleException):
            Worksheet(title="Sheet*1")
        with pytest.raises(SheetTitleException):
            Worksheet(title="Sheet:1")

    def test_empty_title(self) -> None:
        with pytest.raises(SheetTitleException):
            Worksheet(title="")

    def test_too_long(self) -> None:
        with pytest.raises(SheetTitleException):
            Worksheet(title="A" * 32)


class TestWorksheetCell:
    def test_cell_create(self) -> None:
        ws = Worksheet()
        c = ws.cell(1, 1, "hello")
        assert c.value == "hello"
        assert c.row == 1
        assert c.column == 1

    def test_cell_get_existing(self) -> None:
        ws = Worksheet()
        c1 = ws.cell(1, 1, "hello")
        c2 = ws.cell(1, 1)
        assert c1 is c2
        assert c2.value == "hello"

    def test_cell_set_value(self) -> None:
        ws = Worksheet()
        ws.cell(1, 1, "first")
        ws.cell(1, 1, "second")
        assert ws.cell(1, 1).value == "second"


class TestWorksheetGetItem:
    def test_single_cell(self) -> None:
        ws = Worksheet()
        ws["A1"] = "test"
        assert ws["A1"].value == "test"

    def test_range(self) -> None:
        ws = Worksheet()
        ws["A1"] = 1
        ws["B1"] = 2
        ws["A2"] = 3
        ws["B2"] = 4
        result = ws["A1:B2"]
        assert len(result) == 2  # 2 rows
        assert len(result[0]) == 2  # 2 cols
        assert result[0][0].value == 1
        assert result[1][1].value == 4

    def test_row_access(self) -> None:
        ws = Worksheet()
        ws["A1"] = 1
        ws["B1"] = 2
        row = ws[1]
        assert len(row) == 2
        assert row[0].value == 1

    def test_column_access(self) -> None:
        ws = Worksheet()
        ws["A1"] = 1
        ws["A2"] = 2
        ws["B1"] = 3  # extend max_row
        col = ws["A"]
        assert len(col) == 2
        assert col[0].value == 1
        assert col[1].value == 2


class TestWorksheetDimensions:
    def test_empty(self) -> None:
        ws = Worksheet()
        assert ws.max_row == 0
        assert ws.max_column == 0
        assert ws.dimensions == "A1:A1"

    def test_with_data(self) -> None:
        ws = Worksheet()
        ws["A1"] = 1
        ws["C3"] = 2
        assert ws.max_row == 3
        assert ws.max_column == 3
        assert ws.min_row == 1
        assert ws.min_column == 1
        assert ws.dimensions == "A1:C3"


class TestWorksheetIterRows:
    def test_basic(self) -> None:
        ws = Worksheet()
        ws["A1"] = 1
        ws["B1"] = 2
        ws["A2"] = 3
        ws["B2"] = 4
        rows = list(ws.iter_rows())
        assert len(rows) == 2
        assert rows[0][0].value == 1
        assert rows[1][1].value == 4

    def test_values_only(self) -> None:
        ws = Worksheet()
        ws["A1"] = 1
        ws["B1"] = 2
        rows = list(ws.iter_rows(values_only=True))
        assert rows[0] == (1, 2)

    def test_with_bounds(self) -> None:
        ws = Worksheet()
        ws["A1"] = 1
        ws["B2"] = 2
        ws["C3"] = 3
        rows = list(ws.iter_rows(min_row=2, max_row=3, min_col=2, max_col=3, values_only=True))
        assert len(rows) == 2
        assert rows[0][0] == 2


class TestWorksheetIterCols:
    def test_basic(self) -> None:
        ws = Worksheet()
        ws["A1"] = 1
        ws["A2"] = 3
        ws["B1"] = 2
        ws["B2"] = 4
        cols = list(ws.iter_cols())
        assert len(cols) == 2
        assert cols[0][0].value == 1
        assert cols[0][1].value == 3

    def test_values_only(self) -> None:
        ws = Worksheet()
        ws["A1"] = 1
        ws["A2"] = 3
        ws["B1"] = 2
        ws["B2"] = 4
        cols = list(ws.iter_cols(values_only=True))
        assert cols[0] == (1, 3)
        assert cols[1] == (2, 4)


class TestWorksheetAppend:
    def test_append(self) -> None:
        ws = Worksheet()
        ws.append([1, 2, 3])
        ws.append([4, 5, 6])
        assert ws["A1"].value == 1
        assert ws["C1"].value == 3
        assert ws["A2"].value == 4
        assert ws.max_row == 2


class TestMergedCells:
    def test_merge(self) -> None:
        ws = Worksheet()
        ws.merge_cells("A1:B2")
        assert len(ws.merged_cells) == 1
        assert "A1:B2" in ws.merged_cells.ranges

    def test_unmerge(self) -> None:
        ws = Worksheet()
        ws.merge_cells("A1:B2")
        ws.unmerge_cells("A1:B2")
        assert len(ws.merged_cells) == 0


class TestFreezePanes:
    def test_set(self) -> None:
        ws = Worksheet()
        ws.freeze_panes = "A2"
        assert ws.freeze_panes == "A2"


class TestAutoFilter:
    def test_set(self) -> None:
        ws = Worksheet()
        ws.auto_filter.ref = "A1:C10"
        assert ws.auto_filter.ref == "A1:C10"
        assert bool(ws.auto_filter)


class TestDimensionHolder:
    def test_column_dimensions(self) -> None:
        ws = Worksheet()
        ws.column_dimensions[1] = 20.0
        assert ws.column_dimensions[1] == 20.0
        assert ws.column_dimensions[2] is None

    def test_row_dimensions(self) -> None:
        ws = Worksheet()
        ws.row_dimensions[1] = 30.0
        assert ws.row_dimensions[1] == 30.0


class TestRepr:
    def test_repr(self) -> None:
        ws = Worksheet(title="Data")
        assert "Data" in repr(ws)
