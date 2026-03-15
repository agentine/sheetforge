"""Tests for sheetcraft.cell."""

from __future__ import annotations

from datetime import date, datetime, time

from sheetcraft.cell import Cell
from sheetcraft.constants import TYPE_BOOL, TYPE_FORMULA, TYPE_NUMERIC, TYPE_STRING
from sheetcraft.styles.fonts import Font
from sheetcraft.styles.fills import PatternFill
from sheetcraft.styles.borders import Border, Side
from sheetcraft.styles.alignment import Alignment


class TestCellCoordinates:
    def test_default(self) -> None:
        c = Cell()
        assert c.row == 1
        assert c.column == 1
        assert c.column_letter == "A"
        assert c.coordinate == "A1"

    def test_custom_position(self) -> None:
        c = Cell(row=5, column=3)
        assert c.row == 5
        assert c.column == 3
        assert c.column_letter == "C"
        assert c.coordinate == "C5"


class TestCellValue:
    def test_none(self) -> None:
        c = Cell()
        assert c.value is None
        assert c.data_type == "n"

    def test_string(self) -> None:
        c = Cell(value="hello")
        assert c.value == "hello"
        assert c.data_type == TYPE_STRING

    def test_int(self) -> None:
        c = Cell(value=42)
        assert c.value == 42
        assert c.data_type == TYPE_NUMERIC

    def test_float(self) -> None:
        c = Cell(value=3.14)
        assert c.value == 3.14
        assert c.data_type == TYPE_NUMERIC

    def test_bool(self) -> None:
        c = Cell(value=True)
        assert c.value is True
        assert c.data_type == TYPE_BOOL

    def test_bool_false(self) -> None:
        c = Cell(value=False)
        assert c.value is False
        assert c.data_type == TYPE_BOOL

    def test_formula(self) -> None:
        c = Cell(value="=SUM(A1:A10)")
        assert c.value == "=SUM(A1:A10)"
        assert c.data_type == TYPE_FORMULA

    def test_datetime(self) -> None:
        dt = datetime(2024, 1, 15, 10, 30)
        c = Cell(value=dt)
        assert c.value == dt
        assert c.data_type == TYPE_NUMERIC

    def test_date(self) -> None:
        d = date(2024, 1, 15)
        c = Cell(value=d)
        assert c.value == d
        assert c.data_type == TYPE_NUMERIC

    def test_time(self) -> None:
        t = time(10, 30)
        c = Cell(value=t)
        assert c.value == t
        assert c.data_type == TYPE_NUMERIC

    def test_set_value_updates_type(self) -> None:
        c = Cell(value="hello")
        assert c.data_type == TYPE_STRING
        c.value = 42
        assert c.data_type == TYPE_NUMERIC
        c.value = True
        assert c.data_type == TYPE_BOOL


class TestCellStyles:
    def test_number_format(self) -> None:
        c = Cell()
        assert c.number_format == "General"
        c.number_format = "0.00"
        assert c.number_format == "0.00"

    def test_font(self) -> None:
        c = Cell()
        assert c.font is None
        f = Font(bold=True)
        c.font = f
        assert c.font is f
        assert c.font.bold is True

    def test_fill(self) -> None:
        c = Cell()
        assert c.fill is None
        fill = PatternFill(patternType="solid", fgColor="FF0000")
        c.fill = fill
        assert c.fill is fill

    def test_border(self) -> None:
        c = Cell()
        assert c.border is None
        border = Border(left=Side(style="thin"))
        c.border = border
        assert c.border is border

    def test_alignment(self) -> None:
        c = Cell()
        assert c.alignment is None
        align = Alignment(horizontal="center")
        c.alignment = align
        assert c.alignment is align


class TestCellMisc:
    def test_comment(self) -> None:
        c = Cell()
        assert c.comment is None
        c.comment = "A note"
        assert c.comment == "A note"

    def test_hyperlink(self) -> None:
        c = Cell()
        assert c.hyperlink is None
        c.hyperlink = "https://example.com"
        assert c.hyperlink == "https://example.com"

    def test_repr(self) -> None:
        c = Cell(row=1, column=1, value="test")
        assert "A1" in repr(c)
        assert "test" in repr(c)
