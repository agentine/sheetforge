"""Cell class for sheetforge."""

from __future__ import annotations

from datetime import date, datetime, time
from typing import TYPE_CHECKING

from sheetforge.constants import TYPE_BOOL, TYPE_ERROR, TYPE_FORMULA, TYPE_NUMERIC, TYPE_STRING
from sheetforge.styles.alignment import Alignment
from sheetforge.styles.borders import Border
from sheetforge.styles.fills import PatternFill
from sheetforge.styles.fonts import Font
from sheetforge.styles.numbers import FORMAT_GENERAL
from sheetforge.utils import get_column_letter

if TYPE_CHECKING:
    from sheetforge.styles import CellStyle, StyleSheet

# Sentinel for "no value"
_EMPTY: object = object()

CellValue = str | int | float | bool | datetime | date | time | None


class Cell:
    """Represents a single cell in a worksheet."""

    __slots__ = (
        "_row",
        "_column",
        "_value",
        "_data_type",
        "_style_id",
        "_number_format",
        "_font",
        "_fill",
        "_border",
        "_alignment",
        "_comment",
        "_hyperlink",
    )

    def __init__(
        self,
        row: int = 1,
        column: int = 1,
        value: CellValue = None,
        *,
        style_id: int = 0,
    ) -> None:
        self._row = row
        self._column = column
        self._style_id = style_id
        self._number_format: str = FORMAT_GENERAL
        self._font: Font | None = None
        self._fill: PatternFill | None = None
        self._border: Border | None = None
        self._alignment: Alignment | None = None
        self._comment: str | None = None
        self._hyperlink: str | None = None
        self._value: CellValue = None
        self._data_type: str = "n"
        self.value = value  # use property setter

    @property
    def row(self) -> int:
        return self._row

    @property
    def column(self) -> int:
        return self._column

    @property
    def column_letter(self) -> str:
        return get_column_letter(self._column)

    @property
    def coordinate(self) -> str:
        return f"{self.column_letter}{self._row}"

    @property
    def value(self) -> CellValue:
        return self._value

    @value.setter
    def value(self, val: CellValue) -> None:
        self._value = val
        self._data_type = self._infer_type(val)

    @property
    def data_type(self) -> str:
        return self._data_type

    @data_type.setter
    def data_type(self, val: str) -> None:
        self._data_type = val

    @property
    def number_format(self) -> str:
        return self._number_format

    @number_format.setter
    def number_format(self, val: str) -> None:
        self._number_format = val

    @property
    def font(self) -> Font | None:
        return self._font

    @font.setter
    def font(self, val: Font | None) -> None:
        self._font = val

    @property
    def fill(self) -> PatternFill | None:
        return self._fill

    @fill.setter
    def fill(self, val: PatternFill | None) -> None:
        self._fill = val

    @property
    def border(self) -> Border | None:
        return self._border

    @border.setter
    def border(self, val: Border | None) -> None:
        self._border = val

    @property
    def alignment(self) -> Alignment | None:
        return self._alignment

    @alignment.setter
    def alignment(self, val: Alignment | None) -> None:
        self._alignment = val

    @property
    def comment(self) -> str | None:
        return self._comment

    @comment.setter
    def comment(self, val: str | None) -> None:
        self._comment = val

    @property
    def hyperlink(self) -> str | None:
        return self._hyperlink

    @hyperlink.setter
    def hyperlink(self, val: str | None) -> None:
        self._hyperlink = val

    @property
    def style_id(self) -> int:
        return self._style_id

    @style_id.setter
    def style_id(self, val: int) -> None:
        self._style_id = val

    @staticmethod
    def _infer_type(val: CellValue) -> str:
        """Infer the OOXML data type from a Python value."""
        if val is None:
            return "n"  # null
        if isinstance(val, bool):
            return TYPE_BOOL
        if isinstance(val, (int, float)):
            return TYPE_NUMERIC
        if isinstance(val, str):
            if val.startswith("="):
                return TYPE_FORMULA
            return TYPE_STRING
        if isinstance(val, (datetime, date, time)):
            return TYPE_NUMERIC  # dates stored as numbers in xlsx
        return TYPE_STRING

    def __repr__(self) -> str:
        return f"<Cell {self.coordinate}: {self._value!r}>"
