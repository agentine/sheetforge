"""Worksheet class for sheetforge."""

from __future__ import annotations

import re
from typing import Any, Generator, Iterator, TYPE_CHECKING

from sheetforge.cell import Cell, CellValue
from sheetforge.exceptions import CellCoordinatesException, SheetTitleException
from sheetforge.utils import (
    column_index_from_string,
    coordinate_to_tuple,
    get_column_letter,
    range_boundaries,
)

if TYPE_CHECKING:
    pass

_INVALID_TITLE_RE = re.compile(r"[\\*?:/\[\]]")
_MAX_TITLE_LEN = 31


class DimensionHolder:
    """Holds row or column dimension data (height/width)."""

    def __init__(self) -> None:
        self._data: dict[int, float] = {}

    def __getitem__(self, key: int) -> float | None:
        return self._data.get(key)

    def __setitem__(self, key: int, value: float) -> None:
        self._data[key] = value

    def __contains__(self, key: int) -> bool:
        return key in self._data

    def __iter__(self) -> Iterator[int]:
        return iter(self._data)

    def items(self) -> Iterator[tuple[int, float]]:
        yield from self._data.items()

    def as_dict(self) -> dict[int, float]:
        return dict(self._data)


class AutoFilter:
    """Represents an auto-filter on a worksheet."""

    def __init__(self) -> None:
        self.ref: str | None = None

    def __bool__(self) -> bool:
        return self.ref is not None


class MergedCellRange:
    """Tracks merged cell ranges."""

    def __init__(self) -> None:
        self._ranges: list[str] = []

    @property
    def ranges(self) -> list[str]:
        return list(self._ranges)

    def __iter__(self) -> Iterator[str]:
        return iter(self._ranges)

    def __len__(self) -> int:
        return len(self._ranges)

    def __bool__(self) -> bool:
        return len(self._ranges) > 0

    def add(self, range_string: str) -> None:
        self._ranges.append(range_string)

    def remove(self, range_string: str) -> None:
        self._ranges.remove(range_string)


class Worksheet:
    """Represents a single worksheet in a workbook."""

    def __init__(self, title: str = "Sheet") -> None:
        self._title = self._validate_title(title)
        self._cells: dict[tuple[int, int], Cell] = {}
        self._merged_cells = MergedCellRange()
        self._freeze_panes: str | None = None
        self._auto_filter = AutoFilter()
        self._column_dimensions = DimensionHolder()
        self._row_dimensions = DimensionHolder()

    @staticmethod
    def _validate_title(title: str) -> str:
        if not title:
            raise SheetTitleException("Sheet title cannot be empty")
        if len(title) > _MAX_TITLE_LEN:
            raise SheetTitleException(
                f"Sheet title must be <= {_MAX_TITLE_LEN} characters"
            )
        if _INVALID_TITLE_RE.search(title):
            raise SheetTitleException(
                f"Invalid characters in sheet title: {title!r}"
            )
        return title

    @property
    def title(self) -> str:
        return self._title

    @title.setter
    def title(self, value: str) -> None:
        self._title = self._validate_title(value)

    @property
    def max_row(self) -> int:
        if not self._cells:
            return 0
        return max(r for r, _ in self._cells)

    @property
    def max_column(self) -> int:
        if not self._cells:
            return 0
        return max(c for _, c in self._cells)

    @property
    def min_row(self) -> int:
        if not self._cells:
            return 1
        return min(r for r, _ in self._cells)

    @property
    def min_column(self) -> int:
        if not self._cells:
            return 1
        return min(c for _, c in self._cells)

    @property
    def dimensions(self) -> str:
        if not self._cells:
            return "A1:A1"
        return (
            f"{get_column_letter(self.min_column)}{self.min_row}"
            f":{get_column_letter(self.max_column)}{self.max_row}"
        )

    @property
    def merged_cells(self) -> MergedCellRange:
        return self._merged_cells

    @property
    def freeze_panes(self) -> str | None:
        return self._freeze_panes

    @freeze_panes.setter
    def freeze_panes(self, value: str | None) -> None:
        self._freeze_panes = value

    @property
    def auto_filter(self) -> AutoFilter:
        return self._auto_filter

    @property
    def column_dimensions(self) -> DimensionHolder:
        return self._column_dimensions

    @property
    def row_dimensions(self) -> DimensionHolder:
        return self._row_dimensions

    def cell(
        self, row: int, column: int, value: CellValue = None
    ) -> Cell:
        """Get or create a cell at the given position."""
        key = (row, column)
        if key not in self._cells:
            self._cells[key] = Cell(row=row, column=column, value=value)
        elif value is not None:
            self._cells[key].value = value
        return self._cells[key]

    def _get_cell(self, row: int, column: int) -> Cell:
        """Get cell, creating it if needed."""
        return self.cell(row, column)

    def __getitem__(self, key: Any) -> Any:
        """Access cells by coordinate, range, row number, or column letter.

        ws['A1'] -> Cell
        ws['A1:C3'] -> tuple of tuples of Cells
        ws[1] -> tuple of Cells (row)
        ws['A'] -> tuple of Cells (column)
        """
        if isinstance(key, int):
            # Row access
            if self.max_column == 0:
                return ()
            return tuple(
                self._get_cell(key, col)
                for col in range(self.min_column, self.max_column + 1)
            )
        if isinstance(key, str):
            if ":" in key:
                # Range access
                return self._get_range(key)
            if key.isalpha():
                # Column access
                col_idx = column_index_from_string(key)
                if self.max_row == 0:
                    return ()
                return tuple(
                    self._get_cell(row, col_idx)
                    for row in range(self.min_row, self.max_row + 1)
                )
            # Single cell
            row, col = coordinate_to_tuple(key)
            return self._get_cell(row, col)
        raise CellCoordinatesException(f"Invalid cell reference: {key!r}")

    def __setitem__(self, key: str, value: CellValue) -> None:
        """Set a cell value by coordinate string."""
        row, col = coordinate_to_tuple(key)
        self.cell(row, col, value)

    def _get_range(
        self, range_string: str
    ) -> tuple[tuple[Cell, ...], ...]:
        """Return a tuple of tuples of cells for a range."""
        min_col, min_row, max_col, max_row = range_boundaries(range_string)
        return tuple(
            tuple(
                self._get_cell(row, col)
                for col in range(min_col, max_col + 1)
            )
            for row in range(min_row, max_row + 1)
        )

    def iter_rows(
        self,
        min_row: int | None = None,
        max_row: int | None = None,
        min_col: int | None = None,
        max_col: int | None = None,
        values_only: bool = False,
    ) -> Generator[tuple[Any, ...], None, None]:
        """Iterate over rows of cells or values."""
        mr = min_row if min_row is not None else self.min_row
        xr = max_row if max_row is not None else self.max_row
        mc = min_col if min_col is not None else self.min_column
        xc = max_col if max_col is not None else self.max_column
        if xr == 0 or xc == 0:
            return
        for row in range(mr, xr + 1):
            if values_only:
                yield tuple(
                    self._get_cell(row, col).value
                    for col in range(mc, xc + 1)
                )
            else:
                yield tuple(
                    self._get_cell(row, col)
                    for col in range(mc, xc + 1)
                )

    def iter_cols(
        self,
        min_row: int | None = None,
        max_row: int | None = None,
        min_col: int | None = None,
        max_col: int | None = None,
        values_only: bool = False,
    ) -> Generator[tuple[Any, ...], None, None]:
        """Iterate over columns of cells or values."""
        mr = min_row if min_row is not None else self.min_row
        xr = max_row if max_row is not None else self.max_row
        mc = min_col if min_col is not None else self.min_column
        xc = max_col if max_col is not None else self.max_column
        if xr == 0 or xc == 0:
            return
        for col in range(mc, xc + 1):
            if values_only:
                yield tuple(
                    self._get_cell(row, col).value
                    for row in range(mr, xr + 1)
                )
            else:
                yield tuple(
                    self._get_cell(row, col)
                    for row in range(mr, xr + 1)
                )

    def append(self, iterable: Any) -> None:
        """Append a row of values at the bottom of the sheet."""
        row_idx = self.max_row + 1
        for col_idx, value in enumerate(iterable, 1):
            self.cell(row_idx, col_idx, value)

    def merge_cells(self, range_string: str) -> None:
        """Merge cells in the given range."""
        self._merged_cells.add(range_string)

    def unmerge_cells(self, range_string: str) -> None:
        """Unmerge cells in the given range."""
        self._merged_cells.remove(range_string)

    def __repr__(self) -> str:
        return f"<Worksheet [{self._title}]>"
