"""Utility functions for cell references and coordinate conversions."""

from __future__ import annotations

import re
from typing import Iterator

from sheetforge.exceptions import CellCoordinatesException

_COL_RE = re.compile(r"^[A-Z]{1,3}$")
_CELL_RE = re.compile(r"^(\$?)([A-Z]{1,3})(\$?)(\d+)$")
_RANGE_RE = re.compile(
    r"^(\$?[A-Z]{1,3}\$?\d+):(\$?[A-Z]{1,3}\$?\d+)$"
)


def column_index_from_string(col: str) -> int:
    """Convert a column letter to a 1-based index.

    Examples: 'A' -> 1, 'Z' -> 26, 'AA' -> 27, 'AZ' -> 52
    """
    col = col.upper().strip()
    if not _COL_RE.match(col):
        raise CellCoordinatesException(f"Invalid column letter: {col!r}")
    result = 0
    for ch in col:
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result


def get_column_letter(index: int) -> str:
    """Convert a 1-based column index to a column letter.

    Examples: 1 -> 'A', 26 -> 'Z', 27 -> 'AA', 52 -> 'AZ'
    """
    if index < 1:
        raise CellCoordinatesException(
            f"Column index must be >= 1, got {index}"
        )
    if index > 16384:
        raise CellCoordinatesException(
            f"Column index must be <= 16384, got {index}"
        )
    parts: list[str] = []
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        parts.append(chr(ord("A") + remainder))
    return "".join(reversed(parts))


def coordinate_to_tuple(coord: str) -> tuple[int, int]:
    """Convert a cell coordinate string to a (row, column) tuple (1-based).

    Examples: 'A1' -> (1, 1), 'C3' -> (3, 3), 'AA10' -> (10, 27)
    """
    m = _CELL_RE.match(coord.upper().strip())
    if not m:
        raise CellCoordinatesException(f"Invalid cell coordinate: {coord!r}")
    col_str = m.group(2)
    row = int(m.group(4))
    if row < 1:
        raise CellCoordinatesException(f"Row must be >= 1, got {row}")
    return (row, column_index_from_string(col_str))


def tuple_to_coordinate(row: int, column: int) -> str:
    """Convert a (row, column) tuple to a cell coordinate string.

    Examples: (1, 1) -> 'A1', (3, 3) -> 'C3'
    """
    return f"{get_column_letter(column)}{row}"


def absolute_coordinate(coord: str) -> str:
    """Convert a cell coordinate to absolute form.

    Examples: 'A1' -> '$A$1', '$A$1' -> '$A$1', 'A$1' -> '$A$1'
    """
    m = _CELL_RE.match(coord.upper().strip())
    if not m:
        raise CellCoordinatesException(f"Invalid cell coordinate: {coord!r}")
    col_str = m.group(2)
    row = m.group(4)
    return f"${col_str}${row}"


def range_boundaries(
    range_string: str,
) -> tuple[int, int, int, int]:
    """Parse a range string into (min_col, min_row, max_col, max_row).

    Example: 'A1:C3' -> (1, 1, 3, 3)
    """
    m = _RANGE_RE.match(range_string.upper().strip())
    if not m:
        raise CellCoordinatesException(
            f"Invalid range string: {range_string!r}"
        )
    min_row, min_col = coordinate_to_tuple(m.group(1))
    max_row, max_col = coordinate_to_tuple(m.group(2))
    return (min_col, min_row, max_col, max_row)


def rows_from_range(
    range_string: str,
) -> Iterator[tuple[str, ...]]:
    """Yield tuples of cell coordinates for each row in a range.

    Example: 'A1:C2' yields ('A1', 'B1', 'C1'), ('A2', 'B2', 'C2')
    """
    min_col, min_row, max_col, max_row = range_boundaries(range_string)
    for row in range(min_row, max_row + 1):
        yield tuple(
            tuple_to_coordinate(row, col)
            for col in range(min_col, max_col + 1)
        )


def cols_from_range(
    range_string: str,
) -> Iterator[tuple[str, ...]]:
    """Yield tuples of cell coordinates for each column in a range.

    Example: 'A1:C2' yields ('A1', 'A2'), ('B1', 'B2'), ('C1', 'C2')
    """
    min_col, min_row, max_col, max_row = range_boundaries(range_string)
    for col in range(min_col, max_col + 1):
        yield tuple(
            tuple_to_coordinate(row, col)
            for row in range(min_row, max_row + 1)
        )
