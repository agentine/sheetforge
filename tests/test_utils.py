"""Tests for sheetcraft.utils."""

from __future__ import annotations

import pytest

from sheetcraft.utils import (
    absolute_coordinate,
    column_index_from_string,
    cols_from_range,
    coordinate_to_tuple,
    get_column_letter,
    range_boundaries,
    rows_from_range,
    tuple_to_coordinate,
)
from sheetcraft.exceptions import CellCoordinatesException


class TestColumnIndexFromString:
    def test_single_letters(self) -> None:
        assert column_index_from_string("A") == 1
        assert column_index_from_string("B") == 2
        assert column_index_from_string("Z") == 26

    def test_double_letters(self) -> None:
        assert column_index_from_string("AA") == 27
        assert column_index_from_string("AZ") == 52
        assert column_index_from_string("BA") == 53

    def test_triple_letters(self) -> None:
        assert column_index_from_string("AAA") == 703

    def test_case_insensitive(self) -> None:
        assert column_index_from_string("a") == 1
        assert column_index_from_string("aa") == 27

    def test_invalid(self) -> None:
        with pytest.raises(CellCoordinatesException):
            column_index_from_string("")
        with pytest.raises(CellCoordinatesException):
            column_index_from_string("1")
        with pytest.raises(CellCoordinatesException):
            column_index_from_string("A1")


class TestGetColumnLetter:
    def test_single_letters(self) -> None:
        assert get_column_letter(1) == "A"
        assert get_column_letter(2) == "B"
        assert get_column_letter(26) == "Z"

    def test_double_letters(self) -> None:
        assert get_column_letter(27) == "AA"
        assert get_column_letter(52) == "AZ"
        assert get_column_letter(53) == "BA"

    def test_triple_letters(self) -> None:
        assert get_column_letter(703) == "AAA"

    def test_max_column(self) -> None:
        assert get_column_letter(16384) == "XFD"

    def test_invalid(self) -> None:
        with pytest.raises(CellCoordinatesException):
            get_column_letter(0)
        with pytest.raises(CellCoordinatesException):
            get_column_letter(-1)
        with pytest.raises(CellCoordinatesException):
            get_column_letter(16385)

    def test_roundtrip(self) -> None:
        for i in range(1, 100):
            assert column_index_from_string(get_column_letter(i)) == i


class TestCoordinateToTuple:
    def test_simple(self) -> None:
        assert coordinate_to_tuple("A1") == (1, 1)
        assert coordinate_to_tuple("C3") == (3, 3)
        assert coordinate_to_tuple("Z100") == (100, 26)

    def test_double_column(self) -> None:
        assert coordinate_to_tuple("AA10") == (10, 27)

    def test_absolute(self) -> None:
        assert coordinate_to_tuple("$A$1") == (1, 1)
        assert coordinate_to_tuple("$C$3") == (3, 3)

    def test_case_insensitive(self) -> None:
        assert coordinate_to_tuple("a1") == (1, 1)

    def test_invalid(self) -> None:
        with pytest.raises(CellCoordinatesException):
            coordinate_to_tuple("")
        with pytest.raises(CellCoordinatesException):
            coordinate_to_tuple("1A")
        with pytest.raises(CellCoordinatesException):
            coordinate_to_tuple("A")


class TestTupleToCoordinate:
    def test_simple(self) -> None:
        assert tuple_to_coordinate(1, 1) == "A1"
        assert tuple_to_coordinate(3, 3) == "C3"
        assert tuple_to_coordinate(10, 27) == "AA10"


class TestAbsoluteCoordinate:
    def test_relative(self) -> None:
        assert absolute_coordinate("A1") == "$A$1"
        assert absolute_coordinate("C3") == "$C$3"

    def test_already_absolute(self) -> None:
        assert absolute_coordinate("$A$1") == "$A$1"

    def test_mixed(self) -> None:
        assert absolute_coordinate("A$1") == "$A$1"
        assert absolute_coordinate("$A1") == "$A$1"


class TestRangeBoundaries:
    def test_simple(self) -> None:
        assert range_boundaries("A1:C3") == (1, 1, 3, 3)

    def test_larger(self) -> None:
        assert range_boundaries("B2:D10") == (2, 2, 4, 10)

    def test_invalid(self) -> None:
        with pytest.raises(CellCoordinatesException):
            range_boundaries("A1")
        with pytest.raises(CellCoordinatesException):
            range_boundaries("")


class TestRowsFromRange:
    def test_simple(self) -> None:
        rows = list(rows_from_range("A1:C2"))
        assert rows == [("A1", "B1", "C1"), ("A2", "B2", "C2")]

    def test_single_cell(self) -> None:
        rows = list(rows_from_range("A1:A1"))
        assert rows == [("A1",)]


class TestColsFromRange:
    def test_simple(self) -> None:
        cols = list(cols_from_range("A1:C2"))
        assert cols == [("A1", "A2"), ("B1", "B2"), ("C1", "C2")]

    def test_single_cell(self) -> None:
        cols = list(cols_from_range("A1:A1"))
        assert cols == [("A1",)]
