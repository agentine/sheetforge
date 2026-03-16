"""Custom exceptions for sheetforge."""

from __future__ import annotations


class SheetcraftException(Exception):
    """Base exception for all sheetforge errors."""


class InvalidFileException(SheetcraftException):
    """Raised when a file is not a valid .xlsx archive."""


class ReadOnlyWorkbookException(SheetcraftException):
    """Raised when attempting to modify a read-only workbook."""


class WorkbookAlreadySaved(SheetcraftException):
    """Raised when attempting to save a workbook that has already been saved and closed."""


class SheetTitleException(SheetcraftException):
    """Raised when a worksheet title is invalid."""


class CellCoordinatesException(SheetcraftException):
    """Raised when cell coordinates are invalid."""
