"""sheetcraft — Modern Python library for reading and writing Excel .xlsx files."""

from __future__ import annotations

__version__ = "0.1.0"

# These will be importable once implemented in later phases
# from sheetcraft.workbook import Workbook, load_workbook
# from sheetcraft.worksheet import Worksheet
# from sheetcraft.cell import Cell

__all__ = [
    "Workbook",
    "load_workbook",
    "Worksheet",
    "Cell",
]


def __getattr__(name: str) -> object:
    """Lazy imports for classes implemented in later phases."""
    if name == "Workbook":
        from sheetcraft.workbook import Workbook  # type: ignore[import-not-found]

        return Workbook
    if name == "load_workbook":
        from sheetcraft.workbook import load_workbook

        return load_workbook
    if name == "Worksheet":
        from sheetcraft.worksheet import Worksheet  # type: ignore[import-not-found]

        return Worksheet
    if name == "Cell":
        from sheetcraft.cell import Cell

        return Cell
    raise AttributeError(f"module 'sheetcraft' has no attribute {name!r}")
