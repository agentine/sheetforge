"""Style types for sheetforge."""

from __future__ import annotations

from dataclasses import dataclass

from sheetforge.styles.alignment import Alignment, DEFAULT_ALIGNMENT
from sheetforge.styles.borders import Border, DEFAULT_BORDER, Side
from sheetforge.styles.fills import FILL_NONE, FILL_SOLID, GradientFill, PatternFill
from sheetforge.styles.fonts import DEFAULT_FONT, Font
from sheetforge.styles.numbers import (
    BUILTIN_FORMATS,
    FORMAT_GENERAL,
    NumberFormat,
)

__all__ = [
    "Alignment",
    "Border",
    "CellStyle",
    "DEFAULT_ALIGNMENT",
    "DEFAULT_BORDER",
    "DEFAULT_FONT",
    "FILL_NONE",
    "FILL_SOLID",
    "Font",
    "GradientFill",
    "NumberFormat",
    "PatternFill",
    "Side",
    "StyleSheet",
]


@dataclass
class CellStyle:
    """References into StyleSheet arrays for a single cell's formatting."""

    font_id: int = 0
    fill_id: int = 0
    border_id: int = 0
    number_format_id: int = 0
    alignment: Alignment | None = None
    apply_font: bool = False
    apply_fill: bool = False
    apply_border: bool = False
    apply_number_format: bool = False
    apply_alignment: bool = False


@dataclass
class StyleSheet:
    """Aggregated style data from an xlsx file."""

    fonts: list[Font]
    fills: list[PatternFill]
    borders: list[Border]
    number_formats: list[NumberFormat]
    cell_xfs: list[CellStyle]

    @staticmethod
    def default() -> StyleSheet:
        """Return the minimal default stylesheet for a new workbook."""
        return StyleSheet(
            fonts=[DEFAULT_FONT],
            fills=[FILL_NONE, FILL_SOLID],
            borders=[DEFAULT_BORDER],
            number_formats=[],
            cell_xfs=[CellStyle()],
        )

    def get_format_code(self, num_fmt_id: int) -> str:
        """Look up a number format code by ID."""
        for nf in self.number_formats:
            if nf.format_id == num_fmt_id:
                return nf.format_code
        return BUILTIN_FORMATS.get(num_fmt_id, FORMAT_GENERAL)
