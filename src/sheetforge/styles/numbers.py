"""Number format definitions."""

from __future__ import annotations

from dataclasses import dataclass

# Standard format codes (ECMA-376, 18.8.30)
FORMAT_GENERAL = "General"
FORMAT_TEXT = "@"
FORMAT_NUMBER = "0"
FORMAT_NUMBER_00 = "0.00"
FORMAT_NUMBER_COMMA_SEPARATED = "#,##0"
FORMAT_NUMBER_COMMA_SEPARATED_00 = "#,##0.00"
FORMAT_PERCENTAGE = "0%"
FORMAT_PERCENTAGE_00 = "0.00%"
FORMAT_DATE_YYYYMMDD2 = "yyyy-mm-dd"
FORMAT_DATE_DATETIME = "yyyy-mm-dd h:mm:ss"
FORMAT_DATE_DMYSLASH = "d/m/yyyy"
FORMAT_DATE_MDYSLASH = "m/d/yyyy"
FORMAT_DATE_DDMMYYYY = "dd/mm/yyyy"
FORMAT_DATE_TIME1 = "h:mm:ss"
FORMAT_DATE_TIME2 = "h:mm"
FORMAT_DATE_TIME3 = "h:mm AM/PM"
FORMAT_CURRENCY_USD_SIMPLE = '"$"#,##0.00'


@dataclass
class NumberFormat:
    """Represents a custom number format."""

    format_code: str = FORMAT_GENERAL
    format_id: int | None = None


# Built-in format codes mapped by their numeric ID
BUILTIN_FORMATS: dict[int, str] = {
    0: FORMAT_GENERAL,
    1: "0",
    2: "0.00",
    3: "#,##0",
    4: "#,##0.00",
    9: "0%",
    10: "0.00%",
    11: "0.00E+00",
    12: "# ?/?",
    13: "# ??/??",
    14: "mm-dd-yy",
    15: "d-mmm-yy",
    16: "d-mmm",
    17: "mmm-yy",
    18: "h:mm AM/PM",
    19: "h:mm:ss AM/PM",
    20: "h:mm",
    21: "h:mm:ss",
    22: "m/d/yy h:mm",
    37: "#,##0 ;(#,##0)",
    38: "#,##0 ;[Red](#,##0)",
    39: "#,##0.00;(#,##0.00)",
    40: "#,##0.00;[Red](#,##0.00)",
    44: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)',
    45: "mm:ss",
    46: "[h]:mm:ss",
    47: "mmss.0",
    48: "##0.0E+0",
    49: "@",
}

# Date format IDs (used to detect date cells)
DATE_FORMAT_IDS: frozenset[int] = frozenset(range(14, 23)) | frozenset(
    range(45, 48)
)
