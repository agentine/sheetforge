"""Generate xl/worksheets/sheetN.xml for an xlsx archive."""

from __future__ import annotations

import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from datetime import date, datetime, time

from sheetcraft.constants import NS_SPREADSHEETML
from sheetcraft.utils import coordinate_to_tuple, get_column_letter


@dataclass
class WriteCellData:
    """Cell data for writing."""

    row: int
    column: int
    value: str | int | float | bool | datetime | date | time | None = None
    formula: str | None = None
    style_id: int = 0
    shared_string_index: int | None = None


@dataclass
class WriteSheetData:
    """All data needed to write a worksheet XML."""

    cells: list[WriteCellData] = field(default_factory=list)
    merged_cells: list[str] = field(default_factory=list)
    column_widths: dict[int, float] = field(default_factory=dict)
    row_heights: dict[int, float] = field(default_factory=dict)
    auto_filter_ref: str | None = None
    freeze_panes: str | None = None


def _excel_serial(val: datetime | date | time) -> float:
    """Convert a datetime/date/time to an Excel serial date number."""
    if isinstance(val, datetime):
        base_dt = datetime(1899, 12, 30)
        delta = val - base_dt
        return delta.days + delta.seconds / 86400.0
    if isinstance(val, date):
        base_d = date(1899, 12, 30)
        return float((val - base_d).days)
    # time
    return (val.hour * 3600 + val.minute * 60 + val.second) / 86400.0


def _sub(parent: ET.Element, tag: str, attrib: dict[str, str] | None = None) -> ET.Element:
    """Create a SubElement with an explicit attrib dict to satisfy mypy."""
    if attrib:
        return ET.SubElement(parent, tag, attrib)
    return ET.SubElement(parent, tag)


def write_worksheet(data: WriteSheetData) -> bytes:
    """Generate a worksheet XML from WriteSheetData."""
    root = ET.Element("worksheet", xmlns=NS_SPREADSHEETML)

    # Freeze panes
    if data.freeze_panes:
        views_el = _sub(root, "sheetViews")
        view_el = _sub(views_el, "sheetView", {"workbookViewId": "0"})
        freeze_row, freeze_col = coordinate_to_tuple(data.freeze_panes)
        pane_attrs: dict[str, str] = {
            "state": "frozen",
            "topLeftCell": data.freeze_panes,
        }
        if freeze_col > 1:
            pane_attrs["xSplit"] = str(freeze_col - 1)
        if freeze_row > 1:
            pane_attrs["ySplit"] = str(freeze_row - 1)
        _sub(view_el, "pane", pane_attrs)

    # Column widths
    if data.column_widths:
        cols_el = _sub(root, "cols")
        for col_idx in sorted(data.column_widths):
            _sub(
                cols_el,
                "col",
                {
                    "min": str(col_idx),
                    "max": str(col_idx),
                    "width": str(data.column_widths[col_idx]),
                    "customWidth": "1",
                },
            )

    # Group cells by row
    rows: dict[int, list[WriteCellData]] = {}
    for cell in data.cells:
        rows.setdefault(cell.row, []).append(cell)

    sheet_data = _sub(root, "sheetData")
    for row_num in sorted(rows):
        row_attrs: dict[str, str] = {"r": str(row_num)}
        if row_num in data.row_heights:
            row_attrs["ht"] = str(data.row_heights[row_num])
            row_attrs["customHeight"] = "1"
        row_el = _sub(sheet_data, "row", row_attrs)

        for cell in sorted(rows[row_num], key=lambda c: c.column):
            ref = f"{get_column_letter(cell.column)}{cell.row}"
            c_attrs: dict[str, str] = {"r": ref}
            if cell.style_id:
                c_attrs["s"] = str(cell.style_id)

            if cell.formula is not None:
                c_attrs["t"] = "str"
                c_el = _sub(row_el, "c", c_attrs)
                f_el = _sub(c_el, "f")
                f_el.text = cell.formula
            elif cell.shared_string_index is not None:
                c_attrs["t"] = "s"
                c_el = _sub(row_el, "c", c_attrs)
                v_el = _sub(c_el, "v")
                v_el.text = str(cell.shared_string_index)
            elif isinstance(cell.value, bool):
                c_attrs["t"] = "b"
                c_el = _sub(row_el, "c", c_attrs)
                v_el = _sub(c_el, "v")
                v_el.text = "1" if cell.value else "0"
            elif isinstance(cell.value, (int, float)):
                c_el = _sub(row_el, "c", c_attrs)
                v_el = _sub(c_el, "v")
                v_el.text = str(cell.value)
            elif isinstance(cell.value, (datetime, date, time)):
                c_el = _sub(row_el, "c", c_attrs)
                v_el = _sub(c_el, "v")
                v_el.text = str(_excel_serial(cell.value))
            elif isinstance(cell.value, str):
                c_attrs["t"] = "inlineStr"
                c_el = _sub(row_el, "c", c_attrs)
                is_el = _sub(c_el, "is")
                t_el = _sub(is_el, "t")
                t_el.text = cell.value
            else:
                # None / empty
                _sub(row_el, "c", c_attrs)

    # Merged cells
    if data.merged_cells:
        mc_el = _sub(root, "mergeCells", {"count": str(len(data.merged_cells))})
        for mc_ref in data.merged_cells:
            _sub(mc_el, "mergeCell", {"ref": mc_ref})

    # Auto-filter
    if data.auto_filter_ref:
        _sub(root, "autoFilter", {"ref": data.auto_filter_ref})

    result: bytes = ET.tostring(root, encoding="UTF-8", xml_declaration=True)
    return result
