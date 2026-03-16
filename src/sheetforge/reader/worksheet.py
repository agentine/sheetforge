"""Parse xl/worksheets/sheetN.xml from an xlsx archive."""

from __future__ import annotations

import xml.etree.ElementTree as ET
from dataclasses import dataclass, field

from sheetforge.constants import NS_SPREADSHEETML

_NS = f"{{{NS_SPREADSHEETML}}}"


@dataclass
class CellData:
    """Raw cell data parsed from worksheet XML."""

    ref: str
    value: str | None = None
    data_type: str | None = None
    style_id: int = 0
    formula: str | None = None


@dataclass
class RowData:
    """Raw row data parsed from worksheet XML."""

    row_num: int
    cells: list[CellData] = field(default_factory=list)
    height: float | None = None
    custom_height: bool = False
    hidden: bool = False


@dataclass
class ColumnDimension:
    """Column width metadata."""

    min_col: int
    max_col: int
    width: float | None = None
    custom_width: bool = False
    hidden: bool = False
    best_fit: bool = False


@dataclass
class WorksheetData:
    """All data parsed from a worksheet XML file."""

    rows: list[RowData] = field(default_factory=list)
    merged_cells: list[str] = field(default_factory=list)
    column_dimensions: list[ColumnDimension] = field(default_factory=list)
    auto_filter_ref: str | None = None
    freeze_panes: str | None = None


def _bool(val: str | None) -> bool:
    if val is None:
        return False
    return val in ("1", "true")


def read_worksheet(xml_bytes: bytes) -> WorksheetData:
    """Parse a worksheet XML and return structured data."""
    root = ET.fromstring(xml_bytes)
    data = WorksheetData()

    # Parse sheet data (rows and cells)
    sheet_data = root.find(f"{_NS}sheetData")
    if sheet_data is not None:
        for row_el in sheet_data.findall(f"{_NS}row"):
            row_num = int(row_el.get("r", "0"))
            height = row_el.get("ht")
            row = RowData(
                row_num=row_num,
                height=float(height) if height else None,
                custom_height=_bool(row_el.get("customHeight")),
                hidden=_bool(row_el.get("hidden")),
            )
            for cell_el in row_el.findall(f"{_NS}c"):
                ref = cell_el.get("r", "")
                cell_type = cell_el.get("t")
                style_id = int(cell_el.get("s", "0"))

                value: str | None = None
                formula: str | None = None

                v_el = cell_el.find(f"{_NS}v")
                if v_el is not None:
                    value = v_el.text

                f_el = cell_el.find(f"{_NS}f")
                if f_el is not None:
                    formula = f_el.text

                # Inline string
                is_el = cell_el.find(f"{_NS}is")
                if is_el is not None:
                    t_el = is_el.find(f"{_NS}t")
                    if t_el is not None:
                        value = t_el.text
                    cell_type = "inlineStr"

                row.cells.append(
                    CellData(
                        ref=ref,
                        value=value,
                        data_type=cell_type,
                        style_id=style_id,
                        formula=formula,
                    )
                )
            data.rows.append(row)

    # Parse merged cells
    merge_cells_el = root.find(f"{_NS}mergeCells")
    if merge_cells_el is not None:
        for mc in merge_cells_el.findall(f"{_NS}mergeCell"):
            mc_ref = mc.get("ref")
            if mc_ref:
                data.merged_cells.append(mc_ref)

    # Parse column dimensions
    cols_el = root.find(f"{_NS}cols")
    if cols_el is not None:
        for col_el in cols_el.findall(f"{_NS}col"):
            data.column_dimensions.append(
                ColumnDimension(
                    min_col=int(col_el.get("min", "1")),
                    max_col=int(col_el.get("max", "1")),
                    width=float(col_el.get("width", "0")) if col_el.get("width") else None,
                    custom_width=_bool(col_el.get("customWidth")),
                    hidden=_bool(col_el.get("hidden")),
                    best_fit=_bool(col_el.get("bestFit")),
                )
            )

    # Parse auto-filter
    af_el = root.find(f"{_NS}autoFilter")
    if af_el is not None:
        data.auto_filter_ref = af_el.get("ref")

    # Parse freeze panes (sheetViews/sheetView/pane)
    views_el = root.find(f"{_NS}sheetViews")
    if views_el is not None:
        view_el = views_el.find(f"{_NS}sheetView")
        if view_el is not None:
            pane_el = view_el.find(f"{_NS}pane")
            if pane_el is not None:
                state = pane_el.get("state", "")
                if state == "frozen":
                    top_left = pane_el.get("topLeftCell")
                    if top_left:
                        data.freeze_panes = top_left

    return data
