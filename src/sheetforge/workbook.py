"""Workbook class and load_workbook for sheetforge."""

from __future__ import annotations

import os
import zipfile
from typing import Any, Iterator

from sheetforge.cell import Cell
from sheetforge.constants import (
    PATH_SHARED_STRINGS,
    PATH_STYLES,
    PATH_WORKBOOK,
    PATH_WORKBOOK_RELS,
    TYPE_BOOL,
    TYPE_FORMULA,
    TYPE_STRING,
)
from sheetforge.exceptions import (
    InvalidFileException,
    SheetTitleException,
    WorkbookAlreadySaved,
)
from sheetforge.reader.shared_strings import read_shared_strings
from sheetforge.reader.styles import read_styles
from sheetforge.reader.workbook import read_workbook, read_workbook_rels
from sheetforge.reader.worksheet import read_worksheet
from sheetforge.styles import StyleSheet
from sheetforge.styles.numbers import BUILTIN_FORMATS, DATE_FORMAT_IDS
from sheetforge.utils import coordinate_to_tuple
from sheetforge.worksheet import Worksheet
from sheetforge.writer.workbook import write_workbook
from sheetforge.writer.worksheet import WriteCellData, WriteSheetData, write_worksheet

_NEXT_SHEET = 1


class Workbook:
    """Represents an Excel workbook (.xlsx file)."""

    def __init__(self, write_only: bool = False) -> None:
        self._sheets: list[Worksheet] = []
        self._active_index: int = 0
        self._stylesheet: StyleSheet = StyleSheet.default()
        self._write_only = write_only
        self._closed = False
        if not write_only:
            self._sheets.append(Worksheet(title="Sheet"))

    @property
    def active(self) -> Worksheet | None:
        if not self._sheets:
            return None
        return self._sheets[self._active_index]

    @active.setter
    def active(self, ws: Worksheet) -> None:
        try:
            self._active_index = self._sheets.index(ws)
        except ValueError as exc:
            raise ValueError("Worksheet is not part of this workbook") from exc

    @property
    def sheetnames(self) -> list[str]:
        return [ws.title for ws in self._sheets]

    def create_sheet(
        self, title: str | None = None, index: int | None = None
    ) -> Worksheet:
        """Create a new worksheet and add it to the workbook."""
        if title is None:
            title = self._unique_sheet_name()
        else:
            self._check_unique_name(title)
        ws = Worksheet(title=title)
        if index is None:
            self._sheets.append(ws)
        else:
            self._sheets.insert(index, ws)
        return ws

    def remove(self, worksheet: Worksheet) -> None:
        """Remove a worksheet from the workbook."""
        idx = self._sheets.index(worksheet)
        self._sheets.remove(worksheet)
        if self._active_index >= len(self._sheets):
            self._active_index = max(0, len(self._sheets) - 1)

    def remove_sheet(self, title: str) -> None:
        """Remove a worksheet by title."""
        ws = self[title]
        self.remove(ws)

    def __getitem__(self, name: str) -> Worksheet:
        for ws in self._sheets:
            if ws.title == name:
                return ws
        raise KeyError(f"Worksheet '{name}' not found")

    def __contains__(self, name: str) -> bool:
        return any(ws.title == name for ws in self._sheets)

    def __iter__(self) -> Iterator[Worksheet]:
        return iter(self._sheets)

    def __len__(self) -> int:
        return len(self._sheets)

    def save(self, filename: str | os.PathLike[str]) -> None:
        """Save the workbook to a .xlsx file."""
        if self._closed:
            raise WorkbookAlreadySaved("Workbook has already been saved and closed")
        xlsx_bytes = self._build_xlsx()
        with open(filename, "wb") as f:
            f.write(xlsx_bytes)

    def close(self) -> None:
        """Close the workbook."""
        self._closed = True

    def _build_xlsx(self) -> bytes:
        """Build the .xlsx bytes from current state."""
        shared_strings: list[str] = []
        ss_index: dict[str, int] = {}
        sheet_xmls: list[bytes] = []

        for ws in self._sheets:
            sheet_data = self._build_sheet_data(ws, shared_strings, ss_index)
            sheet_xmls.append(write_worksheet(sheet_data))

        return write_workbook(
            sheet_names=[ws.title for ws in self._sheets],
            sheet_xmls=sheet_xmls,
            stylesheet=self._stylesheet,
            shared_strings=shared_strings if shared_strings else None,
            active_sheet=self._active_index,
        )

    @staticmethod
    def _build_sheet_data(
        ws: Worksheet,
        shared_strings: list[str],
        ss_index: dict[str, int],
    ) -> WriteSheetData:
        """Convert a Worksheet into WriteSheetData."""
        data = WriteSheetData()

        for (row, col), cell in sorted(ws._cells.items()):
            val = cell.value
            if isinstance(val, str) and not (isinstance(val, str) and val.startswith("=")):
                # Use shared strings for non-formula strings
                if val not in ss_index:
                    ss_index[val] = len(shared_strings)
                    shared_strings.append(val)
                data.cells.append(
                    WriteCellData(
                        row=row,
                        column=col,
                        value=val,
                        shared_string_index=ss_index[val],
                        style_id=cell.style_id,
                    )
                )
            elif isinstance(val, str) and val.startswith("="):
                data.cells.append(
                    WriteCellData(
                        row=row,
                        column=col,
                        formula=val[1:],  # strip leading =
                        style_id=cell.style_id,
                    )
                )
            else:
                data.cells.append(
                    WriteCellData(
                        row=row,
                        column=col,
                        value=val,
                        style_id=cell.style_id,
                    )
                )

        data.merged_cells = list(ws.merged_cells)
        data.column_widths = ws.column_dimensions.as_dict()
        data.row_heights = ws.row_dimensions.as_dict()
        data.auto_filter_ref = ws.auto_filter.ref
        data.freeze_panes = ws.freeze_panes
        return data

    def _unique_sheet_name(self) -> str:
        """Generate a unique sheet name like 'Sheet', 'Sheet1', etc."""
        existing = set(self.sheetnames)
        if "Sheet" not in existing:
            return "Sheet"
        i = 1
        while f"Sheet{i}" in existing:
            i += 1
        return f"Sheet{i}"

    def _check_unique_name(self, name: str) -> None:
        if name in self.sheetnames:
            raise SheetTitleException(f"Sheet title '{name}' already exists")


def load_workbook(
    filename: str | os.PathLike[str],
    read_only: bool = False,
    data_only: bool = False,
    keep_vba: bool = False,
) -> Workbook:
    """Open an existing .xlsx file and return a Workbook.

    Args:
        filename: Path to the .xlsx file.
        read_only: If True, open in read-only mode (not yet implemented).
        data_only: If True, read cached formula values instead of formulas.
        keep_vba: If True, preserve VBA macros (not yet implemented).
    """
    path = str(filename)
    if not zipfile.is_zipfile(path):
        raise InvalidFileException(f"Not a valid xlsx file: {path}")

    wb = Workbook.__new__(Workbook)
    wb._sheets = []
    wb._active_index = 0
    wb._write_only = False
    wb._closed = False

    with zipfile.ZipFile(path, "r") as zf:
        # Read workbook.xml
        wb_data = read_workbook(zf.read(PATH_WORKBOOK))
        wb._active_index = wb_data.active_sheet

        # Read relationships
        rels = read_workbook_rels(zf.read(PATH_WORKBOOK_RELS))

        # Read shared strings
        shared_strings: list[str] = []
        if PATH_SHARED_STRINGS in zf.namelist():
            shared_strings = read_shared_strings(zf.read(PATH_SHARED_STRINGS))

        # Read styles
        stylesheet = StyleSheet.default()
        if PATH_STYLES in zf.namelist():
            stylesheet = read_styles(zf.read(PATH_STYLES))
        wb._stylesheet = stylesheet

        # Read each worksheet
        for sheet_info in wb_data.sheets:
            target = rels.get(sheet_info.r_id, "")
            if not target:
                continue
            sheet_path = f"xl/{target}"
            if sheet_path not in zf.namelist():
                continue

            ws_data = read_worksheet(zf.read(sheet_path))
            ws = Worksheet(title=sheet_info.name)

            # Populate cells
            for row_data in ws_data.rows:
                if row_data.height is not None:
                    ws.row_dimensions[row_data.row_num] = row_data.height
                for cell_data in row_data.cells:
                    row, col = coordinate_to_tuple(cell_data.ref)
                    cell = ws.cell(row, col)
                    cell.style_id = cell_data.style_id

                    if cell_data.formula and not data_only:
                        cell.value = f"={cell_data.formula}"
                    elif cell_data.data_type == "s" and cell_data.value is not None:
                        # Shared string
                        idx = int(cell_data.value)
                        if idx < len(shared_strings):
                            cell.value = shared_strings[idx]
                    elif cell_data.data_type == "b" and cell_data.value is not None:
                        cell.value = cell_data.value == "1"
                    elif cell_data.data_type == "inlineStr":
                        cell.value = cell_data.value
                    elif cell_data.value is not None:
                        # Numeric or date — check number format
                        try:
                            num_val = float(cell_data.value)
                            # Check if this is a date format
                            if cell_data.style_id < len(stylesheet.cell_xfs):
                                xf = stylesheet.cell_xfs[cell_data.style_id]
                                if xf.number_format_id in DATE_FORMAT_IDS:
                                    cell.value = num_val  # store as number; date conversion in future
                                else:
                                    cell.value = (
                                        int(num_val)
                                        if num_val == int(num_val)
                                        else num_val
                                    )
                            else:
                                cell.value = (
                                    int(num_val)
                                    if num_val == int(num_val)
                                    else num_val
                                )
                        except ValueError:
                            cell.value = cell_data.value

            # Merged cells
            for mc_range in ws_data.merged_cells:
                ws.merge_cells(mc_range)

            # Column dimensions
            for col_dim in ws_data.column_dimensions:
                if col_dim.width is not None:
                    for ci in range(col_dim.min_col, col_dim.max_col + 1):
                        ws.column_dimensions[ci] = col_dim.width

            # Auto-filter
            if ws_data.auto_filter_ref:
                ws.auto_filter.ref = ws_data.auto_filter_ref

            # Freeze panes
            ws.freeze_panes = ws_data.freeze_panes

            wb._sheets.append(ws)

    if not wb._sheets:
        wb._sheets.append(Worksheet(title="Sheet"))
    if wb._active_index >= len(wb._sheets):
        wb._active_index = 0

    return wb
