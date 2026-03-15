"""Tests for sheetcraft.writer modules."""

from __future__ import annotations

import xml.etree.ElementTree as ET
import zipfile
import io
from datetime import datetime, date

from sheetcraft.writer.shared_strings import write_shared_strings
from sheetcraft.writer.styles import write_styles
from sheetcraft.writer.worksheet import (
    WriteCellData,
    WriteSheetData,
    write_worksheet,
)
from sheetcraft.writer.workbook import write_workbook
from sheetcraft.styles import StyleSheet
from sheetcraft.styles.fonts import Font
from sheetcraft.styles.fills import PatternFill
from sheetcraft.constants import NS_SPREADSHEETML

_NS = f"{{{NS_SPREADSHEETML}}}"


class TestWriteSharedStrings:
    def test_basic(self) -> None:
        xml_bytes = write_shared_strings(["Hello", "World"])
        root = ET.fromstring(xml_bytes)
        si_elems = root.findall(f"{_NS}si")
        assert len(si_elems) == 2
        assert si_elems[0].find(f"{_NS}t") is not None
        t0 = si_elems[0].find(f"{_NS}t")
        assert t0 is not None
        assert t0.text == "Hello"

    def test_whitespace_preserve(self) -> None:
        xml_bytes = write_shared_strings(["  spaced  "])
        root = ET.fromstring(xml_bytes)
        t_el = root.find(f".//{_NS}t")
        assert t_el is not None
        assert t_el.get("{http://www.w3.org/XML/1998/namespace}space") == "preserve"

    def test_empty_list(self) -> None:
        xml_bytes = write_shared_strings([])
        root = ET.fromstring(xml_bytes)
        assert root.get("count") == "0"


class TestWriteStyles:
    def test_default_stylesheet(self) -> None:
        ss = StyleSheet.default()
        xml_bytes = write_styles(ss)
        root = ET.fromstring(xml_bytes)
        fonts = root.find(f"{_NS}fonts")
        assert fonts is not None
        assert fonts.get("count") == "1"
        fills = root.find(f"{_NS}fills")
        assert fills is not None
        assert fills.get("count") == "2"

    def test_custom_font(self) -> None:
        ss = StyleSheet.default()
        ss.fonts.append(Font(name="Arial", bold=True, size=14, color="FF0000"))
        xml_bytes = write_styles(ss)
        root = ET.fromstring(xml_bytes)
        fonts = root.find(f"{_NS}fonts")
        assert fonts is not None
        assert fonts.get("count") == "2"


class TestWriteWorksheet:
    def test_basic_cells(self) -> None:
        data = WriteSheetData(
            cells=[
                WriteCellData(row=1, column=1, value=42),
                WriteCellData(row=1, column=2, value="hello", shared_string_index=0),
                WriteCellData(row=2, column=1, value=True),
            ]
        )
        xml_bytes = write_worksheet(data)
        root = ET.fromstring(xml_bytes)
        rows = root.findall(f".//{_NS}row")
        assert len(rows) == 2

    def test_formula(self) -> None:
        data = WriteSheetData(
            cells=[WriteCellData(row=1, column=1, formula="SUM(A2:A10)")]
        )
        xml_bytes = write_worksheet(data)
        root = ET.fromstring(xml_bytes)
        f_el = root.find(f".//{_NS}f")
        assert f_el is not None
        assert f_el.text == "SUM(A2:A10)"

    def test_inline_string(self) -> None:
        data = WriteSheetData(
            cells=[WriteCellData(row=1, column=1, value="text")]
        )
        xml_bytes = write_worksheet(data)
        root = ET.fromstring(xml_bytes)
        c_el = root.find(f".//{_NS}c")
        assert c_el is not None
        assert c_el.get("t") == "inlineStr"

    def test_date_cell(self) -> None:
        data = WriteSheetData(
            cells=[WriteCellData(row=1, column=1, value=date(2024, 1, 15))]
        )
        xml_bytes = write_worksheet(data)
        root = ET.fromstring(xml_bytes)
        v_el = root.find(f".//{_NS}v")
        assert v_el is not None
        assert v_el.text is not None
        assert float(v_el.text) > 45000  # 2024 date serial

    def test_merged_cells(self) -> None:
        data = WriteSheetData(merged_cells=["A1:B2"])
        xml_bytes = write_worksheet(data)
        root = ET.fromstring(xml_bytes)
        mc = root.find(f".//{_NS}mergeCell")
        assert mc is not None
        assert mc.get("ref") == "A1:B2"

    def test_column_widths(self) -> None:
        data = WriteSheetData(column_widths={1: 20.0, 3: 15.5})
        xml_bytes = write_worksheet(data)
        root = ET.fromstring(xml_bytes)
        cols = root.findall(f".//{_NS}col")
        assert len(cols) == 2

    def test_freeze_panes(self) -> None:
        data = WriteSheetData(freeze_panes="A2")
        xml_bytes = write_worksheet(data)
        root = ET.fromstring(xml_bytes)
        pane = root.find(f".//{_NS}pane")
        assert pane is not None
        assert pane.get("state") == "frozen"
        assert pane.get("topLeftCell") == "A2"

    def test_auto_filter(self) -> None:
        data = WriteSheetData(auto_filter_ref="A1:C10")
        xml_bytes = write_worksheet(data)
        root = ET.fromstring(xml_bytes)
        af = root.find(f"{_NS}autoFilter")
        assert af is not None
        assert af.get("ref") == "A1:C10"


class TestWriteWorkbook:
    def test_creates_valid_zip(self) -> None:
        sheet_data = WriteSheetData(
            cells=[
                WriteCellData(row=1, column=1, value=42),
                WriteCellData(row=1, column=2, value="hello", shared_string_index=0),
            ]
        )
        sheet_xml = write_worksheet(sheet_data)
        xlsx_bytes = write_workbook(
            sheet_names=["Sheet1"],
            sheet_xmls=[sheet_xml],
            shared_strings=["hello"],
        )
        with zipfile.ZipFile(io.BytesIO(xlsx_bytes)) as zf:
            names = zf.namelist()
            assert "[Content_Types].xml" in names
            assert "_rels/.rels" in names
            assert "xl/workbook.xml" in names
            assert "xl/_rels/workbook.xml.rels" in names
            assert "xl/styles.xml" in names
            assert "xl/sharedStrings.xml" in names
            assert "xl/worksheets/sheet1.xml" in names

    def test_multiple_sheets(self) -> None:
        sheet1 = write_worksheet(WriteSheetData())
        sheet2 = write_worksheet(WriteSheetData())
        xlsx_bytes = write_workbook(
            sheet_names=["Data", "Summary"],
            sheet_xmls=[sheet1, sheet2],
        )
        with zipfile.ZipFile(io.BytesIO(xlsx_bytes)) as zf:
            assert "xl/worksheets/sheet1.xml" in zf.namelist()
            assert "xl/worksheets/sheet2.xml" in zf.namelist()

    def test_workbook_xml_structure(self) -> None:
        sheet_xml = write_worksheet(WriteSheetData())
        xlsx_bytes = write_workbook(
            sheet_names=["TestSheet"],
            sheet_xmls=[sheet_xml],
        )
        with zipfile.ZipFile(io.BytesIO(xlsx_bytes)) as zf:
            wb_xml = zf.read("xl/workbook.xml")
            root = ET.fromstring(wb_xml)
            sheets = root.findall(f".//{_NS}sheet")
            assert len(sheets) == 1
            assert sheets[0].get("name") == "TestSheet"

    def test_no_shared_strings_when_empty(self) -> None:
        sheet_xml = write_worksheet(WriteSheetData())
        xlsx_bytes = write_workbook(
            sheet_names=["Sheet1"],
            sheet_xmls=[sheet_xml],
        )
        with zipfile.ZipFile(io.BytesIO(xlsx_bytes)) as zf:
            assert "xl/sharedStrings.xml" not in zf.namelist()
