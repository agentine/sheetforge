"""Compatibility tests: verify sheetcraft produces structurally valid .xlsx files."""

from __future__ import annotations

import io
import xml.etree.ElementTree as ET
import zipfile

from sheetcraft import Workbook
from sheetcraft.constants import NS_SPREADSHEETML

_NS = f"{{{NS_SPREADSHEETML}}}"


class TestXlsxStructure:
    def _save_to_bytes(self, wb: Workbook) -> bytes:
        """Save a workbook and return the raw bytes."""
        buf = io.BytesIO()
        # Build xlsx bytes directly
        xlsx = wb._build_xlsx()
        return xlsx

    def test_required_files_present(self) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws["A1"] = "test"
        data = self._save_to_bytes(wb)
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            names = zf.namelist()
            assert "[Content_Types].xml" in names
            assert "_rels/.rels" in names
            assert "xl/workbook.xml" in names
            assert "xl/_rels/workbook.xml.rels" in names
            assert "xl/styles.xml" in names
            assert "xl/worksheets/sheet1.xml" in names

    def test_content_types_valid(self) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws["A1"] = "test"
        data = self._save_to_bytes(wb)
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            ct = zf.read("[Content_Types].xml")
            root = ET.fromstring(ct)
            # Must have Default and Override elements
            defaults = root.findall("{http://schemas.openxmlformats.org/package/2006/content-types}Default")
            overrides = root.findall("{http://schemas.openxmlformats.org/package/2006/content-types}Override")
            assert len(defaults) >= 2
            assert len(overrides) >= 2

    def test_workbook_xml_has_sheets(self) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws["A1"] = "test"
        data = self._save_to_bytes(wb)
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            wb_xml = zf.read("xl/workbook.xml")
            root = ET.fromstring(wb_xml)
            sheets = root.findall(f".//{_NS}sheet")
            assert len(sheets) == 1
            assert sheets[0].get("name") == "Sheet"

    def test_worksheet_has_sheetdata(self) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws["A1"] = "test"
        ws["B1"] = 42
        data = self._save_to_bytes(wb)
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            ws_xml = zf.read("xl/worksheets/sheet1.xml")
            root = ET.fromstring(ws_xml)
            sd = root.find(f"{_NS}sheetData")
            assert sd is not None
            rows = sd.findall(f"{_NS}row")
            assert len(rows) >= 1

    def test_shared_strings_present_for_text(self) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws["A1"] = "Hello"
        ws["B1"] = "World"
        data = self._save_to_bytes(wb)
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            assert "xl/sharedStrings.xml" in zf.namelist()
            ss_xml = zf.read("xl/sharedStrings.xml")
            root = ET.fromstring(ss_xml)
            si_elems = root.findall(f"{_NS}si")
            assert len(si_elems) == 2

    def test_no_shared_strings_for_numeric_only(self) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws["A1"] = 42
        ws["B1"] = 3.14
        data = self._save_to_bytes(wb)
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            assert "xl/sharedStrings.xml" not in zf.namelist()

    def test_styles_xml_valid(self) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws["A1"] = "test"
        data = self._save_to_bytes(wb)
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            styles = zf.read("xl/styles.xml")
            root = ET.fromstring(styles)
            assert root.find(f"{_NS}fonts") is not None
            assert root.find(f"{_NS}fills") is not None
            assert root.find(f"{_NS}borders") is not None
            assert root.find(f"{_NS}cellXfs") is not None

    def test_relationships_valid(self) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws["A1"] = "test"
        data = self._save_to_bytes(wb)
        ns_rel = "{http://schemas.openxmlformats.org/package/2006/relationships}"
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            # Root rels
            root_rels = ET.fromstring(zf.read("_rels/.rels"))
            rels = root_rels.findall(f"{ns_rel}Relationship")
            assert len(rels) >= 1

            # Workbook rels
            wb_rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
            rels = wb_rels.findall(f"{ns_rel}Relationship")
            assert len(rels) >= 2  # at least worksheet + styles

    def test_multi_sheet_structure(self) -> None:
        wb = Workbook()
        wb.create_sheet("Data")
        wb.create_sheet("Config")
        data = self._save_to_bytes(wb)
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            names = zf.namelist()
            assert "xl/worksheets/sheet1.xml" in names
            assert "xl/worksheets/sheet2.xml" in names
            assert "xl/worksheets/sheet3.xml" in names

    def test_merged_cells_in_xml(self) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws["A1"] = "Merged"
        ws.merge_cells("A1:C1")
        data = self._save_to_bytes(wb)
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            ws_xml = zf.read("xl/worksheets/sheet1.xml")
            root = ET.fromstring(ws_xml)
            mc = root.find(f"{_NS}mergeCells")
            assert mc is not None
            assert mc.get("count") == "1"

    def test_zip_is_valid(self) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws["A1"] = "test"
        data = self._save_to_bytes(wb)
        assert zipfile.is_zipfile(io.BytesIO(data))
