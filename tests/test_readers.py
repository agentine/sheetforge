"""Tests for sheetforge.reader modules."""

from __future__ import annotations

from sheetforge.reader.shared_strings import read_shared_strings
from sheetforge.reader.styles import read_styles
from sheetforge.reader.worksheet import read_worksheet
from sheetforge.reader.workbook import read_workbook, read_workbook_rels

_SST_XML = b"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
  <si><t>Hello</t></si>
  <si><t>World</t></si>
  <si>
    <r><rPr><b/></rPr><t>Bold</t></r>
    <r><t> text</t></r>
  </si>
</sst>
"""

_STYLES_XML = b"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>
  </numFmts>
  <fonts count="2">
    <font>
      <sz val="11"/>
      <name val="Calibri"/>
      <family val="2"/>
    </font>
    <font>
      <b/>
      <sz val="12"/>
      <name val="Arial"/>
      <color rgb="FF0000FF"/>
    </font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill>
  </fills>
  <borders count="1">
    <border>
      <left style="thin"><color rgb="FF000000"/></left>
      <right/>
      <top/>
      <bottom/>
    </border>
  </borders>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="164" fontId="1" fillId="1" borderId="0" applyFont="1" applyFill="1" applyNumberFormat="1">
      <alignment horizontal="center" wrapText="1"/>
    </xf>
  </cellXfs>
</styleSheet>
"""

_WORKSHEET_XML = b"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0">
      <pane xSplit="0" ySplit="1" topLeftCell="A2" state="frozen"/>
    </sheetView>
  </sheetViews>
  <cols>
    <col min="1" max="1" width="15" customWidth="1"/>
    <col min="2" max="3" width="10" hidden="1"/>
  </cols>
  <sheetData>
    <row r="1" ht="20" customHeight="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1" t="s"><v>1</v></c>
      <c r="C1" s="1"><v>42</v></c>
    </row>
    <row r="2">
      <c r="A2" t="b"><v>1</v></c>
      <c r="B2"><f>SUM(C1:C1)</f><v>42</v></c>
      <c r="C2" t="inlineStr"><is><t>inline</t></is></c>
    </row>
  </sheetData>
  <mergeCells count="1">
    <mergeCell ref="A3:B3"/>
  </mergeCells>
  <autoFilter ref="A1:C2"/>
</worksheet>
"""

_WORKBOOK_XML = b"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <bookViews>
    <workbookView activeTab="1"/>
  </bookViews>
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Sheet2" sheetId="2" r:id="rId2" state="hidden"/>
  </sheets>
  <definedNames>
    <definedName name="MyRange" localSheetId="0">Sheet1!$A$1:$C$10</definedName>
  </definedNames>
</workbook>
"""

_RELS_XML = b"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Target="worksheets/sheet1.xml"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"/>
  <Relationship Id="rId2" Target="worksheets/sheet2.xml"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"/>
  <Relationship Id="rId3" Target="styles.xml"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"/>
</Relationships>
"""


class TestReadSharedStrings:
    def test_plain_text(self) -> None:
        strings = read_shared_strings(_SST_XML)
        assert strings[0] == "Hello"
        assert strings[1] == "World"

    def test_rich_text(self) -> None:
        strings = read_shared_strings(_SST_XML)
        assert strings[2] == "Bold text"

    def test_count(self) -> None:
        strings = read_shared_strings(_SST_XML)
        assert len(strings) == 3


class TestReadStyles:
    def test_fonts(self) -> None:
        ss = read_styles(_STYLES_XML)
        assert len(ss.fonts) == 2
        assert ss.fonts[0].name == "Calibri"
        assert ss.fonts[0].size == 11.0
        assert ss.fonts[0].bold is False
        assert ss.fonts[1].name == "Arial"
        assert ss.fonts[1].size == 12.0
        assert ss.fonts[1].bold is True
        assert ss.fonts[1].color == "FF0000FF"

    def test_fills(self) -> None:
        ss = read_styles(_STYLES_XML)
        assert len(ss.fills) == 2
        assert ss.fills[0].patternType == "none"
        assert ss.fills[1].patternType == "solid"
        assert ss.fills[1].fgColor == "FFFFFF00"

    def test_borders(self) -> None:
        ss = read_styles(_STYLES_XML)
        assert len(ss.borders) == 1
        assert ss.borders[0].left is not None
        assert ss.borders[0].left.style == "thin"
        assert ss.borders[0].left.color == "FF000000"

    def test_number_formats(self) -> None:
        ss = read_styles(_STYLES_XML)
        assert len(ss.number_formats) == 1
        assert ss.number_formats[0].format_id == 164
        assert ss.number_formats[0].format_code == "yyyy-mm-dd"

    def test_cell_xfs(self) -> None:
        ss = read_styles(_STYLES_XML)
        assert len(ss.cell_xfs) == 2
        xf0 = ss.cell_xfs[0]
        assert xf0.font_id == 0
        assert xf0.alignment is None
        xf1 = ss.cell_xfs[1]
        assert xf1.font_id == 1
        assert xf1.fill_id == 1
        assert xf1.number_format_id == 164
        assert xf1.apply_font is True
        assert xf1.alignment is not None
        assert xf1.alignment.horizontal == "center"
        assert xf1.alignment.wrap_text is True


class TestReadWorksheet:
    def test_rows(self) -> None:
        ws = read_worksheet(_WORKSHEET_XML)
        assert len(ws.rows) == 2
        assert ws.rows[0].row_num == 1
        assert ws.rows[0].height == 20.0
        assert ws.rows[0].custom_height is True

    def test_cells(self) -> None:
        ws = read_worksheet(_WORKSHEET_XML)
        row1 = ws.rows[0]
        assert len(row1.cells) == 3
        assert row1.cells[0].ref == "A1"
        assert row1.cells[0].data_type == "s"
        assert row1.cells[0].value == "0"
        assert row1.cells[2].style_id == 1
        assert row1.cells[2].value == "42"

    def test_formula(self) -> None:
        ws = read_worksheet(_WORKSHEET_XML)
        row2 = ws.rows[1]
        assert row2.cells[1].formula == "SUM(C1:C1)"
        assert row2.cells[1].value == "42"

    def test_inline_string(self) -> None:
        ws = read_worksheet(_WORKSHEET_XML)
        row2 = ws.rows[1]
        assert row2.cells[2].data_type == "inlineStr"
        assert row2.cells[2].value == "inline"

    def test_bool_cell(self) -> None:
        ws = read_worksheet(_WORKSHEET_XML)
        row2 = ws.rows[1]
        assert row2.cells[0].data_type == "b"
        assert row2.cells[0].value == "1"

    def test_merged_cells(self) -> None:
        ws = read_worksheet(_WORKSHEET_XML)
        assert ws.merged_cells == ["A3:B3"]

    def test_column_dimensions(self) -> None:
        ws = read_worksheet(_WORKSHEET_XML)
        assert len(ws.column_dimensions) == 2
        assert ws.column_dimensions[0].min_col == 1
        assert ws.column_dimensions[0].width == 15.0
        assert ws.column_dimensions[0].custom_width is True
        assert ws.column_dimensions[1].hidden is True

    def test_auto_filter(self) -> None:
        ws = read_worksheet(_WORKSHEET_XML)
        assert ws.auto_filter_ref == "A1:C2"

    def test_freeze_panes(self) -> None:
        ws = read_worksheet(_WORKSHEET_XML)
        assert ws.freeze_panes == "A2"


class TestReadWorkbook:
    def test_sheets(self) -> None:
        wb = read_workbook(_WORKBOOK_XML)
        assert len(wb.sheets) == 2
        assert wb.sheets[0].name == "Sheet1"
        assert wb.sheets[0].r_id == "rId1"
        assert wb.sheets[1].name == "Sheet2"
        assert wb.sheets[1].state == "hidden"

    def test_active_sheet(self) -> None:
        wb = read_workbook(_WORKBOOK_XML)
        assert wb.active_sheet == 1

    def test_defined_names(self) -> None:
        wb = read_workbook(_WORKBOOK_XML)
        assert len(wb.defined_names) == 1
        assert wb.defined_names[0].name == "MyRange"
        assert wb.defined_names[0].value == "Sheet1!$A$1:$C$10"
        assert wb.defined_names[0].sheet_id == 0


class TestReadWorkbookRels:
    def test_rels(self) -> None:
        rels = read_workbook_rels(_RELS_XML)
        assert rels["rId1"] == "worksheets/sheet1.xml"
        assert rels["rId2"] == "worksheets/sheet2.xml"
        assert rels["rId3"] == "styles.xml"
