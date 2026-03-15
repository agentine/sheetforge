"""Roundtrip tests: create -> save -> reload -> verify."""

from __future__ import annotations

from sheetcraft import Workbook, load_workbook


class TestRoundtripBasic:
    def test_string_values(self, tmp_xlsx: str) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws["A1"] = "Hello"
        ws["B1"] = "World"
        ws["A2"] = ""
        wb.save(tmp_xlsx)

        wb2 = load_workbook(tmp_xlsx)
        ws2 = wb2.active
        assert ws2 is not None
        assert ws2["A1"].value == "Hello"
        assert ws2["B1"].value == "World"

    def test_numeric_values(self, tmp_xlsx: str) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws["A1"] = 42
        ws["B1"] = 3.14
        ws["C1"] = 0
        ws["D1"] = -100
        wb.save(tmp_xlsx)

        wb2 = load_workbook(tmp_xlsx)
        ws2 = wb2.active
        assert ws2 is not None
        assert ws2["A1"].value == 42
        assert ws2["B1"].value == 3.14
        assert ws2["C1"].value == 0
        assert ws2["D1"].value == -100

    def test_boolean_values(self, tmp_xlsx: str) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws["A1"] = True
        ws["B1"] = False
        wb.save(tmp_xlsx)

        wb2 = load_workbook(tmp_xlsx)
        ws2 = wb2.active
        assert ws2 is not None
        assert ws2["A1"].value is True
        assert ws2["B1"].value is False

    def test_formula_values(self, tmp_xlsx: str) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws["A1"] = 10
        ws["A2"] = 20
        ws["A3"] = "=SUM(A1:A2)"
        wb.save(tmp_xlsx)

        wb2 = load_workbook(tmp_xlsx)
        ws2 = wb2.active
        assert ws2 is not None
        assert ws2["A3"].value == "=SUM(A1:A2)"


class TestRoundtripMultiSheet:
    def test_multiple_sheets(self, tmp_xlsx: str) -> None:
        wb = Workbook()
        ws1 = wb.active
        assert ws1 is not None
        ws1["A1"] = "First"
        ws2 = wb.create_sheet("Second")
        ws2["A1"] = "Second"
        ws3 = wb.create_sheet("Third")
        ws3["A1"] = "Third"
        wb.save(tmp_xlsx)

        wb2 = load_workbook(tmp_xlsx)
        assert wb2.sheetnames == ["Sheet", "Second", "Third"]
        assert wb2["Sheet"]["A1"].value == "First"
        assert wb2["Second"]["A1"].value == "Second"
        assert wb2["Third"]["A1"].value == "Third"

    def test_active_sheet_preserved(self, tmp_xlsx: str) -> None:
        wb = Workbook()
        ws2 = wb.create_sheet("Data")
        wb.active = ws2
        ws2["A1"] = "active"
        wb.save(tmp_xlsx)

        wb2 = load_workbook(tmp_xlsx)
        assert wb2.active is not None
        assert wb2.active.title == "Data"


class TestRoundtripFeatures:
    def test_merged_cells(self, tmp_xlsx: str) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws["A1"] = "Merged"
        ws.merge_cells("A1:C1")
        wb.save(tmp_xlsx)

        wb2 = load_workbook(tmp_xlsx)
        ws2 = wb2.active
        assert ws2 is not None
        assert len(ws2.merged_cells) == 1
        assert "A1:C1" in ws2.merged_cells.ranges

    def test_freeze_panes(self, tmp_xlsx: str) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws["A1"] = "Header"
        ws.freeze_panes = "A2"
        wb.save(tmp_xlsx)

        wb2 = load_workbook(tmp_xlsx)
        ws2 = wb2.active
        assert ws2 is not None
        assert ws2.freeze_panes == "A2"

    def test_auto_filter(self, tmp_xlsx: str) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws["A1"] = "Name"
        ws["B1"] = "Age"
        ws.auto_filter.ref = "A1:B10"
        wb.save(tmp_xlsx)

        wb2 = load_workbook(tmp_xlsx)
        ws2 = wb2.active
        assert ws2 is not None
        assert ws2.auto_filter.ref == "A1:B10"

    def test_append_rows(self, tmp_xlsx: str) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        for i in range(10):
            ws.append([f"row{i}", i, i * 1.5])
        wb.save(tmp_xlsx)

        wb2 = load_workbook(tmp_xlsx)
        ws2 = wb2.active
        assert ws2 is not None
        rows = list(ws2.iter_rows(values_only=True))
        assert len(rows) == 10
        assert rows[0] == ("row0", 0, 0.0)
        assert rows[9][0] == "row9"

    def test_sample_workbook_roundtrip(
        self, sample_workbook: Workbook, tmp_xlsx: str
    ) -> None:
        sample_workbook.save(tmp_xlsx)
        wb2 = load_workbook(tmp_xlsx)
        ws2 = wb2.active
        assert ws2 is not None
        assert ws2["A1"].value == "Name"
        assert ws2["B2"].value == 30
        assert ws2["C2"].value == 95.5
        assert ws2["D2"].value is True
        assert ws2["C4"].value == "=AVERAGE(C2:C3)"

    def test_multi_sheet_roundtrip(
        self, multi_sheet_workbook: Workbook, tmp_xlsx: str
    ) -> None:
        multi_sheet_workbook.save(tmp_xlsx)
        wb2 = load_workbook(tmp_xlsx)
        assert wb2.sheetnames == ["Sheet", "Summary", "Config"]
        assert wb2["Summary"]["A2"].value == 100
