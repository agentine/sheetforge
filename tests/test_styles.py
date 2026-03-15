"""Tests for sheetcraft.styles."""

from __future__ import annotations

from sheetcraft.styles import (
    Alignment,
    Border,
    CellStyle,
    Font,
    GradientFill,
    NumberFormat,
    PatternFill,
    Side,
    StyleSheet,
)
from sheetcraft.styles.numbers import BUILTIN_FORMATS, FORMAT_GENERAL


class TestFont:
    def test_defaults(self) -> None:
        f = Font()
        assert f.name == "Calibri"
        assert f.size == 11.0
        assert f.bold is False
        assert f.italic is False

    def test_custom(self) -> None:
        f = Font(name="Arial", size=14, bold=True, color="FF0000")
        assert f.name == "Arial"
        assert f.size == 14
        assert f.bold is True
        assert f.color == "FF0000"

    def test_copy(self) -> None:
        f = Font(bold=True)
        f2 = f.copy(italic=True)
        assert f2.bold is True
        assert f2.italic is True
        assert f.italic is False


class TestPatternFill:
    def test_defaults(self) -> None:
        pf = PatternFill()
        assert pf.patternType is None
        assert pf.fgColor is None

    def test_solid(self) -> None:
        pf = PatternFill(patternType="solid", fgColor="00FF00")
        assert pf.patternType == "solid"
        assert pf.fgColor == "00FF00"

    def test_copy(self) -> None:
        pf = PatternFill(patternType="solid")
        pf2 = pf.copy(fgColor="FF0000")
        assert pf2.patternType == "solid"
        assert pf2.fgColor == "FF0000"
        assert pf.fgColor is None


class TestGradientFill:
    def test_defaults(self) -> None:
        gf = GradientFill()
        assert gf.type == "linear"
        assert gf.degree == 0.0


class TestSide:
    def test_defaults(self) -> None:
        s = Side()
        assert s.style is None
        assert s.color is None

    def test_thin(self) -> None:
        s = Side(style="thin", color="000000")
        assert s.style == "thin"
        assert s.color == "000000"


class TestBorder:
    def test_defaults(self) -> None:
        b = Border()
        assert b.left is None
        assert b.outline is True

    def test_with_sides(self) -> None:
        b = Border(
            left=Side(style="thin"),
            right=Side(style="thick"),
        )
        assert b.left is not None
        assert b.left.style == "thin"
        assert b.right is not None
        assert b.right.style == "thick"


class TestAlignment:
    def test_defaults(self) -> None:
        a = Alignment()
        assert a.horizontal is None
        assert a.wrap_text is False
        assert a.indent == 0

    def test_custom(self) -> None:
        a = Alignment(horizontal="center", vertical="top", wrap_text=True)
        assert a.horizontal == "center"
        assert a.vertical == "top"
        assert a.wrap_text is True


class TestNumberFormat:
    def test_defaults(self) -> None:
        nf = NumberFormat()
        assert nf.format_code == "General"
        assert nf.format_id is None

    def test_builtin_formats(self) -> None:
        assert BUILTIN_FORMATS[0] == "General"
        assert BUILTIN_FORMATS[1] == "0"
        assert BUILTIN_FORMATS[14] == "mm-dd-yy"
        assert BUILTIN_FORMATS[49] == "@"


class TestStyleSheet:
    def test_default(self) -> None:
        ss = StyleSheet.default()
        assert len(ss.fonts) == 1
        assert len(ss.fills) == 2
        assert len(ss.borders) == 1
        assert len(ss.number_formats) == 0
        assert len(ss.cell_xfs) == 1

    def test_get_format_code_builtin(self) -> None:
        ss = StyleSheet.default()
        assert ss.get_format_code(0) == FORMAT_GENERAL
        assert ss.get_format_code(14) == "mm-dd-yy"

    def test_get_format_code_custom(self) -> None:
        ss = StyleSheet.default()
        ss.number_formats.append(NumberFormat(format_code="yyyy", format_id=164))
        assert ss.get_format_code(164) == "yyyy"

    def test_get_format_code_unknown(self) -> None:
        ss = StyleSheet.default()
        assert ss.get_format_code(999) == FORMAT_GENERAL


class TestCellStyle:
    def test_defaults(self) -> None:
        cs = CellStyle()
        assert cs.font_id == 0
        assert cs.fill_id == 0
        assert cs.border_id == 0
        assert cs.alignment is None
