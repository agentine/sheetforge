"""Parse xl/styles.xml from an xlsx archive."""

from __future__ import annotations

import xml.etree.ElementTree as ET

from sheetforge.constants import NS_SPREADSHEETML
from sheetforge.styles import CellStyle, StyleSheet
from sheetforge.styles.alignment import Alignment
from sheetforge.styles.borders import Border, Side
from sheetforge.styles.fills import PatternFill
from sheetforge.styles.fonts import Font
from sheetforge.styles.numbers import NumberFormat

_NS = f"{{{NS_SPREADSHEETML}}}"


def _text(el: ET.Element | None, attr: str) -> str | None:
    if el is None:
        return None
    return el.get(attr)


def _bool(val: str | None) -> bool:
    if val is None:
        return False
    return val in ("1", "true")


def _float(val: str | None) -> float | None:
    if val is None:
        return None
    return float(val)


def _int(val: str | None) -> int | None:
    if val is None:
        return None
    return int(val)


def _parse_fonts(root: ET.Element) -> list[Font]:
    fonts: list[Font] = []
    fonts_el = root.find(f"{_NS}fonts")
    if fonts_el is None:
        return fonts
    for font_el in fonts_el.findall(f"{_NS}font"):
        name_el = font_el.find(f"{_NS}name")
        sz_el = font_el.find(f"{_NS}sz")
        b_el = font_el.find(f"{_NS}b")
        i_el = font_el.find(f"{_NS}i")
        u_el = font_el.find(f"{_NS}u")
        strike_el = font_el.find(f"{_NS}strike")
        color_el = font_el.find(f"{_NS}color")
        scheme_el = font_el.find(f"{_NS}scheme")
        family_el = font_el.find(f"{_NS}family")

        color: str | None = None
        if color_el is not None:
            color = color_el.get("rgb") or color_el.get("theme") or color_el.get("indexed")

        fonts.append(
            Font(
                name=name_el.get("val") if name_el is not None else None,
                size=_float(sz_el.get("val") if sz_el is not None else None),
                bold=b_el is not None and b_el.get("val", "1") != "0",
                italic=i_el is not None and i_el.get("val", "1") != "0",
                underline=u_el.get("val", "single") if u_el is not None else None,
                strike=strike_el is not None and strike_el.get("val", "1") != "0",
                color=color,
                scheme=scheme_el.get("val") if scheme_el is not None else None,
                family=_int(family_el.get("val") if family_el is not None else None),
            )
        )
    return fonts


def _parse_fills(root: ET.Element) -> list[PatternFill]:
    fills: list[PatternFill] = []
    fills_el = root.find(f"{_NS}fills")
    if fills_el is None:
        return fills
    for fill_el in fills_el.findall(f"{_NS}fill"):
        pf_el = fill_el.find(f"{_NS}patternFill")
        if pf_el is not None:
            fg_el = pf_el.find(f"{_NS}fgColor")
            bg_el = pf_el.find(f"{_NS}bgColor")
            fills.append(
                PatternFill(
                    patternType=pf_el.get("patternType"),
                    fgColor=fg_el.get("rgb") if fg_el is not None else None,
                    bgColor=bg_el.get("rgb") if bg_el is not None else None,
                )
            )
        else:
            fills.append(PatternFill())
    return fills


def _parse_borders(root: ET.Element) -> list[Border]:
    borders: list[Border] = []
    borders_el = root.find(f"{_NS}borders")
    if borders_el is None:
        return borders
    for border_el in borders_el.findall(f"{_NS}border"):

        def _side(tag: str) -> Side | None:
            el = border_el.find(f"{_NS}{tag}")
            if el is None:
                return None
            style = el.get("style")
            color_el = el.find(f"{_NS}color")
            color = color_el.get("rgb") if color_el is not None else None
            return Side(style=style, color=color)

        borders.append(
            Border(
                left=_side("left"),
                right=_side("right"),
                top=_side("top"),
                bottom=_side("bottom"),
                diagonal=_side("diagonal"),
            )
        )
    return borders


def _parse_number_formats(root: ET.Element) -> list[NumberFormat]:
    nfs: list[NumberFormat] = []
    numfmts_el = root.find(f"{_NS}numFmts")
    if numfmts_el is None:
        return nfs
    for nf_el in numfmts_el.findall(f"{_NS}numFmt"):
        fmt_id = _int(nf_el.get("numFmtId"))
        fmt_code = nf_el.get("formatCode", "General")
        nfs.append(NumberFormat(format_code=fmt_code, format_id=fmt_id))
    return nfs


def _parse_cell_xfs(root: ET.Element) -> list[CellStyle]:
    xfs: list[CellStyle] = []
    cellxfs_el = root.find(f"{_NS}cellXfs")
    if cellxfs_el is None:
        return xfs
    for xf_el in cellxfs_el.findall(f"{_NS}xf"):
        alignment: Alignment | None = None
        align_el = xf_el.find(f"{_NS}alignment")
        if align_el is not None:
            alignment = Alignment(
                horizontal=align_el.get("horizontal"),
                vertical=align_el.get("vertical"),
                wrap_text=_bool(align_el.get("wrapText")),
                shrink_to_fit=_bool(align_el.get("shrinkToFit")),
                indent=int(align_el.get("indent", "0")),
                text_rotation=int(align_el.get("textRotation", "0")),
            )
        xfs.append(
            CellStyle(
                font_id=int(xf_el.get("fontId", "0")),
                fill_id=int(xf_el.get("fillId", "0")),
                border_id=int(xf_el.get("borderId", "0")),
                number_format_id=int(xf_el.get("numFmtId", "0")),
                alignment=alignment,
                apply_font=_bool(xf_el.get("applyFont")),
                apply_fill=_bool(xf_el.get("applyFill")),
                apply_border=_bool(xf_el.get("applyBorder")),
                apply_number_format=_bool(xf_el.get("applyNumberFormat")),
                apply_alignment=_bool(xf_el.get("applyAlignment")),
            )
        )
    return xfs


def read_styles(xml_bytes: bytes) -> StyleSheet:
    """Parse styles.xml and return a StyleSheet."""
    root = ET.fromstring(xml_bytes)
    return StyleSheet(
        fonts=_parse_fonts(root),
        fills=_parse_fills(root),
        borders=_parse_borders(root),
        number_formats=_parse_number_formats(root),
        cell_xfs=_parse_cell_xfs(root),
    )
