"""Generate xl/styles.xml for an xlsx archive."""

from __future__ import annotations

import xml.etree.ElementTree as ET

from sheetcraft.constants import NS_SPREADSHEETML
from sheetcraft.styles import CellStyle, StyleSheet
from sheetcraft.styles.borders import Border, Side
from sheetcraft.styles.fills import PatternFill
from sheetcraft.styles.fonts import Font
from sheetcraft.styles.numbers import NumberFormat


def _write_font(parent: ET.Element, font: Font) -> None:
    f_el = ET.SubElement(parent, "font")
    if font.bold:
        ET.SubElement(f_el, "b")
    if font.italic:
        ET.SubElement(f_el, "i")
    if font.strike:
        ET.SubElement(f_el, "strike")
    if font.underline:
        ET.SubElement(f_el, "u", val=font.underline)
    if font.size is not None:
        ET.SubElement(f_el, "sz", val=str(font.size))
    if font.color:
        ET.SubElement(f_el, "color", rgb=font.color)
    if font.name:
        ET.SubElement(f_el, "name", val=font.name)
    if font.family is not None:
        ET.SubElement(f_el, "family", val=str(font.family))
    if font.scheme:
        ET.SubElement(f_el, "scheme", val=font.scheme)


def _write_fill(parent: ET.Element, fill: PatternFill) -> None:
    fill_el = ET.SubElement(parent, "fill")
    pf_el = ET.SubElement(fill_el, "patternFill")
    if fill.patternType:
        pf_el.set("patternType", fill.patternType)
    if fill.fgColor:
        ET.SubElement(pf_el, "fgColor", rgb=fill.fgColor)
    if fill.bgColor:
        ET.SubElement(pf_el, "bgColor", rgb=fill.bgColor)


def _write_side(parent: ET.Element, tag: str, side: Side | None) -> None:
    if side is None or side.style is None:
        ET.SubElement(parent, tag)
        return
    el = ET.SubElement(parent, tag, style=side.style)
    if side.color:
        ET.SubElement(el, "color", rgb=side.color)


def _write_border(parent: ET.Element, border: Border) -> None:
    b_el = ET.SubElement(parent, "border")
    _write_side(b_el, "left", border.left)
    _write_side(b_el, "right", border.right)
    _write_side(b_el, "top", border.top)
    _write_side(b_el, "bottom", border.bottom)
    _write_side(b_el, "diagonal", border.diagonal)


def _write_xf(parent: ET.Element, xf: CellStyle) -> None:
    attrs: dict[str, str] = {
        "numFmtId": str(xf.number_format_id),
        "fontId": str(xf.font_id),
        "fillId": str(xf.fill_id),
        "borderId": str(xf.border_id),
    }
    if xf.apply_font:
        attrs["applyFont"] = "1"
    if xf.apply_fill:
        attrs["applyFill"] = "1"
    if xf.apply_border:
        attrs["applyBorder"] = "1"
    if xf.apply_number_format:
        attrs["applyNumberFormat"] = "1"
    if xf.apply_alignment:
        attrs["applyAlignment"] = "1"
    xf_el = ET.SubElement(parent, "xf", attrs)
    if xf.alignment is not None:
        align_attrs: dict[str, str] = {}
        if xf.alignment.horizontal:
            align_attrs["horizontal"] = xf.alignment.horizontal
        if xf.alignment.vertical:
            align_attrs["vertical"] = xf.alignment.vertical
        if xf.alignment.wrap_text:
            align_attrs["wrapText"] = "1"
        if xf.alignment.shrink_to_fit:
            align_attrs["shrinkToFit"] = "1"
        if xf.alignment.indent:
            align_attrs["indent"] = str(xf.alignment.indent)
        if xf.alignment.text_rotation:
            align_attrs["textRotation"] = str(xf.alignment.text_rotation)
        ET.SubElement(xf_el, "alignment", align_attrs)


def write_styles(stylesheet: StyleSheet) -> bytes:
    """Generate styles.xml from a StyleSheet."""
    root = ET.Element("styleSheet", xmlns=NS_SPREADSHEETML)

    # Number formats
    if stylesheet.number_formats:
        nfs_el = ET.SubElement(
            root, "numFmts", count=str(len(stylesheet.number_formats))
        )
        for nf in stylesheet.number_formats:
            ET.SubElement(
                nfs_el,
                "numFmt",
                numFmtId=str(nf.format_id or 0),
                formatCode=nf.format_code,
            )

    # Fonts
    fonts_el = ET.SubElement(root, "fonts", count=str(len(stylesheet.fonts)))
    for font in stylesheet.fonts:
        _write_font(fonts_el, font)

    # Fills
    fills_el = ET.SubElement(root, "fills", count=str(len(stylesheet.fills)))
    for fill in stylesheet.fills:
        _write_fill(fills_el, fill)

    # Borders
    borders_el = ET.SubElement(
        root, "borders", count=str(len(stylesheet.borders))
    )
    for border in stylesheet.borders:
        _write_border(borders_el, border)

    # cellStyleXfs (required minimal entry)
    cell_style_xfs_el = ET.SubElement(root, "cellStyleXfs", count="1")
    ET.SubElement(
        cell_style_xfs_el,
        "xf",
        numFmtId="0",
        fontId="0",
        fillId="0",
        borderId="0",
    )

    # cellXfs
    xfs_el = ET.SubElement(
        root, "cellXfs", count=str(len(stylesheet.cell_xfs))
    )
    for xf in stylesheet.cell_xfs:
        _write_xf(xfs_el, xf)

    # cellStyles (required minimal entry)
    cell_styles_el = ET.SubElement(root, "cellStyles", count="1")
    ET.SubElement(
        cell_styles_el, "cellStyle", name="Normal", xfId="0", builtinId="0"
    )

    result: bytes = ET.tostring(root, encoding="UTF-8", xml_declaration=True)
    return result
