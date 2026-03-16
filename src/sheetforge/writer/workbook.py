"""Generate a complete .xlsx ZIP archive."""

from __future__ import annotations

import io
import xml.etree.ElementTree as ET
import zipfile

from sheetforge.constants import (
    CT_RELATIONSHIPS,
    CT_SHARED_STRINGS,
    CT_STYLES,
    CT_WORKBOOK,
    CT_WORKSHEET,
    NS_CONTENT_TYPES,
    NS_PACKAGE_RELS,
    NS_SPREADSHEETML,
    PATH_CONTENT_TYPES,
    PATH_RELS,
    PATH_SHARED_STRINGS,
    PATH_STYLES,
    PATH_WORKBOOK,
    PATH_WORKBOOK_RELS,
    REL_OFFICE_DOCUMENT,
    REL_SHARED_STRINGS,
    REL_STYLES,
    REL_WORKSHEET,
)
from sheetforge.styles import StyleSheet
from sheetforge.writer.shared_strings import write_shared_strings
from sheetforge.writer.styles import write_styles


def _content_types_xml(sheet_count: int, has_shared_strings: bool) -> bytes:
    root = ET.Element("Types", xmlns=NS_CONTENT_TYPES)
    ET.SubElement(
        root,
        "Default",
        Extension="rels",
        ContentType=CT_RELATIONSHIPS,
    )
    ET.SubElement(
        root,
        "Default",
        Extension="xml",
        ContentType="application/xml",
    )
    ET.SubElement(root, "Override", PartName="/xl/workbook.xml", ContentType=CT_WORKBOOK)
    ET.SubElement(root, "Override", PartName="/xl/styles.xml", ContentType=CT_STYLES)
    if has_shared_strings:
        ET.SubElement(
            root,
            "Override",
            PartName="/xl/sharedStrings.xml",
            ContentType=CT_SHARED_STRINGS,
        )
    for i in range(1, sheet_count + 1):
        ET.SubElement(
            root,
            "Override",
            PartName=f"/xl/worksheets/sheet{i}.xml",
            ContentType=CT_WORKSHEET,
        )
    result: bytes = ET.tostring(root, encoding="UTF-8", xml_declaration=True)
    return result


def _root_rels_xml() -> bytes:
    root = ET.Element("Relationships", xmlns=NS_PACKAGE_RELS)
    ET.SubElement(
        root,
        "Relationship",
        Id="rId1",
        Type=REL_OFFICE_DOCUMENT,
        Target="xl/workbook.xml",
    )
    result: bytes = ET.tostring(root, encoding="UTF-8", xml_declaration=True)
    return result


def _workbook_xml(
    sheet_names: list[str], active_sheet: int = 0
) -> bytes:
    root = ET.Element(
        "workbook",
        xmlns=NS_SPREADSHEETML,
    )
    root.set(
        "xmlns:r",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    )
    views = ET.SubElement(root, "bookViews")
    ET.SubElement(views, "workbookView", activeTab=str(active_sheet))
    sheets = ET.SubElement(root, "sheets")
    for i, name in enumerate(sheet_names, 1):
        sheet_el = ET.SubElement(
            sheets,
            "sheet",
            name=name,
            sheetId=str(i),
        )
        sheet_el.set("r:id", f"rId{i}")
    result: bytes = ET.tostring(root, encoding="UTF-8", xml_declaration=True)
    return result


def _workbook_rels_xml(
    sheet_count: int, has_shared_strings: bool
) -> bytes:
    root = ET.Element("Relationships", xmlns=NS_PACKAGE_RELS)
    for i in range(1, sheet_count + 1):
        ET.SubElement(
            root,
            "Relationship",
            Id=f"rId{i}",
            Type=REL_WORKSHEET,
            Target=f"worksheets/sheet{i}.xml",
        )
    next_id = sheet_count + 1
    ET.SubElement(
        root,
        "Relationship",
        Id=f"rId{next_id}",
        Type=REL_STYLES,
        Target="styles.xml",
    )
    if has_shared_strings:
        next_id += 1
        ET.SubElement(
            root,
            "Relationship",
            Id=f"rId{next_id}",
            Type=REL_SHARED_STRINGS,
            Target="sharedStrings.xml",
        )
    result: bytes = ET.tostring(root, encoding="UTF-8", xml_declaration=True)
    return result


def write_workbook(
    sheet_names: list[str],
    sheet_xmls: list[bytes],
    stylesheet: StyleSheet | None = None,
    shared_strings: list[str] | None = None,
    active_sheet: int = 0,
) -> bytes:
    """Generate a complete .xlsx ZIP archive and return as bytes.

    Args:
        sheet_names: List of worksheet names.
        sheet_xmls: List of worksheet XML bytes (one per sheet).
        stylesheet: StyleSheet to write. Uses default if None.
        shared_strings: List of shared strings. Omit if not using shared strings.
        active_sheet: Index of the active sheet (default 0).
    """
    if stylesheet is None:
        stylesheet = StyleSheet.default()
    has_ss = shared_strings is not None and len(shared_strings) > 0
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(
            PATH_CONTENT_TYPES,
            _content_types_xml(len(sheet_names), has_ss),
        )
        zf.writestr(PATH_RELS, _root_rels_xml())
        zf.writestr(
            PATH_WORKBOOK, _workbook_xml(sheet_names, active_sheet)
        )
        zf.writestr(
            PATH_WORKBOOK_RELS,
            _workbook_rels_xml(len(sheet_names), has_ss),
        )
        zf.writestr(PATH_STYLES, write_styles(stylesheet))
        if has_ss and shared_strings is not None:
            zf.writestr(PATH_SHARED_STRINGS, write_shared_strings(shared_strings))
        for i, xml_data in enumerate(sheet_xmls, 1):
            zf.writestr(f"xl/worksheets/sheet{i}.xml", xml_data)
    return buf.getvalue()
