"""Generate xl/sharedStrings.xml for an xlsx archive."""

from __future__ import annotations

import xml.etree.ElementTree as ET

from sheetforge.constants import NS_SPREADSHEETML


def write_shared_strings(strings: list[str]) -> bytes:
    """Generate shared strings XML from a list of strings."""
    root = ET.Element(
        "sst",
        xmlns=NS_SPREADSHEETML,
        count=str(len(strings)),
        uniqueCount=str(len(strings)),
    )
    for s in strings:
        si = ET.SubElement(root, "si")
        t = ET.SubElement(si, "t")
        t.text = s
        # Preserve leading/trailing whitespace
        if s != s.strip():
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    result: bytes = ET.tostring(root, encoding="UTF-8", xml_declaration=True)
    return result
