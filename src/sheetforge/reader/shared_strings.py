"""Parse xl/sharedStrings.xml from an xlsx archive."""

from __future__ import annotations

import xml.etree.ElementTree as ET

from sheetforge.constants import NS_SPREADSHEETML

_NS = f"{{{NS_SPREADSHEETML}}}"


def read_shared_strings(xml_bytes: bytes) -> list[str]:
    """Parse shared strings XML and return a list of strings.

    Each <si> element contains either a single <t> (plain text)
    or multiple <r> (rich text run) elements with <t> children.
    """
    root = ET.fromstring(xml_bytes)
    strings: list[str] = []
    for si in root.iter(f"{_NS}si"):
        # Check for simple <t> first
        t_elem = si.find(f"{_NS}t")
        if t_elem is not None and not list(si.iter(f"{_NS}r")):
            strings.append(t_elem.text or "")
            continue
        # Rich text: concatenate all <t> elements inside <r> runs
        parts: list[str] = []
        for r in si.iter(f"{_NS}r"):
            t = r.find(f"{_NS}t")
            if t is not None:
                parts.append(t.text or "")
        strings.append("".join(parts))
    return strings
