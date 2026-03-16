"""Parse xl/workbook.xml and xl/_rels/workbook.xml.rels from an xlsx archive."""

from __future__ import annotations

import xml.etree.ElementTree as ET
from dataclasses import dataclass, field

from sheetforge.constants import NS_PACKAGE_RELS, NS_SPREADSHEETML

_NS_WB = f"{{{NS_SPREADSHEETML}}}"
_NS_REL = f"{{{NS_PACKAGE_RELS}}}"


@dataclass
class SheetInfo:
    """Metadata about a single worksheet."""

    name: str
    sheet_id: int
    r_id: str
    state: str = "visible"


@dataclass
class DefinedName:
    """A defined name (named range) entry."""

    name: str
    value: str
    sheet_id: int | None = None


@dataclass
class WorkbookData:
    """All data parsed from workbook.xml and its relationships."""

    sheets: list[SheetInfo] = field(default_factory=list)
    active_sheet: int = 0
    defined_names: list[DefinedName] = field(default_factory=list)
    rels: dict[str, str] = field(default_factory=dict)


def read_workbook(xml_bytes: bytes) -> WorkbookData:
    """Parse workbook.xml and return structured data."""
    root = ET.fromstring(xml_bytes)
    data = WorkbookData()

    # Parse sheets
    sheets_el = root.find(f"{_NS_WB}sheets")
    if sheets_el is not None:
        ns_r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        for sheet_el in sheets_el.findall(f"{_NS_WB}sheet"):
            name = sheet_el.get("name", "")
            sheet_id = int(sheet_el.get("sheetId", "0"))
            r_id = sheet_el.get(f"{{{ns_r}}}id", "")
            state = sheet_el.get("state", "visible")
            data.sheets.append(
                SheetInfo(name=name, sheet_id=sheet_id, r_id=r_id, state=state)
            )

    # Parse active sheet (bookViews/workbookView/@activeTab)
    views_el = root.find(f"{_NS_WB}bookViews")
    if views_el is not None:
        view_el = views_el.find(f"{_NS_WB}workbookView")
        if view_el is not None:
            active = view_el.get("activeTab")
            if active is not None:
                data.active_sheet = int(active)

    # Parse defined names
    names_el = root.find(f"{_NS_WB}definedNames")
    if names_el is not None:
        for name_el in names_el.findall(f"{_NS_WB}definedName"):
            name = name_el.get("name", "")
            value = name_el.text or ""
            local_id = name_el.get("localSheetId")
            data.defined_names.append(
                DefinedName(
                    name=name,
                    value=value,
                    sheet_id=int(local_id) if local_id is not None else None,
                )
            )

    return data


def read_workbook_rels(xml_bytes: bytes) -> dict[str, str]:
    """Parse workbook.xml.rels and return rId -> target mapping."""
    root = ET.fromstring(xml_bytes)
    rels: dict[str, str] = {}
    for rel in root.findall(f"{_NS_REL}Relationship"):
        r_id = rel.get("Id", "")
        target = rel.get("Target", "")
        rels[r_id] = target
    return rels
