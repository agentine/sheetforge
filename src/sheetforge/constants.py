"""OOXML namespaces, content types, and relationship types for .xlsx files."""

from __future__ import annotations

# ── XML Namespaces ──────────────────────────────────────────────────────────

NS_SPREADSHEETML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_RELATIONSHIPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PACKAGE_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_MARKUP_COMPAT = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS_DRAWING = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_DRAWING_SPREAD = (
    "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
)
NS_DOCPROPS_CORE = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
NS_DOCPROPS_EXTENDED = (
    "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
)
NS_DC = "http://purl.org/dc/elements/1.1/"
NS_DCTERMS = "http://purl.org/dc/terms/"
NS_XSI = "http://www.w3.org/2001/XMLSchema-instance"
NS_VML = "urn:schemas-microsoft-com:vml"

# Namespace map for ET registration
NAMESPACE_MAP: dict[str, str] = {
    NS_SPREADSHEETML: "",
    NS_RELATIONSHIPS: "r",
    NS_PACKAGE_RELS: "",
    NS_CONTENT_TYPES: "",
    NS_MARKUP_COMPAT: "mc",
    NS_DRAWING: "a",
    NS_DRAWING_SPREAD: "xdr",
    NS_DC: "dc",
    NS_DCTERMS: "dcterms",
    NS_XSI: "xsi",
}

# ── Content Types ───────────────────────────────────────────────────────────

CT_WORKBOOK = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
)
CT_WORKSHEET = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
)
CT_STYLES = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
)
CT_SHARED_STRINGS = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
)
CT_RELATIONSHIPS = "application/vnd.openxmlformats-package.relationships+xml"
CT_CORE_PROPERTIES = "application/vnd.openxmlformats-package.core-properties+xml"
CT_EXTENDED_PROPERTIES = (
    "application/vnd.openxmlformats-officedocument.extended-properties+xml"
)
CT_DRAWING = (
    "application/vnd.openxmlformats-officedocument.drawing+xml"
)
CT_CHART = (
    "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
)

# ── Relationship Types ──────────────────────────────────────────────────────

REL_WORKSHEET = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
)
REL_STYLES = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
)
REL_SHARED_STRINGS = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
)
REL_OFFICE_DOCUMENT = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
)
REL_CORE_PROPERTIES = (
    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
)
REL_EXTENDED_PROPERTIES = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
)
REL_DRAWING = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
)
REL_CHART = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
)
REL_HYPERLINK = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
)

# ── Fixed Paths in .xlsx Archive ────────────────────────────────────────────

PATH_CONTENT_TYPES = "[Content_Types].xml"
PATH_RELS = "_rels/.rels"
PATH_WORKBOOK = "xl/workbook.xml"
PATH_WORKBOOK_RELS = "xl/_rels/workbook.xml.rels"
PATH_STYLES = "xl/styles.xml"
PATH_SHARED_STRINGS = "xl/sharedStrings.xml"

# ── Cell Type Constants ─────────────────────────────────────────────────────

TYPE_STRING = "s"
TYPE_NUMERIC = "n"
TYPE_BOOL = "b"
TYPE_FORMULA = "str"
TYPE_INLINE_STRING = "inlineStr"
TYPE_ERROR = "e"
TYPE_NULL = "n"

# ── Default Column Width / Row Height ───────────────────────────────────────

DEFAULT_COLUMN_WIDTH = 8.43
DEFAULT_ROW_HEIGHT = 15.0
