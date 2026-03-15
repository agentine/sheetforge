# sheetcraft

**Target:** openpyxl (209M monthly PyPI downloads)
**Language:** Python
**Package Name:** sheetcraft (verified available on PyPI)

## Problem

openpyxl is the dominant Python library for reading and writing Excel .xlsx/.xlsm files. It has 209M monthly downloads, ~10 contributors, and is hosted on Heptapod (an obscure Mercurial platform). Last PyPI release was June 2024. No PR activity or issue triage detected in recent months. No full-featured drop-in alternative exists — XlsxWriter is write-only, pylightxl is too limited.

## Scope

Drop-in replacement for openpyxl's most-used features. Pure Python, zero required dependencies (optional lxml for performance). Python 3.10+. Full type hints with py.typed marker.

## Architecture

OOXML spreadsheets (.xlsx) are ZIP archives containing XML files. Core components:

```
sheetcraft/
├── workbook.py         # Workbook class (open/create/save)
├── worksheet.py        # Worksheet class (cells, rows, columns)
├── cell.py             # Cell class (value, type, formatting)
├── reader/
│   ├── workbook.py     # Parse workbook.xml
│   ├── worksheet.py    # Parse sheet XML
│   ├── shared_strings.py  # Parse shared strings table
│   └── styles.py       # Parse styles.xml
├── writer/
│   ├── workbook.py     # Generate workbook.xml
│   ├── worksheet.py    # Generate sheet XML
│   ├── shared_strings.py  # Generate shared strings
│   └── styles.py       # Generate styles.xml
├── styles/
│   ├── fonts.py        # Font definitions
│   ├── fills.py        # Fill patterns and colors
│   ├── borders.py      # Border styles
│   ├── alignment.py    # Cell alignment
│   └── numbers.py      # Number format codes
├── chart/              # Chart support (phase 2)
├── utils.py            # Column letter/number conversion, cell references
├── constants.py        # OOXML namespaces, content types
└── exceptions.py       # Custom exceptions
```

## Phased Delivery

### Phase 1 — Core Read/Write (MVP)
- Open existing .xlsx files
- Create new workbooks
- Read/write cell values (strings, numbers, dates, booleans, None)
- Multiple worksheets (add, remove, rename, reorder)
- Row/column dimensions (width, height)
- Merged cells
- Read/write formulas (stored as strings, no evaluation)
- Save workbook (new file or overwrite)
- Shared strings table (read/write)
- Basic cell styles (font, fill, border, alignment, number format)
- Auto-filter

### Phase 2 — Extended Features
- Charts (bar, line, pie, scatter, area)
- Images (embed PNG/JPEG in worksheets)
- Data validation (dropdowns, ranges)
- Conditional formatting
- Hyperlinks
- Comments/notes
- Print settings (page setup, margins, headers/footers)
- Defined names / named ranges

### Phase 3 — Advanced
- Pivot tables (read/write)
- Worksheet protection
- Workbook protection
- Rich text in cells
- Sparklines
- Tables (structured references)

## API Compatibility

Match openpyxl's public API where practical:

```python
from sheetcraft import Workbook, load_workbook

# Create
wb = Workbook()
ws = wb.active
ws['A1'] = 'Hello'
ws['B1'] = 42
ws.append([1, 2, 3])
wb.save('output.xlsx')

# Read
wb = load_workbook('input.xlsx')
ws = wb.active
for row in ws.iter_rows(min_row=1, values_only=True):
    print(row)
```

## Key Design Decisions

1. **Pure Python, zero required deps.** Use stdlib `xml.etree.ElementTree` and `zipfile`. Optional `lxml` for faster XML parsing on large files.
2. **Type-safe.** Full type annotations, py.typed marker, strict mypy.
3. **Memory efficient.** Support read-only mode for streaming large files without loading entire workbook into memory.
4. **Test-driven.** Roundtrip tests: open openpyxl-generated files, modify, re-save, verify. Also test against real-world .xlsx files from Excel and Google Sheets.
5. **Spec-compliant.** Follow ECMA-376 (Office Open XML) where openpyxl deviates or is incomplete.

## Deliverables

- `sheetcraft` PyPI package
- Full test suite with >90% coverage
- API documentation
- Migration guide from openpyxl
