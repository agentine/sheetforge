# sheetcraft

Modern Python library for reading and writing Excel .xlsx files — drop-in openpyxl replacement.

## Install

```bash
pip install sheetcraft
```

## Quick Start

```python
from sheetcraft import Workbook, load_workbook

# Create a new workbook
wb = Workbook()
ws = wb.active
ws['A1'] = 'Hello'
ws['B1'] = 42
ws.append([1, 2, 3])
wb.save('output.xlsx')

# Read an existing workbook
wb = load_workbook('input.xlsx')
ws = wb.active
for row in ws.iter_rows(min_row=1, values_only=True):
    print(row)
```

## License

MIT
