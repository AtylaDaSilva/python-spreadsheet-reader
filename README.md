# python-spreadsheet-reader

A Python library for reading, editing, and saving **Excel (.xlsx)** workbooks. It wraps [openpyxl](https://openpyxl.readthedocs.io/) with a higher-level API for row-by-row reading, cell access, lock detection, and workbook persistence.

## Requirements

- Python **3.14+**
- [Poetry](https://python-poetry.org/) (recommended) or pip

## Installation

### With Poetry

```bash
git clone <repository-url>
cd python-spreadsheet-reader
poetry install
```

### With pip

```bash
pip install .
```

## Quick start

```python
from python_spreadsheet_reader.readers.xlsx import XLSXReader

reader = XLSXReader("path/to/workbook.xlsx")

# Read all rows from the active sheet (values only, keyed by cell coordinates)
sheet_data = reader.read_sheet(cell_values_only=True)

# Read a specific sheet
sheet_data = reader.read_sheet("Sheet2", cell_values_only=True)

# Access and modify cells
reader.load_workbook()
reader.set_cell_value(value="Hello", coords="A1")
reader.set_cell_value(value=42, row=2, col=1)

# Save changes
reader.save_spreadsheet()                          # overwrite original
reader.save_spreadsheet("output/copy.xlsx")      # save to a new path

reader.close_workbook()
```

> **Note:** Close the workbook in Excel or LibreOffice before opening it with this library. If the file is in use, `SpreadsheetIsLockedException` is raised unless you pass `read_locked=True` to `load_workbook()` or `read_sheet()`.

## API overview

### `XLSXReader`

Entry point for working with a `.xlsx` file. Holds the workbook in memory and optionally caches the result of `read_sheet()`.

| Member | Description |
|--------|-------------|
| `workbook_path` | `Path` to the workbook passed to the constructor. |
| `_sheet_data` | Cached row data from the last `read_sheet()` call when `close_workbook=False`. |
| `_workbook` | Loaded `openpyxl.Workbook`, or `None` if closed. |

#### Constructor

```python
XLSXReader(workbook_path: str | Path)
```

#### Public methods

| Method | Returns | Description |
|--------|---------|-------------|
| `load_workbook(...)` | `openpyxl.Workbook` | Opens the workbook with configurable read-only, formula, VBA, link, and rich-text options. |
| `close_workbook()` | `None` | Closes the workbook and clears the sheet data cache. |
| `read_sheet(...)` | `dict[int, dict]` | Reads rows from the active or named sheet. Each key is a 1-based row number; each value is a dict of cells keyed by coordinate (e.g. `"A1"`) or column number. |
| `get_cell(coords=..., row=..., col=..., sheet_name=...)` | `Cell \| MergedCell` | Returns a cell by Excel coordinate or 1-based row/column. |
| `set_cell_value(value, coords=..., row=..., col=..., sheet_name=...)` | `Cell \| MergedCell` | Sets a cell value and returns the updated cell. |
| `save_spreadsheet(workbook_path=None, close_workbook=True)` | `Path` | Saves the workbook. Creates parent directories if needed. |
| `is_spreadsheet_locked()` | `bool` | Checks for Excel (`~$filename`) or LibreOffice (`.~lock.filename#`) lock files. |
| `adicionar_imagem(...)` | — | **Not implemented** — raises `NotImplementedError`. |

#### `read_sheet()` options

| Parameter | Default | Description |
|-----------|---------|-------------|
| `sheet_name` | `None` | Sheet to read; uses the active sheet when omitted. |
| `cell_values_only` | `False` | Return cell values instead of `Cell` objects. |
| `return_cell_coords` | `True` | Use coordinates (`"A1"`) as keys; otherwise use column numbers. |
| `read_only` | `False` | Open in read-optimized mode (no edits). |
| `keep_formulae` | `True` | Return formulas instead of cached values. |
| `read_locked` | `False` | Allow reading while another app has the file open. |
| `close_workbook` | `True` | Close the workbook after reading. |

#### Example `read_sheet()` output

```python
{
    1: {"A1": "ID", "B1": "Title", "C1": "Genre"},
    2: {"A2": 123, "B2": "Alien", "C2": "Science Fiction, Horror"},
    3: {"A3": 124, "B3": "Predator", "C3": "Science Fiction, Action"},
}
```

### Exceptions

All custom exceptions inherit from `SpreadsheetReaderException`.

| Exception | When raised |
|-----------|-------------|
| `NoActiveSpreadsheetException` | The workbook has no active sheet when one is required. |
| `SpreadsheetIsLockedException` | The file is open in Excel or LibreOffice and `read_locked=False`. |

Other errors you may encounter:

- `FileNotFoundError` — workbook path does not exist.
- `ValueError` — unsupported file extension or missing cell coordinates.

## Project structure

```
python-spreadsheet-reader/
├── python_spreadsheet_reader/
│   └── readers/
│       ├── xlsx.py          # XLSXReader
│       └── exceptions.py    # Custom exceptions
├── pyproject.toml
├── poetry.lock
└── README.md
```

## Dependencies

- [openpyxl](https://pypi.org/project/openpyxl/) — read and write `.xlsx` files
- [Pillow](https://pypi.org/project/Pillow/) — image support for openpyxl
- [ipykernel](https://pypi.org/project/ipykernel/) — Jupyter kernel support

## License

See the repository maintainers for licensing information.
