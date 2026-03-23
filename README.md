# excel-cell-mapper

A Python library for mapping Excel spreadsheet cells to Python dictionaries using a declarative DSL.

Instead of writing boilerplate openpyxl code to extract every cell manually, you define a schema once and let `ExcelMapper` do the rest.

## Features

- **Declarative DSL** — describe your mapping as a plain Python dict
- **Flexible cell references** — single cells, ranges, cross-sheet references
- **Rich schema patterns** — static keys, dynamic keys, nested dicts, list ranges, and table schemas
- **Type conversion** — Excel values automatically become `int`, `float`, `bool`, `str`, `datetime`
- **Configurable empty cell handling** — keep as `None`, omit from output, or replace with `""`
- **Custom transformers** — post-process any cell value with your own function
- **Multiple input sources** — file path, `bytes`, or any `BinaryIO` stream
- **Context manager support** — works with `with` statements

## Requirements

- Python 3.13+
- [openpyxl](https://openpyxl.readthedocs.io/) 3.1.5+

## Installation

```bash
pip install excel-cell-mapper
```

Or with [uv](https://github.com/astral-sh/uv):

```bash
uv add excel-cell-mapper
```

## Quick Start

Given an Excel file like:

| | A | B |
|---|---|---|
| 1 | Name | Alice |
| 2 | Age | 30 |
| 3 | Email | alice@example.com |

```python
from excel_cell_mapper import ExcelMapper

with ExcelMapper("form.xlsx") as mapper:
    result = mapper.map({
        "name":  "B1",
        "age":   "B2",
        "email": "B3",
    })

# result == {"name": "Alice", "age": 30, "email": "alice@example.com"}
```

## Schema Patterns

### Static key mapping

Map fixed field names to cell references:

```python
mapper.map({
    "company": "B1",
    "address": "B2",
    "phone":   "B3",
})
```

### Nested dict

Build hierarchical output by nesting schemas:

```python
mapper.map({
    "company": "B1",
    "contact": {
        "name":  "B3",
        "email": "B4",
        "phone": "B5",
    },
})
```

### Dynamic keys

When both key and value are cell references, the key is read from the spreadsheet at runtime:

```python
# Cell A1 contains "Revenue", B1 contains 1_200_000
mapper.map({
    "A1": "B1",   # → {"Revenue": 1200000}
    "A2": "B2",
    "A3": "B3",
})
```

### List from a range

Wrap a range string in a list to flatten the cells into a Python list:

```python
mapper.map({
    "tags": ["A1:A5"],
})
# → {"tags": ["alpha", "beta", "gamma", ...]}
```

### Table mapping

Use `$range` and `$schema` to iterate rows (or columns) and produce a list of dicts:

```python
mapper.map({
    "orders": {
        "$range":  "A2:C6",
        "$schema": {"id": 0, "product": 1, "quantity": 2},
    },
})
# → {"orders": [{"id": 1, "product": "Widget", "quantity": 10}, ...]}
```

Set `"$direction": "column"` to iterate columns instead of rows.
Set `"$skip_empty": True` to silently drop all-empty rows/columns.

### Cross-sheet references

Prefix any cell or range with the sheet name and `!`:

```python
mapper.map({
    "summary": "Summary!B2",
    "detail":  {
        "$range":  "Detail!A2:C20",
        "$schema": {"sku": 0, "qty": 1, "price": 2},
    },
})
```

## Constructor Options

```python
ExcelMapper(
    source,                          # str | Path | bytes | BinaryIO
    *,
    default_sheet=None,              # str | int | None
    empty_cell="none",               # "none" | "omit" | "empty"
    date_format="datetime",          # "datetime" | "iso" | "local"
    transform=None,                  # Callable[[CellValue, CellContext], object] | None
)
```

| Option | Default | Description |
|--------|---------|-------------|
| `default_sheet` | `None` (first sheet) | Sheet name or 0-based index used when a cell reference has no sheet qualifier |
| `empty_cell` | `"none"` | How empty cells appear in output: `"none"` → `None`, `"omit"` → key excluded, `"empty"` → `""` |
| `date_format` | `"datetime"` | Date cell output: `"datetime"` → `datetime` object, `"iso"` → ISO 8601 string, `"local"` → locale string |
| `transform` | `None` | A function `(value, context) -> object` applied to every cell value before it enters the result |

## API Reference

### `mapper.map(schema, *, sheet=None) -> dict`

Resolve *schema* against the workbook and return a dict.
`sheet` overrides `default_sheet` for this call only.

### `mapper.get_cell(cell_ref) -> CellValue`

Read a single cell value directly, bypassing any schema:

```python
value = mapper.get_cell("Sheet2!D7")
```

### `mapper.get_range(range_ref) -> list[list[CellValue]]`

Read a range as a 2-D list of rows:

```python
grid = mapper.get_range("A1:C3")
# [[row0col0, row0col1, row0col2], [row1col0, ...], ...]
```

### `mapper.get_sheet_names() -> list[str]`

Return all sheet names in workbook order.

## Custom Transformers

The `transform` option lets you post-process every resolved cell value.
The callback receives the raw value and a `CellContext` with location metadata:

```python
from excel_cell_mapper import ExcelMapper, CellContext

def trim_strings(value, ctx: CellContext):
    if isinstance(value, str):
        return value.strip()
    return value

with ExcelMapper("report.xlsx", transform=trim_strings) as mapper:
    result = mapper.map({"title": "B1", "notes": "B2"})
```

`CellContext` fields:

| Field | Type | Description |
|-------|------|-------------|
| `cell_ref` | `str` | Bare cell reference, e.g. `"B1"` |
| `sheet_name` | `str` | Name of the sheet |
| `col_index` | `int` | 0-based column index |
| `row_index` | `int` | 0-based row index |

## Error Handling

All exceptions inherit from `ExcelMapperError`:

| Exception | Raised when |
|-----------|-------------|
| `CellNotFoundError` | Cell reference is out of bounds |
| `SheetNotFoundError` | Referenced sheet does not exist |
| `InvalidSchemaError` | Schema structure is malformed |
| `ParseError` | Excel file cannot be parsed |

```python
from excel_cell_mapper import ExcelMapper, SheetNotFoundError, CellNotFoundError

try:
    result = mapper.map({"value": "MissingSheet!A1"})
except SheetNotFoundError as e:
    print(f"Sheet not found: {e}")
except CellNotFoundError as e:
    print(f"Cell out of range: {e}")
```

## Type Conversion

Excel cell types are automatically converted to Python types:

| Excel type | Python type |
|------------|-------------|
| Integer | `int` |
| Decimal | `float` |
| Text | `str` |
| Boolean | `bool` |
| Date / Time | `datetime.datetime` (or `str` if `date_format` is set) |
| Empty | `None` (or `""` / omitted — see `empty_cell`) |

## License

[MIT](LICENSE)
