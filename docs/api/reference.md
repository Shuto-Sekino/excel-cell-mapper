# API Reference

> Status: Draft v0.1

## Installation

```bash
pip install excel-cell-mapper
```

---

## Basic Usage

```python
from excel_cell_mapper import ExcelMapper

mapper = ExcelMapper("data.xlsx")

schema = {
    "name": "B1",
    "age": "B2",
    "email": "B3",
}

result = mapper.map(schema)
# => {"name": "Yamada Taro", "age": 30, "email": "yamada@example.com"}
```

---

## `ExcelMapper` Class

### Constructor

```python
ExcelMapper(
    source: ExcelSource,
    *,
    default_sheet: str | int | None = None,
    empty_cell: Literal["none", "omit", "empty"] = "none",
    date_format: Literal["datetime", "iso", "local"] = "datetime",
    transform: CellTransformer | None = None,
)
```

#### `ExcelSource`

```python
ExcelSource = str | Path | bytes | BinaryIO
```

| Type | Description |
|------|-------------|
| `str` | File path |
| `pathlib.Path` | Path object |
| `bytes` | Binary data |
| `BinaryIO` | File object (e.g., `open("file.xlsx", "rb")`) |

#### Optional Arguments

| Argument | Type | Default | Description |
|----------|------|---------|-------------|
| `default_sheet` | `str \| int \| None` | `None` | Default sheet name or index (0-based) to reference. `None` uses the first sheet. |
| `empty_cell` | `"none" \| "omit" \| "empty"` | `"none"` | How to handle empty cells. `"none"` returns `None`, `"omit"` excludes the key, `"empty"` returns an empty string `""`. |
| `date_format` | `"datetime" \| "iso" \| "local"` | `"datetime"` | How to convert date cells. `"datetime"` returns a `datetime` object, `"iso"` returns an ISO 8601 string. |
| `transform` | `CellTransformer \| None` | `None` | Custom transformer to process cell values. |

---

### `mapper.map(schema, *, sheet=None)`

Converts an Excel file to a dict based on the schema.

```python
def map(
    self,
    schema: Schema,
    *,
    sheet: str | int | None = None,
) -> dict
```

#### Arguments

| Argument | Type | Description |
|----------|------|-------------|
| `schema` | `Schema` | Mapping schema |
| `sheet` | `str \| int \| None` | Sheet specification for this call only. Overrides `default_sheet`. |

#### Return Value

`dict` — Dict converted according to the schema

---

### `mapper.get_cell(cell_ref)`

Retrieves the value of a single cell directly.

```python
def get_cell(self, cell_ref: str) -> CellValue
```

```python
value = mapper.get_cell("Sheet1!B3")
# => "Yamada Taro"
```

---

### `mapper.get_range(range_ref)`

Retrieves the values of a cell range as a 2D list.

```python
def get_range(self, range_ref: str) -> list[list[CellValue]]
```

```python
values = mapper.get_range("A1:C3")
# => [["a", "b", "c"], ["d", "e", "f"], ["g", "h", "i"]]
```

---

### `mapper.get_sheet_names()`

Returns a list of sheet names in the Excel file.

```python
def get_sheet_names(self) -> list[str]
```

---

## Context Manager Support

`ExcelMapper` can be used as a context manager.

```python
with ExcelMapper("data.xlsx") as mapper:
    result = mapper.map(schema)
```

---

## Type Definitions

```python
from typing import Union

# Cell value type
CellValue = Union[str, int, float, bool, datetime.datetime, None]

# Schema type (defined recursively)
Schema = Union[
    str,                        # Cell reference: "B1", "Sheet1!A2"
    list[str],                  # Range reference: ["A1:A10"]
    dict[str, "Schema"],        # Nested dict (or RangeSchema / CellObject)
]

# Cell object syntax (allows explicit sheet specification)
# {"cell": str}                       e.g., {"cell": "B1"}
# {"cell": str, "sheet": str}         e.g., {"cell": "B1", "sheet": "Sheet2"}
#
# The "sheet" key is optional. If specified, it takes precedence over any sheet prefix in "cell".

# Range schema (dict containing $range)
# {
#     "$range": str,
#     "$schema": dict[str, int],
#     "$direction": Literal["row", "column"],  # default: "row"
#     "$skip_empty": bool,                     # default: False
# }
```

---

## Custom Transformer

```python
from excel_cell_mapper import CellContext

def my_transform(value: CellValue, context: CellContext) -> object:
    ...
```

### `CellContext`

```python
@dataclass
class CellContext:
    cell_ref: str      # Cell reference, e.g., "B1"
    sheet_name: str    # Sheet name
    col_index: int     # Column index (0-based)
    row_index: int     # Row index (0-based)
```

---

## Error Handling

```python
from excel_cell_mapper import (
    ExcelMapperError,
    CellNotFoundError,
    SheetNotFoundError,
    InvalidSchemaError,
    ParseError,
)

try:
    result = mapper.map(schema)
except CellNotFoundError as e:
    print(f"Cell {e.cell_ref} not found")
except SheetNotFoundError as e:
    print(f"Sheet {e.sheet_name} not found")
except InvalidSchemaError as e:
    print(f"Invalid schema: {e}")
except ExcelMapperError as e:
    print(f"Mapper error: {e}")
```

### Error Classes

| Class | Inherits From | Description |
|-------|---------------|-------------|
| `ExcelMapperError` | `Exception` | Base error class |
| `CellNotFoundError` | `ExcelMapperError` | The specified cell does not exist. Has a `cell_ref` attribute. |
| `SheetNotFoundError` | `ExcelMapperError` | The specified sheet does not exist. Has a `sheet_name` attribute. |
| `InvalidSchemaError` | `ExcelMapperError` | Invalid schema notation |
| `ParseError` | `ExcelMapperError` | Failed to parse the Excel file |

---

## Dependencies

| Library | Purpose |
|---------|---------|
| `openpyxl` | Reading `.xlsx` files |
