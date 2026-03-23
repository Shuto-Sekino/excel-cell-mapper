# Schema DSL Specification

> Status: Draft v0.1

## Overview

Schemas are written as Python dicts.
By writing cell references as values, the cell values are expanded into the resulting dict.

---

## Cell Reference Notation

### Single Cell Reference

```
"<column><row>"
```

| Example | Description |
|---------|-------------|
| `"A1"` | Column A, Row 1 |
| `"B3"` | Column B, Row 3 |
| `"AA10"` | Column AA, Row 10 |

### Cell Reference with Sheet Specification

```
"<sheet_name>!<column><row>"
```

| Example | Description |
|---------|-------------|
| `"Sheet2!A1"` | Column A, Row 1 of Sheet2 |
| `"CustomerInfo!B5"` | Column B, Row 5 of the "CustomerInfo" sheet |

### Cell Range Reference (for Lists)

```
"<start_cell>:<end_cell>"
```

| Example | Description |
|---------|-------------|
| `"A1:A10"` | Rows 1–10 of column A (vertical list) |
| `"A1:D1"` | Row 1 of columns A–D (horizontal list) |
| `"A1:C3"` | 2D range from A1 to C3 |
| `"Sheet1!A1:A5"` | Range with sheet specification |

---

## Schema Notation

### 1. Static Key + Cell Reference (Most Basic Form)

Specify the dict key name as a string and the cell reference as the value.

```python
schema = {
    "name": "B1",
    "age": "B2",
    "email": "B3",
}
```

**Example output:**
```python
{
    "name": "Yamada Taro",
    "age": 30,
    "email": "yamada@example.com",
}
```

---

### 2. Dynamic Key (Cell Reference → Cell Reference)

When both the key and value are cell references, the value of the key cell is used as the dict key name.

```python
schema = {
    "A1": "B1",
    "A2": "B2",
    "A3": "B3",
}
```

Excel data:

| | A | B |
|-|---|---|
| 1 | name | Yamada Taro |
| 2 | age | 30 |
| 3 | email | yamada@example.com |

**Example output:**
```python
{
    "name": "Yamada Taro",
    "age": 30,
    "email": "yamada@example.com",
}
```

---

### 3. Nested Dict

Nesting dicts within the schema generates hierarchical results.

```python
schema = {
    "name": "B1",
    "address": {
        "prefecture": "B4",
        "city": "B5",
        "zip": "B6",
    },
    "contact": {
        "email": "B7",
        "phone": "B8",
    },
}
```

**Example output:**
```python
{
    "name": "Yamada Taro",
    "address": {
        "prefecture": "Tokyo",
        "city": "Shibuya",
        "zip": "150-0002",
    },
    "contact": {
        "email": "yamada@example.com",
        "phone": "03-1234-5678",
    },
}
```

---

### 4. List (Cell Range)

Specifying a range reference as a list in the value converts that range into a list.

```python
schema = {
    "items": ["A5:A10"],
}
```

**Example output:**
```python
{
    "items": ["Apple", "Orange", "Grape", "Peach", "Pear", "Strawberry"],
}
```

> Handling of `None` and blank cells can be controlled via options (described later)

---

### 5. List of Dicts (Row/Column Range)

Notation for mapping each row (or column) in a range to a single dict.

```python
schema = {
    "products": {
        "$range": "A2:C6",
        "$schema": {
            "id": 0,
            "name": 1,
            "price": 2,
        },
    },
}
```

> Specify the target range in `$range` and the mapping of column indices (0-based) to field names in `$schema`.

**Excel data:**

| A | B | C |
|---|---|---|
| id | name | price |
| 1 | Apple | 100 |
| 2 | Orange | 80 |
| 3 | Grape | 200 |

**Example output:**
```python
{
    "products": [
        {"id": 1, "name": "Apple", "price": 100},
        {"id": 2, "name": "Orange", "price": 80},
        {"id": 3, "name": "Grape", "price": 200},
    ],
}
```

---

### 6. Cell Object Syntax (Explicit Sheet Specification)

Using the `{"cell": "...", "sheet": "..."}` format allows you to specify the cell reference and sheet name separately.

```python
schema = {
    "price": {"cell": "B3"},
    "name":  {"cell": "A1", "sheet": "ProductMaster"},
}
```

| Key | Required | Description |
|-----|----------|-------------|
| `cell` | Required | Cell reference (`"B1"` or `"Sheet1!B1"` format) |
| `sheet` | Optional | Sheet name. If specified, takes precedence over any sheet prefix in `cell`. |

> Equivalent to string-based sheet specification (`"Sheet1!B1"`), but useful when you want to pass the sheet name as a variable.

---

### 7. Mapping Across Multiple Sheets

Using cell references that include sheet names allows values from different sheets to be combined into a single dict.

```python
schema = {
    "customer": {
        "name": "CustomerInfo!B1",
        "id": "CustomerInfo!B2",
    },
    "order": {
        "date": "OrderInfo!C1",
        "total": "OrderInfo!C5",
    },
}
```

---

## Special Directives

Keys with the `$` prefix in a schema are treated as special directives.

| Directive | Description |
|-----------|-------------|
| `$range` | Specifies the target cell range |
| `$schema` | Defines the column/row mapping within the range |
| `$direction` | Traversal direction of the range (`"row"` \| `"column"`, default: `"row"`) |
| `$skip_empty` | Whether to skip rows/columns with empty cells (default: `False`) |

---

## Cell Value Type Conversion

Excel cell values are automatically converted to Python types according to the following rules.

| Excel Type | Python Type |
|------------|-------------|
| Number (integer) | `int` |
| Number (decimal) | `float` |
| String | `str` |
| Date | `datetime.datetime` (can also be a string format via options) |
| Boolean | `bool` |
| Empty cell | `None` (can also be omitted via options) |
| Error value | `None` (or handled by an error handler) |

---

## Validation (Future Extension)

```python
schema = {
    "name": {
        "$cell": "B1",
        "$required": True,
        "$type": str,
    },
    "age": {
        "$cell": "B2",
        "$type": int,
        "$min": 0,
        "$max": 150,
    },
}
```

> Validation is out of scope for v0.1. Planned for a future version.
