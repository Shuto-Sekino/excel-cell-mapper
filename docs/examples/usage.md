# Usage Examples and Recipes

> Status: Draft v0.1

---

## Example 1: Simple Form Conversion

**Excel Layout:**

| | A | B |
|-|---|---|
| 1 | Name | Yamada Taro |
| 2 | Age | 30 |
| 3 | Email | yamada@example.com |
| 4 | Department | Sales |

**Schema:**

```python
from excel_cell_mapper import ExcelMapper

mapper = ExcelMapper("data.xlsx")

schema = {
    "name": "B1",
    "age": "B2",
    "email": "B3",
    "department": "B4",
}

result = mapper.map(schema)
```

**Output:**

```python
{
    "name": "Yamada Taro",
    "age": 30,
    "email": "yamada@example.com",
    "department": "Sales",
}
```

---

## Example 2: Dynamic Keys (Excel with label-value pairs)

**Excel Layout (Column A: labels, Column B: values):**

| | A | B |
|-|---|---|
| 1 | name | Yamada Taro |
| 2 | age | 30 |
| 3 | email | yamada@example.com |

**Schema (using column A cell values as keys):**

```python
schema = {
    "A1": "B1",
    "A2": "B2",
    "A3": "B3",
}

result = mapper.map(schema)
```

**Output:**

```python
{
    "name": "Yamada Taro",
    "age": 30,
    "email": "yamada@example.com",
}
```

---

## Example 3: Nested Dict (Address Information)

**Excel Layout:**

| | A | B |
|-|---|---|
| 1 | Name | Yamada Taro |
| 2 | Prefecture | Tokyo |
| 3 | City | Shibuya |
| 4 | Zip Code | 150-0002 |
| 5 | Phone | 03-1234-5678 |

**Schema:**

```python
schema = {
    "name": "B1",
    "address": {
        "prefecture": "B2",
        "city": "B3",
        "zip": "B4",
    },
    "phone": "B5",
}

result = mapper.map(schema)
```

**Output:**

```python
{
    "name": "Yamada Taro",
    "address": {
        "prefecture": "Tokyo",
        "city": "Shibuya",
        "zip": "150-0002",
    },
    "phone": "03-1234-5678",
}
```

---

## Example 4: List (Item List)

**Excel Layout:**

| | A |
|-|---|
| 1 | Apple |
| 2 | Orange |
| 3 | Grape |
| 4 | Peach |
| 5 | Pear |

**Schema:**

```python
schema = {
    "fruits": ["A1:A5"],
}

result = mapper.map(schema)
```

**Output:**

```python
{
    "fruits": ["Apple", "Orange", "Grape", "Peach", "Pear"],
}
```

---

## Example 5: List of Dicts (Product Table)

**Excel Layout:**

| | A | B | C |
|-|---|---|---|
| 1 | id | name | price |
| 2 | 1 | Apple | 100 |
| 3 | 2 | Orange | 80 |
| 4 | 3 | Grape | 200 |

**Schema (converts A2:C4 excluding the header row into a list of dicts):**

```python
schema = {
    "products": {
        "$range": "A2:C4",
        "$schema": {
            "id": 0,
            "name": 1,
            "price": 2,
        },
    },
}

result = mapper.map(schema)
```

**Output:**

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

## Example 6: Merging Data from Multiple Sheets

**Excel Sheet Structure:**
- `CustomerInfo` sheet: Customer basic data
- `OrderInfo` sheet: Order data

**Schema:**

```python
schema = {
    "customer": {
        "id": "CustomerInfo!B1",
        "name": "CustomerInfo!B2",
        "email": "CustomerInfo!B3",
    },
    "order": {
        "order_id": "OrderInfo!B1",
        "date": "OrderInfo!B2",
        "total": "OrderInfo!B5",
    },
}

result = mapper.map(schema)
```

**Output:**

```python
{
    "customer": {
        "id": "C-001",
        "name": "Yamada Taro",
        "email": "yamada@example.com",
    },
    "order": {
        "order_id": "ORD-2024-001",
        "date": datetime.datetime(2024, 1, 15, 0, 0),
        "total": 15000,
    },
}
```

---

## Example 7: Custom Transformer

Use a transformer when you want to process cell values after retrieval.

```python
from excel_cell_mapper import ExcelMapper, CellContext, CellValue

def my_transform(value: CellValue, context: CellContext) -> object:
    # Round floats to 2 decimal places
    if isinstance(value, float):
        return round(value, 2)
    # Strip whitespace from strings
    if isinstance(value, str):
        return value.strip()
    return value

mapper = ExcelMapper("data.xlsx", transform=my_transform)
result = mapper.map({"price": "B3"})
```

---

## Example 8: Loading from File Path (Typical Usage)

```python
from pathlib import Path
from excel_cell_mapper import ExcelMapper

def process_order_form(file_path: str | Path) -> dict:
    mapper = ExcelMapper(
        file_path,
        default_sheet="OrderForm",
        empty_cell="omit",
        date_format="iso",
    )

    schema = {
        "order_id": "B1",
        "customer": {
            "name": "B3",
            "address": "B4",
            "phone": "B5",
        },
        "items": {
            "$range": "A9:D20",
            "$schema": {"name": 0, "quantity": 1, "unit_price": 2, "subtotal": 3},
            "$skip_empty": True,
        },
        "total_amount": "D22",
    }

    return mapper.map(schema)
```

---

## Example 9: Using as a Context Manager

```python
from excel_cell_mapper import ExcelMapper

schema = {
    "title": "A1",
    "data": {
        "$range": "A3:B10",
        "$schema": {"label": 0, "value": 1},
    },
}

with ExcelMapper("report.xlsx") as mapper:
    result = mapper.map(schema)

print(result)
```

---

## Example 10: Loading from Binary Data (Web API etc.)

```python
# Example upload handling with FastAPI
from fastapi import UploadFile
from excel_cell_mapper import ExcelMapper

async def handle_upload(file: UploadFile) -> dict:
    content = await file.read()

    mapper = ExcelMapper(content)

    schema = {
        "name": "B1",
        "items": ["A2:A10"],
    }

    return mapper.map(schema)
```
