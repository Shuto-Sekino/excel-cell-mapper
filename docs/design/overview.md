# excel-cell-mapper Design Overview

## Project Overview

`excel-cell-mapper` is a Python library that converts Excel files into dicts based on a schema (cell ID mapping).
Users define mappings using a DSL (Domain-Specific Language) with cell IDs, and can extract values from arbitrary Excel cells into structured dicts.

## Problems Solved

- Reduces the manual conversion cost of bringing Excel-based data entry forms into systems
- Enables flexible structured data extraction based on cell positions
- Allows complex structures such as nested dicts and lists to be expressed through schema definitions alone

## Core Concepts

### Mapping by Cell ID

Maps Excel cell references (e.g., `A1`, `B3`, `Sheet2!C5`) to dict fields.

```
Excel cell ──(schema)──▶ dict field
```

### Schema

Schemas are written as Python dicts.

| Notation | Description |
|----------|-------------|
| `{ "fieldName": "B1" }` | Static key + cell reference |
| `{ "A1": "B1" }` | Dynamic key (value of A1) + cell reference (value of B1) |
| `{ "fieldName": { ... } }` | Nested dict |
| `{ "fieldName": ["A1:A10"] }` | List from cell range |

## Documentation Structure

```
docs/
├── design/
│   └── overview.md         # This file (design overview)
├── schema/
│   └── dsl.md              # Schema DSL specification
├── api/
│   └── reference.md        # API reference
└── examples/
    └── usage.md            # Usage examples and recipes
```

## Design Principles

- **Simplicity**: Handles most use cases with minimal schema definitions
- **Pythonic**: API design following Python conventions (type hints, dataclass, exception classes)
- **Extensibility**: Supports adding custom transformers and validators
- **Multi-sheet support**: Supports mappings spanning multiple sheets
