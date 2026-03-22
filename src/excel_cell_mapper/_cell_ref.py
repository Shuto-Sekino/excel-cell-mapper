"""
Utilities for parsing Excel cell/range references.

Supported formats:
  Single cell:  "A1", "B3", "Sheet1!A1", "顧客情報!B5"
  Range:        "A1:C3", "Sheet1!A1:C3"
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from openpyxl.utils import column_index_from_string, get_column_letter

# Excel hard limits
MAX_COL = 16384  # XFD
MAX_ROW = 1048576

# Regex: optional "SheetName!" prefix, then column letters, then digits
# We capture the sheet name as everything before the last "!", to allow
# sheet names that contain spaces / Japanese characters.
_CELL_BARE = r"[A-Za-z]+\d+"
_CELL_BARE_RE = re.compile(r"^([A-Za-z]+)(\d+)$")
_CELL_RE = re.compile(r"^(.*!)?([A-Za-z]+)(\d+)$")
_RANGE_RE = re.compile(r"^(.*!)?([A-Za-z]+)(\d+):([A-Za-z]+)(\d+)$")
_BARE_CELL_RE = re.compile(r"^[A-Za-z]+\d+$")  # no sheet prefix


@dataclass(frozen=True)
class CellAddress:
    sheet: str | None  # None means "use default sheet"
    col: int  # 1-based
    row: int  # 1-based

    @property
    def cell_ref(self) -> str:
        prefix = f"{self.sheet}!" if self.sheet else ""
        return f"{prefix}{get_column_letter(self.col)}{self.row}"


@dataclass(frozen=True)
class RangeAddress:
    sheet: str | None
    col_start: int
    row_start: int
    col_end: int
    row_end: int


def is_bare_cell_ref(s: str) -> bool:
    """Return True if *s* looks like a bare cell ref (e.g. 'A1', 'ZZ10')."""
    return bool(_BARE_CELL_RE.match(s))


def parse_cell_ref(ref: str) -> CellAddress:
    """
    Parse a cell reference string into a CellAddress.
    Raises InvalidSchemaError for invalid formats.
    Raises CellNotFoundError for out-of-Excel-bounds addresses.
    """
    from excel_cell_mapper._exceptions import CellNotFoundError, InvalidSchemaError

    ref.upper()
    # Re-apply the original capitalisation for sheet names
    m = _CELL_RE.match(ref)
    if not m:
        raise InvalidSchemaError(f"Invalid cell reference: {ref!r}")

    sheet_part, col_str, row_str = m.group(1), m.group(2), m.group(3)
    sheet = sheet_part.rstrip("!") if sheet_part else None
    col = column_index_from_string(col_str.upper())
    row = int(row_str)

    if col > MAX_COL or row > MAX_ROW:
        raise CellNotFoundError(ref)

    return CellAddress(sheet=sheet, col=col, row=row)


def parse_range_ref(ref: str) -> RangeAddress:
    """
    Parse a range reference string into a RangeAddress.
    Raises InvalidSchemaError for invalid formats.
    Raises CellNotFoundError for out-of-bounds addresses.
    """
    from excel_cell_mapper._exceptions import CellNotFoundError, InvalidSchemaError

    m = _RANGE_RE.match(ref)
    if not m:
        raise InvalidSchemaError(f"Invalid range reference: {ref!r}")

    sheet_part = m.group(1)
    col_start_str, row_start_str = m.group(2), m.group(3)
    col_end_str, row_end_str = m.group(4), m.group(5)

    sheet = sheet_part.rstrip("!") if sheet_part else None
    col_start = column_index_from_string(col_start_str.upper())
    row_start = int(row_start_str)
    col_end = column_index_from_string(col_end_str.upper())
    row_end = int(row_end_str)

    for col in (col_start, col_end):
        if col > MAX_COL:
            raise CellNotFoundError(ref)
    for row in (row_start, row_end):
        if row > MAX_ROW:
            raise CellNotFoundError(ref)

    return RangeAddress(
        sheet=sheet,
        col_start=col_start,
        row_start=row_start,
        col_end=col_end,
        row_end=row_end,
    )
