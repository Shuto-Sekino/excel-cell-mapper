class ExcelMapperError(Exception):
    """Base error class for excel-cell-mapper."""


class CellNotFoundError(ExcelMapperError):
    def __init__(self, cell_ref: str):
        super().__init__(f"Cell not found or out of bounds: {cell_ref}")
        self.cell_ref = cell_ref


class SheetNotFoundError(ExcelMapperError):
    def __init__(self, sheet_name: str):
        super().__init__(f"Sheet not found: {sheet_name!r}")
        self.sheet_name = sheet_name


class InvalidSchemaError(ExcelMapperError):
    pass


class ParseError(ExcelMapperError):
    pass
