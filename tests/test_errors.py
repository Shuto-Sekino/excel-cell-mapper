"""
エラーハンドリングに関するテスト。
docs/api/reference.md のエラークラス一覧を全てカバーする。
"""
import pytest

from excel_cell_mapper import (
    CellNotFoundError,
    ExcelMapper,
    ExcelMapperError,
    InvalidSchemaError,
    ParseError,
    SheetNotFoundError,
)


# ===========================================================================
# エラー継承関係
# ===========================================================================
class TestErrorHierarchy:
    def test_cell_not_found_is_mapper_error(self):
        assert issubclass(CellNotFoundError, ExcelMapperError)

    def test_sheet_not_found_is_mapper_error(self):
        assert issubclass(SheetNotFoundError, ExcelMapperError)

    def test_invalid_schema_is_mapper_error(self):
        assert issubclass(InvalidSchemaError, ExcelMapperError)

    def test_parse_error_is_mapper_error(self):
        assert issubclass(ParseError, ExcelMapperError)

    def test_mapper_error_is_exception(self):
        assert issubclass(ExcelMapperError, Exception)


# ===========================================================================
# CellNotFoundError
# ===========================================================================
class TestCellNotFoundError:
    def test_out_of_range_cell(self, simple_wb):
        """シート内に存在しない非常に遠いセルを指定すると CellNotFoundError。"""
        mapper = ExcelMapper(simple_wb)
        with pytest.raises(CellNotFoundError) as exc_info:
            mapper.get_cell("ZZZ9999")
        assert exc_info.value.cell_ref == "ZZZ9999"

    def test_cell_not_found_has_cell_ref_attribute(self, simple_wb):
        """例外オブジェクトに cell_ref 属性が含まれる。"""
        mapper = ExcelMapper(simple_wb)
        with pytest.raises(CellNotFoundError) as exc_info:
            mapper.map({"x": "ZZZ9999"})
        assert hasattr(exc_info.value, "cell_ref")

    def test_cell_not_found_in_range(self, simple_wb):
        """範囲参照でも存在しないセルを含む場合は CellNotFoundError。"""
        mapper = ExcelMapper(simple_wb)
        with pytest.raises(CellNotFoundError):
            mapper.get_range("ZZZ9990:ZZZ9999")


# ===========================================================================
# SheetNotFoundError
# ===========================================================================
class TestSheetNotFoundError:
    def test_nonexistent_sheet_in_cell_ref(self, simple_wb):
        """存在しないシートを含むセル参照は SheetNotFoundError。"""
        mapper = ExcelMapper(simple_wb)
        with pytest.raises(SheetNotFoundError) as exc_info:
            mapper.get_cell("存在しないシート!A1")
        assert exc_info.value.sheet_name == "存在しないシート"

    def test_nonexistent_sheet_in_map(self, simple_wb):
        mapper = ExcelMapper(simple_wb)
        with pytest.raises(SheetNotFoundError):
            mapper.map({"x": "A1"}, sheet="存在しないシート")

    def test_nonexistent_default_sheet_by_name(self, simple_wb):
        with pytest.raises(SheetNotFoundError):
            ExcelMapper(simple_wb, default_sheet="存在しないシート")

    def test_nonexistent_default_sheet_by_index(self, simple_wb):
        """範囲外インデックスを default_sheet に指定すると SheetNotFoundError。"""
        with pytest.raises(SheetNotFoundError):
            ExcelMapper(simple_wb, default_sheet=999)

    def test_sheet_not_found_has_sheet_name_attribute(self, simple_wb):
        """例外オブジェクトに sheet_name 属性が含まれる。"""
        mapper = ExcelMapper(simple_wb)
        with pytest.raises(SheetNotFoundError) as exc_info:
            mapper.get_cell("NoSheet!A1")
        assert hasattr(exc_info.value, "sheet_name")


# ===========================================================================
# InvalidSchemaError
# ===========================================================================
class TestInvalidSchemaError:
    def test_schema_value_not_string_or_list_or_dict(self, simple_wb):
        """スキーマ値に不正な型（int など）を渡すと InvalidSchemaError。"""
        mapper = ExcelMapper(simple_wb)
        with pytest.raises(InvalidSchemaError):
            mapper.map({"x": 123})  # type: ignore

    def test_range_list_with_multiple_items(self, simple_wb):
        """リスト記法に複数の要素を渡すと InvalidSchemaError。"""
        mapper = ExcelMapper(simple_wb)
        with pytest.raises(InvalidSchemaError):
            mapper.map({"x": ["A1:A3", "B1:B3"]})

    def test_range_schema_missing_range_key(self, simple_wb):
        """$range なしで $schema だけ指定すると InvalidSchemaError。"""
        mapper = ExcelMapper(simple_wb)
        with pytest.raises(InvalidSchemaError):
            mapper.map({"x": {"$schema": {"id": 0}}})

    def test_range_schema_missing_schema_key(self, simple_wb):
        """$range はあるが $schema がないと InvalidSchemaError。"""
        mapper = ExcelMapper(simple_wb)
        with pytest.raises(InvalidSchemaError):
            mapper.map({"x": {"$range": "A1:C3"}})

    def test_invalid_direction_value(self, simple_wb):
        """$direction に 'row'/'column' 以外を指定すると InvalidSchemaError。"""
        mapper = ExcelMapper(simple_wb)
        with pytest.raises(InvalidSchemaError):
            mapper.map({
                "x": {
                    "$range": "A1:C3",
                    "$schema": {"id": 0},
                    "$direction": "diagonal",  # 不正値
                }
            })

    def test_invalid_cell_ref_format(self, simple_wb):
        """不正なセル参照フォーマット（数字のみなど）は InvalidSchemaError。"""
        mapper = ExcelMapper(simple_wb)
        with pytest.raises(InvalidSchemaError):
            mapper.map({"x": "123"})  # 数字のみは無効

    def test_invalid_range_format(self, simple_wb):
        """不正な範囲参照フォーマット（コロンなしで2セル並べるなど）は InvalidSchemaError。"""
        mapper = ExcelMapper(simple_wb)
        with pytest.raises(InvalidSchemaError):
            mapper.map({"x": ["A1 B3"]})  # コロンなし


# ===========================================================================
# ParseError
# ===========================================================================
class TestParseError:
    def test_invalid_bytes(self):
        """不正なバイト列を渡すと ParseError。"""
        with pytest.raises(ParseError):
            ExcelMapper(b"this is not an xlsx file")

    def test_empty_bytes(self):
        """空のバイト列を渡すと ParseError。"""
        with pytest.raises(ParseError):
            ExcelMapper(b"")

    def test_csv_bytes(self, tmp_path):
        """CSV ファイル（.xlsx ではない）を渡すと ParseError。"""
        csv_content = b"name,age\n\xe5\xb1\xb1\xe7\x94\xb0,30\n"
        with pytest.raises(ParseError):
            ExcelMapper(csv_content)


# ===========================================================================
# invalid empty_cell option
# ===========================================================================
class TestInvalidOptions:
    def test_invalid_empty_cell_option(self, simple_wb):
        with pytest.raises(ValueError):
            ExcelMapper(simple_wb, empty_cell="invalid_value")  # type: ignore

    def test_invalid_date_format_option(self, simple_wb):
        with pytest.raises(ValueError):
            ExcelMapper(simple_wb, date_format="invalid_value")  # type: ignore
