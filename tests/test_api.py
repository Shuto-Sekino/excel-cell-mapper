"""
ExcelMapper の補助APIメソッドに関するテスト。
- get_cell()
- get_range()
- get_sheet_names()
- map_many()
- コンテキストマネージャー
- セル値の型変換
- ソース種別（bytes / Path / BinaryIO）
"""

import datetime
import io

import pytest

from excel_cell_mapper import ExcelMapper


# ===========================================================================
# get_cell()
# ===========================================================================
class TestGetCell:
    def test_returns_string_value(self, simple_wb):
        mapper = ExcelMapper(simple_wb)
        assert mapper.get_cell("B1") == "山田太郎"

    def test_returns_int_value(self, simple_wb):
        mapper = ExcelMapper(simple_wb)
        assert mapper.get_cell("B2") == 30

    def test_returns_none_for_empty_cell(self, empty_cell_wb):
        mapper = ExcelMapper(empty_cell_wb)
        assert mapper.get_cell("B2") is None

    def test_with_sheet_prefix(self, multi_sheet_wb):
        mapper = ExcelMapper(multi_sheet_wb)
        assert mapper.get_cell("顧客情報!B2") == "山田太郎"

    def test_case_insensitive_ref(self, simple_wb):
        """小文字のセル参照でも動作する。"""
        mapper = ExcelMapper(simple_wb)
        assert mapper.get_cell("b1") == "山田太郎"


# ===========================================================================
# get_range()
# ===========================================================================
class TestGetRange:
    def test_vertical_range(self, range_wb):
        mapper = ExcelMapper(range_wb)
        result = mapper.get_range("A1:A5")
        assert result == [["りんご"], ["みかん"], ["ぶどう"], ["もも"], ["なし"]]

    def test_horizontal_range(self, column_range_wb):
        mapper = ExcelMapper(column_range_wb)
        result = mapper.get_range("A1:D1")
        assert result == [["春", "夏", "秋", "冬"]]

    def test_2d_range(self, table_wb):
        mapper = ExcelMapper(table_wb)
        result = mapper.get_range("A1:C2")
        assert result == [
            ["id", "name", "price"],
            [1, "りんご", 100],
        ]

    def test_single_cell_range(self, simple_wb):
        """単一セルの範囲（A1:A1）も 1x1 の2次元リストとして返す。"""
        mapper = ExcelMapper(simple_wb)
        result = mapper.get_range("B1:B1")
        assert result == [["山田太郎"]]

    def test_with_sheet_prefix(self, multi_sheet_wb):
        mapper = ExcelMapper(multi_sheet_wb)
        result = mapper.get_range("顧客情報!B1:B3")
        assert result == [["C-001"], ["山田太郎"], ["yamada@example.com"]]


# ===========================================================================
# get_sheet_names()
# ===========================================================================
class TestGetSheetNames:
    def test_single_sheet(self, simple_wb):
        mapper = ExcelMapper(simple_wb)
        assert mapper.get_sheet_names() == ["Sheet1"]

    def test_multiple_sheets(self, multi_sheet_wb):
        mapper = ExcelMapper(multi_sheet_wb)
        names = mapper.get_sheet_names()
        assert "顧客情報" in names
        assert "注文情報" in names

    def test_sheet_order(self, second_sheet_wb):
        mapper = ExcelMapper(second_sheet_wb)
        names = mapper.get_sheet_names()
        assert names == ["Sheet1", "Sheet2"]


# ===========================================================================
# map_many()
# ===========================================================================
class TestMapMany:
    def test_basic(self, multi_sheet_wb):
        mapper = ExcelMapper(multi_sheet_wb)
        result = mapper.map_many(
            {
                "customer": {"id": "顧客情報!B1", "name": "顧客情報!B2"},
                "order": {"id": "注文情報!B1", "total": "注文情報!B5"},
            }
        )
        assert result == {
            "customer": {"id": "C-001", "name": "山田太郎"},
            "order": {"id": "ORD-001", "total": 15000},
        }

    def test_returns_all_keys(self, simple_wb):
        mapper = ExcelMapper(simple_wb)
        result = mapper.map_many(
            {
                "a": {"name": "B1"},
                "b": {"age": "B2"},
            }
        )
        assert set(result.keys()) == {"a", "b"}

    def test_map_many_with_sheet_option(self, second_sheet_wb):
        """sheet オプションが各スキーマに適用される。"""
        mapper = ExcelMapper(second_sheet_wb)
        result = mapper.map_many(
            {"first": {"value": "A1"}},
            sheet="Sheet2",
        )
        assert result["first"]["value"] == "シート2の値"


# ===========================================================================
# コンテキストマネージャー
# ===========================================================================
class TestContextManager:
    def test_with_statement(self, simple_wb):
        with ExcelMapper(simple_wb) as mapper:
            result = mapper.map({"name": "B1"})
        assert result == {"name": "山田太郎"}

    def test_exits_cleanly_on_exception(self, simple_wb):
        """例外が発生してもコンテキストマネージャーが正常に終了する。"""
        from excel_cell_mapper import SheetNotFoundError

        with pytest.raises(SheetNotFoundError):
            with ExcelMapper(simple_wb) as mapper:
                mapper.map({"name": "B1"}, sheet="存在しないシート")


# ===========================================================================
# セル値の型変換
# ===========================================================================
class TestCellValueTypes:
    def test_integer(self, types_wb):
        mapper = ExcelMapper(types_wb)
        assert mapper.get_cell("A1") == 42
        assert isinstance(mapper.get_cell("A1"), int)

    def test_float(self, types_wb):
        mapper = ExcelMapper(types_wb)
        assert mapper.get_cell("A2") == pytest.approx(3.14)
        assert isinstance(mapper.get_cell("A2"), float)

    def test_string(self, types_wb):
        mapper = ExcelMapper(types_wb)
        assert mapper.get_cell("A3") == "hello"
        assert isinstance(mapper.get_cell("A3"), str)

    def test_boolean(self, types_wb):
        mapper = ExcelMapper(types_wb)
        assert mapper.get_cell("A4") is True
        assert isinstance(mapper.get_cell("A4"), bool)

    def test_datetime(self, types_wb):
        mapper = ExcelMapper(types_wb)
        value = mapper.get_cell("A5")
        assert isinstance(value, datetime.datetime)
        assert value == datetime.datetime(2024, 6, 1, 12, 0, 0)

    def test_empty_cell_returns_none(self, types_wb):
        mapper = ExcelMapper(types_wb)
        assert mapper.get_cell("A6") is None


# ===========================================================================
# ソース種別
# ===========================================================================
class TestSourceTypes:
    def test_bytes(self, simple_wb):
        """bytes を直接渡せる。"""
        mapper = ExcelMapper(simple_wb)
        assert mapper.map({"name": "B1"}) == {"name": "山田太郎"}

    def test_bytesio(self, simple_wb):
        """BytesIO も受け付ける。"""
        mapper = ExcelMapper(io.BytesIO(simple_wb))
        assert mapper.map({"name": "B1"}) == {"name": "山田太郎"}

    def test_file_path_str(self, simple_wb, tmp_path):
        """ファイルパス（str）から読み込める。"""
        p = tmp_path / "data.xlsx"
        p.write_bytes(simple_wb)
        mapper = ExcelMapper(str(p))
        assert mapper.map({"name": "B1"}) == {"name": "山田太郎"}

    def test_file_path_pathlib(self, simple_wb, tmp_path):
        """pathlib.Path から読み込める。"""
        p = tmp_path / "data.xlsx"
        p.write_bytes(simple_wb)
        mapper = ExcelMapper(p)
        assert mapper.map({"name": "B1"}) == {"name": "山田太郎"}

    def test_binary_file_object(self, simple_wb, tmp_path):
        """open() で開いたファイルオブジェクトから読み込める。"""
        p = tmp_path / "data.xlsx"
        p.write_bytes(simple_wb)
        with open(p, "rb") as f:
            mapper = ExcelMapper(f)
            result = mapper.map({"name": "B1"})
        assert result == {"name": "山田太郎"}
