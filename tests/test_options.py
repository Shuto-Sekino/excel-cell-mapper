"""
MapperOptions（コンストラクタのキーワード引数）に関するテスト。
- empty_cell
- date_format
- default_sheet
- transform
"""
import datetime

import pytest

from excel_cell_mapper import ExcelMapper


# ===========================================================================
# empty_cell オプション
# ===========================================================================
class TestEmptyCellOption:
    def test_default_returns_none(self, empty_cell_wb):
        """デフォルト（empty_cell='none'）は空白セルを None で返す。"""
        mapper = ExcelMapper(empty_cell_wb)
        result = mapper.map({"exists": "B1", "missing": "B2"})
        assert result["exists"] == "存在する値"
        assert result["missing"] is None

    def test_none_explicit(self, empty_cell_wb):
        """empty_cell='none' を明示しても同じ挙動。"""
        mapper = ExcelMapper(empty_cell_wb, empty_cell="none")
        result = mapper.map({"missing": "B2"})
        assert result["missing"] is None

    def test_omit_excludes_key(self, empty_cell_wb):
        """empty_cell='omit' は空白セルのキーをdictから除外する。"""
        mapper = ExcelMapper(empty_cell_wb, empty_cell="omit")
        result = mapper.map({"exists": "B1", "missing": "B2"})
        assert result["exists"] == "存在する値"
        assert "missing" not in result

    def test_empty_returns_empty_string(self, empty_cell_wb):
        """empty_cell='empty' は空白セルを空文字列で返す。"""
        mapper = ExcelMapper(empty_cell_wb, empty_cell="empty")
        result = mapper.map({"missing": "B2"})
        assert result["missing"] == ""

    def test_omit_works_in_nested_schema(self, empty_cell_wb):
        """empty_cell='omit' はネストしたスキーマでも機能する。"""
        mapper = ExcelMapper(empty_cell_wb, empty_cell="omit")
        result = mapper.map({
            "outer": {
                "exists": "B1",
                "missing": "B2",
            }
        })
        assert result["outer"] == {"exists": "存在する値"}

    def test_omit_works_in_list(self, empty_cell_wb):
        """empty_cell='omit' のとき、リスト内の None も除外される。"""
        mapper = ExcelMapper(empty_cell_wb, empty_cell="omit")
        # B1:存在する値, B2:None, B3:None
        result = mapper.map({"items": ["B1:B3"]})
        assert result["items"] == ["存在する値"]


# ===========================================================================
# date_format オプション
# ===========================================================================
class TestDateFormatOption:
    def test_default_returns_datetime(self, types_wb):
        """デフォルト（date_format='datetime'）は datetime オブジェクトを返す。"""
        mapper = ExcelMapper(types_wb)
        result = mapper.map({"date": "A5"})
        assert isinstance(result["date"], datetime.datetime)
        assert result["date"] == datetime.datetime(2024, 6, 1, 12, 0, 0)

    def test_datetime_explicit(self, types_wb):
        mapper = ExcelMapper(types_wb, date_format="datetime")
        result = mapper.map({"date": "A5"})
        assert isinstance(result["date"], datetime.datetime)

    def test_iso_returns_string(self, types_wb):
        """date_format='iso' は ISO 8601 文字列を返す。"""
        mapper = ExcelMapper(types_wb, date_format="iso")
        result = mapper.map({"date": "A5"})
        assert isinstance(result["date"], str)
        # ISO 8601 形式であることを確認
        parsed = datetime.datetime.fromisoformat(result["date"])
        assert parsed.year == 2024
        assert parsed.month == 6
        assert parsed.day == 1

    def test_local_returns_string(self, types_wb):
        """date_format='local' は文字列を返す。"""
        mapper = ExcelMapper(types_wb, date_format="local")
        result = mapper.map({"date": "A5"})
        assert isinstance(result["date"], str)


# ===========================================================================
# default_sheet オプション
# ===========================================================================
class TestDefaultSheetOption:
    def test_default_uses_first_sheet(self, second_sheet_wb):
        """default_sheet 未指定は最初のシートを使う。"""
        mapper = ExcelMapper(second_sheet_wb)
        result = mapper.map({"value": "A1"})
        assert result["value"] == "シート1の値"

    def test_default_sheet_by_name(self, second_sheet_wb):
        """default_sheet にシート名を指定できる。"""
        mapper = ExcelMapper(second_sheet_wb, default_sheet="Sheet2")
        result = mapper.map({"value": "A1"})
        assert result["value"] == "シート2の値"

    def test_default_sheet_by_index(self, second_sheet_wb):
        """default_sheet にインデックスを指定できる。"""
        mapper = ExcelMapper(second_sheet_wb, default_sheet=1)
        result = mapper.map({"value": "A1"})
        assert result["value"] == "シート2の値"

    def test_default_sheet_zero_index(self, second_sheet_wb):
        """default_sheet=0 は最初のシートを明示指定。"""
        mapper = ExcelMapper(second_sheet_wb, default_sheet=0)
        result = mapper.map({"value": "A1"})
        assert result["value"] == "シート1の値"

    def test_map_sheet_overrides_default(self, second_sheet_wb):
        """map() の sheet 引数は default_sheet より優先される。"""
        mapper = ExcelMapper(second_sheet_wb, default_sheet="Sheet1")
        result = mapper.map({"value": "A1"}, sheet="Sheet2")
        assert result["value"] == "シート2の値"


# ===========================================================================
# transform オプション
# ===========================================================================
class TestTransformOption:
    def test_transform_applied_to_string(self, simple_wb):
        """トランスフォーマーが文字列セルに適用される。"""
        def upper(value, ctx):
            if isinstance(value, str):
                return value.upper()
            return value

        mapper = ExcelMapper(simple_wb, transform=upper)
        result = mapper.map({"email": "B3"})
        assert result["email"] == "YAMADA@EXAMPLE.COM"

    def test_transform_applied_to_number(self, types_wb):
        """トランスフォーマーが数値セルに適用される。"""
        def double(value, ctx):
            if isinstance(value, (int, float)):
                return value * 2
            return value

        mapper = ExcelMapper(types_wb, transform=double)
        result = mapper.map({"n": "A1"})
        assert result["n"] == 84

    def test_transform_receives_context(self, simple_wb):
        """トランスフォーマーのコンテキストに cell_ref・sheet_name が渡される。"""
        received = []

        def capture(value, ctx):
            received.append({
                "cell_ref": ctx.cell_ref,
                "sheet_name": ctx.sheet_name,
                "col_index": ctx.col_index,
                "row_index": ctx.row_index,
            })
            return value

        mapper = ExcelMapper(simple_wb, transform=capture)
        mapper.map({"name": "B1"})

        assert len(received) == 1
        assert received[0]["cell_ref"] == "B1"
        assert received[0]["sheet_name"] == "Sheet1"
        assert received[0]["col_index"] == 1   # B列 = 1（0始まり）
        assert received[0]["row_index"] == 0   # 1行目 = 0（0始まり）

    def test_transform_applied_before_output(self, types_wb):
        """トランスフォーマーの戻り値がそのままdictの値になる。"""
        def to_str(value, ctx):
            return str(value)

        mapper = ExcelMapper(types_wb, transform=to_str)
        result = mapper.map({"n": "A1"})
        assert result["n"] == "42"

    def test_transform_applied_to_range_elements(self, range_wb):
        """リスト（範囲参照）の各要素にもトランスフォーマーが適用される。"""
        def exclaim(value, ctx):
            if isinstance(value, str):
                return value + "!"
            return value

        mapper = ExcelMapper(range_wb, transform=exclaim)
        result = mapper.map({"fruits": ["A1:A3"]})
        assert result["fruits"] == ["りんご!", "みかん!", "ぶどう!"]
