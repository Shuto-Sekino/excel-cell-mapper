"""
DSL仕様 docs/schema/dsl.md の全パターンをカバーするテスト。
"""

from excel_cell_mapper import ExcelMapper


# ===========================================================================
# 1. 静的キー + セル参照
# ===========================================================================
class TestStaticKeyMapping:
    def test_single_field(self, simple_wb):
        mapper = ExcelMapper(simple_wb)
        result = mapper.map({"name": "B1"})
        assert result == {"name": "山田太郎"}

    def test_multiple_fields(self, simple_wb):
        mapper = ExcelMapper(simple_wb)
        schema = {
            "name": "B1",
            "age": "B2",
            "email": "B3",
            "department": "B4",
        }
        result = mapper.map(schema)
        assert result == {
            "name": "山田太郎",
            "age": 30,
            "email": "yamada@example.com",
            "department": "営業部",
        }

    def test_key_is_not_a_cell_ref_pattern(self, simple_wb):
        """キー名がセル参照パターン（A1形式）に見えない場合は静的キーとして扱う。"""
        mapper = ExcelMapper(simple_wb)
        result = mapper.map({"person_name": "B1"})
        assert result == {"person_name": "山田太郎"}


# ===========================================================================
# 2. 動的キー（セル参照 → セル参照）
# ===========================================================================
class TestDynamicKeyMapping:
    def test_single_dynamic_key(self, simple_wb):
        """A列のセル値をキー名として使用する。"""
        mapper = ExcelMapper(simple_wb)
        result = mapper.map({"A1": "B1"})
        # A1="氏名" → キーは "氏名"
        assert result == {"氏名": "山田太郎"}

    def test_multiple_dynamic_keys(self, simple_wb):
        mapper = ExcelMapper(simple_wb)
        schema = {
            "A1": "B1",
            "A2": "B2",
            "A3": "B3",
        }
        result = mapper.map(schema)
        assert result == {
            "氏名": "山田太郎",
            "年齢": 30,
            "メール": "yamada@example.com",
        }

    def test_dynamic_key_with_english_label(self, table_wb):
        """英字ラベルの動的キーも同様に動作する。"""
        mapper = ExcelMapper(table_wb)
        # A1="id", B1="name", C1="price"（ヘッダー行）
        result = mapper.map({"A1": "A2"})
        # A1="id" → キーは "id", A2=1 → 値は 1
        assert result == {"id": 1}

    def test_dynamic_key_must_both_be_cell_refs(self, simple_wb):
        """キーと値の両方がセル参照パターンのときだけ動的キーとして扱う。"""
        mapper = ExcelMapper(simple_wb)
        # キーが "name"（非セル参照）→ 静的キー
        result = mapper.map({"name": "B1"})
        assert "name" in result
        assert result["name"] == "山田太郎"


# ===========================================================================
# 3. ネストしたdict
# ===========================================================================
class TestNestedMapping:
    def test_one_level_nesting(self, nested_wb):
        mapper = ExcelMapper(nested_wb)
        schema = {
            "name": "B1",
            "address": {
                "prefecture": "B2",
                "city": "B3",
                "zip": "B4",
            },
        }
        result = mapper.map(schema)
        assert result == {
            "name": "山田太郎",
            "address": {
                "prefecture": "東京都",
                "city": "渋谷区",
                "zip": "150-0002",
            },
        }

    def test_sibling_nested_objects(self, nested_wb):
        mapper = ExcelMapper(nested_wb)
        schema = {
            "name": "B1",
            "address": {
                "prefecture": "B2",
                "city": "B3",
            },
            "contact": {
                "phone": "B5",
            },
        }
        result = mapper.map(schema)
        assert result["address"] == {"prefecture": "東京都", "city": "渋谷区"}
        assert result["contact"] == {"phone": "03-1234-5678"}

    def test_deeply_nested(self, nested_wb):
        """3段階以上のネストも動作する。"""
        mapper = ExcelMapper(nested_wb)
        schema = {
            "person": {
                "info": {
                    "name": "B1",
                }
            }
        }
        result = mapper.map(schema)
        assert result == {"person": {"info": {"name": "山田太郎"}}}


# ===========================================================================
# 4. リスト（セル範囲）
# ===========================================================================
class TestListMapping:
    def test_vertical_range(self, range_wb):
        mapper = ExcelMapper(range_wb)
        result = mapper.map({"fruits": ["A1:A5"]})
        assert result == {"fruits": ["りんご", "みかん", "ぶどう", "もも", "なし"]}

    def test_horizontal_range(self, column_range_wb):
        """横方向の範囲も1次元リストとして返す。"""
        mapper = ExcelMapper(column_range_wb)
        result = mapper.map({"seasons": ["A1:D1"]})
        assert result == {"seasons": ["春", "夏", "秋", "冬"]}

    def test_2d_range_returns_flat_list(self, table_wb):
        """2次元範囲の場合は行優先でフラット化したリストを返す。"""
        mapper = ExcelMapper(table_wb)
        result = mapper.map({"data": ["A1:B2"]})
        # A1:id, B1:name, A2:1, B2:りんご
        assert result == {"data": ["id", "name", 1, "りんご"]}

    def test_range_with_sheet_prefix(self, multi_sheet_wb):
        """シート指定付きの範囲参照も動作する。"""
        mapper = ExcelMapper(multi_sheet_wb)
        result = mapper.map({"info": ["顧客情報!B1:B3"]})
        assert result == {"info": ["C-001", "山田太郎", "yamada@example.com"]}


# ===========================================================================
# 5. dictのリスト（$range + $schema）
# ===========================================================================
class TestRangeSchemaMapping:
    def test_basic_row_direction(self, table_wb):
        mapper = ExcelMapper(table_wb)
        schema = {
            "products": {
                "$range": "A2:C4",
                "$schema": {"id": 0, "name": 1, "price": 2},
            }
        }
        result = mapper.map(schema)
        assert result == {
            "products": [
                {"id": 1, "name": "りんご", "price": 100},
                {"id": 2, "name": "みかん", "price": 80},
                {"id": 3, "name": "ぶどう", "price": 200},
            ]
        }

    def test_column_direction(self, column_range_wb):
        """$direction='column' のとき、列ごとにオブジェクトを生成する。"""
        mapper = ExcelMapper(column_range_wb)
        schema = {
            "seasons": {
                "$range": "A1:D1",
                "$schema": {"name": 0},
                "$direction": "column",
            }
        }
        result = mapper.map(schema)
        assert result == {
            "seasons": [
                {"name": "春"},
                {"name": "夏"},
                {"name": "秋"},
                {"name": "冬"},
            ]
        }

    def test_skip_empty_rows(self, table_with_empty_wb):
        """$skip_empty=True で空行をスキップする。"""
        mapper = ExcelMapper(table_with_empty_wb)
        schema = {
            "items": {
                "$range": "A2:C4",
                "$schema": {"id": 0, "name": 1, "price": 2},
                "$skip_empty": True,
            }
        }
        result = mapper.map(schema)
        # 空行（row3）はスキップされ2行分だけ返る
        assert len(result["items"]) == 2
        assert result["items"][0] == {"id": 1, "name": "りんご", "price": 100}
        assert result["items"][1] == {"id": 2, "name": "みかん", "price": 80}

    def test_skip_empty_false_includes_none_rows(self, table_with_empty_wb):
        """$skip_empty=False（デフォルト）では空行もNoneを含むdictとして返す。"""
        mapper = ExcelMapper(table_with_empty_wb)
        schema = {
            "items": {
                "$range": "A2:C4",
                "$schema": {"id": 0, "name": 1, "price": 2},
            }
        }
        result = mapper.map(schema)
        assert len(result["items"]) == 3
        assert result["items"][1] == {"id": None, "name": None, "price": None}

    def test_range_schema_with_sheet_prefix(self, multi_sheet_wb):
        """シート指定付きの $range も動作する。"""
        mapper = ExcelMapper(multi_sheet_wb)
        schema = {
            "info": {
                "$range": "顧客情報!B1:B3",
                "$schema": {"value": 0},
            }
        }
        result = mapper.map(schema)
        assert result["info"] == [
            {"value": "C-001"},
            {"value": "山田太郎"},
            {"value": "yamada@example.com"},
        ]

    def test_range_schema_default_direction_is_row(self, table_wb):
        """$direction 未指定はデフォルトで row。"""
        mapper = ExcelMapper(table_wb)
        schema_with = {
            "products": {
                "$range": "A2:C4",
                "$schema": {"id": 0, "name": 1, "price": 2},
                "$direction": "row",
            }
        }
        schema_without = {
            "products": {
                "$range": "A2:C4",
                "$schema": {"id": 0, "name": 1, "price": 2},
            }
        }
        mapper2 = ExcelMapper(table_wb)
        assert mapper.map(schema_with) == mapper2.map(schema_without)


# ===========================================================================
# 6. 複数シートにまたがるマッピング
# ===========================================================================
class TestMultiSheetMapping:
    def test_explicit_sheet_prefix_in_cell_ref(self, multi_sheet_wb):
        mapper = ExcelMapper(multi_sheet_wb)
        schema = {
            "customer_id": "顧客情報!B1",
            "customer_name": "顧客情報!B2",
            "order_id": "注文情報!B1",
            "total": "注文情報!B5",
        }
        result = mapper.map(schema)
        assert result["customer_id"] == "C-001"
        assert result["customer_name"] == "山田太郎"
        assert result["order_id"] == "ORD-001"
        assert result["total"] == 15000

    def test_nested_with_sheet_prefix(self, multi_sheet_wb):
        mapper = ExcelMapper(multi_sheet_wb)
        schema = {
            "customer": {
                "id": "顧客情報!B1",
                "name": "顧客情報!B2",
            },
            "order": {
                "id": "注文情報!B1",
                "total": "注文情報!B5",
            },
        }
        result = mapper.map(schema)
        assert result == {
            "customer": {"id": "C-001", "name": "山田太郎"},
            "order": {"id": "ORD-001", "total": 15000},
        }

    def test_sheet_prefix_overrides_default_sheet(self, multi_sheet_wb):
        """default_sheet が設定されていても、シート名付き参照は明示シートを使う。"""
        mapper = ExcelMapper(multi_sheet_wb, default_sheet="顧客情報")
        result = mapper.map({"order_id": "注文情報!B1"})
        assert result["order_id"] == "ORD-001"


# ===========================================================================
# 7. セル参照のバリエーション（記法）
# ===========================================================================
class TestCellReferenceNotation:
    def test_lowercase_cell_ref_is_normalized(self, simple_wb):
        """小文字のセル参照（例: b1）も許容する。"""
        mapper = ExcelMapper(simple_wb)
        result = mapper.map({"name": "b1"})
        assert result == {"name": "山田太郎"}

    def test_sheet_name_with_japanese(self, multi_sheet_wb):
        """日本語シート名を含むセル参照が動作する。"""
        mapper = ExcelMapper(multi_sheet_wb)
        result = mapper.map({"name": "顧客情報!B2"})
        assert result == {"name": "山田太郎"}

    def test_sheet_ref_by_index_in_map(self, second_sheet_wb):
        """map() の sheet オプションにインデックスを指定できる。"""
        mapper = ExcelMapper(second_sheet_wb)
        result = mapper.map({"value": "A1"}, sheet=1)
        assert result == {"value": "シート2の値"}

    def test_sheet_ref_by_name_in_map(self, second_sheet_wb):
        """map() の sheet オプションにシート名を指定できる。"""
        mapper = ExcelMapper(second_sheet_wb)
        result = mapper.map({"value": "A1"}, sheet="Sheet2")
        assert result == {"value": "シート2の値"}
