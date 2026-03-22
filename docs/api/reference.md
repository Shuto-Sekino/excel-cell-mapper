# APIリファレンス

> ステータス: 下書き v0.1

## インストール

```bash
pip install excel-cell-mapper
```

---

## 基本的な使い方

```python
from excel_cell_mapper import ExcelMapper

mapper = ExcelMapper("data.xlsx")

schema = {
    "name": "B1",
    "age": "B2",
    "email": "B3",
}

result = mapper.map(schema)
# => {"name": "山田太郎", "age": 30, "email": "yamada@example.com"}
```

---

## `ExcelMapper` クラス

### コンストラクタ

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

| 型 | 説明 |
|----|------|
| `str` | ファイルパス |
| `pathlib.Path` | Pathオブジェクト |
| `bytes` | バイナリデータ |
| `BinaryIO` | ファイルオブジェクト（`open("file.xlsx", "rb")` など） |

#### オプション引数

| 引数 | 型 | デフォルト | 説明 |
|------|----|------------|------|
| `default_sheet` | `str \| int \| None` | `None` | デフォルトで参照するシート名またはインデックス（0始まり）。`None` の場合は最初のシートを使用 |
| `empty_cell` | `"none" \| "omit" \| "empty"` | `"none"` | 空白セルの扱い。`"none"` は `None`、`"omit"` はキーを除外、`"empty"` は空文字列 `""` |
| `date_format` | `"datetime" \| "iso" \| "local"` | `"datetime"` | 日付セルの変換方法。`"datetime"` は `datetime` オブジェクト、`"iso"` はISO 8601文字列 |
| `transform` | `CellTransformer \| None` | `None` | セル値を変換するカスタムトランスフォーマー |

---

### `mapper.map(schema, *, sheet=None)`

スキーマに基づいてExcelをdictに変換します。

```python
def map(
    self,
    schema: Schema,
    *,
    sheet: str | int | None = None,
) -> dict
```

#### 引数

| 引数 | 型 | 説明 |
|------|----|------|
| `schema` | `Schema` | マッピングスキーマ |
| `sheet` | `str \| int \| None` | この呼び出し専用のシート指定。`default_sheet` を上書きする |

#### 戻り値

`dict` — スキーマに従って変換されたdict

---

### `mapper.map_many(schemas, *, sheet=None)`

複数のスキーマをまとめて変換します。

```python
def map_many(
    self,
    schemas: dict[str, Schema],
    *,
    sheet: str | int | None = None,
) -> dict[str, dict]
```

**使用例:**

```python
result = mapper.map_many({
    "customer": {"name": "顧客情報!B1", "id": "顧客情報!B2"},
    "order": {"date": "注文情報!C1", "total": "注文情報!C5"},
})
# => {"customer": {"name": "...", "id": "..."}, "order": {"date": ..., "total": ...}}
```

---

### `mapper.get_cell(cell_ref)`

単一セルの値を直接取得します。

```python
def get_cell(self, cell_ref: str) -> CellValue
```

```python
value = mapper.get_cell("Sheet1!B3")
# => "山田太郎"
```

---

### `mapper.get_range(range_ref)`

セル範囲の値を2次元リストで取得します。

```python
def get_range(self, range_ref: str) -> list[list[CellValue]]
```

```python
values = mapper.get_range("A1:C3")
# => [["a", "b", "c"], ["d", "e", "f"], ["g", "h", "i"]]
```

---

### `mapper.get_sheet_names()`

Excelファイル内のシート名一覧を取得します。

```python
def get_sheet_names(self) -> list[str]
```

---

## コンテキストマネージャー対応

`ExcelMapper` はコンテキストマネージャーとして使用できます。

```python
with ExcelMapper("data.xlsx") as mapper:
    result = mapper.map(schema)
```

---

## 型定義

```python
from typing import Union

# セル参照の値型
CellValue = Union[str, int, float, bool, datetime.datetime, None]

# スキーマ型（再帰的に定義）
Schema = Union[
    str,                        # セル参照 "B1", "Sheet1!A2"
    list[str],                  # 範囲参照 ["A1:A10"]
    dict[str, "Schema"],        # ネストしたdict（または RangeSchema）
]

# レンジスキーマ（$range を含むdict）
# {
#     "$range": str,
#     "$schema": dict[str, int],
#     "$direction": Literal["row", "column"],  # デフォルト "row"
#     "$skip_empty": bool,                     # デフォルト False
# }
```

---

## カスタムトランスフォーマー

```python
from excel_cell_mapper import CellContext

def my_transform(value: CellValue, context: CellContext) -> object:
    ...
```

### `CellContext`

```python
@dataclass
class CellContext:
    cell_ref: str      # セル参照 "B1"
    sheet_name: str    # シート名
    col_index: int     # 列インデックス（0始まり）
    row_index: int     # 行インデックス（0始まり）
```

---

## エラーハンドリング

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
    print(f"セル {e.cell_ref} が見つかりません")
except SheetNotFoundError as e:
    print(f"シート {e.sheet_name} が見つかりません")
except InvalidSchemaError as e:
    print(f"スキーマが不正です: {e}")
except ExcelMapperError as e:
    print(f"マッパーエラー: {e}")
```

### エラークラス一覧

| クラス | 継承元 | 説明 |
|--------|--------|------|
| `ExcelMapperError` | `Exception` | 基底エラークラス |
| `CellNotFoundError` | `ExcelMapperError` | 指定したセルが存在しない。`cell_ref` 属性を持つ |
| `SheetNotFoundError` | `ExcelMapperError` | 指定したシートが存在しない。`sheet_name` 属性を持つ |
| `InvalidSchemaError` | `ExcelMapperError` | スキーマの記法が不正 |
| `ParseError` | `ExcelMapperError` | Excelファイルのパースに失敗 |

---

## 依存ライブラリ（案）

| ライブラリ | 用途 |
|-----------|------|
| `openpyxl` | `.xlsx` ファイルの読み込み |
| `xlrd` | `.xls` ファイルの読み込み（オプション） |
