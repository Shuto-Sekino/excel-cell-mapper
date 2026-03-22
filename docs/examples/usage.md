# 使用例・レシピ集

> ステータス: 下書き v0.1

---

## 例1: シンプルなフォームの変換

**Excelレイアウト:**

| | A | B |
|-|---|---|
| 1 | 氏名 | 山田太郎 |
| 2 | 年齢 | 30 |
| 3 | メール | yamada@example.com |
| 4 | 部署 | 営業部 |

**スキーマ:**

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

**出力:**

```python
{
    "name": "山田太郎",
    "age": 30,
    "email": "yamada@example.com",
    "department": "営業部",
}
```

---

## 例2: 動的キー（ラベルと値が対になったExcel）

**Excelレイアウト（A列がラベル、B列が値）:**

| | A | B |
|-|---|---|
| 1 | name | 山田太郎 |
| 2 | age | 30 |
| 3 | email | yamada@example.com |

**スキーマ（A列のセル値をキーに使用）:**

```python
schema = {
    "A1": "B1",
    "A2": "B2",
    "A3": "B3",
}

result = mapper.map(schema)
```

**出力:**

```python
{
    "name": "山田太郎",
    "age": 30,
    "email": "yamada@example.com",
}
```

---

## 例3: ネストしたdict（住所情報）

**Excelレイアウト:**

| | A | B |
|-|---|---|
| 1 | 氏名 | 山田太郎 |
| 2 | 都道府県 | 東京都 |
| 3 | 市区町村 | 渋谷区 |
| 4 | 郵便番号 | 150-0002 |
| 5 | 電話番号 | 03-1234-5678 |

**スキーマ:**

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

**出力:**

```python
{
    "name": "山田太郎",
    "address": {
        "prefecture": "東京都",
        "city": "渋谷区",
        "zip": "150-0002",
    },
    "phone": "03-1234-5678",
}
```

---

## 例4: リスト（品目一覧）

**Excelレイアウト:**

| | A |
|-|---|
| 1 | りんご |
| 2 | みかん |
| 3 | ぶどう |
| 4 | もも |
| 5 | なし |

**スキーマ:**

```python
schema = {
    "fruits": ["A1:A5"],
}

result = mapper.map(schema)
```

**出力:**

```python
{
    "fruits": ["りんご", "みかん", "ぶどう", "もも", "なし"],
}
```

---

## 例5: dictのリスト（商品テーブル）

**Excelレイアウト:**

| | A | B | C |
|-|---|---|---|
| 1 | id | name | price |
| 2 | 1 | りんご | 100 |
| 3 | 2 | みかん | 80 |
| 4 | 3 | ぶどう | 200 |

**スキーマ（ヘッダー行を除いたA2:C4をdictのリストに変換）:**

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

**出力:**

```python
{
    "products": [
        {"id": 1, "name": "りんご", "price": 100},
        {"id": 2, "name": "みかん", "price": 80},
        {"id": 3, "name": "ぶどう", "price": 200},
    ],
}
```

---

## 例6: 複数シートのデータを結合

**Excelのシート構成:**
- `顧客情報` シート: 顧客の基本データ
- `注文情報` シート: 注文データ

**スキーマ:**

```python
schema = {
    "customer": {
        "id": "顧客情報!B1",
        "name": "顧客情報!B2",
        "email": "顧客情報!B3",
    },
    "order": {
        "order_id": "注文情報!B1",
        "date": "注文情報!B2",
        "total": "注文情報!B5",
    },
}

result = mapper.map(schema)
```

**出力:**

```python
{
    "customer": {
        "id": "C-001",
        "name": "山田太郎",
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

## 例7: カスタムトランスフォーマー

セル値を取得後に加工したい場合はトランスフォーマーを使います。

```python
from excel_cell_mapper import ExcelMapper, CellContext, CellValue

def my_transform(value: CellValue, context: CellContext) -> object:
    # 数値は小数点2桁に丸める
    if isinstance(value, float):
        return round(value, 2)
    # 文字列は前後の空白を除去
    if isinstance(value, str):
        return value.strip()
    return value

mapper = ExcelMapper("data.xlsx", transform=my_transform)
result = mapper.map({"price": "B3"})
```

---

## 例8: ファイルパスからの読み込み（典型的な使い方）

```python
from pathlib import Path
from excel_cell_mapper import ExcelMapper

def process_order_form(file_path: str | Path) -> dict:
    mapper = ExcelMapper(
        file_path,
        default_sheet="注文書",
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

## 例9: コンテキストマネージャーとして使用

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

## 例10: バイナリデータからの読み込み（Web APIなど）

```python
# FastAPIでのアップロード処理例
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
