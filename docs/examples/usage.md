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

```typescript
const schema = {
  name: 'B1',
  age: 'B2',
  email: 'B3',
  department: 'B4',
};
```

**出力:**

```json
{
  "name": "山田太郎",
  "age": 30,
  "email": "yamada@example.com",
  "department": "営業部"
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

```typescript
const schema = {
  'A1': 'B1',
  'A2': 'B2',
  'A3': 'B3',
};
```

**出力:**

```json
{
  "name": "山田太郎",
  "age": 30,
  "email": "yamada@example.com"
}
```

---

## 例3: ネストしたオブジェクト（住所情報）

**Excelレイアウト:**

| | A | B |
|-|---|---|
| 1 | 氏名 | 山田太郎 |
| 2 | 都道府県 | 東京都 |
| 3 | 市区町村 | 渋谷区 |
| 4 | 郵便番号 | 150-0002 |
| 5 | 電話番号 | 03-1234-5678 |

**スキーマ:**

```typescript
const schema = {
  name: 'B1',
  address: {
    prefecture: 'B2',
    city: 'B3',
    zip: 'B4',
  },
  phone: 'B5',
};
```

**出力:**

```json
{
  "name": "山田太郎",
  "address": {
    "prefecture": "東京都",
    "city": "渋谷区",
    "zip": "150-0002"
  },
  "phone": "03-1234-5678"
}
```

---

## 例4: 配列（品目リスト）

**Excelレイアウト:**

| | A |
|-|---|
| 1 | りんご |
| 2 | みかん |
| 3 | ぶどう |
| 4 | もも |
| 5 | なし |

**スキーマ:**

```typescript
const schema = {
  fruits: ['A1:A5'],
};
```

**出力:**

```json
{
  "fruits": ["りんご", "みかん", "ぶどう", "もも", "なし"]
}
```

---

## 例5: オブジェクト配列（商品テーブル）

**Excelレイアウト:**

| | A | B | C |
|-|---|---|---|
| 1 | id | name | price |
| 2 | 1 | りんご | 100 |
| 3 | 2 | みかん | 80 |
| 4 | 3 | ぶどう | 200 |

**スキーマ（ヘッダー行を除いたA2:C4をオブジェクト配列に変換）:**

```typescript
const schema = {
  products: {
    $range: 'A2:C4',
    $schema: {
      id: 0,
      name: 1,
      price: 2,
    },
  },
};
```

**出力:**

```json
{
  "products": [
    { "id": 1, "name": "りんご", "price": 100 },
    { "id": 2, "name": "みかん", "price": 80 },
    { "id": 3, "name": "ぶどう", "price": 200 }
  ]
}
```

---

## 例6: 複数シートのデータを結合

**Excelのシート構成:**
- `顧客情報` シート: 顧客の基本データ
- `注文情報` シート: 注文データ

**スキーマ:**

```typescript
const schema = {
  customer: {
    id: '顧客情報!B1',
    name: '顧客情報!B2',
    email: '顧客情報!B3',
  },
  order: {
    orderId: '注文情報!B1',
    date: '注文情報!B2',
    total: '注文情報!B5',
  },
};
```

**出力:**

```json
{
  "customer": {
    "id": "C-001",
    "name": "山田太郎",
    "email": "yamada@example.com"
  },
  "order": {
    "orderId": "ORD-2024-001",
    "date": "2024-01-15T00:00:00.000Z",
    "total": 15000
  }
}
```

---

## 例7: カスタムトランスフォーマー

セル値を取得後に加工したい場合はトランスフォーマーを使います。

```typescript
import { ExcelMapper } from 'excel-cell-mapper';

const mapper = new ExcelMapper('./data.xlsx', {
  transform: (value, { cellRef }) => {
    // 数値は小数点2桁に丸める
    if (typeof value === 'number') {
      return Math.round(value * 100) / 100;
    }
    // 文字列は前後の空白を除去
    if (typeof value === 'string') {
      return value.trim();
    }
    return value;
  },
});

const result = await mapper.map({ price: 'B3' });
```

---

## 例8: Node.jsでのファイル読み込み

```typescript
import { ExcelMapper } from 'excel-cell-mapper';

async function processOrderForm(filePath: string) {
  const mapper = await ExcelMapper.fromFile(filePath, {
    defaultSheet: '注文書',
    emptyCell: 'omit',
    dateFormat: 'iso',
  });

  const schema = {
    orderId: 'B1',
    customer: {
      name: 'B3',
      address: 'B4',
      phone: 'B5',
    },
    items: {
      $range: 'A9:D20',
      $schema: { name: 0, quantity: 1, unitPrice: 2, subtotal: 3 },
      $skipEmpty: true,
    },
    totalAmount: 'D22',
  };

  return await mapper.map(schema);
}
```

---

## 例9: ブラウザ環境（File API経由）

```typescript
import { ExcelMapper } from 'excel-cell-mapper';

async function handleFileUpload(file: File) {
  const buffer = await file.arrayBuffer();
  const mapper = new ExcelMapper(buffer);

  const schema = {
    title: 'A1',
    data: {
      $range: 'A3:B10',
      $schema: { label: 0, value: 1 },
    },
  };

  const result = await mapper.map(schema);
  console.log(result);
}

// HTML input[type=file] と連携
document.getElementById('upload')?.addEventListener('change', (e) => {
  const file = (e.target as HTMLInputElement).files?.[0];
  if (file) handleFileUpload(file);
});
```
