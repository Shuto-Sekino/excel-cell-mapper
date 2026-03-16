# APIリファレンス

> ステータス: 下書き v0.1

## インストール

```bash
npm install excel-cell-mapper
```

---

## 基本的な使い方

```typescript
import { ExcelMapper } from 'excel-cell-mapper';

const mapper = new ExcelMapper('./data.xlsx');

const schema = {
  name: 'B1',
  age: 'B2',
  email: 'B3',
};

const result = await mapper.map(schema);
// => { name: '山田太郎', age: 30, email: 'yamada@example.com' }
```

---

## `ExcelMapper` クラス

### コンストラクタ

```typescript
new ExcelMapper(source: ExcelSource, options?: MapperOptions)
```

#### `ExcelSource`

```typescript
type ExcelSource =
  | string          // ファイルパス
  | Buffer          // Bufferオブジェクト
  | ArrayBuffer     // ArrayBuffer
  | Uint8Array;     // バイナリデータ
```

#### `MapperOptions`

```typescript
interface MapperOptions {
  /**
   * デフォルトで参照するシート名またはインデックス（0始まり）
   * 未指定の場合は最初のシートを使用
   */
  defaultSheet?: string | number;

  /**
   * 空白セルの値をどう扱うか
   * - 'null'    : null を返す（デフォルト）
   * - 'omit'    : そのフィールドをJSONから除外
   * - 'empty'   : 空文字列 "" を返す
   */
  emptyCell?: 'null' | 'omit' | 'empty';

  /**
   * 日付セルの変換方法
   * - 'date'    : Date オブジェクト（デフォルト）
   * - 'iso'     : ISO 8601 文字列 "2024-01-15T00:00:00.000Z"
   * - 'local'   : ロケール依存の文字列
   */
  dateFormat?: 'date' | 'iso' | 'local';

  /**
   * セル値を変換するカスタムトランスフォーマー
   */
  transform?: CellTransformer;
}
```

---

### `mapper.map(schema, options?)`

スキーマに基づいてExcelをJSONに変換します。

```typescript
mapper.map<T = Record<string, unknown>>(
  schema: Schema,
  options?: MapOptions
): Promise<T>
```

#### `MapOptions`

```typescript
interface MapOptions {
  /**
   * このmap呼び出し専用のシート指定
   * MapperOptionsのdefaultSheetを上書きする
   */
  sheet?: string | number;
}
```

#### 戻り値

`Promise<T>` — スキーマに従って変換されたJSONオブジェクト

---

### `mapper.mapMany(schemas, options?)`

複数のスキーマをまとめて変換します。

```typescript
mapper.mapMany<T extends Record<string, Schema>>(
  schemas: T,
  options?: MapOptions
): Promise<{ [K in keyof T]: unknown }>
```

**使用例:**

```typescript
const result = await mapper.mapMany({
  customer: { name: '顧客情報!B1', id: '顧客情報!B2' },
  order: { date: '注文情報!C1', total: '注文情報!C5' },
});
// => { customer: { name: '...', id: '...' }, order: { date: ..., total: ... } }
```

---

### `mapper.getCell(cellRef)`

単一セルの値を直接取得します。

```typescript
mapper.getCell(cellRef: string): Promise<CellValue>
```

```typescript
const value = await mapper.getCell('Sheet1!B3');
// => '山田太郎'
```

---

### `mapper.getRange(rangeRef)`

セル範囲の値を2次元配列で取得します。

```typescript
mapper.getRange(rangeRef: string): Promise<CellValue[][]>
```

```typescript
const values = await mapper.getRange('A1:C3');
// => [['a', 'b', 'c'], ['d', 'e', 'f'], ['g', 'h', 'i']]
```

---

### `mapper.getSheetNames()`

Excelファイル内のシート名一覧を取得します。

```typescript
mapper.getSheetNames(): Promise<string[]>
```

---

## スタティックAPI

### `ExcelMapper.from(source, options?)`

コンストラクタの代わりにファクトリメソッドで生成できます。

```typescript
const mapper = ExcelMapper.from(buffer, { defaultSheet: 'データ' });
```

### `ExcelMapper.fromFile(filePath, options?)`

ファイルパスからMapperを生成します（Node.js専用）。

```typescript
const mapper = await ExcelMapper.fromFile('./template.xlsx');
```

---

## 型定義

### `Schema`

```typescript
type Schema =
  | string                          // セル参照 "B1", "Sheet1!A2"
  | string[]                        // 範囲参照 ["A1:A10"]
  | { [key: string]: Schema }       // オブジェクト（ネスト可）
  | RangeSchema;                    // レンジスキーマ（オブジェクト配列）
```

### `RangeSchema`

```typescript
interface RangeSchema {
  $range: string;                          // 範囲参照 "A2:C10"
  $schema: { [fieldName: string]: number | string }; // 列インデックスまたは列名
  $direction?: 'row' | 'column';           // 走査方向（デフォルト: 'row'）
  $skipEmpty?: boolean;                    // 空行をスキップ
}
```

### `CellValue`

```typescript
type CellValue = string | number | boolean | Date | null;
```

### `CellTransformer`

```typescript
type CellTransformer = (
  value: CellValue,
  context: {
    cellRef: string;      // セル参照 "B1"
    sheetName: string;    // シート名
    colIndex: number;     // 列インデックス（0始まり）
    rowIndex: number;     // 行インデックス（0始まり）
  }
) => unknown;
```

---

## エラーハンドリング

```typescript
import { ExcelMapper, ExcelMapperError, CellNotFoundError, InvalidSchemaError } from 'excel-cell-mapper';

try {
  const result = await mapper.map(schema);
} catch (err) {
  if (err instanceof CellNotFoundError) {
    console.error(`セル ${err.cellRef} が見つかりません`);
  } else if (err instanceof InvalidSchemaError) {
    console.error(`スキーマが不正です: ${err.message}`);
  } else if (err instanceof ExcelMapperError) {
    console.error(`マッパーエラー: ${err.message}`);
  }
}
```

### エラークラス一覧

| クラス | 説明 |
|--------|------|
| `ExcelMapperError` | 基底エラークラス |
| `CellNotFoundError` | 指定したセルが存在しない |
| `SheetNotFoundError` | 指定したシートが存在しない |
| `InvalidSchemaError` | スキーマの記法が不正 |
| `ParseError` | Excelファイルのパースに失敗 |

---

## TypeScriptによる型推論（将来拡張）

スキーマから戻り値の型を自動推論する機能（v1.0以降での対応を検討）。

```typescript
const schema = {
  name: 'B1',
  age: 'B2',
} as const;

// result の型が { name: CellValue, age: CellValue } に推論される（検討中）
const result = await mapper.map(schema);
```
