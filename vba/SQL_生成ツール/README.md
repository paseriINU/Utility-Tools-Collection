# Oracle SELECT文生成ツール

Excel VBAを使用して、Oracle用のSELECT文を対話的に生成するツールです。SQL初心者でも複雑なクエリを簡単に作成できます。

## 機能

### 基本機能
- **SELECT句**: 取得するカラムの指定、別名(AS)の設定
- **FROM句**: メインテーブルの指定
- **WHERE句**: 抽出条件の設定（AND/OR、括弧による優先順位）
- **ORDER BY句**: 並び順の指定（昇順/降順、NULLS位置）
- **DISTINCT**: 重複行の除外

### 結合（JOIN）
- INNER JOIN（内部結合）
- LEFT JOIN（左外部結合）
- RIGHT JOIN（右外部結合）
- FULL OUTER JOIN（完全外部結合）
- CROSS JOIN（交差結合）

### 集計機能
- **集計関数**: COUNT, SUM, AVG, MAX, MIN, COUNT(DISTINCT)
- **GROUP BY句**: グループ化
- **HAVING句**: 集計結果の絞り込み

### 高度な機能
- **サブクエリ**: SELECT句やWHERE句内でのサブクエリ使用
- **WITH句**: 共通テーブル式（CTE）の定義
- **UNION / UNION ALL**: 複数SELECT結果の結合
- **件数制限**: FETCH FIRST / ROWNUM方式

### 学習支援機能
- **SQLヘルプシート**: 各SQL構文の説明と使用例
- **各セクションの説明**: メインシートに各項目の簡単な説明を表示
- **サンプルデータ**: テーブル定義シートにサンプルテーブル・カラムを登録済み

## 必要な環境

- Microsoft Excel 2010以降
- マクロを有効にする必要があります

## インストール方法

### 方法1: VBAエディタでインポート
1. Excelを開き、`Alt + F11`でVBAエディタを起動
2. 「ファイル」→「ファイルのインポート」を選択
3. `SQLGenerator.bas`ファイルを選択してインポート
4. VBAエディタを閉じる
5. ファイルを`.xlsm`（マクロ有効ブック）として保存

### 方法2: 新規ブックにコードを貼り付け
1. 新規Excelブックを作成
2. `Alt + F11`でVBAエディタを起動
3. 「挿入」→「標準モジュール」を選択
4. `SQLGenerator.bas`の内容を全てコピー＆ペースト
5. ファイルを`.xlsm`として保存

## 使い方

### 初期設定
1. Excelで`Alt + F8`を押してマクロダイアログを開く
2. `InitializeSQLGenerator`を選択して「実行」
3. 必要なシートが自動的に作成されます

### シート構成

| シート名 | 説明 |
|---------|------|
| メイン | SQL生成の入力画面。条件を入力してSQLを生成 |
| テーブル定義 | 使用するテーブルとカラムの一覧を登録 |
| 生成履歴 | 過去に生成したSQLを保存 |
| サブクエリ | サブクエリの定義 |
| WITH句 | 共通テーブル式（CTE）の定義 |
| UNION | UNION結合するSQLの定義 |
| SQLヘルプ | SQL構文の説明と使用例 |

### 基本的な使い方

#### 1. テーブル定義の登録（任意）
「テーブル定義」シートに、使用するテーブルとカラムの情報を登録します。
これにより、メインシートでの入力が楽になります。

#### 2. メインシートでの入力

**メインテーブルの指定**
```
テーブル名: USERS
別名: u
```

**結合テーブルの指定（必要な場合）**
```
結合種別: INNER JOIN
テーブル名: ORDERS
別名: o
結合条件: u.USER_ID = o.USER_ID
```

**取得カラムの指定**
```
テーブル別名: u
カラム名: USER_NAME
別名(AS): ユーザー名
```

**抽出条件の指定**
```
カラム名: STATUS
演算子: =
値: 1
```

#### 3. SQL生成
「SQL生成」ボタンをクリックすると、入力内容に基づいてSQLが生成されます。

### 生成例

#### 基本的なSELECT
```sql
SELECT u.USER_ID,
       u.USER_NAME,
       u.EMAIL
FROM USERS u
WHERE u.STATUS = 1
ORDER BY u.USER_ID ASC;
```

#### JOINと集計を含むSELECT
```sql
SELECT u.USER_ID,
       u.USER_NAME,
       COUNT(*) AS 注文件数,
       SUM(o.AMOUNT) AS 合計金額
FROM USERS u
INNER JOIN ORDERS o ON u.USER_ID = o.USER_ID
WHERE u.STATUS = 1
  AND o.ORDER_DATE >= SYSDATE - 30
GROUP BY u.USER_ID, u.USER_NAME
HAVING SUM(o.AMOUNT) > 10000
ORDER BY SUM(o.AMOUNT) DESC;
```

#### WITH句とサブクエリを含むSELECT
```sql
WITH active_users AS (
    SELECT USER_ID, USER_NAME FROM USERS WHERE STATUS = 1
)
SELECT a.USER_ID,
       a.USER_NAME,
       (
           SELECT COUNT(*) FROM ORDERS WHERE USER_ID = a.USER_ID
       ) AS 注文件数
FROM active_users a
WHERE EXISTS (
    SELECT 1 FROM ORDERS WHERE USER_ID = a.USER_ID
);
```

## SQL構文クイックリファレンス

### JOIN（結合）の種類

| 種別 | 説明 |
|-----|------|
| INNER JOIN | 両方のテーブルに存在するデータのみ取得 |
| LEFT JOIN | 左テーブルの全データ＋右テーブルの一致データ |
| RIGHT JOIN | 右テーブルの全データ＋左テーブルの一致データ |
| FULL OUTER JOIN | 両方のテーブルの全データ |
| CROSS JOIN | 全ての組み合わせ（直積） |

### WHERE句の演算子

| 演算子 | 説明 | 例 |
|-------|------|-----|
| = | 等しい | STATUS = 1 |
| <> | 等しくない | STATUS <> 0 |
| >, <, >=, <= | 大小比較 | AMOUNT > 1000 |
| LIKE | パターン一致 | NAME LIKE '%田中%' |
| IN | リスト内に存在 | STATUS IN (1, 2, 3) |
| BETWEEN | 範囲内 | AGE BETWEEN 20 AND 30 |
| IS NULL | NULLかどうか | DELETE_DATE IS NULL |
| EXISTS | サブクエリに結果が存在 | EXISTS (SELECT 1 FROM ...) |

### 集計関数

| 関数 | 説明 |
|-----|------|
| COUNT(*) | 行数をカウント |
| COUNT(カラム) | NULL以外の件数 |
| COUNT(DISTINCT カラム) | 重複を除いた件数 |
| SUM(カラム) | 合計値 |
| AVG(カラム) | 平均値 |
| MAX(カラム) | 最大値 |
| MIN(カラム) | 最小値 |

## 注意事項

- 生成されたSQLは参考用です。実行前に必ず内容を確認してください
- 件数制限でROWNUM方式を選択した場合、WHERE句に手動で追加が必要です
- サブクエリを使用する場合は、先に「サブクエリ」シートに定義を登録してください
- WITH句やUNIONを使用する場合は、メインシートのオプションで「使用する」を選択してください

## トラブルシューティング

### マクロが実行できない
- Excelのセキュリティ設定でマクロが無効になっている可能性があります
- 「ファイル」→「オプション」→「セキュリティセンター」→「マクロの設定」で確認してください

### シートが作成されない
- `InitializeSQLGenerator`マクロを実行してください
- エラーが発生する場合は、他のシートを閉じて再度実行してください

### 生成されたSQLにエラーがある
- 入力内容を確認してください
- 特にJOIN条件やGROUP BY句の指定が正しいか確認してください

## ファイル一覧

```
vba/SQL_生成ツール/
├── SQLGenerator.bas    # VBAモジュール
└── README.md           # このファイル
```

## ライセンス

このツールは自由に使用・改変できます。
