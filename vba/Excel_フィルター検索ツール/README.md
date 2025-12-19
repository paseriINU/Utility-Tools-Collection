# Excel フィルター検索ツール

複数キーワードでOR条件のフィルター検索を行うVBAツールです。

## 機能

- ユーザーフォームで最大5つのキーワードを入力
- A列・B列に対してOR条件でフィルター適用
- 部分一致検索（ワイルドカード使用）
- フィルタークリア機能
- Enterキーで検索実行可能

## セットアップ手順

### 1. VBAエディタを開く

1. Excelファイルを開く
2. `Alt + F11` でVBAエディタを開く

### 2. 標準モジュールをインポート

1. VBAエディタのメニューから「ファイル」→「ファイルのインポート」
2. `FilterSearch.bas` を選択してインポート

### 3. ユーザーフォームを作成

VBAエディタで以下の手順でフォームを作成します：

1. **フォームを挿入**
   - メニュー「挿入」→「ユーザーフォーム」
   - プロパティウィンドウで `(オブジェクト名)` を `FilterSearchForm` に変更
   - `Caption` を `フィルター検索` に変更

2. **コントロールを配置**

   以下のコントロールを配置してください：

   | コントロール | 名前 | Caption/用途 |
   |-------------|------|-------------|
   | Label | lblTitle | 検索キーワード（OR条件） |
   | Label | lblWord1 | キーワード1: |
   | Label | lblWord2 | キーワード2: |
   | Label | lblWord3 | キーワード3: |
   | Label | lblWord4 | キーワード4: |
   | Label | lblWord5 | キーワード5: |
   | TextBox | txtWord1 | （入力欄1） |
   | TextBox | txtWord2 | （入力欄2） |
   | TextBox | txtWord3 | （入力欄3） |
   | TextBox | txtWord4 | （入力欄4） |
   | TextBox | txtWord5 | （入力欄5） |
   | CommandButton | btnSearch | 検索 |
   | CommandButton | btnClear | クリア |
   | CommandButton | btnClose | 閉じる |

3. **フォームのコードを貼り付け**
   - フォームをダブルクリックしてコードウィンドウを開く
   - `FilterSearchForm.frm` の `Option Explicit` 以降のコードをコピー＆ペースト

### 4. ボタンを配置（任意）

シート上にボタンを配置する場合：

1. 「開発」タブ →「挿入」→「ボタン（フォームコントロール）」
2. シート上にボタンを描画
3. マクロの登録で `ShowFilterSearchForm` を選択

## 設定変更

`FilterSearch.bas` の以下の定数を変更することで、対象シートやフィルター列を変更できます：

```vba
Private Const TARGET_SHEET_NAME As String = "テスト"    ' 対象シート名
Private Const FILTER_COLUMN_A As Long = 1               ' フィルター列1（A列=1）
Private Const FILTER_COLUMN_B As Long = 2               ' フィルター列2（B列=2）
```

## 使い方

1. ボタンを押す（または `ShowFilterSearchForm` マクロを実行）
2. フォームが表示される
3. キーワードを1つ以上入力（空欄は無視されます）
4. 「検索」ボタンをクリック（またはEnterキー）
5. フィルターが適用される

### クリア

「クリア」ボタンをクリックすると：
- フィルターが解除される
- 入力欄がクリアされる

## 注意事項

- 対象シート「テスト」が存在しない場合はエラーメッセージが表示されます
- 3つ以上のキーワードを指定した場合は、AdvancedFilterを使用します
- マクロ有効ブック（.xlsm）として保存してください

## ファイル構成

```
Excel_フィルター検索ツール/
├── FilterSearch.bas         # 標準モジュール（メイン処理）
├── FilterSearchForm.frm     # ユーザーフォーム（コード）
└── README.md               # このファイル
```
