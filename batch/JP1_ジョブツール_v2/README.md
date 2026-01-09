# JP1 ジョブツール v2

JP1/AJS3 Web Console REST APIを使用して、ジョブの実行とログ取得を行うツールです。

## v1からの変更点

- **出力オプション引数を追加**: メインツールに出力方法を指定可能
- **標準ログバッチを簡素化**: ジョブパスと出力オプションを設定するだけ
- **出力先を統一**: `..\02.Output\` フォルダに自動出力

## 概要

2つのメインツールを提供します：

1. **JP1ジョブ実行.bat** - ジョブを即時実行してログを取得
2. **JP1ジョブログ取得.bat** - 実行せずに既存の結果を取得

## ファイル構成

| ファイル | 説明 |
|---------|------|
| `JP1ジョブ実行.bat` | メインツール（実行+ログ取得） |
| `JP1ジョブログ取得.bat` | メインツール（ログ取得のみ） |
| `JP1ジャーナル取得.bat` | 標準ログバッチ（EXEC/GET切替対応） |
| `【JP1ジョブ実行】ジョブ.bat` | 標準ログバッチ（実行専用） |
| `【JP1ジョブログ取得】ジョブ.bat` | 標準ログバッチ（取得専用） |

## 使い方

### 標準ログバッチを使用する場合

1. 標準ログバッチ（例：`【JP1ジョブログ取得】ジョブ.bat`）を開く

2. ジョブパスと出力オプションを編集：

```batch
rem ジョブパス（必須）
set "UNIT_PATH=/JobGroup/Jobnet/Job1"

rem 出力オプション
set "OUTPUT_MODE=/NOTEPAD"
```

3. バッチをダブルクリックで実行

### メインツールを直接使用する場合

```batch
rem 実行してログ取得（メモ帳で開く）
JP1ジョブ実行.bat "/JobGroup/Jobnet/Job1" /NOTEPAD

rem ログ取得のみ（ファイル出力のみ）
JP1ジョブログ取得.bat "/JobGroup/Jobnet/Job1" "" /LOG

rem 2つのジョブを比較（メモ帳で開く）
JP1ジョブログ取得.bat "/JobGroup/Jobnet/Job1" "/JobGroup/Jobnet/Job2" /NOTEPAD
```

## 出力オプション

| オプション | 説明 |
|-----------|------|
| `/LOG` | ログファイル出力のみ（デフォルト） |
| `/NOTEPAD` | メモ帳で開く |
| `/EXCEL` | Excelに貼り付け |
| `/WINMERGE` | WinMergeで比較（TODO: 実装予定） |

### Excel貼り付け設定（/EXCELオプション使用時）

`/EXCEL`オプションを使用する場合は、標準ログバッチで以下の環境変数を設定してください：

```batch
rem Excelファイル名（バッチと同じフォルダに配置）
set "EXCEL_FILE_NAME=ログ貼り付け用.xlsx"

rem 貼り付け先シート名
set "EXCEL_SHEET_NAME=Sheet1"

rem 貼り付け先セル位置
set "EXCEL_PASTE_CELL=A1"
```

## 出力先

```
親フォルダ/
├── 01.Batch/           <- バッチファイルの配置場所
│   ├── JP1ジョブ実行.bat
│   ├── JP1ジョブログ取得.bat
│   └── ...
└── 02.Output/          <- 出力先（自動作成）
    └── 【ジョブ実行結果】【YYYYMMDD_HHMMSS実行分】【終了状態】ジョブネット名_コメント.txt
```

## 必要な環境

- Windows（PowerShell 5.1以降）
- JP1/AJS3 - Web Console（REST API）

## 設定項目

メインツール（JP1ジョブ実行.bat / JP1ジョブログ取得.bat）の設定セクションを編集：

### 接続設定

| 設定項目 | 説明 | デフォルト値 |
|---------|------|-------------|
| `$webConsoleHost` | Web Consoleサーバーのホスト名 | `localhost` |
| `$webConsolePort` | Web Consoleのポート番号 | `22252` |
| `$schedulerService` | スケジューラーサービス名 | `AJSROOT1` |

### 認証設定

事前に資格情報を登録する場合：
```cmd
cmdkey /generic:JP1_WebConsole /user:jp1admin /pass:yourpassword
```

## 終了コード

| コード | 説明 |
|-------|------|
| 0 | 正常終了 |
| 1 | 引数エラー |
| 2 | ユニット未検出 |
| 3 | ユニット種別エラー |
| 4 | ルートジョブネット特定エラー / 実行世代なし |
| 5 | 即時実行登録エラー / 5MB超過エラー |
| 6 | タイムアウト / 詳細取得エラー |
| 8 | 詳細取得エラー / 比較モードで両方取得失敗 |
| 9 | API接続エラー |

## ライセンス

MIT License
