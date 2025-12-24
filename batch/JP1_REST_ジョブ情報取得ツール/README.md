# JP1 REST API ジョブ情報取得ツール

JP1/AJS3 Web Console REST APIを使用して、ジョブ/ジョブネットの状態情報を取得するツールです。

## 概要

以下の処理を実行します：

1. **ユニット一覧取得** - statuses APIでユニット状態とexecIDを取得
2. **実行結果詳細取得** - execResultDetails APIで標準エラー出力を取得

## 必要な環境

- Windows（PowerShell 5.1以降）
- JP1/AJS3 - Web Console がインストール・起動されていること
- Web Console への接続権限（JP1ユーザー）

## 特徴

| 項目 | 説明 |
|------|------|
| 実行方式 | REST API（Web Console経由） |
| 管理者権限 | 不要 |
| WinRM | 不要 |
| 取得情報 | ユニット状態、execID、標準エラー出力 |

## 使い方

### 1. 設定を編集

`JP1_REST_ジョブ情報取得ツール.bat` の設定セクションを編集します：

```powershell
# Web Consoleサーバーのホスト名またはIPアドレス
$webConsoleHost = "localhost"

# Web Consoleのポート番号（HTTP: 22252, HTTPS: 22253）
$webConsolePort = "22252"

# HTTPSを使用する場合は $true に設定
$useHttps = $false

# JP1/AJS3 Managerのホスト名
$managerHost = "localhost"

# スケジューラーサービス名
$schedulerService = "AJSROOT1"

# JP1ユーザー名
$jp1User = "jp1admin"

# JP1パスワード
$jp1Password = "password"

# 取得対象のユニットパス
$unitPath = "/main_unit/jobgroup1/daily_batch"
```

### 2. 実行

バッチファイルをダブルクリックで実行します。

### 3. 実行例

```
================================================================
  JP1 REST API ジョブ情報取得ツール
================================================================

設定内容:
  Web Consoleサーバー : localhost:22252
  Managerホスト       : localhost
  スケジューラー      : AJSROOT1
  JP1ユーザー         : jp1admin
  ユニットパス        : /main_unit/jobgroup1/daily_batch

================================================================
STEP 1: ユニット一覧取得API（execID取得）
================================================================

リクエストURL:
  http://localhost:22252/ajs/api/v1/objects/statuses?manager=localhost&serviceName=AJSROOT1&location=/main_unit/jobgroup1/daily_batch&mode=search

[OK] HTTPステータス: 200

取得したユニット一覧:
  パス: /main_unit/jobgroup1/daily_batch/job1 | execID: @A001 | 状態: 正常終了

================================================================
STEP 2: 実行結果詳細取得API（execResultDetails）
================================================================

対象: /main_unit/jobgroup1/daily_batch/job1 (execID: @A001)

[OK] HTTPステータス: 200

実行結果詳細（標準エラー出力）:
----------------------------------------
(出力なし)
----------------------------------------

================================================================
処理完了
================================================================
```

## 設定項目

| 設定項目 | 説明 | デフォルト値 |
|---------|------|-------------|
| `$webConsoleHost` | Web Consoleサーバーのホスト名 | `localhost` |
| `$webConsolePort` | Web Consoleのポート番号 | `22252` |
| `$useHttps` | HTTPS使用フラグ | `$false` |
| `$managerHost` | JP1/AJS3 Managerのホスト名 | `localhost` |
| `$schedulerService` | スケジューラーサービス名 | `AJSROOT1` |
| `$jp1User` | JP1ユーザー名 | `jp1admin` |
| `$jp1Password` | JP1パスワード | - |
| `$unitPath` | 取得対象のユニットパス | - |
| `$debugMode` | デバッグモード | `$true` |

## 使用API

| API | 用途 |
|-----|------|
| ユニット一覧取得API (7.1.1) | ユニット状態・execID取得 |
| 実行結果詳細取得API (7.1.3) | 標準エラー出力取得 |

詳細は [JP1_AJS3_REST_API.md](../../JP1_AJS3_REST_API.md) を参照してください。

## 注意事項

- JP1/AJS3 - Web Console が必要です
- execResultDetails API は標準エラー出力相当の情報を取得します
- 標準出力の取得には ajsshow コマンド（WinRM経由）が必要です
- パスワードはスクリプト内に平文で記載されます（セキュリティに注意）

## 関連ツール

| ツール | 説明 |
|--------|------|
| [JP1 ジョブツール](../JP1_ジョブツール/) | ajsshowコマンドでジョブ情報取得 |

## トラブルシューティング

### 接続できない

- Web Consoleサービスが起動しているか確認してください
- ファイアウォールでポートがブロックされていないか確認してください
- ホスト名・ポート番号が正しいか確認してください

### 認証エラー

- JP1ユーザー名・パスワードが正しいか確認してください
- JP1ユーザーに参照権限があるか確認してください

### ユニットが取得できない

- ユニットパスが正しいか確認してください（`/`から始まるフルパス）
- 実行履歴が存在するか確認してください

## ライセンス

MIT License
