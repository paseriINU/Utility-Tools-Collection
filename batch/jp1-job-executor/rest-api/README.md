# JP1ジョブネット起動ツール - REST API版

JP1/AJS3のREST APIを使用してジョブネットを起動するPowerShellツールです。

---

## 📋 概要

このツールは、JP1/AJS3 - Manager バージョン10以降で提供されるREST APIを使用して、
HTTP/HTTPS経由でジョブネットを起動します。

---

## ✅ 必要な環境

### ローカルPC（実行する側）

- Windows OS
- PowerShell 5.1以降
- **JP1のインストール不要**

### JP1サーバ（リモート側）

- **JP1/AJS3 - Manager バージョン10以降**
- REST APIサービスが有効
- ネットワークでポート22250（デフォルト）が開放

---

## 🔧 事前準備

### JP1サーバ側の設定

JP1/AJS3のREST APIを有効化する必要があります。

#### 1. REST APIサービスの確認

JP1サーバで以下のサービスが起動しているか確認：

```
サービス名: JP1/AJS3 Web Application Server
```

#### 2. ファイアウォール設定

ポート22250（デフォルト）を開放：

```powershell
# ファイアウォールルールを追加
New-NetFirewallRule -DisplayName "JP1 REST API" -Direction Inbound -Protocol TCP -LocalPort 22250 -Action Allow
```

#### 3. REST APIの動作確認

ブラウザで以下にアクセス（ローカルPCから）：

```
http://192.168.1.100:22250/ajs3web/
```

**正常な場合**: JP1のログイン画面が表示されます

---

## 🚀 使い方

### 方法1: PowerShellスクリプト直接実行

#### 1. PowerShellスクリプトを実行

```powershell
.\Start-JP1Job.ps1 `
    -JP1Host "192.168.1.100" `
    -JP1User "jp1admin" `
    -JobnetPath "/main_unit/jobgroup1/daily_batch"
```

#### 2. パスワード入力

実行時にJP1パスワードの入力を求められます。

#### オプション指定例

```powershell
# HTTPSを使用する場合
.\Start-JP1Job.ps1 `
    -JP1Host "192.168.1.100" `
    -JP1Port 22250 `
    -JP1User "jp1admin" `
    -JobnetPath "/main_unit/jobgroup1/daily_batch" `
    -UseSSL

# パスワードを事前に渡す場合（非推奨）
$securePassword = ConvertTo-SecureString "password123" -AsPlainText -Force
.\Start-JP1Job.ps1 `
    -JP1Host "192.168.1.100" `
    -JP1User "jp1admin" `
    -JP1Password $securePassword `
    -JobnetPath "/main_unit/jobgroup1/daily_batch"
```

---

### 方法2: バッチファイル経由（簡単）

#### 1. バッチファイルを編集

`jp1_start_api.bat`をテキストエディタで開き、設定を編集：

```batch
set JP1_HOST=192.168.1.100
set JP1_PORT=22250
set JP1_USER=jp1admin
set JOBNET_PATH=/main_unit/jobgroup1/daily_batch
set USE_SSL=false
```

#### 2. 実行

```cmd
jp1_start_api.bat
```

実行時にJP1パスワードの入力を求められます。

---

## 📖 実行例

### 成功時の出力

```
========================================
JP1ジョブネット起動（REST API版）
========================================

JP1ホスト      : 192.168.1.100
ポート番号      : 22250
プロトコル      : http
JP1ユーザー    : jp1admin
ジョブネットパス: /main_unit/jobgroup1/daily_batch

認証トークンを取得中...
認証成功

ジョブネットを起動中...

========================================
ジョブネットの起動に成功しました
========================================

ジョブネット: /main_unit/jobgroup1/daily_batch
実行ID      : 20251202123456-0001
ホスト      : 192.168.1.100

ログアウト中...
ログアウト完了

処理が完了しました。
```

---

## ⚙️ パラメータ詳細

### Start-JP1Job.ps1のパラメータ

| パラメータ | 必須 | 説明 | デフォルト |
|-----------|------|------|-----------|
| JP1Host | ✅ | JP1サーバのホスト名/IP | - |
| JP1Port | ❌ | REST APIのポート番号 | 22250 |
| JP1User | ✅ | JP1ユーザー名 | - |
| JP1Password | ❌ | JP1パスワード（SecureString） | 実行時入力 |
| JobnetPath | ✅ | ジョブネットのフルパス | - |
| UseSSL | ❌ | HTTPS接続を使用 | false |

---

## 🔐 認証の仕組み

REST API版では以下の流れで認証します：

1. **ログイン** (`/ajs3web/api/auth/login`)
   - JP1ユーザー名とパスワードで認証
   - 認証トークンを取得

2. **ジョブネット起動** (`/ajs3web/api/jobnets/{path}/executions`)
   - 認証トークンをヘッダーに含めて送信
   - ジョブネット起動要求

3. **ログアウト** (`/ajs3web/api/auth/logout`)
   - 認証トークンを無効化

---

## 🌐 REST APIエンドポイント

### ベースURL

```
http://{JP1_HOST}:{JP1_PORT}/ajs3web/api
```

### 使用するエンドポイント

| エンドポイント | メソッド | 用途 |
|---------------|---------|------|
| `/auth/login` | POST | 認証トークン取得 |
| `/jobnets/{path}/executions` | POST | ジョブネット起動 |
| `/auth/logout` | POST | ログアウト |

---

## ⚠️ 注意事項

### セキュリティ

- **HTTPS推奨**: 本番環境では`-UseSSL`を使用
- **パスワード**: スクリプトにパスワードを直接記載しない
- **証明書**: 自己署名証明書の場合、証明書検証を無効化しています

### REST API制限

- **バージョン要件**: JP1/AJS3 - Manager バージョン10以降
- **同時接続数**: REST APIの同時接続数に制限あり
- **タイムアウト**: 長時間実行されるジョブネットには不向き

### JP1環境

- **実行権限**: JP1ユーザーにジョブネット実行権限が必要
- **本番環境**: 本番実行前に必ずテスト環境で確認

---

## 🐛 トラブルシューティング

### エラー: 「認証に失敗しました」

**原因1**: REST APIサービスが起動していない

**対処法**:
```
JP1サーバでサービスを確認：
サービス → JP1/AJS3 Web Application Server → 開始
```

---

**原因2**: ポート番号が間違っている

**対処法**:
```
デフォルト: 22250
確認方法: JP1/AJS3の設定ファイルを確認
```

---

**原因3**: ユーザー名またはパスワードが間違っている

**対処法**:
```
JP1ユーザー情報を確認
```

---

### エラー: 「接続できません」

**原因1**: ネットワーク接続の問題

**対処法**:
```powershell
# ローカルPCから確認
Test-NetConnection -ComputerName 192.168.1.100 -Port 22250
```

**正常な場合の出力**:
```
TcpTestSucceeded : True
```

---

**原因2**: ファイアウォールでブロックされている

**対処法**（JP1サーバで実行）:
```powershell
# ファイアウォールルールを確認
Get-NetFirewallRule | Where-Object {$_.LocalPort -eq 22250}

# ルールがない場合は追加
New-NetFirewallRule -DisplayName "JP1 REST API" -Direction Inbound -Protocol TCP -LocalPort 22250 -Action Allow
```

---

### エラー: 「ジョブネットの起動に失敗しました」

**原因1**: ジョブネットパスが間違っている

**対処法**:
```
正しい形式: /main_unit/jobgroup1/daily_batch
- 先頭に / が必要
- 大文字小文字を正確に
```

---

**原因2**: ジョブネットが存在しない

**対処法**:
```
JP1/AJS3 - ViewまたはWeb UIでジョブネットの存在を確認
```

---

**原因3**: 実行権限がない

**対処法**:
```
JP1ユーザーにジョブネット実行権限を付与
```

---

### エラー: 「証明書エラー」（HTTPS使用時）

**原因**: 自己署名証明書が使用されている

**対処法**:
- スクリプト内で証明書検証を無効化しているため、通常は問題なし
- 本番環境では正式な証明書を使用することを推奨

---

## 💡 応用例

### 複数のジョブネットを順次起動

PowerShellで複数実行：

```powershell
# ジョブネットのリスト
$jobnets = @(
    "/main_unit/jobs/morning_batch",
    "/main_unit/jobs/noon_batch",
    "/main_unit/jobs/evening_batch"
)

# 順次実行
foreach ($jobnet in $jobnets) {
    Write-Host "起動中: $jobnet"
    .\Start-JP1Job.ps1 -JP1Host "192.168.1.100" -JP1User "jp1admin" -JobnetPath $jobnet
    Start-Sleep -Seconds 5
}
```

### JSON設定ファイルから読み込み

設定ファイル（settings.json）：

```json
{
    "jp1Host": "192.168.1.100",
    "jp1Port": 22250,
    "jp1User": "jp1admin",
    "jobnetPath": "/main_unit/jobgroup1/daily_batch",
    "useSSL": false
}
```

実行スクリプト：

```powershell
$settings = Get-Content -Path "settings.json" | ConvertFrom-Json

.\Start-JP1Job.ps1 `
    -JP1Host $settings.jp1Host `
    -JP1Port $settings.jp1Port `
    -JP1User $settings.jp1User `
    -JobnetPath $settings.jobnetPath
```

---

## 📊 REST API版のメリット・デメリット

### メリット ✅

- ローカルにJP1のインストール不要
- HTTP/HTTPS通信で分かりやすい
- ファイアウォール設定が簡単（1ポートのみ）
- 他のシステムとの連携が容易（APIベース）
- プログラミング言語を選ばない

### デメリット ⚠️

- JP1/AJS3 - Manager バージョン10以降が必要
- REST APIの知識が必要
- 同時接続数に制限あり
- エラーハンドリングが複雑

---

## 📚 参考資料

- [JP1/AJS3 REST API リファレンス](https://www.hitachi.co.jp/Prod/comp/soft1/manual/pc/d3K2211/AJSK06/index.html)
- [JP1/AJS3 REST API 認証](https://www.hitachi.co.jp/Prod/comp/soft1/manual/pc/d3K2211/AJSK06/EU210050.HTM)
- [ジョブネット実行API](https://www.hitachi.co.jp/Prod/comp/soft1/manual/pc/d3K2211/AJSK06/EU210053.HTM)

---

**更新日**: 2025-12-02
