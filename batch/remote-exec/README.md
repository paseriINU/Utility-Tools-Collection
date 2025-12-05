# リモートバッチ実行ツール (Remote Batch Executor)

## 概要

リモートのWindowsサーバ上でバッチファイルをPowerShell Remotingで実行するツールです。
ハイブリッド版として、.bat形式で保存されており、ダブルクリックで実行可能です。

## ✨ 主な機能

- ✅ **ダブルクリックで実行可能** - .bat形式のハイブリッドスクリプト
- ✅ **管理者権限の自動昇格** - 必要に応じてUACプロンプトで再起動
- ✅ **WinRM設定の自動構成と復元** - TrustedHostsの追加・削除を自動化
- ✅ **環境選択機能** - tst1t/tst2t など複数環境に対応
- ✅ **実行結果のリアルタイム表示** - 標準出力・エラー出力を画面表示
- ✅ **終了コードの取得** - バッチの成功/失敗を判定可能
- ✅ **引数のサポート** - バッチファイルに引数を渡すことが可能
- ✅ **ログファイルの自動保存** - 実行結果を日時付きログに保存
- ✅ **ネットワークパス対応** - UNCパスから実行可能
- ✅ **SSL/HTTPS対応** - セキュアな通信にも対応

## 📁 ファイル構成

```
remote-exec/
├── Invoke-RemoteBatch-Hybrid.bat  # リモートバッチ実行ツール（本体）
├── README.md                      # このファイル
└── .gitignore                     # Git除外設定
```

## 🚀 クイックスタート

### 1. リモートサーバの準備（初回のみ）

リモートサーバでWinRMを有効化します：

```powershell
# 管理者権限のPowerShellで実行
Enable-PSRemoting -Force

# ファイアウォールの設定を確認
Get-NetFirewallRule -Name "WINRM-HTTP-In-TCP" | Format-List
```

または簡易設定：

```powershell
winrm quickconfig
```

### 2. スクリプトの編集

`Invoke-RemoteBatch-Hybrid.bat` をテキストエディタで開き、設定セクションを編集：

```powershell
$Config = @{
    # リモートサーバのIPアドレスまたはホスト名
    ComputerName = "192.168.1.100"

    # ユーザー名（空の場合は実行時に入力）
    UserName = "Administrator"

    # パスワード（空の場合は実行時に入力、推奨）
    Password = ""

    # 実行するバッチファイルのパス（リモートサーバ上）
    # {env} は環境選択時に置換されます
    BatchPath = "C:\Scripts\{env}\backup.bat"

    # バッチファイルに渡す引数（オプション）
    Arguments = ""

    # HTTPS接続を使用する場合は $true
    UseSSL = $false
}
```

### 3. 実行

ダブルクリックで実行：

```cmd
Invoke-RemoteBatch-Hybrid.bat
```

または、コマンドプロンプトから：

```cmd
cd batch\remote-exec
Invoke-RemoteBatch-Hybrid.bat
```

### 4. 実行フロー

1. **管理者権限チェック** → 必要に応じてUACプロンプト表示
2. **環境選択**（BatchPathに{env}が含まれる場合）
   ```
   実行環境を選択してください:
     1. tst1t
     2. tst2t
   選択 (1-2):
   ```
3. **WinRM設定の自動構成**
4. **パスワード入力**（Passwordが空の場合）
5. **リモートサーバに接続してバッチ実行**
6. **実行結果の表示とログ保存**
7. **WinRM設定の復元**

## 🔧 主な機能の詳細

### 環境選択機能

`BatchPath` に `{env}` を含めると、実行時に環境を選択できます：

```powershell
BatchPath = "C:\Scripts\{env}\backup.bat"
```

実行時の選択肢：
1. tst1t
2. tst2t

選択すると、`{env}` が選択した環境名に置換されます。

**環境選択機能を使用しない場合**は、固定値を設定：

```powershell
BatchPath = "C:\Scripts\production\daily_batch.bat"
```

### WinRM設定の自動構成

スクリプトは以下を自動的に行います：

1. **WinRMサービスの起動確認**
   - 停止している場合のみ起動
   - 終了時に元の状態（停止）に戻す

2. **TrustedHostsの設定**
   - 接続先をTrustedHostsに追加
   - 終了時に元の設定に復元

3. **設定の復元**
   - スクリプト終了時（正常終了・エラー時共に）に元の状態に復元

### ログファイルの保存

実行結果は自動的にログファイルに保存されます：

- **保存先**: `log\RemoteBatch_yyyyMMdd_HHmmss.log`
- **UNCパスから実行時**: `%TEMP%\RemoteBatchLogs\RemoteBatch_yyyyMMdd_HHmmss.log`
- **ファイル名例**: `RemoteBatch_20251203_153045.log`

### 引数のサポート

バッチファイルに引数を渡すことができます：

```powershell
Arguments = "/full /backup"
```

リモートサーバ上では以下のように実行されます：

```cmd
C:\Scripts\backup.bat /full /backup
```

### SSL/HTTPS接続

セキュアな通信を使用する場合：

```powershell
UseSSL = $true
```

ポート 5986 が使用されます（通常は 5985）。

## ⚙️ 必要な環境

### ローカルPC（実行元）
- Windows 10 / Windows 11 / Windows Server 2016以降
- PowerShell 5.1以降
- 管理者権限（WinRM設定のため）

### リモートサーバ（実行先）
- Windows Server 2012 R2以降 / Windows 10以降
- WinRMサービスが有効化されていること
- PowerShell Remotingが有効
- ファイアウォールでポート5985（HTTP）または5986（HTTPS）が開放
- 管理者権限を持つアカウント

## 🔐 セキュリティ注意事項

1. **パスワードの管理**
   - スクリプト内にパスワードを記載しないことを強く推奨
   - `Password = ""` のまま実行時に入力する方法を推奨
   - パスワードを設定ファイルに記述する場合：
     ```powershell
     Password = "YourPassword123"  # ⚠️ 平文保存になります
     ```

2. **ネットワークセキュリティ**
   - 信頼できるネットワーク内でのみ使用
   - インターネット経由の場合はVPNを使用
   - 本番環境ではHTTPS（ポート5986）を使用推奨

3. **TrustedHosts設定**
   - スクリプトは実行後に自動的に設定を復元します
   - 手動で設定する場合は、必要最小限のホストのみを追加

4. **管理者権限**
   - このスクリプトは管理者権限で実行されます
   - TrustedHosts設定を変更するため、信頼できる接続先のみ指定してください
   - 信頼できるスクリプトのみを実行してください

## 🐛 トラブルシューティング

### エラー: "WinRM設定の自動構成に失敗しました"

**原因**: 管理者権限で実行されていない

**解決方法**:
- スクリプトは自動的に管理者権限で再起動を試みます
- UACプロンプトが表示されたら「はい」を選択してください
- または、スクリプトを右クリック → 「管理者として実行」

### エラー: "リモート実行に失敗しました"

**原因1**: リモートサーバでWinRMが有効化されていない

**解決方法**:
```powershell
# リモートサーバで実行
winrm quickconfig
# または
Enable-PSRemoting -Force
```

**原因2**: ファイアウォールでポートが開いていない

**解決方法**:
- HTTP: ポート 5985 を開放
- HTTPS: ポート 5986 を開放

**原因3**: 認証情報が正しくない

**解決方法**:
- ユーザー名とパスワードを確認
- リモートサーバの管理者アカウントを使用しているか確認
- ドメイン環境の場合は `ドメイン名\ユーザー名` 形式で指定

### 接続テスト

リモートサーバへの接続をテストするには：

```powershell
# ネットワーク接続確認
Test-Connection -ComputerName 192.168.1.100

# WinRM接続テスト
Test-WSMan -ComputerName 192.168.1.100
```

### TrustedHostsの設定が復元されない

**原因**: スクリプトが異常終了した

**解決方法**: 手動で設定を確認・復元
```powershell
# 現在の設定を確認
Get-Item WSMan:\localhost\Client\TrustedHosts

# 空に戻す
Set-Item WSMan:\localhost\Client\TrustedHosts -Value "" -Force
```

## 🔧 手動でWinRMを設定する場合

スクリプトを使用せず、手動で設定する場合：

```powershell
# WinRMサービスの起動
Start-Service WinRM

# TrustedHostsの設定
Set-Item WSMan:\localhost\Client\TrustedHosts -Value "192.168.1.100" -Force

# 確認
Get-Item WSMan:\localhost\Client\TrustedHosts
```

## 📝 ライセンス

このツールはMITライセンスの下で公開されています。
個人・商用問わず自由に使用・改変できます。

---

**作成日:** 2025-12-05
**バージョン:** 2.0
**更新内容:** PowerShell Remoting版に統一、タスクスケジューラ版とWinRM版を削除

## バージョン履歴

- **v2.0** (2025-12-05)
  - ハイブリッド.bat形式に統合
  - 管理者権限の自動昇格機能を追加
  - WinRM自動設定・復元機能を追加
  - 環境選択機能を追加
  - エラー時の自動復元機能を追加
  - タスクスケジューラ版とWinRM版を削除し、PowerShell Remoting版に統一
