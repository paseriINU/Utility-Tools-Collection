# リモートバッチ実行ツール - PowerShell Remoting版

## 概要

PowerShell Remotingを使用して、リモートのWindowsサーバ上でバッチファイルを実行するツールです。
純粋なPowerShellスクリプトとして実装されており、PowerShellの全機能を活用できます。

### WinRM版（バッチ版）との違い

| 項目 | PowerShell Remoting版 | WinRM版（バッチ版） |
|-----|---------------------|------------------|
| **実装言語** | PowerShell (.ps1) | バッチ (.bat) |
| **柔軟性** | ✅ 高い（PowerShellの全機能） | ⚠️ 制限あり |
| **引数渡し** | ✅ 完全対応 | ⚠️ 基本的な対応 |
| **終了コード取得** | ✅ 取得可能 | ❌ 取得不可 |
| **エラーハンドリング** | ✅ 詳細なエラー情報 | ⚠️ 基本的 |
| **カスタマイズ** | ✅ 容易 | ⚠️ やや困難 |
| **Get-Credential対応** | ✅ 対応 | ❌ 非対応 |
| **実行方法** | PowerShell or バッチ経由 | バッチのみ |

## 必要な環境

### ローカル（実行元）
- Windows 10 / Windows 11 / Windows Server 2016以降
- PowerShell 5.1以降

### リモートサーバ（実行先）
- Windows Server 2012以降 / Windows 10以降
- **WinRMサービスが有効化されていること**
- PowerShell Remotingが有効

### 必要な権限
- リモートサーバの**管理者権限**を持つアカウント

## セットアップ

### 1. リモートサーバ側の設定

WinRM版と同じ設定が必要です。詳細は[WinRM版のREADME](../winrm/README.md)を参照してください。

```powershell
# WinRMクイック設定
winrm quickconfig

# サービス確認
Get-Service WinRM

# ファイアウォール確認
Get-NetFirewallRule -Name "WINRM-HTTP-In-TCP"
```

### 2. ローカルPC側の設定

#### ① 実行ポリシーの設定

```powershell
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```

#### ② TrustedHosts設定（ワークグループ環境の場合）

```powershell
# 特定のIPアドレスを追加
Set-Item WSMan:\localhost\Client\TrustedHosts -Value "192.168.1.100"

# 複数追加
Set-Item WSMan:\localhost\Client\TrustedHosts -Value "192.168.1.100,192.168.1.101"
```

## 使い方

### 方法1: PowerShellから直接実行

#### 基本的な使い方

```powershell
.\Invoke-RemoteBatch.ps1 -ComputerName "192.168.1.100" -UserName "Administrator" -BatchPath "C:\Scripts\backup.bat"
```

#### Get-Credentialを使用

```powershell
$cred = Get-Credential
.\Invoke-RemoteBatch.ps1 -ComputerName "192.168.1.100" -Credential $cred -BatchPath "C:\Scripts\test.bat"
```

#### 引数を渡す

```powershell
.\Invoke-RemoteBatch.ps1 -ComputerName "192.168.1.100" -UserName "Administrator" -BatchPath "C:\Scripts\process.bat" -Arguments "param1 param2"
```

#### ログファイルに保存

```powershell
.\Invoke-RemoteBatch.ps1 -ComputerName "192.168.1.100" -UserName "Administrator" -BatchPath "C:\Scripts\backup.bat" -OutputLog "backup_result.log"
```

#### HTTPSを使用（ポート5986）

```powershell
.\Invoke-RemoteBatch.ps1 -ComputerName "server.example.com" -UserName "Administrator" -BatchPath "C:\Scripts\secure.bat" -UseSSL
```

---

### 方法2: バッチファイル経由で実行

PowerShellスクリプトを直接実行するのが難しい場合、バッチファイル経由で実行できます。

#### 基本版

1. **`remote_exec_ps.bat`** を編集
   ```batch
   set REMOTE_SERVER=192.168.1.100
   set REMOTE_USER=Administrator
   set REMOTE_BATCH_PATH=C:\Scripts\backup.bat
   ```

2. 実行
   ```cmd
   remote_exec_ps.bat
   ```

#### 設定ファイル版

1. **`config_ps.ini.sample`** を **`config_ps.ini`** にコピー
   ```cmd
   copy config_ps.ini.sample config_ps.ini
   ```

2. **`config_ps.ini`** を編集
   ```ini
   [Server]
   REMOTE_SERVER=192.168.1.100
   REMOTE_USER=Administrator
   REMOTE_BATCH_PATH=C:\Scripts\backup.bat

   [Options]
   BATCH_ARGUMENTS=
   OUTPUT_LOG=remote_exec_output.log
   USE_SSL=0
   ```

3. 実行
   ```cmd
   remote_exec_ps_config.bat
   ```

## パラメータ詳細

### PowerShellスクリプトのパラメータ

| パラメータ | 必須 | 説明 | 例 |
|----------|------|------|-----|
| `-ComputerName` | ✅ | リモートサーバのコンピュータ名またはIPアドレス | `"192.168.1.100"` |
| `-Credential` | ❌ | PSCredentialオブジェクト | `$cred` |
| `-UserName` | ❌ | ユーザー名（Credentialと排他） | `"Administrator"` |
| `-Password` | ❌ | パスワード（SecureString） | - |
| `-BatchPath` | ✅ | 実行するバッチファイルのフルパス | `"C:\Scripts\test.bat"` |
| `-Arguments` | ❌ | バッチファイルに渡す引数 | `"arg1 arg2"` |
| `-OutputLog` | ❌ | 結果を保存するログファイル | `"result.log"` |
| `-UseSSL` | ❌ | HTTPSを使用（スイッチ） | - |

## 実行例

### 成功時の出力

```
========================================
PowerShell Remoting - リモートバッチ実行
========================================

リモートサーバ: 192.168.1.100
実行ユーザー  : Administrator
実行ファイル  : C:\Scripts\test.bat
出力ログ      : result.log
プロトコル    : HTTP (ポート 5985)

リモートサーバに接続中...
✓ 接続成功

========================================
バッチファイル実行中...
========================================

テストスクリプト開始
現在の日時: 2025/12/01 10:30:00
処理を実行中...
処理完了

========================================
実行完了
終了コード: 0
========================================

実行結果をログファイルに保存中...
✓ ログ保存完了: result.log

処理が正常に完了しました。
```

## PowerShell Remotingの特徴

### メリット

1. **強力なエラーハンドリング**
   - Try-Catchで詳細なエラー情報を取得
   - エラーの種類を判別して適切なメッセージを表示

2. **柔軟な認証**
   - Get-Credentialでセキュアに認証情報を入力
   - PSCredentialオブジェクトの再利用可能

3. **終了コードの取得**
   - バッチファイルの終了コード（`$LASTEXITCODE`）を取得
   - スクリプトの終了コードとして返す

4. **引数の完全サポート**
   - バッチファイルに複数の引数を渡せる
   - スペースを含む引数も正しく処理

5. **詳細ログ**
   - `-Verbose` スイッチで詳細ログを表示可能
   - 実行結果を構造化してログ保存

### PowerShellならではの使い方

#### 複数サーバへの一括実行

```powershell
$servers = @("192.168.1.100", "192.168.1.101", "192.168.1.102")
$cred = Get-Credential

foreach ($server in $servers) {
    Write-Host "=== $server ===" -ForegroundColor Cyan
    .\Invoke-RemoteBatch.ps1 -ComputerName $server -Credential $cred -BatchPath "C:\Scripts\check.bat"
}
```

#### 並列実行

```powershell
$servers = @("192.168.1.100", "192.168.1.101", "192.168.1.102")
$cred = Get-Credential

$jobs = foreach ($server in $servers) {
    Start-Job -ScriptBlock {
        param($s, $c)
        & "C:\Tools\Invoke-RemoteBatch.ps1" -ComputerName $s -Credential $c -BatchPath "C:\Scripts\backup.bat"
    } -ArgumentList $server, $cred
}

# 全ジョブの完了を待つ
$jobs | Wait-Job | Receive-Job
```

#### 条件分岐

```powershell
$result = .\Invoke-RemoteBatch.ps1 -ComputerName "192.168.1.100" -UserName "Administrator" -BatchPath "C:\Scripts\check.bat"

if ($LASTEXITCODE -eq 0) {
    Write-Host "✓ 成功: 次の処理を実行" -ForegroundColor Green
    # 次の処理
} else {
    Write-Host "✗ 失敗: アラート送信" -ForegroundColor Red
    # アラート処理
}
```

#### ログ解析

```powershell
.\Invoke-RemoteBatch.ps1 -ComputerName "192.168.1.100" -UserName "Administrator" -BatchPath "C:\Scripts\backup.bat" -OutputLog "backup.log"

# ログファイルを解析
$log = Get-Content "backup.log"
if ($log -match "エラー") {
    Write-Host "エラーが検出されました" -ForegroundColor Red
    # エラー処理
}
```

## トラブルシューティング

### エラー: "用語 'Invoke-RemoteBatch.ps1' は、コマンドレットの名前として認識されません"

**原因**: スクリプトへのパスが正しくない

**対処法**:
```powershell
# フルパスで指定
C:\Tools\remote-exec\powershell-remoting\Invoke-RemoteBatch.ps1 -ComputerName ...

# または、スクリプトのディレクトリに移動
cd C:\Tools\remote-exec\powershell-remoting
.\Invoke-RemoteBatch.ps1 -ComputerName ...
```

---

### エラー: "このシステムではスクリプトの実行が無効になっているため..."

**原因**: 実行ポリシーが制限されている

**対処法**:
```powershell
# 一時的に実行を許可
powershell -ExecutionPolicy Bypass -File ".\Invoke-RemoteBatch.ps1" -ComputerName ...

# または、ポリシーを変更（管理者権限必要）
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```

---

### エラー: "WinRM クライアントは要求を処理できません"

WinRM版のトラブルシューティングを参照：[WinRM版README](../winrm/README.md#トラブルシューティング)

---

### デバッグモード

詳細な情報を表示：

```powershell
.\Invoke-RemoteBatch.ps1 -ComputerName "192.168.1.100" -UserName "Administrator" -BatchPath "C:\Scripts\test.bat" -Verbose
```

## 高度な使い方

### カスタムスクリプトブロック

`Invoke-RemoteBatch.ps1` を改造して、より複雑な処理を実行できます。

```powershell
# スクリプトを編集して、リモートで複数のコマンドを実行
$scriptBlock = {
    param($batchPath)

    # 環境変数を設定
    $env:CUSTOM_VAR = "CustomValue"

    # バッチ実行前の処理
    Write-Host "実行前: $(Get-Date)"

    # バッチ実行
    $output = & cmd.exe /c $batchPath 2>&1

    # バッチ実行後の処理
    Write-Host "実行後: $(Get-Date)"

    return $output
}
```

### 資格情報の保存（Windows Credential Manager）

```powershell
# 資格情報を保存（初回のみ）
cmdkey /generic:RemoteServer /user:Administrator /pass:YourPassword

# 保存した資格情報を使用（カスタムコード必要）
# または、Get-Credentialで保存した資格情報を読み込むサードパーティツールを使用
```

## セキュリティベストプラクティス

### 1. HTTPS使用を推奨（機密性が高い場合）

リモートサーバでHTTPSリスナー作成：
```powershell
$cert = New-SelfSignedCertificate -DnsName "server.example.com" -CertStoreLocation Cert:\LocalMachine\My
New-Item -Path WSMan:\localhost\Listener -Transport HTTPS -Address * -CertificateThumbPrint $cert.Thumbprint -Force
```

実行時に `-UseSSL` を指定：
```powershell
.\Invoke-RemoteBatch.ps1 -ComputerName "server.example.com" -UserName "Administrator" -BatchPath "C:\Scripts\secure.bat" -UseSSL
```

### 2. Get-Credentialで資格情報を安全に入力

```powershell
$cred = Get-Credential -Message "リモートサーバの管理者アカウント"
.\Invoke-RemoteBatch.ps1 -ComputerName "192.168.1.100" -Credential $cred -BatchPath "C:\Scripts\backup.bat"
```

### 3. 最小権限の原則

専用の管理者アカウントを作成し、必要最小限の権限のみ付与

## 比較表：3つの方法

| 方法 | 実行結果 | 終了コード | 引数渡し | 柔軟性 | セットアップ | 推奨用途 |
|-----|---------|----------|---------|--------|------------|---------|
| **PowerShell Remoting** | ✅ 取得可能 | ✅ 取得可能 | ✅ 完全対応 | ⭐⭐⭐⭐⭐ | ⭐⭐⭐⭐ | PowerShellに精通している場合 |
| **WinRM（バッチ）** | ✅ 取得可能 | ❌ 取得不可 | ⚠️ 基本的 | ⭐⭐⭐ | ⭐⭐⭐⭐ | 結果確認が必要でバッチで完結したい場合 |
| **タスクスケジューラ** | ❌ 取得不可 | ❌ 取得不可 | ⚠️ 基本的 | ⭐⭐ | ⭐⭐ | 単純な起動のみでOKな場合 |

## ライセンス

このツールはMITライセンスの下で公開されています。

## 参考情報

### PowerShell Remotingコマンド

```powershell
# セッション作成
$session = New-PSSession -ComputerName "192.168.1.100" -Credential $cred

# コマンド実行
Invoke-Command -Session $session -ScriptBlock { Get-Process }

# セッションクローズ
Remove-PSSession -Session $session

# 接続テスト
Test-WSMan -ComputerName "192.168.1.100"
```

### 関連リンク

- [Microsoft Docs: PowerShell Remoting](https://docs.microsoft.com/ja-jp/powershell/scripting/learn/remoting/running-remote-commands)
- [about_Remote](https://docs.microsoft.com/ja-jp/powershell/module/microsoft.powershell.core/about/about_remote)
- [Invoke-Command](https://docs.microsoft.com/ja-jp/powershell/module/microsoft.powershell.core/invoke-command)

---

**作成日:** 2025-12-01
**バージョン:** 1.0
