# リモートバッチ実行ツール - WinRM版

## 概要

WinRM（Windows Remote Management）を使用して、リモートのWindowsサーバ上でバッチファイルを実行し、**実行結果をリアルタイムで取得**するツールです。

### タスクスケジューラ版との違い

| 項目 | WinRM版 | タスクスケジューラ版 |
|-----|---------|------------------|
| **実行結果の取得** | ✅ リアルタイムで取得可能 | ❌ 取得不可（ログファイル必要） |
| **標準出力の表示** | ✅ 画面に表示される | ❌ 表示されない |
| **エラー確認** | ✅ 即座に確認可能 | ❌ リモートサーバで確認必要 |
| **セットアップ** | ⚠️ やや複雑（WinRM設定必要） | ✅ 簡単 |
| **使用ポート** | 5985 (HTTP) / 5986 (HTTPS) | 135, 445 |

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

### 1. リモートサーバ側の設定（重要）

リモートサーバで**管理者権限**のPowerShellを開き、以下を実行：

#### ① WinRMの有効化

```powershell
# WinRMクイック設定（自動設定）
winrm quickconfig

# または手動設定
Enable-PSRemoting -Force
```

実行すると以下が設定されます：
- WinRMサービスの起動と自動起動設定
- HTTPリスナーの作成（ポート5985）
- ファイアウォール例外の追加

#### ② ファイアウォール確認

```powershell
# WinRM用のファイアウォール規則を確認
Get-NetFirewallRule -Name "WINRM-HTTP-In-TCP" | Select-Object Name, Enabled, Direction

# 無効の場合は有効化
Enable-NetFirewallRule -Name "WINRM-HTTP-In-TCP"
```

#### ③ WinRMサービスの確認

```powershell
# サービス状態を確認
Get-Service WinRM

# 停止している場合は起動
Start-Service WinRM
Set-Service WinRM -StartupType Automatic
```

### 2. ローカルPC側の設定

#### ① 実行ポリシーの設定（初回のみ）

管理者権限のPowerShellで実行：

```powershell
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```

#### ② TrustedHosts設定（ワークグループ環境の場合）

ドメイン環境でない場合、リモートサーバを信頼済みホストに追加：

```powershell
# 現在の設定を確認
Get-Item WSMan:\localhost\Client\TrustedHosts

# 特定のIPアドレスを追加
Set-Item WSMan:\localhost\Client\TrustedHosts -Value "192.168.1.100"

# 複数追加する場合（カンマ区切り）
Set-Item WSMan:\localhost\Client\TrustedHosts -Value "192.168.1.100,192.168.1.101"

# すべてを信頼（セキュリティリスクあり）
Set-Item WSMan:\localhost\Client\TrustedHosts -Value "*"
```

⚠️ **注意**: ドメイン環境ではTrustedHosts設定は不要です。

## 使い方

### 方法1: 基本版（毎回パスワード入力）

1. **`remote_exec_winrm.bat`** をテキストエディタで開く
2. 以下の設定項目を編集：
   ```batch
   set REMOTE_SERVER=192.168.1.100          ← サーバ名またはIPアドレス
   set REMOTE_USER=Administrator            ← ユーザー名
   set REMOTE_BATCH_PATH=C:\Scripts\target_script.bat  ← 実行するバッチファイル
   set OUTPUT_LOG=%~dp0remote_exec_output.log          ← 結果保存先
   ```
3. **`remote_exec_winrm.bat`** をダブルクリックまたはCMDから実行
4. パスワードを入力
5. **実行結果が画面にリアルタイムで表示されます**

### 方法2: 設定ファイル版

1. **`config_winrm.ini.sample`** を **`config_winrm.ini`** にコピー
   ```cmd
   copy config_winrm.ini.sample config_winrm.ini
   ```

2. **`config_winrm.ini`** を編集：
   ```ini
   [Server]
   REMOTE_SERVER=192.168.1.100
   REMOTE_USER=Administrator
   REMOTE_BATCH_PATH=C:\Scripts\backup.bat

   [Options]
   OUTPUT_LOG=remote_exec_output.log
   ```

3. **`remote_exec_winrm_config.bat`** を実行
4. パスワードを入力
5. 実行結果が画面に表示され、ログファイルにも保存されます

## 実行例

### 成功時の出力例

```
========================================
リモートバッチ実行ツール (WinRM版)
========================================

リモートサーバ: 192.168.1.100
実行ユーザー  : Administrator
実行ファイル  : C:\Scripts\test.bat
出力ログ      : remote_exec_output.log

パスワードを入力してください：
********

リモートサーバに接続中...

接続確認中...
接続成功

========================================
バッチファイル実行結果：
========================================

テストスクリプト開始
ファイル一覧：
 Volume in drive C has no label.
 Directory of C:\Scripts

2025/12/01  10:00    <DIR>          .
2025/12/01  10:00    <DIR>          ..
2025/12/01  09:30               123 test.bat
テストスクリプト完了

========================================
実行完了
========================================

結果をログファイルに保存しました: remote_exec_output.log

処理が完了しました。
```

## WinRMの仕組み

### 通信フロー

```
[ローカルPC]                           [リモートサーバ]
    │                                       │
    │  ① PowerShell Remotingセッション確立   │
    ├──────────────────────────────────────>│
    │  認証（Kerberos/NTLM）                 │  WinRMサービス
    │  ポート: 5985 (HTTP) / 5986 (HTTPS)   │  が受信
    │                                       │
    │  ② Invoke-Command実行                 │
    ├──────────────────────────────────────>│
    │  スクリプトブロック送信                 │  PowerShellホスト
    │                                       │  プロセス起動
    │                                       │  └─> cmd.exe起動
    │  ③ 実行結果をリアルタイム取得           │      └─> batch実行
    │<──────────────────────────────────────┤
    │  標準出力・標準エラー出力               │
    │                                       │
    │  ④ セッションクローズ                  │
    ├──────────────────────────────────────>│
    │                                       │
```

### 使用プロトコル

- **SOAP over HTTP/HTTPS**
- **WS-Management（Web Services-Management）**
- 認証: Kerberos（ドメイン環境）/ NTLM（ワークグループ）

### セキュリティ

#### HTTP vs HTTPS

| プロトコル | ポート | 暗号化 | 用途 |
|-----------|-------|--------|-----|
| HTTP | 5985 | ❌ 認証のみ暗号化<br>データは平文 | ローカルネットワーク内 |
| HTTPS | 5986 | ✅ 完全暗号化 | インターネット経由・機密性高 |

⚠️ デフォルトではHTTP（5985）を使用します。HTTPSを使用する場合は証明書の設定が必要です。

## トラブルシューティング

### エラー: "WinRM クライアントは要求を処理できません"

**原因1: WinRMが有効化されていない**

リモートサーバで確認：
```powershell
winrm enumerate winrm/config/Listener
```

何も表示されない場合は有効化：
```powershell
winrm quickconfig
```

---

**原因2: TrustedHostsに登録されていない（ワークグループ環境）**

ローカルPCで確認：
```powershell
Get-Item WSMan:\localhost\Client\TrustedHosts
```

追加：
```powershell
Set-Item WSMan:\localhost\Client\TrustedHosts -Value "192.168.1.100"
```

---

**原因3: ファイアウォールでブロックされている**

リモートサーバで確認：
```powershell
Test-NetConnection -ComputerName localhost -Port 5985
```

ファイアウォール規則を有効化：
```powershell
Enable-NetFirewallRule -Name "WINRM-HTTP-In-TCP"
```

### エラー: "アクセスが拒否されました"

**対処法1: 管理者権限の確認**
- 使用しているアカウントがリモートサーバの管理者グループに所属しているか確認

**対処法2: UAC設定（ローカルアカウント使用時）**

リモートサーバのレジストリで以下を設定：
```
HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System
LocalAccountTokenFilterPolicy = 1 (DWORD)
```

PowerShellで設定：
```powershell
New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" `
                 -Name "LocalAccountTokenFilterPolicy" `
                 -Value 1 `
                 -PropertyType DWORD `
                 -Force
```

### エラー: "接続がタイムアウトしました"

**対処法：**

1. ネットワーク接続確認
   ```cmd
   ping 192.168.1.100
   ```

2. ポート接続確認
   ```powershell
   Test-NetConnection -ComputerName 192.168.1.100 -Port 5985
   ```

3. リモートサーバのWinRMサービス確認
   ```powershell
   # リモートサーバで実行
   Get-Service WinRM
   ```

### デバッグ方法

詳細なエラー情報を取得：

```powershell
$ErrorActionPreference = "Stop"
$VerbosePreference = "Continue"

$password = ConvertTo-SecureString "YourPassword" -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential("Administrator", $password)

Test-WSMan -ComputerName 192.168.1.100 -Credential $credential -Authentication Default
```

## 応用例

### 1. リモートサーバの情報取得

リモートサーバで実行するバッチ（`C:\Scripts\info.bat`）:
```batch
@echo off
echo ========== システム情報 ==========
systeminfo | findstr /C:"OS Name" /C:"OS Version"
echo.
echo ========== ディスク使用状況 ==========
wmic logicaldisk get caption,freespace,size
echo.
echo ========== 実行中のプロセス ==========
tasklist /FI "STATUS eq running" | more
```

実行：
```cmd
remote_exec_winrm.bat
```

### 2. 複数サーバの一括実行

```batch
@echo off
echo サーバ1実行中...
call remote_exec_winrm_config.bat server1_config.ini
echo.

echo サーバ2実行中...
call remote_exec_winrm_config.bat server2_config.ini
echo.

echo サーバ3実行中...
call remote_exec_winrm_config.bat server3_config.ini
echo.

echo すべてのサーバで実行完了
```

### 3. バッチ実行 + 結果の自動判定

```batch
@echo off
call remote_exec_winrm.bat

if errorlevel 1 (
    echo バッチ実行に失敗しました
    echo アラートメールを送信...
    rem メール送信処理など
) else (
    echo バッチ実行成功
)
```

## セキュリティベストプラクティス

### 1. HTTPS使用を推奨（機密性が高い場合）

リモートサーバでHTTPSリスナー作成：
```powershell
# 自己署名証明書作成
$cert = New-SelfSignedCertificate -DnsName "server.example.com" -CertStoreLocation Cert:\LocalMachine\My

# HTTPSリスナー作成
New-Item -Path WSMan:\localhost\Listener -Transport HTTPS -Address * -CertificateThumbPrint $cert.Thumbprint -Force
```

### 2. パスワードを設定ファイルに保存しない

- 必ず実行時に入力
- または、Credential Managerを使用

### 3. 最小権限の原則

- 専用の管理者アカウントを作成
- 必要最小限の権限のみ付与

## 比較表：リモート実行方法まとめ

| 方法 | 実行結果 | セットアップ | ポート | 追加ツール | 推奨用途 |
|-----|---------|------------|--------|-----------|---------|
| **WinRM** | ✅ 取得可能 | ⭐⭐⭐⭐ | 5985/5986 | 不要 | 結果確認が必要な処理 |
| **タスクスケジューラ** | ❌ 取得不可 | ⭐⭐ | 135/445 | 不要 | 単純な起動のみ |
| **PsExec** | ✅ 取得可能 | ⭐⭐⭐ | 135/445 | 要DL | レガシー環境 |

## ライセンス

このツールはMITライセンスの下で公開されています。

## 参考情報

### WinRM設定コマンド

```powershell
# WinRM設定の確認
winrm get winrm/config

# リスナー一覧
winrm enumerate winrm/config/Listener

# 接続テスト
Test-WSMan -ComputerName 192.168.1.100
```

### 関連リンク

- [Microsoft Docs: Windows Remote Management](https://docs.microsoft.com/ja-jp/windows/win32/winrm/portal)
- [PowerShell Remoting](https://docs.microsoft.com/ja-jp/powershell/scripting/learn/remoting/running-remote-commands)
- [WinRM セキュリティ](https://docs.microsoft.com/ja-jp/windows/win32/winrm/authentication-for-remote-connections)

---

**作成日:** 2025-12-01
**バージョン:** 1.0
