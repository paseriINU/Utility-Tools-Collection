# Windows Server 2022 WinRM設定手順

LinuxからWindows Server 2022へWinRM接続を行うための、Windows側の詳細設定手順です。

## 目次

1. [事前確認](#事前確認)
2. [WinRMサービスの有効化](#winrmサービスの有効化)
3. [認証設定](#認証設定)
4. [ファイアウォール設定](#ファイアウォール設定)
5. [TrustedHosts設定](#trustedhosts設定)
6. [HTTPSの設定（推奨）](#httpsの設定推奨)
7. [動作確認](#動作確認)
8. [セキュリティ強化](#セキュリティ強化)
9. [トラブルシューティング](#トラブルシューティング)

## 事前確認

### 必要な権限

- **管理者権限**でPowerShellを実行する必要があります

### システム要件

- Windows Server 2022（またはWindows 10/11 Pro以上）
- PowerShell 5.1以降
- .NET Framework 4.5以降

### 管理者としてPowerShellを起動

1. スタートメニューを開く
2. "PowerShell"と入力
3. "Windows PowerShell"を右クリック
4. "管理者として実行"を選択

## WinRMサービスの有効化

### 1. クイック設定（推奨）

```powershell
# WinRMのクイック設定（対話なし）
winrm quickconfig -force

# 実行結果の例:
# WinRM サービスは既に実行中です。
# WinRM は既に管理用に設定されています。
```

この設定により以下が自動的に実行されます：
- WinRMサービスの起動
- WinRMサービスの自動起動設定
- HTTPリスナーの作成（ポート5985）
- ファイアウォール例外の作成

### 2. 手動設定

詳細に制御したい場合は、以下のコマンドを個別に実行します：

```powershell
# WinRMサービスの起動タイプを自動に設定
Set-Service -Name WinRM -StartupType Automatic

# WinRMサービスの起動
Start-Service -Name WinRM

# サービス状態の確認
Get-Service -Name WinRM | Format-List

# 出力例:
# Name                : WinRM
# DisplayName         : Windows Remote Management (WS-Management)
# Status              : Running
# DependentServices   : {}
# ServicesDependedOn  : {RPCSS, HTTP}
# CanPauseAndContinue : False
# CanShutdown         : True
# CanStop             : True
# ServiceType         : Win32ShareProcess
```

### 3. リスナーの作成

```powershell
# HTTPリスナーの確認
winrm enumerate winrm/config/listener

# HTTPリスナーが存在しない場合は作成
New-Item -Path WSMan:\localhost\Listener -Transport HTTP -Address * -Force

# リスナーの詳細確認
Get-ChildItem WSMan:\localhost\Listener
```

## 認証設定

### 1. NTLM認証の有効化（推奨）

```powershell
# NTLM認証を有効化
Set-Item WSMan:\localhost\Service\Auth\Negotiate -Value $true

# 設定確認
Get-Item WSMan:\localhost\Service\Auth\Negotiate
```

### 2. 基本認証の有効化（開発環境のみ）

**警告**: 基本認証はパスワードがBase64エンコードされるだけで、暗号化されません。HTTPSと併用しない場合はセキュリティリスクがあります。

```powershell
# 基本認証を有効化（開発環境のみ）
Set-Item WSMan:\localhost\Service\Auth\Basic -Value $true

# 設定確認
Get-Item WSMan:\localhost\Service\Auth\Basic
```

### 3. 認証設定の確認

```powershell
# すべての認証方式の確認
winrm get winrm/config/service/auth

# 出力例:
# Auth
#     Basic = true
#     Kerberos = true
#     Negotiate = true
#     Certificate = false
#     CredSSP = false
#     CbtHardeningLevel = Relaxed
```

## ファイアウォール設定

### 1. WinRM HTTPポートの許可

```powershell
# ファイアウォールルールの作成（ポート5985）
New-NetFirewallRule -Name "WinRM-HTTP-In" `
    -DisplayName "Windows Remote Management (HTTP-In)" `
    -Description "Inbound rule for WinRM via HTTP" `
    -Protocol TCP `
    -LocalPort 5985 `
    -Direction Inbound `
    -Action Allow `
    -Enabled True `
    -Profile Any

# ルールの確認
Get-NetFirewallRule -Name "WinRM-HTTP-In" | Format-List
```

### 2. 特定のIPからのみ接続を許可（推奨）

```powershell
# 特定のLinuxサーバIPからのみ許可
New-NetFirewallRule -Name "WinRM-HTTP-In-FromLinux" `
    -DisplayName "Windows Remote Management (HTTP-In from Linux)" `
    -Description "Inbound rule for WinRM via HTTP from specific Linux server" `
    -Protocol TCP `
    -LocalPort 5985 `
    -Direction Inbound `
    -Action Allow `
    -Enabled True `
    -Profile Any `
    -RemoteAddress "192.168.1.10"

# 複数のIPを許可する場合
New-NetFirewallRule -Name "WinRM-HTTP-In-FromLinux" `
    -DisplayName "Windows Remote Management (HTTP-In from Linux)" `
    -Protocol TCP `
    -LocalPort 5985 `
    -Direction Inbound `
    -Action Allow `
    -Enabled True `
    -Profile Any `
    -RemoteAddress @("192.168.1.10", "192.168.1.11", "192.168.1.12")
```

### 3. ファイアウォールの状態確認

```powershell
# WinRM関連のファイアウォールルールを確認
Get-NetFirewallRule | Where-Object { $_.DisplayName -like "*WinRM*" } | Format-Table Name, DisplayName, Enabled, Direction, Action

# 特定のポートが開いているか確認
Get-NetFirewallRule | Where-Object { $_.LocalPort -eq 5985 }
```

## TrustedHosts設定

WinRMクライアント（Linuxサーバ）を信頼済みホストとして登録します。

### 1. すべてのホストを許可（開発環境のみ）

```powershell
# すべてのホストからの接続を許可
Set-Item WSMan:\localhost\Client\TrustedHosts -Value "*" -Force

# 設定確認
Get-Item WSMan:\localhost\Client\TrustedHosts
```

### 2. 特定のIPのみ許可（推奨）

```powershell
# 特定のLinuxサーバのIPのみ許可
Set-Item WSMan:\localhost\Client\TrustedHosts -Value "192.168.1.10" -Force

# 現在の設定確認
Get-Item WSMan:\localhost\Client\TrustedHosts
```

### 3. 複数のIPを許可

```powershell
# 複数のIPをカンマ区切りで指定
Set-Item WSMan:\localhost\Client\TrustedHosts -Value "192.168.1.10,192.168.1.11,192.168.1.12" -Force

# または既存の設定に追加
$current = (Get-Item WSMan:\localhost\Client\TrustedHosts).Value
$newHosts = "$current,192.168.1.20"
Set-Item WSMan:\localhost\Client\TrustedHosts -Value $newHosts -Force
```

### 4. TrustedHostsのリセット

```powershell
# TrustedHostsをクリア
Set-Item WSMan:\localhost\Client\TrustedHosts -Value "" -Force

# 確認
Get-Item WSMan:\localhost\Client\TrustedHosts
```

## HTTPSの設定（推奨）

本番環境ではHTTPS（ポート5986）を使用することを強く推奨します。

### 1. 自己署名証明書の作成

```powershell
# 自己署名証明書の作成
$cert = New-SelfSignedCertificate `
    -DnsName "winserver.example.com" `
    -CertStoreLocation "Cert:\LocalMachine\My" `
    -KeyLength 2048 `
    -KeyAlgorithm RSA `
    -HashAlgorithm SHA256 `
    -KeyUsage DigitalSignature, KeyEncipherment `
    -Type SSLServerAuthentication `
    -NotAfter (Get-Date).AddYears(5)

# 証明書の確認
$cert | Format-List Subject, Thumbprint, NotAfter

# 証明書のサムプリント表示
Write-Host "証明書のサムプリント: $($cert.Thumbprint)"
```

### 2. HTTPSリスナーの作成

```powershell
# HTTPSリスナーを作成
New-Item -Path WSMan:\localhost\Listener `
    -Transport HTTPS `
    -Address * `
    -CertificateThumbPrint $cert.Thumbprint `
    -Force

# リスナーの確認
winrm enumerate winrm/config/listener
```

### 3. HTTPSファイアウォールルールの追加

```powershell
# ファイアウォールでHTTPS（ポート5986）を許可
New-NetFirewallRule -Name "WinRM-HTTPS-In" `
    -DisplayName "Windows Remote Management (HTTPS-In)" `
    -Description "Inbound rule for WinRM via HTTPS" `
    -Protocol TCP `
    -LocalPort 5986 `
    -Direction Inbound `
    -Action Allow `
    -Enabled True `
    -Profile Any

# または特定のIPからのみ許可
New-NetFirewallRule -Name "WinRM-HTTPS-In-FromLinux" `
    -DisplayName "Windows Remote Management (HTTPS-In from Linux)" `
    -Protocol TCP `
    -LocalPort 5986 `
    -Direction Inbound `
    -Action Allow `
    -Enabled True `
    -Profile Any `
    -RemoteAddress "192.168.1.10"
```

### 4. 証明書のエクスポート（Linux側で使用）

```powershell
# 証明書を公開鍵のみエクスポート（Linux側に配布）
Export-Certificate -Cert $cert -FilePath "C:\Temp\winrm-cert.cer"

Write-Host "証明書をC:\Temp\winrm-cert.cerにエクスポートしました"
Write-Host "このファイルをLinuxサーバにコピーして使用してください"
```

## 動作確認

### 1. ローカルテスト

```powershell
# WinRMサービスの状態確認
Get-Service -Name WinRM

# WinRMリスナーの確認
winrm enumerate winrm/config/listener

# 自分自身への接続テスト
Test-WSMan -ComputerName localhost

# 認証テスト（NTLMを使用）
$cred = Get-Credential -Message "Windowsのユーザー名とパスワードを入力"
Test-WSMan -ComputerName localhost -Credential $cred -Authentication Negotiate
```

### 2. リモートコマンド実行テスト

```powershell
# 自分自身にリモート接続してコマンド実行
$session = New-PSSession -ComputerName localhost -Credential (Get-Credential)
Invoke-Command -Session $session -ScriptBlock { Get-Date }
Remove-PSSession -Session $session
```

### 3. WinRM設定の全体確認

```powershell
# 全体設定の表示
winrm get winrm/config

# サービス設定の表示
winrm get winrm/config/service
winrm get winrm/config/service/auth

# クライアント設定の表示
winrm get winrm/config/client

# リスナー設定の表示
winrm get winrm/config/listener?Address=*+Transport=HTTP
```

### 4. ポート疎通確認

```powershell
# ポート5985がリスニング状態か確認
netstat -an | findstr :5985

# 期待される出力:
# TCP    0.0.0.0:5985          0.0.0.0:0              LISTENING
# TCP    [::]:5985             [::]:0                 LISTENING
```

## セキュリティ強化

### 1. 最小権限の原則

WinRM専用の管理者アカウントを作成します：

```powershell
# 専用の管理者アカウント作成
$username = "WinRMAdmin"
$password = ConvertTo-SecureString "StrongP@ssw0rd123!" -AsPlainText -Force
New-LocalUser -Name $username -Password $password -Description "WinRM専用管理者"

# Administratorsグループに追加
Add-LocalGroupMember -Group "Administrators" -Member $username

# Remote Management Usersグループに追加
Add-LocalGroupMember -Group "Remote Management Users" -Member $username
```

### 2. 基本認証の無効化（本番環境）

```powershell
# 基本認証を無効化
Set-Item WSMan:\localhost\Service\Auth\Basic -Value $false

# NTLM認証のみ有効
Set-Item WSMan:\localhost\Service\Auth\Negotiate -Value $true

# 設定確認
winrm get winrm/config/service/auth
```

### 3. HTTPリスナーの削除（HTTPS専用）

```powershell
# HTTPリスナーを削除（HTTPS専用にする場合）
Remove-Item -Path WSMan:\localhost\Listener\* -Recurse -Force

# HTTPSリスナーのみ再作成
New-Item -Path WSMan:\localhost\Listener `
    -Transport HTTPS `
    -Address * `
    -CertificateThumbPrint $cert.Thumbprint `
    -Force

# HTTPファイアウォールルールを無効化
Disable-NetFirewallRule -Name "WinRM-HTTP-In"
```

### 4. ログの有効化

```powershell
# WinRMイベントログの有効化
wevtutil set-log "Microsoft-Windows-WinRM/Operational" /enabled:true

# ログの確認
Get-WinEvent -LogName "Microsoft-Windows-WinRM/Operational" -MaxEvents 10
```

### 5. タイムアウト設定

```powershell
# セッションタイムアウトの設定（ミリ秒）
Set-Item WSMan:\localhost\MaxTimeoutms -Value 300000  # 5分

# アイドルタイムアウトの設定（ミリ秒）
Set-Item WSMan:\localhost\Shell\IdleTimeout -Value 180000  # 3分

# 設定確認
Get-Item WSMan:\localhost\MaxTimeoutms
Get-Item WSMan:\localhost\Shell\IdleTimeout
```

## トラブルシューティング

### 問題1: WinRMサービスが起動しない

```powershell
# 依存サービスの確認
Get-Service -Name RPCSS, HTTP

# 依存サービスが停止している場合は起動
Start-Service -Name RPCSS
Start-Service -Name HTTP
Start-Service -Name WinRM

# イベントログでエラー確認
Get-WinEvent -LogName "System" -MaxEvents 20 | Where-Object { $_.ProviderName -eq "Service Control Manager" }
```

### 問題2: ファイアウォールルールが機能しない

```powershell
# ファイアウォールが有効か確認
Get-NetFirewallProfile | Format-Table Name, Enabled

# WinRM関連ルールの詳細確認
Get-NetFirewallRule -Name "WinRM-HTTP-In" | Get-NetFirewallPortFilter
Get-NetFirewallRule -Name "WinRM-HTTP-In" | Get-NetFirewallAddressFilter

# ファイアウォールログの有効化
Set-NetFirewallProfile -Profile Domain,Public,Private -LogBlocked True -LogAllowed True
```

### 問題3: TrustedHostsが反映されない

```powershell
# WinRMサービスを再起動
Restart-Service -Name WinRM

# TrustedHostsを再設定
Set-Item WSMan:\localhost\Client\TrustedHosts -Value "*" -Force -Confirm:$false

# 設定の詳細確認
Get-Item WSMan:\localhost\Client\TrustedHosts | Format-List *
```

### 問題4: 証明書エラー

```powershell
# 証明書の一覧表示
Get-ChildItem Cert:\LocalMachine\My

# 特定の証明書の詳細確認
$cert = Get-ChildItem Cert:\LocalMachine\My | Where-Object { $_.Subject -like "*winserver*" }
$cert | Format-List *

# 証明書の有効期限確認
$cert.NotAfter
```

### 問題5: 認証エラー

```powershell
# 認証設定の確認
winrm get winrm/config/service/auth

# すべての認証方式を有効化（テスト用）
Set-Item WSMan:\localhost\Service\Auth\Basic -Value $true
Set-Item WSMan:\localhost\Service\Auth\Negotiate -Value $true
Set-Item WSMan:\localhost\Service\Auth\Kerberos -Value $true

# WinRMサービスを再起動
Restart-Service -Name WinRM
```

### デバッグログの有効化

```powershell
# 詳細ログの有効化
Set-Item WSMan:\localhost\Service\EnableCompatibilityHttpListener -Value $true

# デバッグトレースの有効化
winrm set winrm/config/service '@{EnableCompatibilityHttpListener="true"}'

# ログの確認
Get-WinEvent -LogName "Microsoft-Windows-WinRM/Operational" -MaxEvents 50 | Format-Table TimeCreated, Id, Message -AutoSize
```

## 設定の削除・リセット

### WinRM設定の完全リセット

```powershell
# WinRMサービスを停止
Stop-Service -Name WinRM

# すべてのリスナーを削除
Remove-Item -Path WSMan:\localhost\Listener\* -Recurse -Force

# TrustedHostsをクリア
Set-Item WSMan:\localhost\Client\TrustedHosts -Value "" -Force

# ファイアウォールルールを削除
Remove-NetFirewallRule -Name "WinRM-HTTP-In" -ErrorAction SilentlyContinue
Remove-NetFirewallRule -Name "WinRM-HTTPS-In" -ErrorAction SilentlyContinue

# WinRMサービスを無効化
Set-Service -Name WinRM -StartupType Disabled
```

### 再セットアップ

```powershell
# 設定をリセットした後、再度セットアップ
winrm quickconfig -force
```

## 参考コマンド一覧

```powershell
# サービス管理
Get-Service -Name WinRM
Start-Service -Name WinRM
Stop-Service -Name WinRM
Restart-Service -Name WinRM

# 設定確認
winrm get winrm/config
winrm enumerate winrm/config/listener
Get-Item WSMan:\localhost\Client\TrustedHosts

# テスト
Test-WSMan -ComputerName localhost
Test-NetConnection -ComputerName localhost -Port 5985

# ファイアウォール
Get-NetFirewallRule | Where-Object { $_.DisplayName -like "*WinRM*" }
New-NetFirewallRule -Name "WinRM-HTTP-In" -Protocol TCP -LocalPort 5985 -Action Allow

# ログ
Get-WinEvent -LogName "Microsoft-Windows-WinRM/Operational" -MaxEvents 10
```

## まとめ

この手順に従ってWindows Server 2022のWinRM設定を行うことで、LinuxからのWinRM接続が可能になります。

**開発環境の最小設定**:
1. `winrm quickconfig -force`
2. `Set-Item WSMan:\localhost\Client\TrustedHosts -Value "*" -Force`
3. ファイアウォールでポート5985を許可

**本番環境の推奨設定**:
1. HTTPS接続（ポート5986）を使用
2. 特定のIPからのみ接続を許可
3. 専用の管理者アカウントを使用
4. 基本認証を無効化
5. ログを有効化して監視
