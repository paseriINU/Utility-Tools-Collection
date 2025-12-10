<# :
@echo off
chcp 65001 >nul
title サーバ構成情報収集ツール
setlocal

rem UNCパス対応（PushD/PopDで自動マッピング）
pushd "%~dp0"

powershell -NoProfile -ExecutionPolicy Bypass -Command "$scriptDir=('%~dp0' -replace '\\$',''); try { iex ((gc '%~f0' -Encoding UTF8) -join \"`n\") } finally { Set-Location C:\ }"
set EXITCODE=%ERRORLEVEL%

popd

pause
exit /b %EXITCODE%
: #>

<#
.SYNOPSIS
    Windowsサーバ構成情報収集ツール（ネットワーク・セキュリティ重視）

.DESCRIPTION
    WinRM実行検討のため、ローカルサーバのネットワーク・セキュリティ設定を
    収集してExcelファイルに出力します。

.NOTES
    作成日: 2025-12-10
    バージョン: 1.0

    収集項目:
    - OS基本情報
    - ネットワーク設定（IP、DNS、ゲートウェイ）
    - WinRM設定（サービス状態、TrustedHosts、認証設定）
    - ファイアウォール設定（WinRM関連ルール）
    - 開いているポート（LISTENING）
    - セキュリティ関連レジストリ
    - Windowsサービス一覧
#>

# ==============================================================================
# ■ 設定セクション（ここを編集してください）
# ==============================================================================

$Config = @{
    # 出力先フォルダ（空の場合はデスクトップ）
    OutputFolder = ""

    # 収集項目の有効/無効
    CollectOSInfo = $true
    CollectNetworkConfig = $true
    CollectWinRMConfig = $true
    CollectFirewallRules = $true
    CollectOpenPorts = $true
    CollectRegistrySettings = $true
    CollectServices = $true
    CollectLocalUsers = $true
}

# ==============================================================================
# ■ メイン処理（以下は編集不要）
# ==============================================================================

$ErrorActionPreference = "Continue"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ヘッダー表示
Write-Host ""
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host "  サーバ構成情報収集ツール" -ForegroundColor Cyan
Write-Host "  （ネットワーク・セキュリティ重視）" -ForegroundColor Cyan
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host ""

# 出力先設定
$outputFolder = if ($Config.OutputFolder -ne "") { $Config.OutputFolder } else { "$env:USERPROFILE\Desktop" }
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$computerName = $env:COMPUTERNAME
$excelPath = Join-Path $outputFolder "ServerConfig_${computerName}_${timestamp}.xlsx"
$csvFolder = Join-Path $outputFolder "ServerConfig_${computerName}_${timestamp}"

Write-Host "コンピュータ名: $computerName" -ForegroundColor White
Write-Host "出力先: $outputFolder" -ForegroundColor White
Write-Host ""

# データ収集用ハッシュテーブル
$collectedData = @{}

#region OS基本情報
if ($Config.CollectOSInfo) {
    Write-Host "[収集中] OS基本情報..." -ForegroundColor Cyan

    $osInfo = Get-CimInstance -ClassName Win32_OperatingSystem
    $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem

    $collectedData["OS情報"] = @(
        [PSCustomObject]@{ 項目 = "コンピュータ名"; 値 = $computerName }
        [PSCustomObject]@{ 項目 = "OS名"; 値 = $osInfo.Caption }
        [PSCustomObject]@{ 項目 = "OSバージョン"; 値 = $osInfo.Version }
        [PSCustomObject]@{ 項目 = "OSビルド"; 値 = $osInfo.BuildNumber }
        [PSCustomObject]@{ 項目 = "サービスパック"; 値 = $osInfo.ServicePackMajorVersion }
        [PSCustomObject]@{ 項目 = "OSアーキテクチャ"; 値 = $osInfo.OSArchitecture }
        [PSCustomObject]@{ 項目 = "インストール日"; 値 = $osInfo.InstallDate.ToString("yyyy-MM-dd HH:mm:ss") }
        [PSCustomObject]@{ 項目 = "最終起動日時"; 値 = $osInfo.LastBootUpTime.ToString("yyyy-MM-dd HH:mm:ss") }
        [PSCustomObject]@{ 項目 = "ドメイン/ワークグループ"; 値 = $computerSystem.Domain }
        [PSCustomObject]@{ 項目 = "ドメイン参加"; 値 = if ($computerSystem.PartOfDomain) { "はい" } else { "いいえ（ワークグループ）" } }
        [PSCustomObject]@{ 項目 = "物理メモリ(GB)"; 値 = [math]::Round($computerSystem.TotalPhysicalMemory / 1GB, 2) }
        [PSCustomObject]@{ 項目 = "プロセッサ数"; 値 = $computerSystem.NumberOfProcessors }
        [PSCustomObject]@{ 項目 = "論理プロセッサ数"; 値 = $computerSystem.NumberOfLogicalProcessors }
    )

    Write-Host "  [OK] OS基本情報: $($collectedData["OS情報"].Count) 件" -ForegroundColor Green
}
#endregion

#region ネットワーク設定
if ($Config.CollectNetworkConfig) {
    Write-Host "[収集中] ネットワーク設定..." -ForegroundColor Cyan

    $networkAdapters = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled -eq $true }

    $networkData = @()
    foreach ($adapter in $networkAdapters) {
        $networkData += [PSCustomObject]@{
            アダプタ名 = $adapter.Description
            MACアドレス = $adapter.MACAddress
            IPアドレス = ($adapter.IPAddress -join ", ")
            サブネットマスク = ($adapter.IPSubnet -join ", ")
            デフォルトゲートウェイ = ($adapter.DefaultIPGateway -join ", ")
            DNSサーバ = ($adapter.DNSServerSearchOrder -join ", ")
            DHCP有効 = if ($adapter.DHCPEnabled) { "はい" } else { "いいえ" }
            DHCPサーバ = $adapter.DHCPServer
        }
    }

    $collectedData["ネットワーク設定"] = $networkData

    # ネットワークプロファイル
    try {
        $networkProfiles = Get-NetConnectionProfile -ErrorAction SilentlyContinue
        $profileData = @()
        foreach ($profile in $networkProfiles) {
            $profileData += [PSCustomObject]@{
                インターフェース = $profile.InterfaceAlias
                ネットワーク名 = $profile.Name
                ネットワークカテゴリ = $profile.NetworkCategory.ToString()
                IPv4接続 = $profile.IPv4Connectivity.ToString()
                IPv6接続 = $profile.IPv6Connectivity.ToString()
            }
        }
        $collectedData["ネットワークプロファイル"] = $profileData
    } catch {
        Write-Host "  [警告] ネットワークプロファイル取得に失敗" -ForegroundColor Yellow
    }

    Write-Host "  [OK] ネットワーク設定: $($collectedData["ネットワーク設定"].Count) 件" -ForegroundColor Green
}
#endregion

#region WinRM設定
if ($Config.CollectWinRMConfig) {
    Write-Host "[収集中] WinRM設定..." -ForegroundColor Cyan

    $winrmData = @()

    # WinRMサービス状態
    $winrmService = Get-Service -Name WinRM -ErrorAction SilentlyContinue
    $winrmData += [PSCustomObject]@{
        カテゴリ = "サービス"
        項目 = "WinRMサービス状態"
        値 = if ($winrmService) { $winrmService.Status.ToString() } else { "未インストール" }
        備考 = if ($winrmService.Status -eq "Running") { "正常" } else { "要確認" }
    }
    $winrmData += [PSCustomObject]@{
        カテゴリ = "サービス"
        項目 = "WinRMスタートアップ種別"
        値 = if ($winrmService) { $winrmService.StartType.ToString() } else { "-" }
        備考 = ""
    }

    # WinRM設定取得
    try {
        # TrustedHosts
        $trustedHosts = (Get-Item WSMan:\localhost\Client\TrustedHosts -ErrorAction SilentlyContinue).Value
        $winrmData += [PSCustomObject]@{
            カテゴリ = "クライアント設定"
            項目 = "TrustedHosts"
            値 = if ($trustedHosts) { $trustedHosts } else { "(未設定)" }
            備考 = if ($trustedHosts -eq "*") { "全ホスト許可（セキュリティ注意）" } else { "" }
        }

        # 認証設定
        $authBasic = (Get-Item WSMan:\localhost\Client\Auth\Basic -ErrorAction SilentlyContinue).Value
        $authDigest = (Get-Item WSMan:\localhost\Client\Auth\Digest -ErrorAction SilentlyContinue).Value
        $authKerberos = (Get-Item WSMan:\localhost\Client\Auth\Kerberos -ErrorAction SilentlyContinue).Value
        $authNegotiate = (Get-Item WSMan:\localhost\Client\Auth\Negotiate -ErrorAction SilentlyContinue).Value
        $authCredSSP = (Get-Item WSMan:\localhost\Client\Auth\CredSSP -ErrorAction SilentlyContinue).Value

        $winrmData += [PSCustomObject]@{ カテゴリ = "クライアント認証"; 項目 = "Basic認証"; 値 = $authBasic; 備考 = if ($authBasic -eq "true") { "平文パスワード（非推奨）" } else { "" } }
        $winrmData += [PSCustomObject]@{ カテゴリ = "クライアント認証"; 項目 = "Digest認証"; 値 = $authDigest; 備考 = "" }
        $winrmData += [PSCustomObject]@{ カテゴリ = "クライアント認証"; 項目 = "Kerberos認証"; 値 = $authKerberos; 備考 = "ドメイン環境推奨" }
        $winrmData += [PSCustomObject]@{ カテゴリ = "クライアント認証"; 項目 = "Negotiate認証"; 値 = $authNegotiate; 備考 = "" }
        $winrmData += [PSCustomObject]@{ カテゴリ = "クライアント認証"; 項目 = "CredSSP認証"; 値 = $authCredSSP; 備考 = "" }

        # サービス側設定
        $svcAuthBasic = (Get-Item WSMan:\localhost\Service\Auth\Basic -ErrorAction SilentlyContinue).Value
        $svcAuthKerberos = (Get-Item WSMan:\localhost\Service\Auth\Kerberos -ErrorAction SilentlyContinue).Value
        $svcAuthNegotiate = (Get-Item WSMan:\localhost\Service\Auth\Negotiate -ErrorAction SilentlyContinue).Value
        $svcAllowUnencrypted = (Get-Item WSMan:\localhost\Service\AllowUnencrypted -ErrorAction SilentlyContinue).Value

        $winrmData += [PSCustomObject]@{ カテゴリ = "サービス認証"; 項目 = "Basic認証"; 値 = $svcAuthBasic; 備考 = if ($svcAuthBasic -eq "true") { "平文パスワード（非推奨）" } else { "" } }
        $winrmData += [PSCustomObject]@{ カテゴリ = "サービス認証"; 項目 = "Kerberos認証"; 値 = $svcAuthKerberos; 備考 = "" }
        $winrmData += [PSCustomObject]@{ カテゴリ = "サービス認証"; 項目 = "Negotiate認証"; 値 = $svcAuthNegotiate; 備考 = "" }
        $winrmData += [PSCustomObject]@{ カテゴリ = "サービス設定"; 項目 = "暗号化なし許可"; 値 = $svcAllowUnencrypted; 備考 = if ($svcAllowUnencrypted -eq "true") { "セキュリティリスク" } else { "" } }

        # リスナー設定
        $listeners = Get-ChildItem WSMan:\localhost\Listener -ErrorAction SilentlyContinue
        foreach ($listener in $listeners) {
            $address = (Get-Item "WSMan:\localhost\Listener\$($listener.Name)\Address" -ErrorAction SilentlyContinue).Value
            $transport = (Get-Item "WSMan:\localhost\Listener\$($listener.Name)\Transport" -ErrorAction SilentlyContinue).Value
            $port = (Get-Item "WSMan:\localhost\Listener\$($listener.Name)\Port" -ErrorAction SilentlyContinue).Value
            $enabled = (Get-Item "WSMan:\localhost\Listener\$($listener.Name)\Enabled" -ErrorAction SilentlyContinue).Value

            $winrmData += [PSCustomObject]@{
                カテゴリ = "リスナー"
                項目 = "リスナー ($($listener.Name))"
                値 = "Transport=$transport, Port=$port, Address=$address, Enabled=$enabled"
                備考 = if ($transport -eq "HTTP") { "ポート5985" } else { "ポート5986（HTTPS）" }
            }
        }

    } catch {
        $winrmData += [PSCustomObject]@{
            カテゴリ = "エラー"
            項目 = "WinRM設定取得"
            値 = "取得失敗"
            備考 = "WinRMサービスが停止している可能性があります"
        }
    }

    $collectedData["WinRM設定"] = $winrmData
    Write-Host "  [OK] WinRM設定: $($collectedData["WinRM設定"].Count) 件" -ForegroundColor Green
}
#endregion

#region ファイアウォール設定
if ($Config.CollectFirewallRules) {
    Write-Host "[収集中] ファイアウォール設定..." -ForegroundColor Cyan

    $firewallData = @()

    # ファイアウォールプロファイル状態
    try {
        $fwProfiles = Get-NetFirewallProfile -ErrorAction SilentlyContinue
        foreach ($profile in $fwProfiles) {
            $firewallData += [PSCustomObject]@{
                種別 = "プロファイル"
                名前 = $profile.Name
                有効 = if ($profile.Enabled) { "はい" } else { "いいえ" }
                既定の受信 = $profile.DefaultInboundAction.ToString()
                既定の送信 = $profile.DefaultOutboundAction.ToString()
                ポート = "-"
                プロトコル = "-"
                備考 = ""
            }
        }
    } catch {
        Write-Host "  [警告] ファイアウォールプロファイル取得に失敗" -ForegroundColor Yellow
    }

    # WinRM関連ルール（受信）
    try {
        $winrmRules = Get-NetFirewallRule -ErrorAction SilentlyContinue | Where-Object {
            $_.DisplayName -like "*WinRM*" -or
            $_.DisplayName -like "*Windows Remote Management*" -or
            $_.DisplayName -like "*リモート管理*"
        }

        foreach ($rule in $winrmRules) {
            $portFilter = Get-NetFirewallPortFilter -AssociatedNetFirewallRule $rule -ErrorAction SilentlyContinue
            $firewallData += [PSCustomObject]@{
                種別 = "WinRMルール"
                名前 = $rule.DisplayName
                有効 = if ($rule.Enabled -eq "True") { "はい" } else { "いいえ" }
                既定の受信 = $rule.Direction.ToString()
                既定の送信 = $rule.Action.ToString()
                ポート = $portFilter.LocalPort
                プロトコル = $portFilter.Protocol
                備考 = $rule.Profile
            }
        }

        # ポート5985/5986のルール
        $portRules = Get-NetFirewallPortFilter -ErrorAction SilentlyContinue | Where-Object {
            $_.LocalPort -eq "5985" -or $_.LocalPort -eq "5986"
        }

        foreach ($portRule in $portRules) {
            $rule = Get-NetFirewallRule -AssociatedNetFirewallPortFilter $portRule -ErrorAction SilentlyContinue
            if ($rule -and $rule.DisplayName -notlike "*WinRM*" -and $rule.DisplayName -notlike "*Windows Remote Management*") {
                $firewallData += [PSCustomObject]@{
                    種別 = "ポート5985/5986ルール"
                    名前 = $rule.DisplayName
                    有効 = if ($rule.Enabled -eq "True") { "はい" } else { "いいえ" }
                    既定の受信 = $rule.Direction.ToString()
                    既定の送信 = $rule.Action.ToString()
                    ポート = $portRule.LocalPort
                    プロトコル = $portRule.Protocol
                    備考 = ""
                }
            }
        }

    } catch {
        Write-Host "  [警告] ファイアウォールルール取得に失敗" -ForegroundColor Yellow
    }

    $collectedData["ファイアウォール"] = $firewallData
    Write-Host "  [OK] ファイアウォール設定: $($collectedData["ファイアウォール"].Count) 件" -ForegroundColor Green
}
#endregion

#region 開いているポート
if ($Config.CollectOpenPorts) {
    Write-Host "[収集中] 開いているポート..." -ForegroundColor Cyan

    $portData = @()

    try {
        $tcpConnections = Get-NetTCPConnection -State Listen -ErrorAction SilentlyContinue | Sort-Object LocalPort

        foreach ($conn in $tcpConnections) {
            $process = Get-Process -Id $conn.OwningProcess -ErrorAction SilentlyContinue
            $processName = if ($process) { $process.ProcessName } else { "不明" }

            # WinRM関連ポートをハイライト
            $remark = ""
            if ($conn.LocalPort -eq 5985) { $remark = "WinRM HTTP" }
            elseif ($conn.LocalPort -eq 5986) { $remark = "WinRM HTTPS" }
            elseif ($conn.LocalPort -eq 3389) { $remark = "RDP" }
            elseif ($conn.LocalPort -eq 445) { $remark = "SMB" }
            elseif ($conn.LocalPort -eq 135) { $remark = "RPC" }
            elseif ($conn.LocalPort -eq 139) { $remark = "NetBIOS" }

            $portData += [PSCustomObject]@{
                プロトコル = "TCP"
                ローカルアドレス = $conn.LocalAddress
                ローカルポート = $conn.LocalPort
                プロセス名 = $processName
                プロセスID = $conn.OwningProcess
                備考 = $remark
            }
        }
    } catch {
        Write-Host "  [警告] TCPポート取得に失敗" -ForegroundColor Yellow
    }

    $collectedData["開いているポート"] = $portData
    Write-Host "  [OK] 開いているポート: $($collectedData["開いているポート"].Count) 件" -ForegroundColor Green
}
#endregion

#region レジストリ設定
if ($Config.CollectRegistrySettings) {
    Write-Host "[収集中] セキュリティ関連レジストリ..." -ForegroundColor Cyan

    $registryData = @()

    # 確認するレジストリキー一覧
    $registryKeys = @(
        @{ Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"; Name = "LocalAccountTokenFilterPolicy"; Description = "リモートUAC（0=有効、1=無効）"; WinRMRelated = $true }
        @{ Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"; Name = "EnableLUA"; Description = "UAC有効（1=有効、0=無効）"; WinRMRelated = $false }
        @{ Path = "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa"; Name = "LmCompatibilityLevel"; Description = "LAN Manager認証レベル"; WinRMRelated = $true }
        @{ Path = "HKLM:\SYSTEM\CurrentControlSet\Services\WinRM"; Name = "Start"; Description = "WinRMサービス起動種別（2=自動、3=手動、4=無効）"; WinRMRelated = $true }
        @{ Path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WinRM\Client"; Name = "AllowBasic"; Description = "Basic認証許可（クライアント）"; WinRMRelated = $true }
        @{ Path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WinRM\Service"; Name = "AllowBasic"; Description = "Basic認証許可（サービス）"; WinRMRelated = $true }
        @{ Path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WinRM\Service"; Name = "AllowUnencryptedTraffic"; Description = "暗号化なしトラフィック許可"; WinRMRelated = $true }
        @{ Path = "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server"; Name = "fDenyTSConnections"; Description = "RDP接続拒否（0=許可、1=拒否）"; WinRMRelated = $false }
        @{ Path = "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp"; Name = "PortNumber"; Description = "RDPポート番号"; WinRMRelated = $false }
        @{ Path = "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters"; Name = "AutoShareServer"; Description = "管理共有自動作成"; WinRMRelated = $false }
        @{ Path = "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters"; Name = "AutoShareWks"; Description = "管理共有自動作成（WKS）"; WinRMRelated = $false }
    )

    foreach ($regKey in $registryKeys) {
        $value = $null
        $exists = $false

        try {
            if (Test-Path $regKey.Path) {
                $regValue = Get-ItemProperty -Path $regKey.Path -Name $regKey.Name -ErrorAction SilentlyContinue
                if ($null -ne $regValue) {
                    $value = $regValue.($regKey.Name)
                    $exists = $true
                }
            }
        } catch {
            $value = "取得エラー"
        }

        $registryData += [PSCustomObject]@{
            パス = $regKey.Path
            名前 = $regKey.Name
            値 = if ($exists) { $value } else { "(未設定)" }
            説明 = $regKey.Description
            WinRM関連 = if ($regKey.WinRMRelated) { "はい" } else { "いいえ" }
        }
    }

    $collectedData["レジストリ設定"] = $registryData
    Write-Host "  [OK] レジストリ設定: $($collectedData["レジストリ設定"].Count) 件" -ForegroundColor Green
}
#endregion

#region Windowsサービス
if ($Config.CollectServices) {
    Write-Host "[収集中] 関連Windowsサービス..." -ForegroundColor Cyan

    # WinRM/リモート管理関連サービス
    $relatedServices = @(
        "WinRM", "RemoteRegistry", "RemoteAccess", "RpcSs", "RpcEptMapper",
        "LanmanServer", "LanmanWorkstation", "Netlogon", "TermService",
        "WinHttpAutoProxySvc", "iphlpsvc", "PolicyAgent", "IKEEXT"
    )

    $serviceData = @()

    foreach ($svcName in $relatedServices) {
        $svc = Get-Service -Name $svcName -ErrorAction SilentlyContinue
        if ($svc) {
            $serviceData += [PSCustomObject]@{
                サービス名 = $svc.Name
                表示名 = $svc.DisplayName
                状態 = $svc.Status.ToString()
                スタートアップ種別 = $svc.StartType.ToString()
                備考 = switch ($svc.Name) {
                    "WinRM" { "Windows Remote Management" }
                    "RemoteRegistry" { "リモートレジストリ" }
                    "TermService" { "リモートデスクトップ" }
                    "LanmanServer" { "ファイル共有（Server）" }
                    "LanmanWorkstation" { "ファイル共有（Client）" }
                    default { "" }
                }
            }
        }
    }

    $collectedData["関連サービス"] = $serviceData
    Write-Host "  [OK] 関連サービス: $($collectedData["関連サービス"].Count) 件" -ForegroundColor Green
}
#endregion

#region ローカルユーザー/グループ
if ($Config.CollectLocalUsers) {
    Write-Host "[収集中] ローカルユーザー/グループ..." -ForegroundColor Cyan

    $userData = @()

    try {
        $localUsers = Get-LocalUser -ErrorAction SilentlyContinue
        foreach ($user in $localUsers) {
            $userData += [PSCustomObject]@{
                種別 = "ローカルユーザー"
                名前 = $user.Name
                有効 = if ($user.Enabled) { "はい" } else { "いいえ" }
                説明 = $user.Description
                最終ログオン = if ($user.LastLogon) { $user.LastLogon.ToString("yyyy-MM-dd HH:mm:ss") } else { "-" }
                パスワード期限切れ = if ($user.PasswordExpires) { $user.PasswordExpires.ToString("yyyy-MM-dd") } else { "なし" }
            }
        }

        # 管理者グループのメンバー
        $adminGroup = Get-LocalGroupMember -Group "Administrators" -ErrorAction SilentlyContinue
        foreach ($member in $adminGroup) {
            $userData += [PSCustomObject]@{
                種別 = "Administratorsメンバー"
                名前 = $member.Name
                有効 = "-"
                説明 = $member.ObjectClass
                最終ログオン = "-"
                パスワード期限切れ = "-"
            }
        }

        # Remote Desktop Usersグループのメンバー
        $rdpGroup = Get-LocalGroupMember -Group "Remote Desktop Users" -ErrorAction SilentlyContinue
        foreach ($member in $rdpGroup) {
            $userData += [PSCustomObject]@{
                種別 = "Remote Desktop Usersメンバー"
                名前 = $member.Name
                有効 = "-"
                説明 = $member.ObjectClass
                最終ログオン = "-"
                パスワード期限切れ = "-"
            }
        }

    } catch {
        Write-Host "  [警告] ローカルユーザー/グループ取得に失敗" -ForegroundColor Yellow
    }

    $collectedData["ユーザー_グループ"] = $userData
    Write-Host "  [OK] ユーザー/グループ: $($collectedData["ユーザー_グループ"].Count) 件" -ForegroundColor Green
}
#endregion

#region Excel出力
Write-Host ""
Write-Host "========================================================================" -ForegroundColor Yellow
Write-Host " Excel出力中..." -ForegroundColor Yellow
Write-Host "========================================================================" -ForegroundColor Yellow
Write-Host ""

# Excelが利用可能か確認
$excelAvailable = $false
try {
    $excelApp = New-Object -ComObject Excel.Application -ErrorAction Stop
    $excelAvailable = $true
    $excelApp.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApp) | Out-Null
} catch {
    $excelAvailable = $false
}

if ($excelAvailable) {
    Write-Host "Excelを使用して出力します..." -ForegroundColor Cyan

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Add()

        # 既存のシートを削除（Sheet1以外）
        while ($workbook.Sheets.Count -gt 1) {
            $workbook.Sheets.Item($workbook.Sheets.Count).Delete()
        }

        $sheetIndex = 0
        foreach ($sheetName in $collectedData.Keys) {
            $data = $collectedData[$sheetName]

            if ($data.Count -eq 0) { continue }

            if ($sheetIndex -eq 0) {
                $sheet = $workbook.Sheets.Item(1)
                $sheet.Name = $sheetName.Substring(0, [Math]::Min(31, $sheetName.Length))
            } else {
                $sheet = $workbook.Sheets.Add([System.Reflection.Missing]::Value, $workbook.Sheets.Item($workbook.Sheets.Count))
                $sheet.Name = $sheetName.Substring(0, [Math]::Min(31, $sheetName.Length))
            }

            # ヘッダー出力
            $properties = $data[0].PSObject.Properties.Name
            $col = 1
            foreach ($prop in $properties) {
                $sheet.Cells.Item(1, $col) = $prop
                $sheet.Cells.Item(1, $col).Font.Bold = $true
                $sheet.Cells.Item(1, $col).Interior.ColorIndex = 15
                $col++
            }

            # データ出力
            $row = 2
            foreach ($item in $data) {
                $col = 1
                foreach ($prop in $properties) {
                    $value = $item.$prop
                    $sheet.Cells.Item($row, $col) = if ($null -ne $value) { $value.ToString() } else { "" }
                    $col++
                }
                $row++
            }

            # 列幅自動調整
            $sheet.UsedRange.Columns.AutoFit() | Out-Null

            $sheetIndex++
        }

        # 保存
        $workbook.SaveAs($excelPath, 51) # 51 = xlOpenXMLWorkbook (.xlsx)
        $workbook.Close()
        $excel.Quit()

        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()

        Write-Host "[OK] Excel出力完了: $excelPath" -ForegroundColor Green

    } catch {
        Write-Host "[エラー] Excel出力に失敗: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "CSV形式で出力します..." -ForegroundColor Yellow
        $excelAvailable = $false
    }
}

# Excelが使えない場合はCSV出力
if (-not $excelAvailable) {
    Write-Host "CSVで出力します（Excelで開けます）..." -ForegroundColor Cyan

    New-Item -ItemType Directory -Path $csvFolder -Force | Out-Null

    foreach ($sheetName in $collectedData.Keys) {
        $data = $collectedData[$sheetName]
        if ($data.Count -gt 0) {
            $csvPath = Join-Path $csvFolder "$sheetName.csv"
            $data | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
            Write-Host "  [OK] $csvPath" -ForegroundColor Green
        }
    }

    Write-Host ""
    Write-Host "[OK] CSV出力完了: $csvFolder" -ForegroundColor Green
}
#endregion

#region 完了
Write-Host ""
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host " 収集完了" -ForegroundColor Cyan
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "収集日時: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor White
Write-Host "コンピュータ: $computerName" -ForegroundColor White
Write-Host ""

foreach ($key in $collectedData.Keys) {
    Write-Host "  $key : $($collectedData[$key].Count) 件" -ForegroundColor Gray
}

Write-Host ""

# 出力ファイルを開くか確認
$openFile = Read-Host "出力ファイルを開きますか？ (y/n)"
if ($openFile -eq "y") {
    if ($excelAvailable -and (Test-Path $excelPath)) {
        Start-Process $excelPath
    } elseif (Test-Path $csvFolder) {
        Start-Process explorer.exe $csvFolder
    }
}
#endregion

Write-Host ""
exit 0
