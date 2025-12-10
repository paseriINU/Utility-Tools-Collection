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
    Windowsサーバ構成情報収集ツール（完全版）

.DESCRIPTION
    ローカルサーバの構成情報を包括的に収集してExcelファイルに出力します。

.NOTES
    作成日: 2025-12-10
    バージョン: 2.0

    収集項目:
    - OS基本情報
    - ハードウェア情報（CPU、メモリ、ディスク）
    - ネットワーク設定（IP、DNS、ゲートウェイ）
    - WinRM設定（サービス状態、TrustedHosts、認証設定）
    - ファイアウォール設定
    - 開いているポート（LISTENING）
    - セキュリティ関連レジストリ
    - Windowsサービス一覧（全サービス）
    - ローカルユーザー/グループ
    - インストール済みソフトウェア
    - 共有フォルダ
    - タスクスケジューラ
#>

# ==============================================================================
# ■ 設定セクション（ここを編集してください）
# ==============================================================================

$Config = @{
    # 出力先フォルダ（空の場合はデスクトップ）
    OutputFolder = ""

    # 収集項目の有効/無効
    CollectOSInfo = $true
    CollectHardwareInfo = $true
    CollectNetworkConfig = $true
    CollectWinRMConfig = $true
    CollectFirewallRules = $true
    CollectOpenPorts = $true
    CollectRegistrySettings = $true
    CollectAllServices = $true
    CollectLocalUsers = $true
    CollectInstalledSoftware = $true
    CollectSharedFolders = $true
    CollectScheduledTasks = $true
}

# ==============================================================================
# ■ メイン処理（以下は編集不要）
# ==============================================================================

$ErrorActionPreference = "Continue"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ヘッダー表示
Write-Host ""
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host "  サーバ構成情報収集ツール（完全版）" -ForegroundColor Cyan
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

# データ収集用ハッシュテーブル（順序付き）
$collectedData = [ordered]@{}

#region OS基本情報
if ($Config.CollectOSInfo) {
    Write-Host "[収集中] OS基本情報..." -ForegroundColor Cyan

    $osInfo = Get-CimInstance -ClassName Win32_OperatingSystem
    $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem

    $collectedData["01_OS情報"] = @(
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
        [PSCustomObject]@{ 項目 = "システムディレクトリ"; 値 = $osInfo.SystemDirectory }
        [PSCustomObject]@{ 項目 = "Windowsディレクトリ"; 値 = $osInfo.WindowsDirectory }
        [PSCustomObject]@{ 項目 = "シリアル番号"; 値 = $osInfo.SerialNumber }
    )

    Write-Host "  [OK] OS基本情報: $($collectedData["01_OS情報"].Count) 件" -ForegroundColor Green
}
#endregion

#region ハードウェア情報
if ($Config.CollectHardwareInfo) {
    Write-Host "[収集中] ハードウェア情報..." -ForegroundColor Cyan

    $hardwareData = @()

    # CPU情報
    $cpuInfo = Get-CimInstance -ClassName Win32_Processor
    foreach ($cpu in $cpuInfo) {
        $hardwareData += [PSCustomObject]@{ カテゴリ = "CPU"; 項目 = "プロセッサ名"; 値 = $cpu.Name }
        $hardwareData += [PSCustomObject]@{ カテゴリ = "CPU"; 項目 = "コア数"; 値 = $cpu.NumberOfCores }
        $hardwareData += [PSCustomObject]@{ カテゴリ = "CPU"; 項目 = "論理プロセッサ数"; 値 = $cpu.NumberOfLogicalProcessors }
        $hardwareData += [PSCustomObject]@{ カテゴリ = "CPU"; 項目 = "最大クロック(MHz)"; 値 = $cpu.MaxClockSpeed }
        $hardwareData += [PSCustomObject]@{ カテゴリ = "CPU"; 項目 = "ソケット"; 値 = $cpu.SocketDesignation }
    }

    # メモリ情報
    $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
    $osInfo = Get-CimInstance -ClassName Win32_OperatingSystem
    $hardwareData += [PSCustomObject]@{ カテゴリ = "メモリ"; 項目 = "物理メモリ合計(GB)"; 値 = [math]::Round($computerSystem.TotalPhysicalMemory / 1GB, 2) }
    $hardwareData += [PSCustomObject]@{ カテゴリ = "メモリ"; 項目 = "利用可能メモリ(GB)"; 値 = [math]::Round($osInfo.FreePhysicalMemory / 1MB, 2) }
    $hardwareData += [PSCustomObject]@{ カテゴリ = "メモリ"; 項目 = "仮想メモリ合計(GB)"; 値 = [math]::Round($osInfo.TotalVirtualMemorySize / 1MB, 2) }
    $hardwareData += [PSCustomObject]@{ カテゴリ = "メモリ"; 項目 = "利用可能仮想メモリ(GB)"; 値 = [math]::Round($osInfo.FreeVirtualMemory / 1MB, 2) }

    # メモリスロット情報
    $memorySlots = Get-CimInstance -ClassName Win32_PhysicalMemory -ErrorAction SilentlyContinue
    $slotNum = 1
    foreach ($slot in $memorySlots) {
        $hardwareData += [PSCustomObject]@{ カテゴリ = "メモリスロット$slotNum"; 項目 = "容量(GB)"; 値 = [math]::Round($slot.Capacity / 1GB, 2) }
        $hardwareData += [PSCustomObject]@{ カテゴリ = "メモリスロット$slotNum"; 項目 = "速度(MHz)"; 値 = $slot.Speed }
        $hardwareData += [PSCustomObject]@{ カテゴリ = "メモリスロット$slotNum"; 項目 = "製造元"; 値 = $slot.Manufacturer }
        $slotNum++
    }

    $collectedData["02_ハードウェア"] = $hardwareData

    # ディスク情報（別シート）
    $diskData = @()
    $disks = Get-CimInstance -ClassName Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 }
    foreach ($disk in $disks) {
        $usedSpace = $disk.Size - $disk.FreeSpace
        $usagePercent = if ($disk.Size -gt 0) { [math]::Round(($usedSpace / $disk.Size) * 100, 1) } else { 0 }
        $diskData += [PSCustomObject]@{
            ドライブ = $disk.DeviceID
            ボリューム名 = $disk.VolumeName
            ファイルシステム = $disk.FileSystem
            "合計容量(GB)" = [math]::Round($disk.Size / 1GB, 2)
            "使用容量(GB)" = [math]::Round($usedSpace / 1GB, 2)
            "空き容量(GB)" = [math]::Round($disk.FreeSpace / 1GB, 2)
            "使用率(%)" = $usagePercent
            状態 = if ($usagePercent -gt 90) { "警告:容量不足" } elseif ($usagePercent -gt 80) { "注意" } else { "正常" }
        }
    }
    $collectedData["03_ディスク"] = $diskData

    Write-Host "  [OK] ハードウェア情報: $($collectedData["02_ハードウェア"].Count) 件" -ForegroundColor Green
    Write-Host "  [OK] ディスク情報: $($collectedData["03_ディスク"].Count) 件" -ForegroundColor Green
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
            DNSドメイン = $adapter.DNSDomain
        }
    }

    $collectedData["04_ネットワーク設定"] = $networkData

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
        $collectedData["05_ネットワークプロファイル"] = $profileData
    } catch {
        Write-Host "  [警告] ネットワークプロファイル取得に失敗" -ForegroundColor Yellow
    }

    Write-Host "  [OK] ネットワーク設定: $($collectedData["04_ネットワーク設定"].Count) 件" -ForegroundColor Green
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

    $collectedData["06_WinRM設定"] = $winrmData
    Write-Host "  [OK] WinRM設定: $($collectedData["06_WinRM設定"].Count) 件" -ForegroundColor Green
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

    # WinRM関連ルール
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

    $collectedData["07_ファイアウォール"] = $firewallData
    Write-Host "  [OK] ファイアウォール設定: $($collectedData["07_ファイアウォール"].Count) 件" -ForegroundColor Green
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

            # 主要ポートの備考
            $remark = switch ($conn.LocalPort) {
                5985 { "WinRM HTTP" }
                5986 { "WinRM HTTPS" }
                3389 { "RDP" }
                445 { "SMB" }
                135 { "RPC" }
                139 { "NetBIOS" }
                80 { "HTTP" }
                443 { "HTTPS" }
                22 { "SSH" }
                21 { "FTP" }
                25 { "SMTP" }
                53 { "DNS" }
                1433 { "SQL Server" }
                1521 { "Oracle" }
                3306 { "MySQL" }
                5432 { "PostgreSQL" }
                default { "" }
            }

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

    $collectedData["08_開いているポート"] = $portData
    Write-Host "  [OK] 開いているポート: $($collectedData["08_開いているポート"].Count) 件" -ForegroundColor Green
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

    $collectedData["09_レジストリ設定"] = $registryData
    Write-Host "  [OK] レジストリ設定: $($collectedData["09_レジストリ設定"].Count) 件" -ForegroundColor Green
}
#endregion

#region Windowsサービス（全サービス）
if ($Config.CollectAllServices) {
    Write-Host "[収集中] Windowsサービス（全サービス）..." -ForegroundColor Cyan

    $serviceData = @()

    $allServices = Get-Service | Sort-Object Name
    foreach ($svc in $allServices) {
        $serviceData += [PSCustomObject]@{
            サービス名 = $svc.Name
            表示名 = $svc.DisplayName
            状態 = $svc.Status.ToString()
            スタートアップ種別 = $svc.StartType.ToString()
        }
    }

    $collectedData["10_Windowsサービス"] = $serviceData
    Write-Host "  [OK] Windowsサービス: $($collectedData["10_Windowsサービス"].Count) 件" -ForegroundColor Green
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

    $collectedData["11_ユーザー_グループ"] = $userData
    Write-Host "  [OK] ユーザー/グループ: $($collectedData["11_ユーザー_グループ"].Count) 件" -ForegroundColor Green
}
#endregion

#region インストール済みソフトウェア
if ($Config.CollectInstalledSoftware) {
    Write-Host "[収集中] インストール済みソフトウェア..." -ForegroundColor Cyan

    $softwareData = @()

    try {
        # 32bit/64bitの両方から取得
        $regPaths = @(
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
            "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
        )

        foreach ($regPath in $regPaths) {
            $software = Get-ItemProperty $regPath -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName }
            foreach ($sw in $software) {
                # 重複チェック
                if (-not ($softwareData | Where-Object { $_.名前 -eq $sw.DisplayName })) {
                    $softwareData += [PSCustomObject]@{
                        名前 = $sw.DisplayName
                        バージョン = $sw.DisplayVersion
                        発行元 = $sw.Publisher
                        インストール日 = $sw.InstallDate
                        インストール場所 = $sw.InstallLocation
                    }
                }
            }
        }

        $softwareData = $softwareData | Sort-Object 名前
    } catch {
        Write-Host "  [警告] ソフトウェア一覧取得に失敗" -ForegroundColor Yellow
    }

    $collectedData["12_インストール済ソフト"] = $softwareData
    Write-Host "  [OK] インストール済みソフトウェア: $($collectedData["12_インストール済ソフト"].Count) 件" -ForegroundColor Green
}
#endregion

#region 共有フォルダ
if ($Config.CollectSharedFolders) {
    Write-Host "[収集中] 共有フォルダ..." -ForegroundColor Cyan

    $shareData = @()

    try {
        $shares = Get-CimInstance -ClassName Win32_Share
        foreach ($share in $shares) {
            $shareData += [PSCustomObject]@{
                共有名 = $share.Name
                パス = $share.Path
                説明 = $share.Description
                種別 = switch ($share.Type) {
                    0 { "ディスクドライブ" }
                    1 { "プリンター" }
                    2 { "デバイス" }
                    3 { "IPC" }
                    2147483648 { "管理共有(ディスク)" }
                    2147483649 { "管理共有(プリンター)" }
                    2147483650 { "管理共有(デバイス)" }
                    2147483651 { "管理共有(IPC)" }
                    default { $share.Type }
                }
                状態 = $share.Status
            }
        }
    } catch {
        Write-Host "  [警告] 共有フォルダ取得に失敗" -ForegroundColor Yellow
    }

    $collectedData["13_共有フォルダ"] = $shareData
    Write-Host "  [OK] 共有フォルダ: $($collectedData["13_共有フォルダ"].Count) 件" -ForegroundColor Green
}
#endregion

#region タスクスケジューラ
if ($Config.CollectScheduledTasks) {
    Write-Host "[収集中] タスクスケジューラ..." -ForegroundColor Cyan

    $taskData = @()

    try {
        $tasks = Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object { $_.State -ne "Disabled" -and $_.TaskPath -notlike "\Microsoft\*" }
        foreach ($task in $tasks) {
            $taskInfo = Get-ScheduledTaskInfo -TaskName $task.TaskName -TaskPath $task.TaskPath -ErrorAction SilentlyContinue
            $taskData += [PSCustomObject]@{
                タスク名 = $task.TaskName
                パス = $task.TaskPath
                状態 = $task.State.ToString()
                最終実行 = if ($taskInfo.LastRunTime -and $taskInfo.LastRunTime -ne [DateTime]::MinValue) { $taskInfo.LastRunTime.ToString("yyyy-MM-dd HH:mm:ss") } else { "-" }
                最終結果 = $taskInfo.LastTaskResult
                次回実行 = if ($taskInfo.NextRunTime -and $taskInfo.NextRunTime -ne [DateTime]::MinValue) { $taskInfo.NextRunTime.ToString("yyyy-MM-dd HH:mm:ss") } else { "-" }
                説明 = $task.Description
            }
        }
    } catch {
        Write-Host "  [警告] タスクスケジューラ取得に失敗" -ForegroundColor Yellow
    }

    $collectedData["14_タスクスケジューラ"] = $taskData
    Write-Host "  [OK] タスクスケジューラ: $($collectedData["14_タスクスケジューラ"].Count) 件" -ForegroundColor Green
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

            if ($null -eq $data -or $data.Count -eq 0) { continue }

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
            Write-Host "  [OK] $sheetName" -ForegroundColor Gray
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

        Write-Host ""
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
        if ($null -ne $data -and $data.Count -gt 0) {
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
    $count = if ($null -ne $collectedData[$key]) { $collectedData[$key].Count } else { 0 }
    Write-Host "  $key : $count 件" -ForegroundColor Gray
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
