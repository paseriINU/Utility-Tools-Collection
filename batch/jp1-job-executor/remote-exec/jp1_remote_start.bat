<# :
@echo off
setlocal
chcp 65001 >nul
powershell -NoProfile -ExecutionPolicy Bypass -Command "iex ((gc '%~f0') -join \"`n\")"
exit /b %ERRORLEVEL%
: #> | sv -name _ > $null

<#
.SYNOPSIS
    JP1ジョブネット起動ツール（リモート実行版・ハイブリッド）

.DESCRIPTION
    PowerShell Remotingを使用してリモートサーバでajsentryを実行します。
    WinRM設定を自動的に構成・復元します。

.NOTES
    作成日: 2025-12-02
    バージョン: 3.0

    使い方:
    1. 下記の「■ 設定セクション」を編集
    2. このファイルをダブルクリックで実行
#>

# ==============================================================================
# ■ 設定セクション（ここを編集してください）
# ==============================================================================

$Config = @{
    # JP1/AJS3が稼働しているリモートサーバ
    JP1Server = "192.168.1.100"

    # リモートサーバのユーザー名（Windowsログインユーザー）
    RemoteUser = "Administrator"

    # リモートサーバのパスワード（空の場合は実行時に入力）
    RemotePassword = ""

    # JP1ユーザー名
    JP1User = "jp1admin"

    # JP1パスワード（空の場合は実行時に入力）
    JP1Password = ""

    # 起動するジョブネットのフルパス
    JobnetPath = "/main_unit/jobgroup1/daily_batch"

    # ajsentryコマンドのパス（リモートサーバ上）
    AjsentryPath = "C:\Program Files\HITACHI\JP1AJS3\bin\ajsentry.exe"

    # HTTPS接続を使用する場合は $true
    UseSSL = $false
}

# ==============================================================================
# ■ メイン処理（以下は編集不要）
# ==============================================================================

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Clear-Host

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "JP1ジョブネット起動ツール" -ForegroundColor Cyan
Write-Host "（リモート実行版）" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

#region WinRM設定の保存と自動設定
$originalTrustedHosts = $null
$winrmConfigChanged = $false

try {
    Write-Host "WinRM設定を確認中..." -ForegroundColor Cyan

    # 現在のTrustedHostsを取得
    try {
        $originalTrustedHosts = (Get-Item WSMan:\localhost\Client\TrustedHosts -ErrorAction SilentlyContinue).Value
        Write-Verbose "現在のTrustedHosts: $originalTrustedHosts"
    } catch {
        $originalTrustedHosts = ""
        Write-Verbose "TrustedHostsは未設定です"
    }

    # WinRMサービスの起動確認
    $winrmService = Get-Service -Name WinRM -ErrorAction SilentlyContinue
    if ($winrmService.Status -ne 'Running') {
        Write-Host "WinRMサービスを起動中..." -ForegroundColor Yellow
        Start-Service -Name WinRM -ErrorAction Stop
        Write-Host "[OK] WinRMサービスを起動しました" -ForegroundColor Green
    } else {
        Write-Host "[OK] WinRMサービスは起動済みです" -ForegroundColor Green
    }

    # 接続先がTrustedHostsに含まれているか確認
    $needsConfig = $true
    if ($originalTrustedHosts) {
        $trustedList = $originalTrustedHosts -split ','
        if ($trustedList -contains $Config.JP1Server -or $trustedList -contains '*') {
            Write-Host "[OK] 接続先は既にTrustedHostsに登録されています" -ForegroundColor Green
            $needsConfig = $false
        }
    }

    # 必要に応じてTrustedHostsに追加
    if ($needsConfig) {
        Write-Host "接続先をTrustedHostsに追加中..." -ForegroundColor Yellow

        if ([string]::IsNullOrEmpty($originalTrustedHosts)) {
            Set-Item WSMan:\localhost\Client\TrustedHosts -Value $Config.JP1Server -Force
        } else {
            Set-Item WSMan:\localhost\Client\TrustedHosts -Value "$originalTrustedHosts,$($Config.JP1Server)" -Force
        }

        $winrmConfigChanged = $true
        Write-Host "[OK] TrustedHostsに追加しました: $($Config.JP1Server)" -ForegroundColor Green
    }

    Write-Host ""
} catch {
    Write-Host "[警告] WinRM設定の自動構成に失敗しました" -ForegroundColor Yellow
    Write-Host "エラー: $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host "手動でWinRM設定を行ってください" -ForegroundColor Yellow
    Write-Host ""
}
#endregion

Write-Host "JP1サーバ      : $($Config.JP1Server)" -ForegroundColor White
Write-Host "リモートユーザー: $($Config.RemoteUser)" -ForegroundColor White
Write-Host "JP1ユーザー    : $($Config.JP1User)" -ForegroundColor White
Write-Host "ジョブネットパス: $($Config.JobnetPath)" -ForegroundColor White
Write-Host ""

# JP1パスワード入力
if ([string]::IsNullOrEmpty($Config.JP1Password)) {
    Write-Host "[注意] JP1パスワードが設定されていません。" -ForegroundColor Yellow
    $Config.JP1Password = Read-Host "JP1パスワードを入力してください" -AsSecureString
    $Config.JP1Password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($Config.JP1Password))
    Write-Host ""
}

# 実行確認
Write-Host "ジョブネットを起動しますか？ (Y/N)" -ForegroundColor Cyan
$confirmation = Read-Host
if ($confirmation -ne "Y" -and $confirmation -ne "y") {
    Write-Host "処理をキャンセルしました。" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Enterキーを押して終了..." -ForegroundColor Gray
    $null = Read-Host
    exit 0
}

Write-Host ""

#region 認証情報の準備
if ($Config.RemoteUser) {
    if ($Config.RemotePassword) {
        $securePassword = ConvertTo-SecureString $Config.RemotePassword -AsPlainText -Force
        $Credential = New-Object System.Management.Automation.PSCredential($Config.RemoteUser, $securePassword)
    } else {
        Write-Host "リモートサーバの認証情報を入力してください" -ForegroundColor Cyan
        $Credential = Get-Credential -UserName $Config.RemoteUser -Message "JP1サーバ($($Config.JP1Server))の認証情報を入力"
    }
} else {
    Write-Host "リモートサーバの認証情報を入力してください" -ForegroundColor Cyan
    $Credential = Get-Credential -Message "JP1サーバ($($Config.JP1Server))の認証情報を入力"
}

if ($null -eq $Credential) {
    Write-Host "[エラー] 認証情報の入力がキャンセルされました。" -ForegroundColor Red
    Write-Host ""
    Write-Host "Enterキーを押して終了..." -ForegroundColor Gray
    $null = Read-Host
    exit 1
}
#endregion

#region セッションオプションの設定
$sessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck
#endregion

Write-Host "========================================" -ForegroundColor Yellow
Write-Host "リモート接続してジョブネット起動中..." -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

# メイン処理（WinRM設定復元用のfinallyブロック付き）
try {
    Write-Host "リモートサーバに接続中..." -ForegroundColor Cyan

    $sessionParams = @{
        ComputerName = $Config.JP1Server
        Credential = $Credential
        SessionOption = $sessionOption
        ErrorAction = "Stop"
    }

    if ($Config.UseSSL) {
        $sessionParams.UseSSL = $true
    }

    $session = New-PSSession @sessionParams
    Write-Host "[OK] 接続成功" -ForegroundColor Green
    Write-Host ""

    Write-Host "ajsentryコマンドを実行中..." -ForegroundColor Cyan

    $scriptBlock = {
        param($ajsPath, $jp1User, $jp1Pass, $jobnetPath)

        if (-not (Test-Path $ajsPath)) {
            throw "ajsentryが見つかりません: $ajsPath"
        }

        $output = & $ajsPath -h localhost -u $jp1User -p $jp1Pass -F $jobnetPath 2>&1
        $exitCode = $LASTEXITCODE

        @{
            ExitCode = $exitCode
            Output = $output
        }
    }

    $result = Invoke-Command -Session $session -ScriptBlock $scriptBlock -ArgumentList $Config.AjsentryPath, $Config.JP1User, $Config.JP1Password, $Config.JobnetPath

    Remove-PSSession -Session $session

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    if ($result.ExitCode -eq 0) {
        Write-Host "ジョブネットの起動に成功しました" -ForegroundColor Green
    } else {
        Write-Host "ジョブネットの起動に失敗しました" -ForegroundColor Red
        Write-Host "エラーコード: $($result.ExitCode)" -ForegroundColor Red
    }
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""

    Write-Host "実行結果:" -ForegroundColor White
    $result.Output | ForEach-Object {
        Write-Host $_ -ForegroundColor White
    }

    if ($result.ExitCode -eq 0) {
        Write-Host ""
        Write-Host "ジョブネット: $($Config.JobnetPath)" -ForegroundColor Green
        Write-Host "サーバ      : $($Config.JP1Server)" -ForegroundColor Green
    } else {
        Write-Host ""
        Write-Host "追加の確認事項：" -ForegroundColor Yellow
        Write-Host "- ajsentryのパスが正しいか: $($Config.AjsentryPath)" -ForegroundColor Yellow
        Write-Host "- JP1ユーザー名、パスワードが正しいか" -ForegroundColor Yellow
        Write-Host "- ジョブネットパスが正しいか" -ForegroundColor Yellow
        Write-Host "- JP1/AJS3サービスが起動しているか" -ForegroundColor Yellow
    }

    $exitCode = $result.ExitCode

} catch {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "[エラー] リモート実行に失敗しました" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "エラー詳細:" -ForegroundColor Yellow
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""

    Write-Host "以下を確認してください：" -ForegroundColor Yellow
    Write-Host "- リモートサーバのWinRMサービスが有効か" -ForegroundColor Yellow
    Write-Host "- PowerShell Remotingが有効か（Enable-PSRemoting）" -ForegroundColor Yellow
    Write-Host "- ファイアウォールで5985/5986ポートが開いているか" -ForegroundColor Yellow
    Write-Host "- ネットワーク接続が正常か" -ForegroundColor Yellow

    if ($session) {
        Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    }

    $exitCode = 1
}
} finally {
    #region WinRM設定の復元
    if ($winrmConfigChanged) {
        Write-Host ""
        Write-Host "WinRM設定を復元中..." -ForegroundColor Cyan

        try {
            if ([string]::IsNullOrEmpty($originalTrustedHosts)) {
                Set-Item WSMan:\localhost\Client\TrustedHosts -Value "" -Force
                Write-Host "[OK] TrustedHostsを元の状態（空）に復元しました" -ForegroundColor Green
            } else {
                Set-Item WSMan:\localhost\Client\TrustedHosts -Value $originalTrustedHosts -Force
                Write-Host "[OK] TrustedHostsを元の状態に復元しました" -ForegroundColor Green
            }
        } catch {
            Write-Host "[警告] TrustedHostsの復元に失敗しました" -ForegroundColor Yellow
            Write-Host "エラー: $($_.Exception.Message)" -ForegroundColor Yellow
            Write-Host "手動で復元してください: Set-Item WSMan:\localhost\Client\TrustedHosts -Value '$originalTrustedHosts' -Force" -ForegroundColor Yellow
        }
    }
    #endregion
}

Write-Host ""
Write-Host "Enterキーを押して終了..." -ForegroundColor Gray
$null = Read-Host

exit $exitCode
