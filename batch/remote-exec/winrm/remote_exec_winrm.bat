<# :
@echo off
setlocal

rem 管理者権限チェック
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo 管理者権限が必要です。管理者として再起動します...
    powershell -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

powershell -NoProfile -ExecutionPolicy Bypass -Command "$scriptDir=('%~dp0' -replace '\\$',''); iex ((gc '%~f0') -join \"`n\")"
set EXITCODE=%ERRORLEVEL%
pause
exit /b %EXITCODE%
: #>

<#
.SYNOPSIS
    リモートWindowsサーバ上でバッチファイルをWinRMで実行（ハイブリッド版）

.DESCRIPTION
    PowerShellのInvoke-Commandを使用してリモート実行し、結果を取得します。
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
    # リモートサーバのコンピュータ名またはIPアドレス
    ComputerName = "192.168.1.100"

    # リモートサーバの管理者ユーザー名
    UserName = "Administrator"

    # パスワード（空の場合は実行時に入力を求められます）
    Password = ""

    # リモートサーバで実行するバッチファイルのフルパス
    BatchPath = "C:\Scripts\target_script.bat"

    # HTTPS接続を使用する場合は $true
    UseSSL = $false
}

# ==============================================================================
# ■ メイン処理（以下は編集不要）
# ==============================================================================

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# 出力ログファイル名を自動生成（日時付き）
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
if (-not $scriptDir) {
    $scriptDir = $PSScriptRoot
    if (-not $scriptDir) {
        $scriptDir = (Get-Location).Path
    }
}

# ネットワークパス対応: UNCパスの場合はローカルのTempフォルダを使用
if ($scriptDir -like "\\*") {
    $logDir = "$env:TEMP\RemoteExecLogs"
} else {
    $logDir = "$scriptDir\log"
}

# logフォルダが存在しない場合は作成
if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}

$OutputLog = "$logDir\remote_exec_output_$timestamp.log"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "リモートバッチ実行ツール (WinRM版)" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

#region WinRM設定の保存と自動設定
$originalTrustedHosts = $null
$winrmConfigChanged = $false
$winrmServiceWasStarted = $false

try {
    Write-Host "WinRM設定を確認中..." -ForegroundColor Cyan

    # WinRMサービスの起動確認（TrustedHosts取得前に実行）
    $winrmService = Get-Service -Name WinRM -ErrorAction SilentlyContinue
    if ($winrmService.Status -ne 'Running') {
        Write-Host "  WinRMサービスを起動中..." -ForegroundColor Yellow
        Start-Service -Name WinRM -ErrorAction Stop -Confirm:$false
        $winrmServiceWasStarted = $true
        Write-Host "  [OK] WinRMサービスを起動しました（終了時に停止します）" -ForegroundColor Green
    } else {
        Write-Host "  [OK] WinRMサービスは起動済みです" -ForegroundColor Green
    }

    # 現在のTrustedHostsを取得（復元用）
    try {
        $originalTrustedHosts = (Get-Item WSMan:\localhost\Client\TrustedHosts -ErrorAction SilentlyContinue).Value
        Write-Verbose "現在のTrustedHosts: $originalTrustedHosts"
    } catch {
        $originalTrustedHosts = ""
        Write-Verbose "TrustedHostsは未設定です"
    }

    # 接続先がTrustedHostsに含まれているか確認
    $needsConfig = $true
    if ($originalTrustedHosts) {
        $trustedList = $originalTrustedHosts -split ','
        if ($trustedList -contains $Config.ComputerName -or $trustedList -contains '*') {
            Write-Host "  [OK] 接続先は既にTrustedHostsに登録されています" -ForegroundColor Green
            $needsConfig = $false
        }
    }

    # 必要に応じてTrustedHostsに追加
    if ($needsConfig) {
        Write-Host ""
        Write-Host "  接続先をTrustedHostsに追加中..." -ForegroundColor Yellow

        if ([string]::IsNullOrEmpty($originalTrustedHosts)) {
            Set-Item WSMan:\localhost\Client\TrustedHosts -Value $Config.ComputerName -Force -Confirm:$false
        } else {
            Set-Item WSMan:\localhost\Client\TrustedHosts -Value "$originalTrustedHosts,$($Config.ComputerName)" -Force -Confirm:$false
        }

        $winrmConfigChanged = $true
        Write-Host "  [OK] TrustedHostsに追加しました: $($Config.ComputerName)" -ForegroundColor Green
    }

    Write-Host ""
} catch {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "[エラー] WinRM設定の自動構成に失敗しました" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "エラー詳細:" -ForegroundColor Yellow
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
    Write-Host "このスクリプトは管理者権限で実行する必要があります。" -ForegroundColor Yellow
    Write-Host ""
    exit 1
}
#endregion

#region 認証情報の準備
if ($Config.UserName) {
    if ($Config.Password) {
        $securePassword = ConvertTo-SecureString $Config.Password -AsPlainText -Force
        $Credential = New-Object System.Management.Automation.PSCredential($Config.UserName, $securePassword)
    } else {
        Write-Host "ユーザー名: $($Config.UserName)" -ForegroundColor Cyan
        $securePassword = Read-Host "パスワードを入力してください" -AsSecureString
        $Credential = New-Object System.Management.Automation.PSCredential($Config.UserName, $securePassword)
        Write-Host ""
    }
} else {
    Write-Host "認証情報を入力してください" -ForegroundColor Yellow
    $Credential = Get-Credential -Message "リモートサーバの認証情報を入力"
    Write-Host ""
}
#endregion

#region セッションオプションの設定
$sessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck
#endregion

Write-Host "リモートサーバ: $($Config.ComputerName)" -ForegroundColor White
Write-Host "実行ユーザー  : $($Credential.UserName)" -ForegroundColor White
Write-Host "実行ファイル  : $($Config.BatchPath)" -ForegroundColor White
Write-Host "出力ログ      : $OutputLog" -ForegroundColor White
Write-Host "プロトコル    : " -NoNewline -ForegroundColor White
if ($Config.UseSSL) {
    Write-Host "HTTPS (ポート 5986)" -ForegroundColor Green
} else {
    Write-Host "HTTP (ポート 5985)" -ForegroundColor Yellow
}
Write-Host ""

# メイン処理（WinRM設定復元用のfinallyブロック付き）
try {
    Write-Host "リモートサーバに接続中..." -ForegroundColor Cyan

    $sessionParams = @{
        ComputerName = $Config.ComputerName
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

    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "バッチファイル実行結果" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host ""

    $scriptBlock = {
        param($batchPath)

        if (-not (Test-Path $batchPath)) {
            throw "バッチファイルが見つかりません: $batchPath"
        }

        $output = & cmd.exe /c $batchPath 2>&1
        $exitCode = $LASTEXITCODE

        @{
            Output = $output
            ExitCode = $exitCode
        }
    }

    $result = Invoke-Command -Session $session -ScriptBlock $scriptBlock -ArgumentList $Config.BatchPath

    # 出力を表示
    $result.Output | ForEach-Object {
        Write-Host $_ -ForegroundColor White
    }

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "実行完了" -ForegroundColor Green
    Write-Host "終了コード: $($result.ExitCode)" -ForegroundColor $(if ($result.ExitCode -eq 0) { "Green" } else { "Red" })
    Write-Host "========================================" -ForegroundColor Yellow

    # ログファイル保存
    if ($Config.OutputLog) {
        Write-Host ""
        Write-Host "実行結果をログファイルに保存中..." -ForegroundColor Cyan

        $logContent = @"
========================================
リモートバッチ実行結果 (WinRM版)
========================================
実行日時: $(Get-Date -Format "yyyy/MM/dd HH:mm:ss")
リモートサーバ: $($Config.ComputerName)
実行ユーザー: $($Credential.UserName)
実行ファイル: $($Config.BatchPath)
終了コード: $($result.ExitCode)

========================================
実行結果:
========================================
$($result.Output | Out-String)
"@

        $logContent | Out-File -FilePath $OutputLog -Encoding UTF8
        Write-Host "[OK] ログ保存完了: $OutputLog" -ForegroundColor Green
    }

    Remove-PSSession -Session $session
    Write-Host ""
    Write-Host "処理が正常に完了しました。" -ForegroundColor Green

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
    Write-Host "エラー種類: $($_.Exception.GetType().FullName)" -ForegroundColor Gray

    if ($_.Exception.InnerException) {
        Write-Host "内部エラー: $($_.Exception.InnerException.Message)" -ForegroundColor Gray
    }

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "トラブルシューティング:" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "1. リモートサーバでWinRMが有効か確認:" -ForegroundColor White
    Write-Host "   winrm quickconfig" -ForegroundColor Gray
    Write-Host ""
    Write-Host "2. ファイアウォールでポートが開いているか確認:" -ForegroundColor White
    Write-Host "   HTTP: 5985 / HTTPS: 5986" -ForegroundColor Gray
    Write-Host ""
    Write-Host "3. TrustedHostsの設定（ワークグループ環境の場合）:" -ForegroundColor White
    Write-Host "   Set-Item WSMan:\localhost\Client\TrustedHosts -Value '$($Config.ComputerName)'" -ForegroundColor Gray
    Write-Host ""
    Write-Host "4. 接続テスト:" -ForegroundColor White
    Write-Host "   Test-WSMan -ComputerName $($Config.ComputerName)" -ForegroundColor Gray
    Write-Host ""

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
                Set-Item WSMan:\localhost\Client\TrustedHosts -Value "" -Force -Confirm:$false
                Write-Host "[OK] TrustedHostsを元の状態（空）に復元しました" -ForegroundColor Green
            } else {
                Set-Item WSMan:\localhost\Client\TrustedHosts -Value $originalTrustedHosts -Force -Confirm:$false
                Write-Host "[OK] TrustedHostsを元の状態に復元しました" -ForegroundColor Green
            }
        } catch {
            Write-Host "[警告] TrustedHostsの復元に失敗しました" -ForegroundColor Yellow
            Write-Host "エラー: $($_.Exception.Message)" -ForegroundColor Yellow
            Write-Host "手動で復元してください: Set-Item WSMan:\localhost\Client\TrustedHosts -Value '$originalTrustedHosts' -Force" -ForegroundColor Yellow
        }
    }

    # WinRMサービスを停止（このスクリプトで起動した場合のみ）
    if ($winrmServiceWasStarted) {
        Write-Host ""
        Write-Host "WinRMサービスを停止中..." -ForegroundColor Cyan

        try {
            Stop-Service -Name WinRM -Force -Confirm:$false -ErrorAction Stop
            Write-Host "[OK] WinRMサービスを元の状態（停止）に復元しました" -ForegroundColor Green
        } catch {
            Write-Host "[警告] WinRMサービスの停止に失敗しました" -ForegroundColor Yellow
            Write-Host "エラー: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    #endregion
}

# バッチ側でpauseが実行されるため、ここでは何もしない
Write-Host ""
exit $exitCode
