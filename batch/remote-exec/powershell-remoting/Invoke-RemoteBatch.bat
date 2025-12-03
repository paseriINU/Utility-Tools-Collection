<# :
@echo off
setlocal
chcp 65001 >nul

:: 引数チェック
if "%~1"=="" (
    echo 使い方: %~nx0 -ComputerName ^<サーバ名^> -UserName ^<ユーザー名^> -BatchPath ^<バッチパス^> [-Arguments ^<引数^>] [-OutputLog ^<ログパス^>] [-UseSSL]
    echo.
    echo 例: %~nx0 -ComputerName "192.168.1.100" -UserName "Administrator" -BatchPath "C:\Scripts\test.bat"
    pause
    exit /b 1
)

:: PowerShellスクリプトを実行
powershell -NoProfile -ExecutionPolicy Bypass -Command "iex ((gc '%~f0') -join \"`n\")" %*
exit /b %ERRORLEVEL%
: #> | sv -name _ > $null

<#
.SYNOPSIS
    リモートWindowsサーバ上でバッチファイルをPowerShell Remotingで実行

.DESCRIPTION
    PowerShell Remotingを使用してリモートサーバでバッチファイルを実行し、
    実行結果をリアルタイムで取得します。
    WinRM版（バッチファイル経由）よりも柔軟で、PowerShellの機能をフルに活用できます。

.PARAMETER ComputerName
    リモートサーバのコンピュータ名またはIPアドレス

.PARAMETER Credential
    認証情報（PSCredentialオブジェクト）

.PARAMETER UserName
    ユーザー名（Credentialと併用不可）

.PARAMETER Password
    パスワード（SecureString）（Credentialと併用不可）

.PARAMETER BatchPath
    リモートサーバで実行するバッチファイルのフルパス

.PARAMETER Arguments
    バッチファイルに渡す引数（オプション）

.PARAMETER OutputLog
    実行結果を保存するローカルファイルパス（オプション）

.PARAMETER UseSSL
    HTTPS（ポート5986）を使用する場合はこのスイッチを指定

.EXAMPLE
    .\Invoke-RemoteBatch.bat -ComputerName "192.168.1.100" -UserName "Administrator" -BatchPath "C:\Scripts\backup.bat"

.EXAMPLE
    $cred = Get-Credential
    .\Invoke-RemoteBatch.bat -ComputerName "server01" -Credential $cred -BatchPath "C:\Scripts\test.bat" -OutputLog "result.log"

.EXAMPLE
    # 引数付きでバッチを実行
    .\Invoke-RemoteBatch.bat -ComputerName "192.168.1.100" -UserName "Administrator" -BatchPath "C:\Scripts\process.bat" -Arguments "param1 param2"

.NOTES
    作成日: 2025-12-02
    バージョン: 2.0
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$ComputerName,

    [Parameter(Mandatory=$false)]
    [PSCredential]$Credential,

    [Parameter(Mandatory=$false)]
    [string]$UserName,

    [Parameter(Mandatory=$false)]
    [SecureString]$Password,

    [Parameter(Mandatory=$true)]
    [string]$BatchPath,

    [Parameter(Mandatory=$false)]
    [string]$Arguments,

    [Parameter(Mandatory=$false)]
    [string]$OutputLog,

    [Parameter(Mandatory=$false)]
    [switch]$UseSSL
)

# エラー時は停止
$ErrorActionPreference = "Stop"

# UTF-8出力設定
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

#region WinRM設定の保存と自動設定
$originalTrustedHosts = $null
$winrmConfigChanged = $false

try {
    Write-Host "WinRM設定を確認中..." -ForegroundColor Cyan

    # 現在のTrustedHostsを取得（復元用）
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
        if ($trustedList -contains $ComputerName -or $trustedList -contains '*') {
            Write-Host "[OK] 接続先は既にTrustedHostsに登録されています" -ForegroundColor Green
            $needsConfig = $false
        }
    }

    # 必要に応じてTrustedHostsに追加
    if ($needsConfig) {
        Write-Host "接続先をTrustedHostsに追加中..." -ForegroundColor Yellow

        if ([string]::IsNullOrEmpty($originalTrustedHosts)) {
            # 既存設定なし
            Set-Item WSMan:\localhost\Client\TrustedHosts -Value $ComputerName -Force
        } else {
            # 既存設定あり
            Set-Item WSMan:\localhost\Client\TrustedHosts -Value "$originalTrustedHosts,$ComputerName" -Force
        }

        $winrmConfigChanged = $true
        Write-Host "[OK] TrustedHostsに追加しました: $ComputerName" -ForegroundColor Green
    }

    Write-Host ""
} catch {
    Write-Host "[警告] WinRM設定の自動構成に失敗しました" -ForegroundColor Yellow
    Write-Host "エラー: $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host "手動でWinRM設定を行ってください" -ForegroundColor Yellow
    Write-Host ""
}
#endregion

#region 認証情報の準備
if ($Credential) {
    # Credentialオブジェクトが指定されている場合はそれを使用
    Write-Verbose "Credentialオブジェクトを使用します"
} elseif ($UserName) {
    # UserNameが指定されている場合
    if (-not $Password) {
        # パスワードが指定されていない場合は入力を求める
        Write-Host "ユーザー名: $UserName" -ForegroundColor Cyan
        $Password = Read-Host "パスワードを入力してください" -AsSecureString
    }
    $Credential = New-Object System.Management.Automation.PSCredential($UserName, $Password)
} else {
    # 認証情報が何も指定されていない場合
    Write-Host "認証情報を入力してください" -ForegroundColor Yellow
    $Credential = Get-Credential -Message "リモートサーバの認証情報を入力"
}
#endregion

#region セッションオプションの設定
$sessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck
#endregion

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "PowerShell Remoting - リモートバッチ実行" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "リモートサーバ: $ComputerName" -ForegroundColor White
Write-Host "実行ユーザー  : $($Credential.UserName)" -ForegroundColor White
Write-Host "実行ファイル  : $BatchPath" -ForegroundColor White
if ($Arguments) {
    Write-Host "引数          : $Arguments" -ForegroundColor White
}
if ($OutputLog) {
    Write-Host "出力ログ      : $OutputLog" -ForegroundColor White
}
Write-Host "プロトコル    : " -NoNewline -ForegroundColor White
if ($UseSSL) {
    Write-Host "HTTPS (ポート 5986)" -ForegroundColor Green
} else {
    Write-Host "HTTP (ポート 5985)" -ForegroundColor Yellow
}
Write-Host ""

# メイン処理（WinRM設定復元用のfinallyブロック付き）
try {
    #region リモートセッションの確立
    Write-Host "リモートサーバに接続中..." -ForegroundColor Cyan

    $sessionParams = @{
        ComputerName = $ComputerName
        Credential = $Credential
        SessionOption = $sessionOption
        ErrorAction = "Stop"
    }

    if ($UseSSL) {
        $sessionParams.UseSSL = $true
    }

    $session = New-PSSession @sessionParams
    Write-Host "[OK] 接続成功" -ForegroundColor Green
    Write-Host ""
    #endregion

    #region バッチファイルの実行
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "バッチファイル実行中..." -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host ""

    $scriptBlock = {
        param($batchPath, $batchArgs)

        # バッチファイルの存在確認
        if (-not (Test-Path $batchPath)) {
            throw "バッチファイルが見つかりません: $batchPath"
        }

        # バッチファイルを実行
        if ($batchArgs) {
            $output = & cmd.exe /c "$batchPath $batchArgs" 2>&1
        } else {
            $output = & cmd.exe /c $batchPath 2>&1
        }

        # 終了コードを保存
        $exitCode = $LASTEXITCODE

        # 結果を返す
        @{
            Output = $output
            ExitCode = $exitCode
        }
    }

    $result = Invoke-Command -Session $session -ScriptBlock $scriptBlock -ArgumentList $BatchPath, $Arguments

    # 出力を表示
    $result.Output | ForEach-Object {
        Write-Host $_ -ForegroundColor White
    }

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "実行完了" -ForegroundColor Green
    Write-Host "終了コード: $($result.ExitCode)" -ForegroundColor $(if ($result.ExitCode -eq 0) { "Green" } else { "Red" })
    Write-Host "========================================" -ForegroundColor Yellow
    #endregion

    #region ログファイル保存
    if ($OutputLog) {
        Write-Host ""
        Write-Host "実行結果をログファイルに保存中..." -ForegroundColor Cyan

        $logContent = @"
========================================
PowerShell Remoting - リモートバッチ実行結果
========================================
実行日時: $(Get-Date -Format "yyyy/MM/dd HH:mm:ss")
リモートサーバ: $ComputerName
実行ユーザー: $($Credential.UserName)
実行ファイル: $BatchPath
引数: $Arguments
終了コード: $($result.ExitCode)

========================================
実行結果:
========================================
$($result.Output | Out-String)
"@

        $logContent | Out-File -FilePath $OutputLog -Encoding UTF8
        Write-Host "[OK] ログ保存完了: $OutputLog" -ForegroundColor Green
    }
    #endregion

    #region セッションのクローズ
    Remove-PSSession -Session $session
    Write-Host ""
    Write-Host "処理が正常に完了しました。" -ForegroundColor Green
    #endregion

    # 終了コードを返す
    exit $result.ExitCode

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
    Write-Host "   Set-Item WSMan:\localhost\Client\TrustedHosts -Value '$ComputerName'" -ForegroundColor Gray
    Write-Host ""
    Write-Host "4. 接続テスト:" -ForegroundColor White
    Write-Host "   Test-WSMan -ComputerName $ComputerName" -ForegroundColor Gray
    Write-Host ""

    # セッションが残っている場合はクリーンアップ
    if ($session) {
        Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    }

    exit 1
}
} finally {
    #region WinRM設定の復元
    if ($winrmConfigChanged) {
        Write-Host ""
        Write-Host "WinRM設定を復元中..." -ForegroundColor Cyan

        try {
            if ([string]::IsNullOrEmpty($originalTrustedHosts)) {
                # 元々空だった場合は空に戻す
                Set-Item WSMan:\localhost\Client\TrustedHosts -Value "" -Force
                Write-Host "[OK] TrustedHostsを元の状態（空）に復元しました" -ForegroundColor Green
            } else {
                # 元の値に戻す
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
