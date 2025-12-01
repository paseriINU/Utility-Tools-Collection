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
    .\Invoke-RemoteBatch.ps1 -ComputerName "192.168.1.100" -UserName "Administrator" -BatchPath "C:\Scripts\backup.bat"

.EXAMPLE
    $cred = Get-Credential
    .\Invoke-RemoteBatch.ps1 -ComputerName "server01" -Credential $cred -BatchPath "C:\Scripts\test.bat" -OutputLog "result.log"

.EXAMPLE
    # 引数付きでバッチを実行
    .\Invoke-RemoteBatch.ps1 -ComputerName "192.168.1.100" -UserName "Administrator" -BatchPath "C:\Scripts\process.bat" -Arguments "param1 param2"

.NOTES
    作成日: 2025-12-01
    バージョン: 1.0
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
    Write-Host "✓ 接続成功" -ForegroundColor Green
    Write-Host ""
    #endregion

    #region バッチファイルの実行
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "バッチファイル実行中..." -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host ""

    $scriptBlock = {
        param($batchPath, $args)

        # バッチファイルの存在確認
        if (-not (Test-Path $batchPath)) {
            throw "バッチファイルが見つかりません: $batchPath"
        }

        # バッチファイルを実行
        if ($args) {
            $output = & cmd.exe /c "$batchPath $args" 2>&1
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
        Write-Host "✓ ログ保存完了: $OutputLog" -ForegroundColor Green
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
