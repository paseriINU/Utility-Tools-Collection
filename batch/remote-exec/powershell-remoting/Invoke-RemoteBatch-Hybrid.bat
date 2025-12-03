<# :
@echo off
setlocal
chcp 65001 >nul
powershell -NoProfile -ExecutionPolicy Bypass -Command "iex ((gc '%~f0') -join \"`n\")"
exit /b %ERRORLEVEL%
: #> | sv -name _ > $null

<#
.SYNOPSIS
    リモートWindowsサーバ上でバッチファイルをPowerShell Remotingで実行（ハイブリッド版）

.DESCRIPTION
    .batとして実行可能なハイブリッドスクリプトです。
    ダブルクリックで実行でき、設定を内部に記述できます。

.NOTES
    作成日: 2025-12-02
    バージョン: 2.0

    使い方:
    1. 下記の「■ 設定セクション」を編集
    2. このファイルをダブルクリックで実行（.bat として実行されます）
#>

# ==============================================================================
# ■ 設定セクション（ここを編集してください）
# ==============================================================================

# リモートサーバの設定
$Config = @{
    # リモートサーバのコンピュータ名またはIPアドレス
    ComputerName = "192.168.1.100"

    # 認証情報（以下のいずれかを設定）
    # 方法1: ユーザー名とパスワードを直接指定（セキュリティ注意）
    UserName = "Administrator"
    Password = ""  # 空の場合は実行時に入力を求められます

    # 方法2: 実行時に認証情報を入力（UserNameを空にする）
    # UserName = ""
    # Password = ""

    # 実行するバッチファイルのパス（リモートサーバ上のパス）
    # {env}の部分が実行時に選択した環境（tst1t/tst2t）に置換されます
    BatchPath = "C:\Scripts\{env}\test.bat"

    # バッチファイルに渡す引数（オプション、不要な場合は空文字）
    Arguments = ""

    # 実行結果を保存するローカルファイルパス
    # 空の場合は自動的に日時付きファイル名で保存されます
    # 例: RemoteBatch_20250102_153045.log
    OutputLog = ""

    # SSL/HTTPS接続を使用する場合は $true（通常は $false）
    UseSSL = $false
}

# ==============================================================================
# ■ メイン処理（以下は編集不要）
# ==============================================================================

# エラー時は停止
$ErrorActionPreference = "Stop"

# UTF-8出力設定
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# 出力ログファイルのデフォルト設定（空の場合）
if (-not $Config.OutputLog) {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $scriptDir = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
    if (-not $scriptDir) {
        $scriptDir = Get-Location
    }
    $Config.OutputLog = Join-Path $scriptDir "RemoteBatch_$timestamp.log"
}

# ヘッダー表示
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "PowerShell Remoting - リモートバッチ実行" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

#region 環境選択
if ($Config.BatchPath -like "*{env}*") {
    Write-Host "実行環境を選択してください:" -ForegroundColor Cyan
    Write-Host "  1. tst1t"
    Write-Host "  2. tst2t"
    Write-Host ""

    $envChoice = Read-Host "選択 (1-2)"

    switch ($envChoice) {
        "1" { $selectedEnv = "tst1t" }
        "2" { $selectedEnv = "tst2t" }
        default {
            Write-Host "[エラー] 無効な選択です" -ForegroundColor Red
            Write-Host "Enterキーを押して終了..." -ForegroundColor Gray
            $null = Read-Host
            exit 1
        }
    }

    # BatchPathの{env}を選択した環境に置換
    $Config.BatchPath = $Config.BatchPath -replace '\{env\}', $selectedEnv
    Write-Host ""
    Write-Host "選択された環境: $selectedEnv" -ForegroundColor Green
    Write-Host ""
}
#endregion

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
        if ($trustedList -contains $Config.ComputerName -or $trustedList -contains '*') {
            Write-Host "[OK] 接続先は既にTrustedHostsに登録されています" -ForegroundColor Green
            $needsConfig = $false
        }
    }

    # 必要に応じてTrustedHostsに追加
    if ($needsConfig) {
        Write-Host "接続先をTrustedHostsに追加中..." -ForegroundColor Yellow

        if ([string]::IsNullOrEmpty($originalTrustedHosts)) {
            # 既存設定なし
            Set-Item WSMan:\localhost\Client\TrustedHosts -Value $Config.ComputerName -Force
        } else {
            # 既存設定あり
            Set-Item WSMan:\localhost\Client\TrustedHosts -Value "$originalTrustedHosts,$($Config.ComputerName)" -Force
        }

        $winrmConfigChanged = $true
        Write-Host "[OK] TrustedHostsに追加しました: $($Config.ComputerName)" -ForegroundColor Green
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
if ($Config.UserName) {
    # UserNameが指定されている場合
    if ($Config.Password) {
        # パスワードが指定されている場合
        $securePassword = ConvertTo-SecureString $Config.Password -AsPlainText -Force
        $Credential = New-Object System.Management.Automation.PSCredential($Config.UserName, $securePassword)
    } else {
        # パスワードが指定されていない場合は入力を求める
        Write-Host "ユーザー名: $($Config.UserName)" -ForegroundColor Cyan
        $securePassword = Read-Host "パスワードを入力してください" -AsSecureString
        $Credential = New-Object System.Management.Automation.PSCredential($Config.UserName, $securePassword)
    }
} else {
    # 認証情報が何も指定されていない場合
    Write-Host "認証情報を入力してください" -ForegroundColor Yellow
    $Credential = Get-Credential -Message "リモートサーバの認証情報を入力"
}
#endregion

#region セッションオプションの設定
$sessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck
#endregion

# 設定情報を表示
Write-Host "リモートサーバ: $($Config.ComputerName)" -ForegroundColor White
Write-Host "実行ユーザー  : $($Credential.UserName)" -ForegroundColor White
Write-Host "実行ファイル  : $($Config.BatchPath)" -ForegroundColor White
if ($Config.Arguments) {
    Write-Host "引数          : $($Config.Arguments)" -ForegroundColor White
}
if ($Config.OutputLog) {
    Write-Host "出力ログ      : $($Config.OutputLog)" -ForegroundColor White
}
Write-Host "プロトコル    : " -NoNewline -ForegroundColor White
if ($Config.UseSSL) {
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

    $result = Invoke-Command -Session $session -ScriptBlock $scriptBlock -ArgumentList $Config.BatchPath, $Config.Arguments

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
    if ($Config.OutputLog) {
        Write-Host ""
        Write-Host "実行結果をログファイルに保存中..." -ForegroundColor Cyan

        $logContent = @"
========================================
PowerShell Remoting - リモートバッチ実行結果
========================================
実行日時: $(Get-Date -Format "yyyy/MM/dd HH:mm:ss")
リモートサーバ: $($Config.ComputerName)
実行ユーザー: $($Credential.UserName)
実行ファイル: $($Config.BatchPath)
引数: $($Config.Arguments)
終了コード: $($result.ExitCode)

========================================
実行結果:
========================================
$($result.Output | Out-String)
"@

        $logContent | Out-File -FilePath $Config.OutputLog -Encoding UTF8
        Write-Host "[OK] ログ保存完了: $($Config.OutputLog)" -ForegroundColor Green
    }
    #endregion

    #region セッションのクローズ
    Remove-PSSession -Session $session
    Write-Host ""
    Write-Host "処理が正常に完了しました。" -ForegroundColor Green
    #endregion

    # 終了コードを返す
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

    # セッションが残っている場合はクリーンアップ
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

# 実行後にウィンドウを閉じないようにする
Write-Host ""
Write-Host "Enterキーを押して終了..." -ForegroundColor Gray
$null = Read-Host

exit $exitCode
