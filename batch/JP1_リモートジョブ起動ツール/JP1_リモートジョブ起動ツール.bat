<# :
@echo off
chcp 65001 >nul
title JP1 リモートジョブ起動ツール
setlocal

rem 管理者権限チェック
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo 管理者権限が必要です。管理者として再起動します...
    powershell -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

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

    # ajsstatusコマンドのパス（リモートサーバ上）
    AjsstatusPath = "C:\Program Files\HITACHI\JP1AJS3\bin\ajsstatus.exe"

    # ajsshowコマンドのパス（リモートサーバ上）
    AjsshowPath = "C:\Program Files\HITACHI\JP1AJS3\bin\ajsshow.exe"

    # HTTPS接続を使用する場合は $true
    UseSSL = $false

    # ジョブ完了を待つ場合は $true（起動のみの場合は $false）
    WaitForCompletion = $true

    # 完了待ちの最大時間（秒）。0の場合は無制限
    WaitTimeoutSeconds = 3600

    # 状態確認の間隔（秒）
    PollingIntervalSeconds = 10
}

# ==============================================================================
# ■ メイン処理（以下は編集不要）
# ==============================================================================

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "JP1ジョブネット起動ツール" -ForegroundColor Cyan
Write-Host "（リモート実行版）" -ForegroundColor Cyan
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

    # 現在のTrustedHostsを取得
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
        if ($trustedList -contains $Config.JP1Server -or $trustedList -contains '*') {
            Write-Host "  [OK] 接続先は既にTrustedHostsに登録されています" -ForegroundColor Green
            $needsConfig = $false
        }
    }

    # 必要に応じてTrustedHostsに追加
    if ($needsConfig) {
        Write-Host ""
        Write-Host "  接続先をTrustedHostsに追加中..." -ForegroundColor Yellow

        if ([string]::IsNullOrEmpty($originalTrustedHosts)) {
            Set-Item WSMan:\localhost\Client\TrustedHosts -Value $Config.JP1Server -Force -Confirm:$false
        } else {
            Set-Item WSMan:\localhost\Client\TrustedHosts -Value "$originalTrustedHosts,$($Config.JP1Server)" -Force -Confirm:$false
        }

        $winrmConfigChanged = $true
        Write-Host "  [OK] TrustedHostsに追加しました: $($Config.JP1Server)" -ForegroundColor Green
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

Write-Host "JP1サーバ      : $($Config.JP1Server)" -ForegroundColor White
Write-Host "リモートユーザー: $($Config.RemoteUser)" -ForegroundColor White
Write-Host "JP1ユーザー    : $($Config.JP1User)" -ForegroundColor White
Write-Host "ジョブネットパス: $($Config.JobnetPath)" -ForegroundColor White
Write-Host "完了待ち       : $(if ($Config.WaitForCompletion) { '有効' } else { '無効' })" -ForegroundColor White
if ($Config.WaitForCompletion) {
    $timeoutDisplay = if ($Config.WaitTimeoutSeconds -eq 0) { "無制限" } else { "$($Config.WaitTimeoutSeconds)秒" }
    Write-Host "タイムアウト   : $timeoutDisplay" -ForegroundColor White
}
Write-Host ""

# JP1パスワード入力
if ([string]::IsNullOrEmpty($Config.JP1Password)) {
    Write-Host "[注意] JP1パスワードが設定されていません。" -ForegroundColor Yellow
    $Config.JP1Password = Read-Host "JP1パスワードを入力してください" -AsSecureString
    $Config.JP1Password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($Config.JP1Password))
    Write-Host ""
}

# 実行確認
Write-Host "ジョブネットを起動しますか？ (y/n)" -ForegroundColor Cyan
$confirmation = Read-Host
if ($confirmation -ne "Y" -and $confirmation -ne "y") {
    Write-Host "処理をキャンセルしました。" -ForegroundColor Yellow
    Write-Host ""
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

    #region ジョブネット起動
    $scriptBlockEntry = {
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

    $result = Invoke-Command -Session $session -ScriptBlock $scriptBlockEntry -ArgumentList $Config.AjsentryPath, $Config.JP1User, $Config.JP1Password, $Config.JobnetPath

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

    Write-Host "ajsentry出力:" -ForegroundColor White
    $result.Output | ForEach-Object {
        Write-Host "  $_" -ForegroundColor White
    }
    #endregion

    #region ジョブ完了待ち
    $jobFinalStatus = $null
    $jobEndNormally = $false

    if ($result.ExitCode -eq 0 -and $Config.WaitForCompletion) {
        Write-Host ""
        Write-Host "========================================" -ForegroundColor Yellow
        Write-Host "ジョブ完了を待機中..." -ForegroundColor Yellow
        Write-Host "========================================" -ForegroundColor Yellow
        Write-Host ""

        $startTime = Get-Date
        $isRunning = $true
        $timedOut = $false

        # ajsstatusでジョブネットの状態を監視
        $scriptBlockStatus = {
            param($ajsStatusPath, $jp1User, $jp1Pass, $jobnetPath)

            if (-not (Test-Path $ajsStatusPath)) {
                throw "ajsstatusが見つかりません: $ajsStatusPath"
            }

            $output = & $ajsStatusPath -h localhost -u $jp1User -p $jp1Pass -F $jobnetPath 2>&1
            $exitCode = $LASTEXITCODE

            @{
                ExitCode = $exitCode
                Output = $output
                RawOutput = ($output -join "`n")
            }
        }

        while ($isRunning) {
            # タイムアウトチェック
            if ($Config.WaitTimeoutSeconds -gt 0) {
                $elapsed = (Get-Date) - $startTime
                if ($elapsed.TotalSeconds -ge $Config.WaitTimeoutSeconds) {
                    $timedOut = $true
                    break
                }
            }

            # 状態確認
            $statusResult = Invoke-Command -Session $session -ScriptBlock $scriptBlockStatus -ArgumentList $Config.AjsstatusPath, $Config.JP1User, $Config.JP1Password, $Config.JobnetPath

            $statusOutput = $statusResult.RawOutput.ToLower()

            # 状態判定（JP1/AJS3の典型的なステータス）
            # 実行中: "now running", "running", "wait"
            # 正常終了: "end normally", "ended normally", "normal end"
            # 異常終了: "ended abnormally", "abnormal end", "abend", "killed", "interrupted"
            if ($statusOutput -match "ended abnormally|abnormal end|abend|killed|interrupted|failed") {
                $isRunning = $false
                $jobEndNormally = $false
                $jobFinalStatus = "異常終了"
            } elseif ($statusOutput -match "end normally|ended normally|normal end|completed") {
                $isRunning = $false
                $jobEndNormally = $true
                $jobFinalStatus = "正常終了"
            } elseif ($statusOutput -match "now running|running|wait|queued|executing") {
                # 実行中 - 継続して待機
                $elapsed = (Get-Date) - $startTime
                $elapsedStr = "{0:mm\:ss}" -f $elapsed
                Write-Host "`r  状態: 実行中... (経過時間: $elapsedStr)  " -NoNewline -ForegroundColor Cyan
            } elseif ($statusOutput -match "not registered|not found|does not exist") {
                $isRunning = $false
                $jobEndNormally = $false
                $jobFinalStatus = "ジョブネットが見つかりません"
            } else {
                # その他のステータス（終了と判断）
                # 終了コードで判断
                if ($statusResult.ExitCode -eq 0) {
                    $isRunning = $false
                    $jobEndNormally = $true
                    $jobFinalStatus = "終了（ステータス不明）"
                } else {
                    # 引き続き実行中と判断して待機
                    $elapsed = (Get-Date) - $startTime
                    $elapsedStr = "{0:mm\:ss}" -f $elapsed
                    Write-Host "`r  状態: 確認中... (経過時間: $elapsedStr)  " -NoNewline -ForegroundColor Gray
                }
            }

            if ($isRunning) {
                Start-Sleep -Seconds $Config.PollingIntervalSeconds
            }
        }

        Write-Host ""
        Write-Host ""

        if ($timedOut) {
            Write-Host "========================================" -ForegroundColor Red
            Write-Host "[タイムアウト] 完了待ちがタイムアウトしました" -ForegroundColor Red
            Write-Host "  タイムアウト時間: $($Config.WaitTimeoutSeconds)秒" -ForegroundColor Red
            Write-Host "========================================" -ForegroundColor Red
            $jobFinalStatus = "タイムアウト"
            $jobEndNormally = $false
        } else {
            Write-Host "========================================" -ForegroundColor $(if ($jobEndNormally) { "Green" } else { "Red" })
            Write-Host "ジョブネット実行結果: $jobFinalStatus" -ForegroundColor $(if ($jobEndNormally) { "Green" } else { "Red" })
            Write-Host "========================================" -ForegroundColor $(if ($jobEndNormally) { "Green" } else { "Red" })
        }
    }
    #endregion

    #region 詳細メッセージ取得
    if ($result.ExitCode -eq 0) {
        Write-Host ""
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host "ジョブ詳細情報を取得中..." -ForegroundColor Cyan
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host ""

        $scriptBlockShow = {
            param($ajsShowPath, $jp1User, $jp1Pass, $jobnetPath)

            if (-not (Test-Path $ajsShowPath)) {
                return @{
                    ExitCode = -1
                    Output = @("ajsshowが見つかりません: $ajsShowPath")
                    Available = $false
                }
            }

            # -i オプションで詳細情報を取得、-E で実行結果詳細
            $output = & $ajsShowPath -h localhost -u $jp1User -p $jp1Pass -F $jobnetPath -E 2>&1
            $exitCode = $LASTEXITCODE

            @{
                ExitCode = $exitCode
                Output = $output
                Available = $true
            }
        }

        $showResult = Invoke-Command -Session $session -ScriptBlock $scriptBlockShow -ArgumentList $Config.AjsshowPath, $Config.JP1User, $Config.JP1Password, $Config.JobnetPath

        if ($showResult.Available) {
            Write-Host "詳細情報 (ajsshow -E):" -ForegroundColor Yellow
            Write-Host "----------------------------------------" -ForegroundColor Gray
            $showResult.Output | ForEach-Object {
                Write-Host "  $_" -ForegroundColor White
            }
            Write-Host "----------------------------------------" -ForegroundColor Gray
        } else {
            Write-Host "[情報] ajsshowコマンドが利用できません" -ForegroundColor Yellow
            Write-Host "  パス: $($Config.AjsshowPath)" -ForegroundColor Gray
        }
    }
    #endregion

    Remove-PSSession -Session $session

    #region 最終結果表示
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "処理サマリー" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  ジョブネット: $($Config.JobnetPath)" -ForegroundColor White
    Write-Host "  サーバ      : $($Config.JP1Server)" -ForegroundColor White
    Write-Host "  起動結果    : $(if ($result.ExitCode -eq 0) { '成功' } else { '失敗' })" -ForegroundColor $(if ($result.ExitCode -eq 0) { "Green" } else { "Red" })

    if ($Config.WaitForCompletion -and $result.ExitCode -eq 0) {
        Write-Host "  実行結果    : $jobFinalStatus" -ForegroundColor $(if ($jobEndNormally) { "Green" } else { "Red" })
    }

    if ($result.ExitCode -ne 0) {
        Write-Host ""
        Write-Host "追加の確認事項：" -ForegroundColor Yellow
        Write-Host "  - ajsentryのパスが正しいか: $($Config.AjsentryPath)" -ForegroundColor Yellow
        Write-Host "  - JP1ユーザー名、パスワードが正しいか" -ForegroundColor Yellow
        Write-Host "  - ジョブネットパスが正しいか" -ForegroundColor Yellow
        Write-Host "  - JP1/AJS3サービスが起動しているか" -ForegroundColor Yellow
    }
    #endregion

    # 最終終了コード決定
    if ($result.ExitCode -ne 0) {
        $exitCode = $result.ExitCode
    } elseif ($Config.WaitForCompletion -and -not $jobEndNormally) {
        $exitCode = 1  # ジョブが異常終了
    } else {
        $exitCode = 0
    }

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
