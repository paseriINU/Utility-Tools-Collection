<# :
@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion

rem ============================================================================
rem JP1 ジョブ実行バッチ
rem   - スタンドアローン実行またはExcel VBAからの呼び出しに対応
rem   - 設定は下記の設定セクションを編集してください
rem ============================================================================

rem 引数がなければ使用方法を表示
if "%~1"=="" (
    echo JP1 ジョブ実行バッチ
    echo.
    echo 使用方法:
    echo   JP1_ジョブ実行.bat [ジョブネットパス] [オプション]
    echo.
    echo オプション:
    echo   -IsHold         保留中フラグ（true/false、デフォルト: false）
    echo   -WaitCompletion 完了待ちフラグ（true/false、デフォルト: true）
    echo   -LogFile        ログファイルパス（省略時はLogsフォルダに自動生成）
    echo.
    echo 例:
    echo   JP1_ジョブ実行.bat /main_unit/daily_batch
    echo   JP1_ジョブ実行.bat /main_unit/daily_batch -IsHold true -WaitCompletion false
    echo.
    echo ※設定はこのバッチファイルの「設定セクション」を編集してください
    pause
    exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -Command "$scriptPath='%~f0'; iex ((gc $scriptPath -Encoding UTF8) -join \"`n\")" -- %*
exit /b %ERRORLEVEL%
: #>

#==============================================================================
# ★★★ 設定セクション（ここを編集してください）★★★
#==============================================================================

# 実行モード: "ローカル" または "リモート"
$CONFIG_Mode = "リモート"

# JP1サーバ（リモートモード時のみ使用）
$CONFIG_JP1Server = "192.168.1.100"

# リモートユーザー（リモートモード時のみ使用）
$CONFIG_RemoteUser = "Administrator"

# リモートパスワード（リモートモード時のみ使用）
# ※空の場合は実行時に入力を求めます
$CONFIG_RemotePassword = ""

# JP1ユーザー
$CONFIG_JP1User = "jp1admin"

# JP1パスワード
# ※空の場合は実行時に入力を求めます
$CONFIG_JP1Password = ""

# タイムアウト（秒）0=無制限
$CONFIG_Timeout = 0

# ポーリング間隔（秒）
$CONFIG_PollingInterval = 10

#==============================================================================
# ★★★ 設定セクション ここまで ★★★
#==============================================================================

#==============================================================================
# パラメータ定義（コマンドライン引数）
#==============================================================================
param(
    [Parameter(Position=0)]
    [string]$JobnetPath = "",
    [string]$IsHold = "false",
    [string]$WaitCompletion = "true",
    [string]$LogFile = ""
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

#==============================================================================
# パスワード入力（空の場合）
#==============================================================================
if ($CONFIG_JP1Password -eq "") {
    $secureJP1Pass = Read-Host "JP1パスワードを入力してください" -AsSecureString
    $CONFIG_JP1Password = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureJP1Pass)
    )
    if ($CONFIG_JP1Password -eq "") {
        Write-Host "[ERROR] JP1パスワードが入力されていません" -ForegroundColor Red
        exit 1
    }
}

if ($CONFIG_Mode -eq "リモート" -and $CONFIG_RemotePassword -eq "") {
    $secureRemotePass = Read-Host "リモートパスワードを入力してください" -AsSecureString
    $CONFIG_RemotePassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureRemotePass)
    )
    if ($CONFIG_RemotePassword -eq "") {
        Write-Host "[ERROR] リモートパスワードが入力されていません" -ForegroundColor Red
        exit 1
    }
}

#==============================================================================
# ログファイル設定
#==============================================================================
if ($LogFile -eq "") {
    $scriptDir = Split-Path -Parent $scriptPath
    $logFolder = Join-Path $scriptDir "Logs"
    if (-not (Test-Path $logFolder)) {
        New-Item -ItemType Directory -Path $logFolder -Force | Out-Null
    }
    $LogFile = Join-Path $logFolder ("JP1_実行ログ_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".txt")
}

#==============================================================================
# ログ出力関数
#==============================================================================
function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format 'yyyy/MM/dd HH:mm:ss'
    $logLine = "[$timestamp] $Message"
    Add-Content -Path $LogFile -Value $logLine -Encoding UTF8
}

# ログヘッダー出力
Add-Content -Path $LogFile -Value "================================================================================" -Encoding UTF8
Add-Content -Path $LogFile -Value "JP1 ジョブ実行バッチ - 実行ログ" -Encoding UTF8
Add-Content -Path $LogFile -Value "================================================================================" -Encoding UTF8
Add-Content -Path $LogFile -Value ("開始日時: " + (Get-Date -Format "yyyy/MM/dd HH:mm:ss")) -Encoding UTF8
Add-Content -Path $LogFile -Value ("実行モード: " + $CONFIG_Mode) -Encoding UTF8
Add-Content -Path $LogFile -Value "================================================================================" -Encoding UTF8
Add-Content -Path $LogFile -Value "" -Encoding UTF8

#==============================================================================
# JP1コマンドパス検出
#==============================================================================
function Get-JP1BinPath {
    $paths = @(
        'C:\Program Files\HITACHI\JP1AJS3\bin',
        'C:\Program Files\Hitachi\JP1AJS2\bin'
    )
    foreach ($p in $paths) {
        if (Test-Path "$p\ajsentry.exe") {
            return $p
        }
    }
    return $null
}

#==============================================================================
# ローカルモード実行
#==============================================================================
function Invoke-LocalMode {
    Write-Log '--------------------------------------------------------------------------------'
    Write-Log "ジョブネット: $JobnetPath"
    Write-Log '実行モード: ローカル'
    Write-Log '--------------------------------------------------------------------------------'

    # JP1コマンドパスの検出
    $jp1BinPath = Get-JP1BinPath
    if (-not $jp1BinPath) {
        Write-Log '[ERROR] JP1コマンドが見つかりません'
        Write-Output "RESULT_STATUS:起動失敗"
        Write-Output "RESULT_MESSAGE:JP1コマンドが見つかりません。このPCにJP1/AJS3がインストールされているか確認してください。"
        return
    }
    Write-Log "JP1コマンドパス: $jp1BinPath"

    # 保留解除（必要な場合）
    if ($IsHold -eq "true") {
        Write-Log '[実行] ajsrelease - 保留解除'
        Write-Log "コマンド: ajsrelease.exe -h localhost -u $CONFIG_JP1User -p ***** -F $JobnetPath"
        $releaseOutput = & "$jp1BinPath\ajsrelease.exe" -h localhost -u $CONFIG_JP1User -p $CONFIG_JP1Password -F $JobnetPath 2>&1
        Write-Log "結果: $($releaseOutput -join ' ')"
        if ($LASTEXITCODE -ne 0) {
            Write-Log '[ERROR] 保留解除失敗'
            Write-Output "RESULT_STATUS:保留解除失敗"
            Write-Output "RESULT_MESSAGE:$($releaseOutput -join ' ')"
            return
        }
        Write-Log '[成功] 保留解除完了'
    }

    # ajsentry実行
    Write-Log '[実行] ajsentry - ジョブ起動'
    Write-Log "コマンド: ajsentry.exe -h localhost -u $CONFIG_JP1User -p ***** -F $JobnetPath"
    $output = & "$jp1BinPath\ajsentry.exe" -h localhost -u $CONFIG_JP1User -p $CONFIG_JP1Password -F $JobnetPath 2>&1
    Write-Log "結果: $($output -join ' ')"
    $exitCode = $LASTEXITCODE

    if ($exitCode -ne 0) {
        Write-Log '[ERROR] ジョブ起動失敗'
        Write-Output "RESULT_STATUS:起動失敗"
        Write-Output "RESULT_MESSAGE:$($output -join ' ')"
        return
    }
    Write-Log '[成功] ジョブ起動完了'

    # 完了待ち
    if ($WaitCompletion -eq "true") {
        Write-Log '[待機] ジョブ完了待ち開始...'
        $startTime = Get-Date
        $isRunning = $true
        $pollCount = 0

        while ($isRunning) {
            $pollCount++
            if ($CONFIG_Timeout -gt 0 -and ((Get-Date) - $startTime).TotalSeconds -ge $CONFIG_Timeout) {
                Write-Log '[TIMEOUT] タイムアウトしました'
                Write-Output "RESULT_STATUS:タイムアウト"
                break
            }

            $statusResult = & "$jp1BinPath\ajsstatus.exe" -h localhost -u $CONFIG_JP1User -p $CONFIG_JP1Password -F $JobnetPath 2>&1
            $statusStr = ($statusResult -join ' ').ToLower()
            Write-Log "[ポーリング $pollCount] ステータス: $($statusResult -join ' ')"

            if ($statusStr -match 'ended abnormally|abnormal end|abend|killed|failed') {
                Write-Log '[完了] 異常終了'
                Write-Output "RESULT_STATUS:異常終了"
                $isRunning = $false
            } elseif ($statusStr -match 'end normally|ended normally|normal end|completed') {
                Write-Log '[完了] 正常終了'
                Write-Output "RESULT_STATUS:正常終了"
                $isRunning = $false
            } else {
                Start-Sleep -Seconds $CONFIG_PollingInterval
            }
        }

        # 詳細取得
        Write-Log '[実行] ajsshow - 詳細取得'
        $ajsshowPath = "$jp1BinPath\ajsshow.exe"
        if (Test-Path $ajsshowPath) {
            $showResult = & $ajsshowPath -h localhost -u $CONFIG_JP1User -p $CONFIG_JP1Password -F $JobnetPath -E 2>&1
            Write-Log "詳細: $($showResult -join ' ')"
            Write-Output "RESULT_MESSAGE:$($showResult -join ' ')"
        }
    } else {
        Write-Log '[完了] 起動成功（完了待ちなし）'
        Write-Output "RESULT_STATUS:起動成功"
        Write-Output "RESULT_MESSAGE:$($output -join ' ')"
    }

    Write-Log "[完了] ログファイル: $LogFile"
}

#==============================================================================
# リモートモード実行
#==============================================================================
function Invoke-RemoteMode {
    Write-Log '--------------------------------------------------------------------------------'
    Write-Log "ジョブネット: $JobnetPath"
    Write-Log "接続先: $CONFIG_JP1Server (リモートモード)"
    Write-Log '--------------------------------------------------------------------------------'

    # 認証情報
    $securePass = ConvertTo-SecureString $CONFIG_RemotePassword -AsPlainText -Force
    $cred = New-Object System.Management.Automation.PSCredential($CONFIG_RemoteUser, $securePass)

    # WinRM設定の保存と自動設定
    $originalTrustedHosts = $null
    $winrmConfigChanged = $false
    $winrmServiceWasStarted = $false

    try {
        Write-Log '[準備] WinRM設定を確認中...'
        # 現在のTrustedHostsを取得
        $originalTrustedHosts = (Get-Item WSMan:\localhost\Client\TrustedHosts -ErrorAction SilentlyContinue).Value

        # WinRMサービスの起動確認
        $winrmService = Get-Service -Name WinRM -ErrorAction SilentlyContinue
        if ($winrmService.Status -ne 'Running') {
            Write-Log '[準備] WinRMサービスを起動'
            Start-Service -Name WinRM -ErrorAction Stop
            $winrmServiceWasStarted = $true
        }

        # TrustedHostsに接続先を追加（必要な場合のみ）
        if ($originalTrustedHosts -notmatch $CONFIG_JP1Server) {
            Write-Log '[準備] TrustedHostsに接続先を追加'
            if ($originalTrustedHosts) {
                Set-Item WSMan:\localhost\Client\TrustedHosts -Value "$originalTrustedHosts,$CONFIG_JP1Server" -Force -Confirm:$false
            } else {
                Set-Item WSMan:\localhost\Client\TrustedHosts -Value $CONFIG_JP1Server -Force -Confirm:$false
            }
            $winrmConfigChanged = $true
        }

        # リモートセッション作成
        Write-Log '[接続] リモートセッション作成中...'
        $session = New-PSSession -ComputerName $CONFIG_JP1Server -Credential $cred -ErrorAction Stop
        Write-Log '[接続] セッション確立完了'

        # 保留解除（必要な場合）
        if ($IsHold -eq "true") {
            Write-Log '[実行] ajsrelease - 保留解除（リモート）'
            $releaseResult = Invoke-Command -Session $session -ScriptBlock {
                param($jp1User, $jp1Pass, $jobnetPath)
                $ajsreleasePath = 'C:\Program Files\HITACHI\JP1AJS3\bin\ajsrelease.exe'
                if (-not (Test-Path $ajsreleasePath)) { $ajsreleasePath = 'C:\Program Files\Hitachi\JP1AJS2\bin\ajsrelease.exe' }
                $output = & $ajsreleasePath -h localhost -u $jp1User -p $jp1Pass -F $jobnetPath 2>&1
                @{ ExitCode = $LASTEXITCODE; Output = ($output -join ' ') }
            } -ArgumentList $CONFIG_JP1User, $CONFIG_JP1Password, $JobnetPath
            Write-Log "結果: $($releaseResult.Output)"

            if ($releaseResult.ExitCode -ne 0) {
                Write-Log '[ERROR] 保留解除失敗'
                Write-Output "RESULT_STATUS:保留解除失敗"
                Write-Output "RESULT_MESSAGE:$($releaseResult.Output)"
                Remove-PSSession $session
                return
            }
            Write-Log '[成功] 保留解除完了'
        }

        # ajsentry実行
        Write-Log '[実行] ajsentry - ジョブ起動（リモート）'
        $entryResult = Invoke-Command -Session $session -ScriptBlock {
            param($jp1User, $jp1Pass, $jobnetPath)
            $ajsentryPath = 'C:\Program Files\HITACHI\JP1AJS3\bin\ajsentry.exe'
            if (-not (Test-Path $ajsentryPath)) { $ajsentryPath = 'C:\Program Files\Hitachi\JP1AJS2\bin\ajsentry.exe' }
            $output = & $ajsentryPath -h localhost -u $jp1User -p $jp1Pass -F $jobnetPath 2>&1
            @{ ExitCode = $LASTEXITCODE; Output = ($output -join ' ') }
        } -ArgumentList $CONFIG_JP1User, $CONFIG_JP1Password, $JobnetPath
        Write-Log "結果: $($entryResult.Output)"

        if ($entryResult.ExitCode -ne 0) {
            Write-Log '[ERROR] ジョブ起動失敗'
            Write-Output "RESULT_STATUS:起動失敗"
            Write-Output "RESULT_MESSAGE:$($entryResult.Output)"
            Remove-PSSession $session
            return
        }
        Write-Log '[成功] ジョブ起動完了'

        # 完了待ち
        if ($WaitCompletion -eq "true") {
            Write-Log '[待機] ジョブ完了待ち開始...'
            $startTime = Get-Date
            $isRunning = $true
            $pollCount = 0

            while ($isRunning) {
                $pollCount++
                if ($CONFIG_Timeout -gt 0 -and ((Get-Date) - $startTime).TotalSeconds -ge $CONFIG_Timeout) {
                    Write-Log '[TIMEOUT] タイムアウトしました'
                    Write-Output "RESULT_STATUS:タイムアウト"
                    break
                }

                $statusResult = Invoke-Command -Session $session -ScriptBlock {
                    param($jp1User, $jp1Pass, $jobnetPath)
                    $ajsstatusPath = 'C:\Program Files\HITACHI\JP1AJS3\bin\ajsstatus.exe'
                    if (-not (Test-Path $ajsstatusPath)) { $ajsstatusPath = 'C:\Program Files\Hitachi\JP1AJS2\bin\ajsstatus.exe' }
                    & $ajsstatusPath -h localhost -u $jp1User -p $jp1Pass -F $jobnetPath 2>&1
                } -ArgumentList $CONFIG_JP1User, $CONFIG_JP1Password, $JobnetPath

                $statusStr = ($statusResult -join ' ').ToLower()
                Write-Log "[ポーリング $pollCount] ステータス: $($statusResult -join ' ')"
                if ($statusStr -match 'ended abnormally|abnormal end|abend|killed|failed') {
                    Write-Log '[完了] 異常終了'
                    Write-Output "RESULT_STATUS:異常終了"
                    $isRunning = $false
                } elseif ($statusStr -match 'end normally|ended normally|normal end|completed') {
                    Write-Log '[完了] 正常終了'
                    Write-Output "RESULT_STATUS:正常終了"
                    $isRunning = $false
                } else {
                    Start-Sleep -Seconds $CONFIG_PollingInterval
                }
            }

            # 詳細取得
            Write-Log '[実行] ajsshow - 詳細取得（リモート）'
            $showResult = Invoke-Command -Session $session -ScriptBlock {
                param($jp1User, $jp1Pass, $jobnetPath)
                $ajsshowPath = 'C:\Program Files\HITACHI\JP1AJS3\bin\ajsshow.exe'
                if (-not (Test-Path $ajsshowPath)) { $ajsshowPath = 'C:\Program Files\Hitachi\JP1AJS2\bin\ajsshow.exe' }
                if (Test-Path $ajsshowPath) { & $ajsshowPath -h localhost -u $jp1User -p $jp1Pass -F $jobnetPath -E 2>&1 }
            } -ArgumentList $CONFIG_JP1User, $CONFIG_JP1Password, $JobnetPath
            Write-Log "詳細: $($showResult -join ' ')"
            Write-Output "RESULT_MESSAGE:$($showResult -join ' ')"
        } else {
            Write-Log '[完了] 起動成功（完了待ちなし）'
            Write-Output "RESULT_STATUS:起動成功"
            Write-Output "RESULT_MESSAGE:$($entryResult.Output)"
        }

        Write-Log '[クリーンアップ] セッション終了'
        Remove-PSSession $session
    } catch {
        Write-Log "[EXCEPTION] $($_.Exception.Message)"
        Write-Output "RESULT_STATUS:起動失敗"
        Write-Output "RESULT_MESSAGE:$($_.Exception.Message)"
    } finally {
        # WinRM設定の復元
        Write-Log '[クリーンアップ] WinRM設定を復元中...'
        if ($winrmConfigChanged) {
            if ($originalTrustedHosts) {
                Set-Item WSMan:\localhost\Client\TrustedHosts -Value $originalTrustedHosts -Force -Confirm:$false -ErrorAction SilentlyContinue
            } else {
                Clear-Item WSMan:\localhost\Client\TrustedHosts -Force -Confirm:$false -ErrorAction SilentlyContinue
            }
        }
        if ($winrmServiceWasStarted) {
            Stop-Service -Name WinRM -Force -ErrorAction SilentlyContinue
        }
        Write-Log "[完了] ログファイル: $LogFile"
    }
}

#==============================================================================
# メイン処理
#==============================================================================
try {
    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host "  JP1 ジョブ実行バッチ" -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "ジョブネット: $JobnetPath"
    Write-Host "実行モード: $CONFIG_Mode"
    Write-Host "ログファイル: $LogFile"
    Write-Host ""

    if ($CONFIG_Mode -eq "ローカル") {
        Invoke-LocalMode
    } else {
        Invoke-RemoteMode
    }

    Write-Host ""
    Write-Host "完了しました。ログファイル: $LogFile" -ForegroundColor Green
} catch {
    Write-Log "[EXCEPTION] $($_.Exception.Message)"
    Write-Output "RESULT_STATUS:起動失敗"
    Write-Output "RESULT_MESSAGE:$($_.Exception.Message)"
    Write-Host "[ERROR] $($_.Exception.Message)" -ForegroundColor Red
}
