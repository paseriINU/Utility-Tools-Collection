<# :
@echo off
chcp 932 >nul
setlocal

powershell -NoProfile -ExecutionPolicy Bypass -Command "$host.UI.RawUI.WindowTitle='JP1 ジョブログ取得ツール'; iex ((gc '%~f0' -Encoding UTF8) -join \"`n\")"
set EXITCODE=%ERRORLEVEL%
pause
exit /b %EXITCODE%
: #>

<#
.SYNOPSIS
    JP1ジョブログ取得ツール（ローカル実行版）

.DESCRIPTION
    JP1/AJS3の指定されたジョブの標準出力（スプール）を取得し、
    クリップボードにコピーします。

.NOTES
    作成日: 2025-12-21
    バージョン: 1.0

    使い方:
    1. 下記の「設定セクション」を編集
    2. このファイルをダブルクリックで実行
    3. 取得したログがクリップボードにコピーされます
#>

# ==============================================================================
# ■ 設定セクション（ここを編集してください）
# ==============================================================================

$Config = @{
    # スケジューラーサービス名（デフォルト: AJSROOT1）
    SchedulerService = "AJSROOT1"

    # 取得対象のジョブのフルパス（ジョブネット内のジョブを指定）
    # 例: "/main_unit/jobgroup1/daily_batch/job1"
    JobPath = "/main_unit/jobgroup1/daily_batch/job1"

    # JP1ユーザー名（空の場合は現在のログインユーザーで実行）
    JP1User = ""

    # JP1パスワード（空の場合は実行時に入力、JP1Userが空の場合は不要）
    JP1Password = ""

    # ajsshowコマンドのパス
    AjsshowPath = "C:\Program Files (x86)\HITACHI\JP1AJS2\bin\ajsshow.exe"

    # 取得するスプールの種類（stdout=標準出力、stderr=標準エラー出力、both=両方）
    SpoolType = "stdout"

    # クリップボードにコピーするか
    CopyToClipboard = $true

    # コンソールにも出力するか
    ShowInConsole = $true
}

# ==============================================================================
# ■ メイン処理（以下は編集不要）
# ==============================================================================

$ErrorActionPreference = "Stop"
# JP1コマンドの出力はShift_JIS（CP932）のためエンコーディングを設定
[Console]::OutputEncoding = [System.Text.Encoding]::GetEncoding(932)

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  JP1 ジョブログ取得ツール" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

#region 設定確認
Write-Host "設定内容:" -ForegroundColor Yellow
Write-Host "  スケジューラーサービス: $($Config.SchedulerService)" -ForegroundColor White
Write-Host "  ジョブパス            : $($Config.JobPath)" -ForegroundColor White
Write-Host "  スプール種類          : $($Config.SpoolType)" -ForegroundColor White
Write-Host ""
#endregion

#region コマンドパス確認
if (-not (Test-Path $Config.AjsshowPath)) {
    Write-Host "[エラー] ajsshowコマンドが見つかりません" -ForegroundColor Red
    Write-Host "  パス: $($Config.AjsshowPath)" -ForegroundColor Red
    Write-Host ""
    Write-Host "以下のパスを確認してください:" -ForegroundColor Yellow
    Write-Host "  - C:\Program Files (x86)\HITACHI\JP1AJS2\bin\ajsshow.exe" -ForegroundColor Gray
    Write-Host "  - C:\Program Files\HITACHI\JP1AJS2\bin\ajsshow.exe" -ForegroundColor Gray
    Write-Host "  - C:\Program Files\HITACHI\JP1AJS3\bin\ajsshow.exe" -ForegroundColor Gray
    Write-Host "  - C:\Program Files (x86)\HITACHI\JP1AJS3\bin\ajsshow.exe" -ForegroundColor Gray
    exit 1
}
#endregion

#region JP1認証情報の準備
$authParams = @()
if (-not [string]::IsNullOrEmpty($Config.JP1User)) {
    $authParams += "-u"
    $authParams += $Config.JP1User

    if ([string]::IsNullOrEmpty($Config.JP1Password)) {
        Write-Host "[注意] JP1パスワードが設定されていません。" -ForegroundColor Yellow
        $securePass = Read-Host "JP1パスワードを入力してください" -AsSecureString
        $Config.JP1Password = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
            [Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePass)
        )
        Write-Host ""
    }

    $authParams += "-p"
    $authParams += $Config.JP1Password
}
#endregion

#region スプール取得
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "ジョブのスプールを取得中..." -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

try {
    # ajsshow -B オプションでスプールを取得
    # -B stdout: 標準出力
    # -B stderr: 標準エラー出力
    $spoolContent = ""
    $spoolResults = @()

    $spoolTypes = @()
    switch ($Config.SpoolType) {
        "stdout" { $spoolTypes = @("stdout") }
        "stderr" { $spoolTypes = @("stderr") }
        "both"   { $spoolTypes = @("stdout", "stderr") }
        default  { $spoolTypes = @("stdout") }
    }

    foreach ($type in $spoolTypes) {
        Write-Host "  取得中: $type ..." -ForegroundColor Cyan

        $cmdArgs = @(
            "-F", $Config.SchedulerService
            "-B", $type
            "-g", "001"  # 最新の実行世代を指定
            $Config.JobPath
        )

        if ($authParams.Count -gt 0) {
            $cmdArgs = $authParams + $cmdArgs
        }

        $result = & $Config.AjsshowPath @cmdArgs 2>&1
        $exitCode = $LASTEXITCODE

        if ($exitCode -eq 0) {
            $output = ($result | Out-String).Trim()
            if (-not [string]::IsNullOrEmpty($output)) {
                $spoolResults += @{
                    Type = $type
                    Content = $output
                }
                Write-Host "  [OK] $type を取得しました" -ForegroundColor Green
            } else {
                Write-Host "  [情報] $type は空です" -ForegroundColor Gray
            }
        } else {
            Write-Host "  [警告] $type の取得に失敗しました (終了コード: $exitCode)" -ForegroundColor Yellow
            # エラーメッセージを表示
            $errorOutput = ($result | Out-String).Trim()
            if (-not [string]::IsNullOrEmpty($errorOutput)) {
                Write-Host "  エラー: $errorOutput" -ForegroundColor Red
            }
        }
    }

    Write-Host ""

    # 結果の整形
    if ($spoolResults.Count -gt 0) {
        $formattedContent = @()

        foreach ($spool in $spoolResults) {
            if ($spoolResults.Count -gt 1) {
                $formattedContent += "===== $($spool.Type.ToUpper()) ====="
            }
            $formattedContent += $spool.Content
            if ($spoolResults.Count -gt 1) {
                $formattedContent += ""
            }
        }

        $spoolContent = $formattedContent -join "`n"

        # コンソールに出力
        if ($Config.ShowInConsole) {
            Write-Host "========================================" -ForegroundColor Cyan
            Write-Host "取得したスプール内容:" -ForegroundColor Cyan
            Write-Host "========================================" -ForegroundColor Cyan
            Write-Host ""
            Write-Host $spoolContent -ForegroundColor White
            Write-Host ""
        }

        # クリップボードにコピー
        if ($Config.CopyToClipboard) {
            $spoolContent | Set-Clipboard
            Write-Host "========================================" -ForegroundColor Green
            Write-Host "[OK] スプール内容をクリップボードにコピーしました" -ForegroundColor Green
            Write-Host "========================================" -ForegroundColor Green
        }

        $exitCode = 0
    } else {
        Write-Host "========================================" -ForegroundColor Red
        Write-Host "[エラー] スプールを取得できませんでした" -ForegroundColor Red
        Write-Host "========================================" -ForegroundColor Red
        Write-Host ""
        Write-Host "以下を確認してください:" -ForegroundColor Yellow
        Write-Host "  - ジョブパスが正しいか: $($Config.JobPath)" -ForegroundColor Yellow
        Write-Host "  - ジョブが実行済みか" -ForegroundColor Yellow
        Write-Host "  - JP1ユーザーに権限があるか" -ForegroundColor Yellow
        Write-Host "  - スプールが保存されているか" -ForegroundColor Yellow
        $exitCode = 1
    }

} catch {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "[エラー] スプール取得中にエラーが発生しました" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "エラー詳細:" -ForegroundColor Yellow
    Write-Host $_.Exception.Message -ForegroundColor Red
    $exitCode = 1
}
#endregion

Write-Host ""
exit $exitCode
