<# :
@echo off
chcp 932 >nul
title JP1 ジョブログ取得ツール
setlocal

powershell -NoProfile -ExecutionPolicy Bypass -Command "iex ((gc '%~f0' -Encoding UTF8) -join \"`n\")"
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
    jpqjobgetコマンドを使用してスプールを取得します。

.NOTES
    作成日: 2025-12-21
    バージョン: 2.0

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

#region コマンドパス検索
$ajsshowPath = $null
$jpqjobgetPath = $null

$searchPaths = @(
    "C:\Program Files (x86)\Hitachi\JP1AJS2\bin",
    "C:\Program Files\Hitachi\JP1AJS2\bin",
    "C:\Program Files (x86)\HITACHI\JP1AJS3\bin",
    "C:\Program Files\HITACHI\JP1AJS3\bin"
)

foreach ($basePath in $searchPaths) {
    $ajsPath = Join-Path $basePath "ajsshow.exe"
    $jpqPath = Join-Path $basePath "jpqjobget.exe"

    if ((Test-Path $ajsPath) -and (-not $ajsshowPath)) {
        $ajsshowPath = $ajsPath
    }
    if ((Test-Path $jpqPath) -and (-not $jpqjobgetPath)) {
        $jpqjobgetPath = $jpqPath
    }
}

if (-not $ajsshowPath) {
    Write-Host "[エラー] ajsshowコマンドが見つかりません" -ForegroundColor Red
    Write-Host ""
    Write-Host "以下のパスを確認してください:" -ForegroundColor Yellow
    foreach ($basePath in $searchPaths) {
        Write-Host "  - $basePath\ajsshow.exe" -ForegroundColor Gray
    }
    exit 1
}

if (-not $jpqjobgetPath) {
    Write-Host "[エラー] jpqjobgetコマンドが見つかりません" -ForegroundColor Red
    Write-Host ""
    Write-Host "以下のパスを確認してください:" -ForegroundColor Yellow
    foreach ($basePath in $searchPaths) {
        Write-Host "  - $basePath\jpqjobget.exe" -ForegroundColor Gray
    }
    exit 1
}

Write-Host "コマンドパス:" -ForegroundColor Yellow
Write-Host "  ajsshow  : $ajsshowPath" -ForegroundColor White
Write-Host "  jpqjobget: $jpqjobgetPath" -ForegroundColor White
Write-Host ""
#endregion

#region 設定確認
Write-Host "設定内容:" -ForegroundColor Yellow
Write-Host "  スケジューラーサービス: $($Config.SchedulerService)" -ForegroundColor White
Write-Host "  ジョブパス            : $($Config.JobPath)" -ForegroundColor White
Write-Host "  スプール種類          : $($Config.SpoolType)" -ForegroundColor White
Write-Host ""
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

#region ジョブ情報の取得（ajsshow）
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "ジョブ情報を取得中..." -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

try {
    # ajsshow -g 1 -i '%## %I' でジョブ名とジョブ番号を取得
    # %## = ジョブ名、%I = ジョブ番号
    $cmdArgs = @(
        "-F", $Config.SchedulerService
        "-g", "1"
        "-i", "%## %I"
        $Config.JobPath
    )

    if ($authParams.Count -gt 0) {
        $cmdArgs = $authParams + $cmdArgs
    }

    Write-Host "実行コマンド: ajsshow $($cmdArgs -join ' ')" -ForegroundColor Gray
    $result = & $ajsshowPath @cmdArgs 2>&1
    $exitCode = $LASTEXITCODE

    Write-Host "ajsshow結果:" -ForegroundColor Gray
    $result | ForEach-Object { Write-Host "  $_" -ForegroundColor Gray }
    Write-Host ""

    if ($exitCode -ne 0) {
        Write-Host "[エラー] ジョブ情報の取得に失敗しました (終了コード: $exitCode)" -ForegroundColor Red
        Write-Host ""
        Write-Host "以下を確認してください:" -ForegroundColor Yellow
        Write-Host "  - ジョブパスが正しいか: $($Config.JobPath)" -ForegroundColor Yellow
        Write-Host "  - ジョブが実行済みか" -ForegroundColor Yellow
        Write-Host "  - JP1ユーザーに権限があるか" -ForegroundColor Yellow
        exit 1
    }

    # ジョブ番号を抽出（出力形式: "ジョブ名 ジョブ番号"）
    $output = ($result | Out-String).Trim()
    $jobNo = $null

    # 各行を解析してジョブ番号を取得
    foreach ($line in $output -split "`n") {
        $line = $line.Trim()
        if ($line -match '\s+(\d+)\s*$') {
            $jobNo = $matches[1]
            break
        }
    }

    if (-not $jobNo) {
        Write-Host "[エラー] ジョブ番号を取得できませんでした" -ForegroundColor Red
        Write-Host "ajsshow出力: $output" -ForegroundColor Gray
        exit 1
    }

    Write-Host "[OK] ジョブ番号: $jobNo" -ForegroundColor Green
    Write-Host ""

} catch {
    Write-Host "[エラー] ジョブ情報取得中にエラーが発生しました" -ForegroundColor Red
    Write-Host "エラー詳細: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
#endregion

#region スプール取得（jpqjobget）
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "スプールを取得中..." -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

try {
    $spoolContent = ""
    $spoolResults = @()
    $tempDir = [System.IO.Path]::GetTempPath()

    $spoolTypes = @()
    switch ($Config.SpoolType) {
        "stdout" { $spoolTypes = @(@{Type="stdout"; Option="-oso"}) }
        "stderr" { $spoolTypes = @(@{Type="stderr"; Option="-ose"}) }
        "both"   { $spoolTypes = @(@{Type="stderr"; Option="-ose"}, @{Type="stdout"; Option="-oso"}) }
        default  { $spoolTypes = @(@{Type="stdout"; Option="-oso"}) }
    }

    foreach ($spool in $spoolTypes) {
        Write-Host "  取得中: $($spool.Type) ..." -ForegroundColor Cyan

        # 一時ファイルを作成
        $tempFile = Join-Path $tempDir "jp1_spool_$($spool.Type)_$(Get-Date -Format 'yyyyMMddHHmmss').txt"

        $cmdArgs = @(
            "-j", $jobNo
            $spool.Option, $tempFile
        )

        Write-Host "  実行コマンド: jpqjobget $($cmdArgs -join ' ')" -ForegroundColor Gray
        $result = & $jpqjobgetPath @cmdArgs 2>&1
        $exitCode = $LASTEXITCODE

        if ($exitCode -eq 0 -and (Test-Path $tempFile)) {
            $content = Get-Content -Path $tempFile -Raw -Encoding Default
            if (-not [string]::IsNullOrEmpty($content)) {
                $spoolResults += @{
                    Type = $spool.Type
                    Content = $content.Trim()
                }
                Write-Host "  [OK] $($spool.Type) を取得しました" -ForegroundColor Green
            } else {
                Write-Host "  [情報] $($spool.Type) は空です" -ForegroundColor Gray
            }
            # 一時ファイルを削除
            Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
        } else {
            Write-Host "  [警告] $($spool.Type) の取得に失敗しました (終了コード: $exitCode)" -ForegroundColor Yellow
            $errorOutput = ($result | Out-String).Trim()
            if (-not [string]::IsNullOrEmpty($errorOutput)) {
                Write-Host "  詳細: $errorOutput" -ForegroundColor Gray
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
