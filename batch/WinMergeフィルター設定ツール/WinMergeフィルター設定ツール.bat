<# :
@echo off
chcp 65001 >nul
title WinMerge フィルター設定ツール
setlocal

rem UNCパス対応
pushd "%~dp0"

powershell -NoProfile -ExecutionPolicy Bypass -Command "$scriptDir=('%~dp0' -replace '\\$',''); try { iex ((gc '%~f0' -Encoding UTF8) -join \"`n\") } finally { Set-Location C:\ }"
set EXITCODE=%ERRORLEVEL%

popd

pause
exit /b %EXITCODE%
: #>

# ============================================================
#  WinMerge フィルター設定ツール
#  Git管理ファイルを除外するフィルターをWinMergeに組み込みます
# ============================================================

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  WinMerge フィルター設定ツール" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

#region 設定
# 組み込むフィルターファイル（このバッチと同じフォルダに配置）
$filterFiles = @(
    "GitFiles.flt"
)
#endregion

#region WinMerge確認
$winmergePath = $null
$winmergeLocations = @(
    "${env:ProgramFiles}\WinMerge\WinMergeU.exe",
    "${env:ProgramFiles(x86)}\WinMerge\WinMergeU.exe"
)

foreach ($path in $winmergeLocations) {
    if (Test-Path $path) {
        $winmergePath = $path
        break
    }
}

if (-not $winmergePath) {
    Write-Host "[エラー] WinMergeがインストールされていません" -ForegroundColor Red
    Write-Host ""
    Write-Host "WinMergeをインストールしてから再度実行してください。" -ForegroundColor Yellow
    Write-Host "ダウンロード: https://winmerge.org/downloads/" -ForegroundColor Gray
    Write-Host ""
    exit 1
}

Write-Host "[OK] WinMergeが見つかりました" -ForegroundColor Green
Write-Host "  パス: $winmergePath" -ForegroundColor Gray
Write-Host ""
#endregion

#region フィルターディレクトリ
$filterDir = "$env:APPDATA\WinMerge\Filters"

if (-not (Test-Path $filterDir)) {
    New-Item -ItemType Directory -Path $filterDir -Force | Out-Null
    Write-Host "[作成] フィルターディレクトリを作成しました" -ForegroundColor Green
    Write-Host "  パス: $filterDir" -ForegroundColor Gray
    Write-Host ""
}
#endregion

#region フィルターファイルのコピー
Write-Host "フィルターファイルを組み込みます..." -ForegroundColor Yellow
Write-Host ""

$copiedCount = 0
$skippedCount = 0
$errorCount = 0

foreach ($filterFile in $filterFiles) {
    $sourcePath = Join-Path $scriptDir $filterFile
    $destPath = Join-Path $filterDir $filterFile

    if (-not (Test-Path $sourcePath)) {
        Write-Host "  [エラー] $filterFile が見つかりません" -ForegroundColor Red
        $errorCount++
        continue
    }

    # 既存ファイルの確認
    if (Test-Path $destPath) {
        # 内容を比較
        $sourceContent = Get-Content $sourcePath -Raw -Encoding UTF8
        $destContent = Get-Content $destPath -Raw -Encoding UTF8

        if ($sourceContent -eq $destContent) {
            Write-Host "  [スキップ] $filterFile (既に同じ内容)" -ForegroundColor Yellow
            $skippedCount++
            continue
        }

        Write-Host "  [情報] $filterFile が既に存在します（内容が異なる）" -ForegroundColor Yellow
        $overwrite = Read-Host "  上書きしますか？ (y/n)"
        if ($overwrite -ne "y" -and $overwrite -ne "Y") {
            Write-Host "  [スキップ] $filterFile" -ForegroundColor Yellow
            $skippedCount++
            continue
        }
    }

    try {
        Copy-Item -Path $sourcePath -Destination $destPath -Force
        Write-Host "  [成功] $filterFile を組み込みました" -ForegroundColor Green
        $copiedCount++
    } catch {
        Write-Host "  [エラー] $filterFile のコピーに失敗: $($_.Exception.Message)" -ForegroundColor Red
        $errorCount++
    }
}
#endregion

#region 結果表示
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "処理結果:" -ForegroundColor Cyan
Write-Host "  組み込み: $copiedCount 件" -ForegroundColor Green
Write-Host "  スキップ: $skippedCount 件" -ForegroundColor Yellow
Write-Host "  エラー: $errorCount 件" -ForegroundColor $(if ($errorCount -gt 0) { "Red" } else { "Gray" })
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

if ($copiedCount -gt 0 -or $skippedCount -gt 0) {
    Write-Host "フィルターの有効化方法:" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  1. WinMergeを起動" -ForegroundColor White
    Write-Host "  2. メニュー: ツール > フィルター" -ForegroundColor White
    Write-Host "  3. 「Git管理ファイル除外フィルター」にチェック" -ForegroundColor White
    Write-Host "  4. OK をクリック" -ForegroundColor White
    Write-Host ""
    Write-Host "[ヒント] フォルダ比較時にフィルターが自動適用されます" -ForegroundColor Gray
    Write-Host ""

    # WinMergeを開くか確認
    $openWinMerge = Read-Host "WinMergeを起動してフィルター設定を確認しますか？ (y/n)"
    if ($openWinMerge -eq "y" -or $openWinMerge -eq "Y") {
        Start-Process $winmergePath
        Write-Host ""
        Write-Host "[起動] WinMergeを起動しました" -ForegroundColor Green
    }
}
#endregion

if ($errorCount -gt 0) {
    exit 1
}
