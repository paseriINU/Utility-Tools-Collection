<# :
@echo off
chcp 65001 >nul
title makeファイル検証ツール
setlocal
powershell -NoProfile -ExecutionPolicy Bypass -Command "$scriptDir=('%~dp0' -replace '\\$',''); iex ((gc '%~f0' -Encoding UTF8) -join \"`n\")"
set EXITCODE=%ERRORLEVEL%
pause
exit /b %EXITCODE%
: #>

#region 設定
# 対象フォルダ（空欄の場合はスクリプトと同じフォルダ）
$TARGET_FOLDER = ""

# 例外ファイル（makeファイル名と異なっていてもOKな.oファイル）
# 拡張子なしで指定（例: "common" は common.o にマッチ）
$EXCEPTION_FILES = @(
    "common",
    "util",
    "shared"
)

# 部分一致を許可するか（library.mk に library_sub.o があってもOK）
$ALLOW_PARTIAL_MATCH = $true
#endregion

#region メイン処理

# タイトル表示
Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  makeファイル .o 検証ツール" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

# 対象フォルダの決定
if ([string]::IsNullOrWhiteSpace($TARGET_FOLDER)) {
    $TARGET_FOLDER = $scriptDir
}

# フォルダ存在チェック
if (-not (Test-Path $TARGET_FOLDER -PathType Container)) {
    Write-Host "[エラー] 対象フォルダが存在しません: $TARGET_FOLDER" -ForegroundColor Red
    exit 1
}

Write-Host "対象フォルダ: $TARGET_FOLDER" -ForegroundColor White
Write-Host "例外ファイル: $($EXCEPTION_FILES -join ', ')" -ForegroundColor Gray
Write-Host "部分一致許可: $(if ($ALLOW_PARTIAL_MATCH) { 'はい' } else { 'いいえ' })" -ForegroundColor Gray
Write-Host ""

# .mkファイルを取得（サブフォルダも含む）
$mkFiles = Get-ChildItem -Path $TARGET_FOLDER -Filter "*.mk" -File -Recurse -ErrorAction SilentlyContinue

if ($mkFiles.Count -eq 0) {
    Write-Host "[情報] .mkファイルが見つかりませんでした。" -ForegroundColor Yellow
    exit 0
}

# 結果を格納するリスト
$allMismatches = @()
$totalMkFiles = 0
$totalOFiles = 0
$totalOkCount = 0
$totalNgCount = 0
$totalExcludedCount = 0

foreach ($mkFile in $mkFiles) {
    $totalMkFiles++
    $mkBaseName = [System.IO.Path]::GetFileNameWithoutExtension($mkFile.Name)
    
    Write-Host "----------------------------------------" -ForegroundColor DarkGray
    # 相対パスを表示
    $relativePath = $mkFile.FullName.Replace($TARGET_FOLDER, "").TrimStart("\", "/")
    Write-Host "[検証中] $relativePath" -ForegroundColor White
    
    # makeファイルの内容を読み込み
    try {
        $content = Get-Content -Path $mkFile.FullName -Raw -Encoding UTF8 -ErrorAction Stop
    } catch {
        try {
            $content = Get-Content -Path $mkFile.FullName -Raw -ErrorAction Stop
        } catch {
            Write-Host "  [警告] ファイルを読み込めませんでした" -ForegroundColor Yellow
            continue
        }
    }
    
    # .oファイルを抽出（パス付きでもファイル名部分のみ取得）
    # パターン: 空白以外の文字列で .o で終わるもの
    $oFileMatches = [regex]::Matches($content, '([^\s\(\)]+)\.o\b')
    
    if ($oFileMatches.Count -eq 0) {
        Write-Host "  [情報] .oファイルの参照が見つかりませんでした" -ForegroundColor Gray
        continue
    }
    
    # ユニークな.oファイル名を抽出
    $oFileNames = @{}
    foreach ($match in $oFileMatches) {
        $fullMatch = $match.Groups[1].Value
        # パスから最後のファイル名部分を抽出
        if ($fullMatch -match '[/\\]([^/\\]+)$') {
            $oName = $Matches[1]
        } else {
            $oName = $fullMatch
        }
        # $変数展開を含む場合はスキップ（例: $(OBJ)）
        if ($oName -notmatch '^\$') {
            $oFileNames[$oName] = $true
        }
    }
    
    $fileMismatches = @()
    
    foreach ($oName in $oFileNames.Keys | Sort-Object) {
        $totalOFiles++
        
        # 例外リストに含まれているかチェック
        if ($EXCEPTION_FILES -contains $oName) {
            Write-Host "  [除外] $oName.o" -ForegroundColor DarkGray
            $totalExcludedCount++
            continue
        }
        
        # 一致判定
        $isMatch = $false
        
        if ($ALLOW_PARTIAL_MATCH) {
            # 部分一致: makeファイル名が.oファイル名に含まれているか
            if ($oName -like "$mkBaseName*" -or $oName -like "*$mkBaseName*") {
                $isMatch = $true
            }
        } else {
            # 完全一致
            if ($oName -eq $mkBaseName) {
                $isMatch = $true
            }
        }
        
        if ($isMatch) {
            Write-Host "  [OK] $oName.o" -ForegroundColor Green
            $totalOkCount++
        } else {
            Write-Host "  [NG] $oName.o" -ForegroundColor Red
            $totalNgCount++
            $fileMismatches += $oName
        }
    }
    
    if ($fileMismatches.Count -gt 0) {
        $allMismatches += [PSCustomObject]@{
            MakeFile = $relativePath
            Mismatches = $fileMismatches
        }
    }
}

# サマリー表示
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "検証結果サマリー" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "検証makeファイル数: $totalMkFiles" -ForegroundColor White
Write-Host "検証.oファイル数:   $totalOFiles" -ForegroundColor White
Write-Host "  OK:    $totalOkCount" -ForegroundColor Green
Write-Host "  NG:    $totalNgCount" -ForegroundColor $(if ($totalNgCount -gt 0) { 'Red' } else { 'Green' })
Write-Host "  除外:  $totalExcludedCount" -ForegroundColor Gray
Write-Host ""

if ($allMismatches.Count -gt 0) {
    Write-Host "不一致一覧:" -ForegroundColor Yellow
    foreach ($item in $allMismatches) {
        Write-Host "  $($item.MakeFile):" -ForegroundColor White
        foreach ($mismatch in $item.Mismatches) {
            Write-Host "    - $mismatch.o" -ForegroundColor Red
        }
    }
    exit 1
} else {
    Write-Host "すべてのファイルが一致しました。" -ForegroundColor Green
    exit 0
}

#endregion
