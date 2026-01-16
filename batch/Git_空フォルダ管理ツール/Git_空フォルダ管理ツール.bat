<# :
@echo off
chcp 65001 >nul
title Git 空フォルダ管理ツール
setlocal

powershell -NoProfile -ExecutionPolicy Bypass -Command "iex ((gc '%~f0' -Encoding UTF8) -join \"`n\")"
set EXITCODE=%ERRORLEVEL%
pause
exit /b %EXITCODE%
: #>

# ============================================================
#  Git 空フォルダ管理ツール
#  空のフォルダを検出し、.gitignoreファイルを作成します
# ============================================================

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Git 空フォルダ管理ツール" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

#region 設定
# 対象フォルダ（空の場合はカレントディレクトリ）
$targetPath = ""

# .gitignoreファイルの内容
$gitignoreContent = @"
# このファイルは空フォルダをGitで管理するために自動生成されました
# This file was auto-generated to keep this empty folder in Git

# このフォルダ内のすべてのファイルを無視（.gitignore自身を除く）
*
!.gitignore
"@

# 除外するフォルダ名（正規表現パターン）
$excludePatterns = @(
    "^\.git$",           # .gitフォルダ
    "^node_modules$",    # node_modules
    "^\.vs$",            # Visual Studio
    "^\.vscode$",        # VS Code
    "^__pycache__$",     # Python cache
    "^\.idea$"           # JetBrains IDE
)
#endregion

#region メイン処理

# 対象パスの決定
if ([string]::IsNullOrWhiteSpace($targetPath)) {
    $targetPath = Get-Location
}

# パスの存在確認
if (-not (Test-Path $targetPath -PathType Container)) {
    Write-Host "[エラー] 指定されたパスが存在しません: $targetPath" -ForegroundColor Red
    exit 1
}

$targetPath = (Resolve-Path $targetPath).Path
Write-Host "対象フォルダ: $targetPath" -ForegroundColor Yellow
Write-Host ""

# 除外パターンをまとめた正規表現
$excludeRegex = ($excludePatterns -join "|")

# 空フォルダを検出する関数
function Get-EmptyFolders {
    param (
        [string]$Path
    )

    $emptyFolders = @()

    # すべてのサブフォルダを取得（除外パターンに一致するフォルダはスキップ）
    $allFolders = Get-ChildItem -Path $Path -Directory -Recurse -Force -ErrorAction SilentlyContinue | Where-Object {
        $folderName = $_.Name
        $shouldExclude = $false

        # 除外パターンチェック
        foreach ($pattern in $excludePatterns) {
            if ($folderName -match $pattern) {
                $shouldExclude = $true
                break
            }
        }

        # パス内に.gitが含まれる場合も除外
        if ($_.FullName -match "[\\/]\.git[\\/]") {
            $shouldExclude = $true
        }

        -not $shouldExclude
    }

    foreach ($folder in $allFolders) {
        # フォルダ内のアイテム数を確認（隠しファイル含む）
        $items = Get-ChildItem -Path $folder.FullName -Force -ErrorAction SilentlyContinue

        if ($items.Count -eq 0) {
            $emptyFolders += $folder.FullName
        }
    }

    return $emptyFolders
}

# 空フォルダを検出
Write-Host "空フォルダを検索中..." -ForegroundColor Gray
$emptyFolders = Get-EmptyFolders -Path $targetPath

if ($emptyFolders.Count -eq 0) {
    Write-Host ""
    Write-Host "[結果] 空のフォルダは見つかりませんでした。" -ForegroundColor Green
    exit 0
}

# 検出結果を表示
Write-Host ""
Write-Host "検出された空フォルダ: $($emptyFolders.Count) 件" -ForegroundColor Yellow
Write-Host "----------------------------------------" -ForegroundColor Gray

foreach ($folder in $emptyFolders) {
    $relativePath = $folder.Replace($targetPath, "").TrimStart("\", "/")
    if ([string]::IsNullOrWhiteSpace($relativePath)) {
        $relativePath = "(ルートフォルダ)"
    }
    Write-Host "  $relativePath" -ForegroundColor White
}

Write-Host "----------------------------------------" -ForegroundColor Gray
Write-Host ""

# 確認プロンプト
$confirm = Read-Host "これらのフォルダに .gitignore を作成しますか？ (y/n)"
if ($confirm -ne "y" -and $confirm -ne "Y") {
    Write-Host ""
    Write-Host "キャンセルしました。" -ForegroundColor Yellow
    exit 0
}

Write-Host ""

# .gitignoreを作成
$createdCount = 0
$skippedCount = 0
$errorCount = 0

foreach ($folder in $emptyFolders) {
    $gitignorePath = Join-Path $folder ".gitignore"
    $relativePath = $folder.Replace($targetPath, "").TrimStart("\", "/")
    if ([string]::IsNullOrWhiteSpace($relativePath)) {
        $relativePath = "(ルートフォルダ)"
    }

    try {
        # 既に.gitignoreが存在するかチェック
        if (Test-Path $gitignorePath) {
            Write-Host "  [スキップ] $relativePath\.gitignore (既に存在)" -ForegroundColor Yellow
            $skippedCount++
        } else {
            # .gitignoreを作成
            $gitignoreContent | Out-File -FilePath $gitignorePath -Encoding UTF8 -NoNewline
            Write-Host "  [作成] $relativePath\.gitignore" -ForegroundColor Green
            $createdCount++
        }
    } catch {
        Write-Host "  [エラー] $relativePath\.gitignore - $($_.Exception.Message)" -ForegroundColor Red
        $errorCount++
    }
}

# 結果サマリー
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "処理結果:" -ForegroundColor Cyan
Write-Host "  作成: $createdCount 件" -ForegroundColor Green
Write-Host "  スキップ: $skippedCount 件" -ForegroundColor Yellow
Write-Host "  エラー: $errorCount 件" -ForegroundColor $(if ($errorCount -gt 0) { "Red" } else { "Gray" })
Write-Host "========================================" -ForegroundColor Cyan

if ($errorCount -gt 0) {
    exit 1
}

#endregion
