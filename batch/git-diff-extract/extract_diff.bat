<# :
@echo off
chcp 65001 >nul
title Git 差分ファイル抽出ツール
setlocal

rem UNCパス対応（PushD/PopDで自動マッピング）
pushd "%~dp0"

powershell -NoProfile -ExecutionPolicy Bypass -Command "$scriptDir=('%~dp0' -replace '\\$',''); try { iex ((gc '%~f0' -Encoding UTF8) -join \"`n\") } finally { Set-Location C:\ }"
set EXITCODE=%ERRORLEVEL%

popd

pause
exit /b %EXITCODE%
: #>

# =============================================================================
# Git Diff Extract Tool (PowerShell)
# Gitブランチ間の差分ファイルを抽出してフォルダ構造を保ったままコピー
# =============================================================================

# タイトル表示
Write-Host ""
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host "  Git 差分ファイル抽出ツール" -ForegroundColor Cyan
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host ""

# UTF-8出力設定
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Gitコマンドの存在確認
$gitCommand = Get-Command git -ErrorAction SilentlyContinue
if (-not $gitCommand) {
    Write-Host ""
    Write-Host "========================================================================" -ForegroundColor Red
    Write-Host "  [エラー] Gitがインストールされていません" -ForegroundColor Red
    Write-Host "========================================================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "このスクリプトを実行するには、Gitがインストールされている必要があります。" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Gitのインストール方法:" -ForegroundColor Cyan
    Write-Host "  1. https://git-scm.com/download/win にアクセス" -ForegroundColor White
    Write-Host "  2. 「Download for Windows」をクリック" -ForegroundColor White
    Write-Host "  3. インストーラーをダウンロードして実行" -ForegroundColor White
    Write-Host "  4. インストール後、コマンドプロンプトを再起動" -ForegroundColor White
    Write-Host ""
    exit 1
}

#region 設定セクション
# Gitプロジェクトのパス
$GIT_PROJECT_PATH = "C:\Users\$env:USERNAME\source\Git\project"

# 削除されたファイルも含めるか（$true=含める, $false=含めない）
$INCLUDE_DELETED = $false
#endregion

#region Gitリポジトリ確認
# パス存在確認
if (-not (Test-Path $GIT_PROJECT_PATH)) {
    Write-Host "[エラー] Gitプロジェクトフォルダが見つかりません: $GIT_PROJECT_PATH" -ForegroundColor Red
    Write-Host ""
    Write-Host "フォルダが存在するか確認してください。" -ForegroundColor Yellow
    exit 1
}

Write-Host "Gitプロジェクトパス: $GIT_PROJECT_PATH" -ForegroundColor White
Set-Location $GIT_PROJECT_PATH
Write-Host ""

# Gitリポジトリ確認
git rev-parse --git-dir 2>&1 | Out-Null
if ($LASTEXITCODE -ne 0) {
    Write-Host "[エラー] このフォルダはGit管理下にありません: $GIT_PROJECT_PATH" -ForegroundColor Red
    exit 1
}

# リポジトリのルートディレクトリを取得
$REPO_ROOT = git rev-parse --show-toplevel
$REPO_ROOT = $REPO_ROOT -replace '/', '\'

Write-Host "リポジトリルート: $REPO_ROOT" -ForegroundColor White
Write-Host ""
#endregion

#region ブランチ選択
# すべてのブランチ一覧を取得
Write-Host "ブランチ一覧を取得中..." -ForegroundColor Yellow
Write-Host ""

$allBranches = git branch --format="%(refname:short)" | ForEach-Object { $_.Trim() }

if ($allBranches.Count -eq 0) {
    Write-Host "[エラー] ブランチが見つかりません" -ForegroundColor Red
    exit 1
}

# 比較元ブランチの選択
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host "  比較元ブランチ（基準ブランチ）を選択してください" -ForegroundColor Cyan
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host ""

for ($i = 0; $i -lt $allBranches.Count; $i++) {
    $displayNum = $i + 1
    $branchName = $allBranches[$i]
    Write-Host " $displayNum. $branchName"
}
Write-Host ""

$maxNum = $allBranches.Count
$baseSelection = Read-Host "番号を選択してください (1-$maxNum)"

# 入力検証
if (-not $baseSelection -or $baseSelection -notmatch '^\d+$' -or [int]$baseSelection -lt 1 -or [int]$baseSelection -gt $maxNum) {
    Write-Host "[エラー] 無効な選択です" -ForegroundColor Red
    exit 1
}

$BASE_BRANCH = $allBranches[[int]$baseSelection - 1]
Write-Host ""
Write-Host "[選択] 比較元ブランチ: $BASE_BRANCH" -ForegroundColor Green
Write-Host ""

# 比較先ブランチの選択
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host "  比較先ブランチ（差分を取得したいブランチ）を選択してください" -ForegroundColor Cyan
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host ""

for ($i = 0; $i -lt $allBranches.Count; $i++) {
    $displayNum = $i + 1
    $branchName = $allBranches[$i]

    # 比較元ブランチと同じ場合はマーク
    if ($branchName -eq $BASE_BRANCH) {
        Write-Host " $displayNum. $branchName [比較元]" -ForegroundColor Gray
    } else {
        Write-Host " $displayNum. $branchName"
    }
}
Write-Host ""

$targetSelection = Read-Host "番号を選択してください (1-$maxNum)"

# 入力検証
if (-not $targetSelection -or $targetSelection -notmatch '^\d+$' -or [int]$targetSelection -lt 1 -or [int]$targetSelection -gt $maxNum) {
    Write-Host "[エラー] 無効な選択です" -ForegroundColor Red
    exit 1
}

$TARGET_BRANCH = $allBranches[[int]$targetSelection - 1]

# 同じブランチが選択された場合の警告
if ($BASE_BRANCH -eq $TARGET_BRANCH) {
    Write-Host "[警告] 比較元と比較先が同じブランチです" -ForegroundColor Yellow
    $continue = Read-Host "続行しますか? (y/n)"
    if ($continue -ne "y") {
        Write-Host "処理を中止しました" -ForegroundColor Yellow
        exit 0
    }
}

Write-Host ""
Write-Host "[選択] 比較先ブランチ: $TARGET_BRANCH" -ForegroundColor Green
Write-Host ""

# 出力先フォルダ（デスクトップ + タイムスタンプ）
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$OUTPUT_DIR = "$env:USERPROFILE\Desktop\git_diff_$timestamp"

Write-Host "比較元ブランチ  : $BASE_BRANCH" -ForegroundColor White
Write-Host "比較先ブランチ  : $TARGET_BRANCH" -ForegroundColor White
Write-Host "出力先フォルダ  : $OUTPUT_DIR" -ForegroundColor White
Write-Host ""
#endregion

#region 出力先フォルダ確認
if (Test-Path $OUTPUT_DIR) {
    Write-Host "[警告] 出力先フォルダ '$OUTPUT_DIR' は既に存在します" -ForegroundColor Yellow
    $overwrite = Read-Host "上書きしますか? (y/n)"

    if ($overwrite -ne "y") {
        Write-Host "処理を中止しました" -ForegroundColor Yellow
        exit 0
    }

    Write-Host "既存のフォルダをクリア中..." -ForegroundColor Yellow
    Remove-Item -Path $OUTPUT_DIR -Recurse -Force
}

New-Item -ItemType Directory -Path $OUTPUT_DIR -Force | Out-Null
#endregion

#region 差分ファイル取得
Write-Host "差分ファイルを検出中..." -ForegroundColor Cyan
Write-Host ""

# 差分ファイルリストを取得
if ($INCLUDE_DELETED) {
    # 削除されたファイルも含める
    $diffFiles = git diff --name-only "$BASE_BRANCH...$TARGET_BRANCH"
} else {
    # 削除されたファイルを除外（追加・変更のみ）
    $diffFiles = git diff --name-only --diff-filter=ACMR "$BASE_BRANCH...$TARGET_BRANCH"
}

if (-not $diffFiles -or $diffFiles.Count -eq 0) {
    Write-Host "[情報] 差分ファイルが見つかりませんでした" -ForegroundColor Yellow
    Write-Host "2つのブランチは同じ内容です" -ForegroundColor Yellow
    exit 0
}

$FILE_COUNT = ($diffFiles | Measure-Object).Count
Write-Host "検出された差分ファイル数: $FILE_COUNT 個" -ForegroundColor Green
Write-Host ""
Write-Host "ファイルをコピー中..." -ForegroundColor Cyan
Write-Host ""
#endregion

#region ファイルコピー
$COPY_COUNT = 0
$ERROR_COUNT = 0
$SKIP_COUNT = 0

foreach ($file in $diffFiles) {
    # Unixスタイルのパスをバックスラッシュに変換
    $filePath = $file -replace '/', '\'

    # フルパス
    $sourceFile = Join-Path $REPO_ROOT $filePath
    $destFile = Join-Path $OUTPUT_DIR $filePath

    # ファイルの存在確認（削除されたファイルはスキップ）
    if (Test-Path $sourceFile) {
        # コピー先のディレクトリを作成
        $destDir = Split-Path -Path $destFile -Parent
        if (-not (Test-Path $destDir)) {
            New-Item -ItemType Directory -Path $destDir -Force | Out-Null
        }

        # ファイルをコピー
        try {
            Copy-Item -Path $sourceFile -Destination $destFile -Force -ErrorAction Stop
            Write-Host "[コピー] $filePath" -ForegroundColor Green
            $COPY_COUNT++
        } catch {
            Write-Host "[エラー] $filePath" -ForegroundColor Red
            $ERROR_COUNT++
        }
    } else {
        Write-Host "[削除済] $filePath (スキップ)" -ForegroundColor Gray
        $SKIP_COUNT++
    }
}
#endregion

#region 結果表示
Write-Host ""
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host " 処理完了" -ForegroundColor Cyan
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "コピーしたファイル数: $COPY_COUNT 個" -ForegroundColor Green

if ($SKIP_COUNT -gt 0) {
    Write-Host "スキップ          : $SKIP_COUNT 個" -ForegroundColor Gray
}

if ($ERROR_COUNT -gt 0) {
    Write-Host "エラー            : $ERROR_COUNT 個" -ForegroundColor Red
}

Write-Host "出力先: $OUTPUT_DIR" -ForegroundColor White
Write-Host ""

# 出力先フォルダを開く
$openFolder = Read-Host "出力先フォルダを開きますか? (y/n)"
if ($openFolder -eq "y") {
    explorer $OUTPUT_DIR
}
#endregion

exit 0
