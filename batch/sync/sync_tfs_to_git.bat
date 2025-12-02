@echo off
:: UTF-8モードに設定（日本語ブランチ名対応）
chcp 65001 >nul
setlocal enabledelayedexpansion

:: =============================================================================
:: TFS to Git Sync Script
:: TFSのファイルをGitリポジトリに同期します
:: =============================================================================

:: Git日本語表示設定
git config --global core.quotepath false >nul 2>&1

echo ========================================
echo TFS to Git 同期スクリプト
echo ========================================
echo.

:: TFSとGitのディレクトリパスを設定（固定値）
set "TFS_DIR=C:\Users\%username%\source"
set "GIT_REPO_DIR=C:\Users\%username%\source\Git\project"

:: パスの存在確認
if not exist "%TFS_DIR%" (
    echo [エラー] TFSディレクトリが見つかりません: %TFS_DIR%
    pause
    exit /b 1
)

if not exist "%GIT_REPO_DIR%" (
    echo [エラー] Gitディレクトリが見つかりません: %GIT_REPO_DIR%
    pause
    exit /b 1
)

if not exist "%GIT_REPO_DIR%\.git" (
    echo [エラー] 指定されたディレクトリはGitリポジトリではありません: %GIT_REPO_DIR%
    pause
    exit /b 1
)

echo.
echo TFSディレクトリ: %TFS_DIR%
echo Gitディレクトリ: %GIT_REPO_DIR%
echo.

:: Gitディレクトリに移動
cd /d "%GIT_REPO_DIR%"

:: 現在のブランチを表示
echo ----------------------------------------
echo 現在のGitブランチ:
echo ----------------------------------------
git branch
echo.

:BRANCH_MENU
echo ブランチ操作を選択してください:
echo  1. このまま続行
echo  2. ブランチを切り替える
echo  3. 終了
echo.
set /p BRANCH_CHOICE="選択 (1-3): "

if "%BRANCH_CHOICE%"=="1" goto START_SYNC
if "%BRANCH_CHOICE%"=="2" goto SWITCH_BRANCH
if "%BRANCH_CHOICE%"=="3" exit /b 0

echo 無効な選択です。
goto BRANCH_MENU

:SWITCH_BRANCH
echo.
echo ----------------------------------------
echo 利用可能なブランチ:
echo ----------------------------------------
git branch -a
echo.
set /p NEW_BRANCH="切り替え先のブランチ名を入力してください: "
git checkout "%NEW_BRANCH%"
if errorlevel 1 (
    echo [エラー] ブランチの切り替えに失敗しました
    pause
    goto BRANCH_MENU
)
echo ブランチを切り替えました: %NEW_BRANCH%
echo.
goto BRANCH_MENU

:START_SYNC
echo.
echo ========================================
echo 同期処理を開始します
echo ========================================
echo.

:: PowerShellスクリプトを実行（ファイル末尾に埋め込まれたコード）
powershell -NoProfile -ExecutionPolicy Bypass -Command "$content = Get-Content '%~f0' -Raw; $start = $content.IndexOf('<#SYNC_LOGIC#>') + 14; $end = $content.IndexOf('<#/SYNC_LOGIC#>'); if ($start -gt 13 -and $end -gt $start) { $psCode = $content.Substring($start, $end - $start); Invoke-Expression $psCode } else { Write-Error 'PowerShell sync logic not found'; exit 1 }" -TfsDir "%TFS_DIR%" -GitDir "%GIT_REPO_DIR%"

if errorlevel 1 (
    echo [エラー] PowerShellスクリプトの実行に失敗しました
    pause
    exit /b 1
)

echo.
echo ========================================
echo Gitステータスを確認してください
echo ========================================
git status

echo.
echo ----------------------------------------
echo 次の操作を選択してください:
echo 1. 変更をコミットする
echo 2. 何もせず終了
echo ----------------------------------------
set /p COMMIT_CHOICE="選択 (1-2): "

if "%COMMIT_CHOICE%"=="1" goto DO_COMMIT
goto END

:DO_COMMIT
echo.
set /p COMMIT_MSG="コミットメッセージを入力してください: "
git add -A
git commit -m "%COMMIT_MSG%"
if errorlevel 1 (
    echo [警告] コミットに失敗しました、または変更がありませんでした
) else (
    echo コミットが完了しました
)

:END
echo.
echo 処理が完了しました。
pause
exit /b 0

rem =============================================================================
rem PowerShell Sync Logic (埋め込み版)
rem =============================================================================
<#SYNC_LOGIC#>
param(
    [Parameter(Mandatory=$true)]
    [string]$TfsDir,

    [Parameter(Mandatory=$true)]
    [string]$GitDir
)

# UTF-8出力設定
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host "差分チェック中..." -ForegroundColor Cyan
Write-Host ""

# TFSとGitのファイル一覧を取得
Write-Verbose "TFSディレクトリをスキャン中: $TfsDir"
$tfsFiles = Get-ChildItem -Path $TfsDir -Recurse -File -ErrorAction SilentlyContinue

Write-Verbose "Gitディレクトリをスキャン中: $GitDir"
$gitFiles = Get-ChildItem -Path $GitDir -Recurse -File -ErrorAction SilentlyContinue | Where-Object {
    $_.FullName -notlike '*\.git\*'
}

# ファイルを相対パスでハッシュテーブルに格納
$tfsFileDict = @{}
foreach ($file in $tfsFiles) {
    $relativePath = $file.FullName.Substring($TfsDir.Length).TrimStart('\')
    $tfsFileDict[$relativePath] = $file
}

$gitFileDict = @{}
foreach ($file in $gitFiles) {
    $relativePath = $file.FullName.Substring($GitDir.Length).TrimStart('\')
    $gitFileDict[$relativePath] = $file
}

# 統計カウンタ
$copiedCount = 0
$deletedCount = 0
$identicalCount = 0

Write-Host "=== ファイル差分チェック ===" -ForegroundColor Yellow
Write-Host ""

# TFSファイルをチェック（更新 & 新規追加）
foreach ($relativePath in $tfsFileDict.Keys) {
    $tfsFile = $tfsFileDict[$relativePath]
    $gitFilePath = Join-Path $GitDir $relativePath

    if (Test-Path $gitFilePath) {
        # ファイルが両方に存在 → MD5ハッシュで比較
        try {
            $tfsHash = (Get-FileHash -Path $tfsFile.FullName -Algorithm MD5).Hash
            $gitHash = (Get-FileHash -Path $gitFilePath -Algorithm MD5).Hash

            if ($tfsHash -ne $gitHash) {
                # ハッシュが異なる → 更新
                Write-Host "[更新] " -ForegroundColor Yellow -NoNewline
                Write-Host $relativePath

                $targetDir = Split-Path -Path $gitFilePath -Parent
                if (-not (Test-Path $targetDir)) {
                    New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
                }

                Copy-Item -Path $tfsFile.FullName -Destination $gitFilePath -Force
                $copiedCount++
            } else {
                # ハッシュが同じ → 変更なし
                $identicalCount++
            }
        } catch {
            Write-Warning "ファイルハッシュ取得エラー: $relativePath - $_"
        }
    } else {
        # Gitに存在しない → 新規追加
        Write-Host "[新規] " -ForegroundColor Green -NoNewline
        Write-Host $relativePath

        $targetDir = Split-Path -Path $gitFilePath -Parent
        if (-not (Test-Path $targetDir)) {
            New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
        }

        Copy-Item -Path $tfsFile.FullName -Destination $gitFilePath -Force
        $copiedCount++
    }
}

Write-Host ""
Write-Host "=== Gitのみに存在するファイル (削除対象) ===" -ForegroundColor Yellow
Write-Host ""

# Gitのみのファイルをチェック（削除）
foreach ($relativePath in $gitFileDict.Keys) {
    if (-not $tfsFileDict.ContainsKey($relativePath)) {
        $gitFile = $gitFileDict[$relativePath]
        Write-Host "[削除] " -ForegroundColor Red -NoNewline
        Write-Host $relativePath

        try {
            Remove-Item -Path $gitFile.FullName -Force
            $deletedCount++
        } catch {
            Write-Warning "ファイル削除エラー: $relativePath - $_"
        }
    }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "同期完了" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "更新/新規ファイル: $copiedCount" -ForegroundColor Green
Write-Host "削除ファイル: $deletedCount" -ForegroundColor Red
Write-Host "変更なし: $identicalCount" -ForegroundColor Gray
Write-Host ""

# 終了コード: 0=成功
exit 0
<#/SYNC_LOGIC#>
