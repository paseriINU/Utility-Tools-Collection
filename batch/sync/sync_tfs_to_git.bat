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
set "GIT_DIR=C:\Users\%username%\source\Git\project"

:: パスの存在確認
if not exist "%TFS_DIR%" (
    echo [エラー] TFSディレクトリが見つかりません: %TFS_DIR%
    pause
    exit /b 1
)

if not exist "%GIT_DIR%" (
    echo [エラー] Gitディレクトリが見つかりません: %GIT_DIR%
    pause
    exit /b 1
)

if not exist "%GIT_DIR%\.git" (
    echo [エラー] 指定されたディレクトリはGitリポジトリではありません: %GIT_DIR%
    pause
    exit /b 1
)

echo.
echo TFSディレクトリ: %TFS_DIR%
echo Gitディレクトリ: %GIT_DIR%
echo.

:: Gitディレクトリに移動
cd /d "%GIT_DIR%"

:: 現在のブランチを表示
echo ----------------------------------------
echo 現在のGitブランチ:
echo ----------------------------------------
git branch
echo.

:BRANCH_MENU
echo ブランチ操作を選択してください:
echo 1. このまま続行
echo 2. ブランチを切り替える
echo 3. 終了
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

:: PowerShellスクリプトを実行
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
"$tfsDir = '%TFS_DIR%'; ^
$gitDir = '%GIT_DIR%'; ^
Write-Host '差分チェック中...' -ForegroundColor Cyan; ^
Write-Host ''; ^
$tfsFiles = Get-ChildItem -Path $tfsDir -Recurse -File; ^
$gitFiles = Get-ChildItem -Path $gitDir -Recurse -File | Where-Object { $_.FullName -notlike '*\.git\*' }; ^
$tfsFileDict = @{}; ^
foreach ($file in $tfsFiles) { ^
    $relativePath = $file.FullName.Substring($tfsDir.Length).TrimStart('\'); ^
    $tfsFileDict[$relativePath] = $file; ^
}; ^
$gitFileDict = @{}; ^
foreach ($file in $gitFiles) { ^
    $relativePath = $file.FullName.Substring($gitDir.Length).TrimStart('\'); ^
    $gitFileDict[$relativePath] = $file; ^
}; ^
$copiedCount = 0; ^
$deletedCount = 0; ^
$identicalCount = 0; ^
Write-Host '=== ファイル差分チェック ===' -ForegroundColor Yellow; ^
Write-Host ''; ^
foreach ($relativePath in $tfsFileDict.Keys) { ^
    $tfsFile = $tfsFileDict[$relativePath]; ^
    $gitFilePath = Join-Path $gitDir $relativePath; ^
    if (Test-Path $gitFilePath) { ^
        $tfsHash = (Get-FileHash -Path $tfsFile.FullName -Algorithm MD5).Hash; ^
        $gitHash = (Get-FileHash -Path $gitFilePath -Algorithm MD5).Hash; ^
        if ($tfsHash -ne $gitHash) { ^
            Write-Host '[更新] ' -ForegroundColor Yellow -NoNewline; ^
            Write-Host $relativePath; ^
            $targetDir = Split-Path -Path $gitFilePath -Parent; ^
            if (-not (Test-Path $targetDir)) { ^
                New-Item -ItemType Directory -Path $targetDir -Force | Out-Null; ^
            }; ^
            Copy-Item -Path $tfsFile.FullName -Destination $gitFilePath -Force; ^
            $copiedCount++; ^
        } else { ^
            $identicalCount++; ^
        } ^
    } else { ^
        Write-Host '[新規] ' -ForegroundColor Green -NoNewline; ^
        Write-Host $relativePath; ^
        $targetDir = Split-Path -Path $gitFilePath -Parent; ^
        if (-not (Test-Path $targetDir)) { ^
            New-Item -ItemType Directory -Path $targetDir -Force | Out-Null; ^
        }; ^
        Copy-Item -Path $tfsFile.FullName -Destination $gitFilePath -Force; ^
        $copiedCount++; ^
    } ^
}; ^
Write-Host ''; ^
Write-Host '=== Gitのみに存在するファイル (削除対象) ===' -ForegroundColor Yellow; ^
Write-Host ''; ^
foreach ($relativePath in $gitFileDict.Keys) { ^
    if (-not $tfsFileDict.ContainsKey($relativePath)) { ^
        $gitFile = $gitFileDict[$relativePath]; ^
        Write-Host '[削除] ' -ForegroundColor Red -NoNewline; ^
        Write-Host $relativePath; ^
        Remove-Item -Path $gitFile.FullName -Force; ^
        $deletedCount++; ^
    } ^
}; ^
Write-Host ''; ^
Write-Host '========================================' -ForegroundColor Cyan; ^
Write-Host '同期完了' -ForegroundColor Cyan; ^
Write-Host '========================================' -ForegroundColor Cyan; ^
Write-Host ''; ^
Write-Host \"更新/新規ファイル: $copiedCount\" -ForegroundColor Green; ^
Write-Host \"削除ファイル: $deletedCount\" -ForegroundColor Red; ^
Write-Host \"変更なし: $identicalCount\" -ForegroundColor Gray; ^
Write-Host ''"

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
