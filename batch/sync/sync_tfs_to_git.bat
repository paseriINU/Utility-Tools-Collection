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

:: PowerShellスクリプトを実行
set "SCRIPT_DIR=%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%sync_logic.ps1" -TfsDir "%TFS_DIR%" -GitDir "%GIT_DIR%"

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
