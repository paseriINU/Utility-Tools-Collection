@echo off
rem ====================================================================
rem ローカルブランチ削除ツール（シンプル版）
rem ====================================================================

setlocal enabledelayedexpansion

cls
echo ========================================
echo ローカルブランチ削除ツール
echo ========================================
echo.

rem Gitリポジトリのパスに移動
set GIT_PROJECT_PATH=C:\Users\%username%\source\Git\project

if not exist "%GIT_PROJECT_PATH%" (
    echo [エラー] Gitプロジェクトフォルダが見つかりません。
    echo パス: %GIT_PROJECT_PATH%
    pause
    exit /b 1
)

echo Gitプロジェクトパス: %GIT_PROJECT_PATH%
cd /d "%GIT_PROJECT_PATH%"
echo.

rem Gitリポジトリ確認
git rev-parse --git-dir >nul 2>&1
if errorlevel 1 (
    echo [エラー] このフォルダはGitリポジトリではありません。
    pause
    exit /b 1
)

rem 現在のブランチを取得
for /f "delims=" %%b in ('git branch --show-current') do set CURRENT_BRANCH=%%b

echo 現在のブランチ: %CURRENT_BRANCH%
echo.
echo ローカルブランチ一覧:
echo.

rem ブランチ一覧を表示（現在のブランチ以外）
set INDEX=0
set TEMP_FILE=%TEMP%\git_local_branches_%RANDOM%.txt
git branch > "%TEMP_FILE%"

for /f "usebackq delims=" %%b in ("%TEMP_FILE%") do (
    set "LINE=%%b"
    set "BRANCH_NAME=!LINE:* =!"
    set "BRANCH_NAME=!BRANCH_NAME:  =!"
    set "BRANCH_NAME=!BRANCH_NAME: =!"

    if not "!BRANCH_NAME!"=="%CURRENT_BRANCH%" (
        set /a INDEX+=1
        set "BRANCH[!INDEX!]=!BRANCH_NAME!"
        echo [!INDEX!] !BRANCH_NAME!
        set BRANCH_COUNT=!INDEX!
    )
)

del "%TEMP_FILE%"

if %INDEX% EQU 0 (
    echo 削除可能なブランチがありません。
    pause
    exit /b 0
)

echo.
echo [0] キャンセル
echo.
set /p BRANCH_NUM="削除するブランチ番号 (1-%BRANCH_COUNT%, 0=キャンセル): "

if "%BRANCH_NUM%"=="0" exit /b 0
if %BRANCH_NUM% LSS 1 goto INVALID
if %BRANCH_NUM% GTR %BRANCH_COUNT% goto INVALID

rem 選択されたブランチ
set "SELECTED=!BRANCH[%BRANCH_NUM%]!"

echo.
echo ブランチ: !SELECTED!
echo.

rem 保護ブランチチェック
if "!SELECTED!"=="main" goto PROTECTED
if "!SELECTED!"=="master" goto PROTECTED
if "!SELECTED!"=="develop" goto PROTECTED

echo [1] 通常の削除 (マージ済みのみ)
echo [2] 強制削除 (マージされていなくても削除)
echo [0] キャンセル
echo.
set /p DELETE_MODE="選択 (1-2, 0=キャンセル): "

if "%DELETE_MODE%"=="0" exit /b 0
if "%DELETE_MODE%"=="1" goto NORMAL
if "%DELETE_MODE%"=="2" goto FORCE

echo 無効な選択です。
pause
exit /b 1

:NORMAL
choice /M "削除しますか"
if errorlevel 2 exit /b 0

echo.
echo 削除中...
git branch -d !SELECTED!

if errorlevel 1 (
    echo [エラー] 削除に失敗しました。
    echo マージされていない可能性があります。
    pause
    exit /b 1
)

echo.
echo 削除しました: !SELECTED!
pause
exit /b 0

:FORCE
echo.
echo [警告] 強制削除します。マージされていない変更は失われます。
echo.
choice /M "本当に削除しますか"
if errorlevel 2 exit /b 0

echo.
echo 強制削除中...
git branch -D !SELECTED!

if errorlevel 1 (
    echo [エラー] 削除に失敗しました。
    pause
    exit /b 1
)

echo.
echo 強制削除しました: !SELECTED!
pause
exit /b 0

:INVALID
echo 無効な番号です。
pause
exit /b 1

:PROTECTED
echo.
echo [警告] main/master/develop は保護されています。
pause
exit /b 1

endlocal
