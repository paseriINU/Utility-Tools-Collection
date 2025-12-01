@echo off
rem ====================================================================
rem リモートブランチ削除ツール（シンプル版）
rem ====================================================================

setlocal enabledelayedexpansion

cls
echo ========================================
echo リモートブランチ削除ツール
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

rem リモート名を取得
for /f "delims=" %%r in ('git remote') do set REMOTE_NAME=%%r

if not defined REMOTE_NAME (
    echo [エラー] リモートリポジトリが設定されていません。
    pause
    exit /b 1
)

echo リモート名: %REMOTE_NAME%
echo.
echo リモートブランチ一覧:
echo.

rem ブランチ一覧を表示
set INDEX=0
set TEMP_FILE=%TEMP%\git_remote_branches_%RANDOM%.txt
git branch -r | findstr /v "HEAD" > "%TEMP_FILE%"

for /f "usebackq delims=" %%b in ("%TEMP_FILE%") do (
    set /a INDEX+=1
    set "BRANCH[!INDEX!]=%%b"
    set "DISPLAY=%%b"
    set "DISPLAY=!DISPLAY:  =!"
    set "DISPLAY=!DISPLAY: =!"
    echo [!INDEX!] !DISPLAY!
    set BRANCH_COUNT=!INDEX!
)

del "%TEMP_FILE%"

if %INDEX% EQU 0 (
    echo リモートブランチが見つかりません。
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
set "SELECTED=!SELECTED:  =!"
set "SELECTED=!SELECTED: =!"
set "SELECTED=!SELECTED:%REMOTE_NAME%/=!"

echo.
echo ブランチ: %REMOTE_NAME%/!SELECTED!
echo.

rem 保護ブランチチェック
if "!SELECTED!"=="main" goto PROTECTED
if "!SELECTED!"=="master" goto PROTECTED
if "!SELECTED!"=="develop" goto PROTECTED

choice /M "削除しますか"
if errorlevel 2 exit /b 0

echo.
echo 削除中...
git push %REMOTE_NAME% --delete !SELECTED!

if errorlevel 1 (
    echo [エラー] 削除に失敗しました。
    pause
    exit /b 1
)

echo.
echo 削除しました: %REMOTE_NAME%/!SELECTED!
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
