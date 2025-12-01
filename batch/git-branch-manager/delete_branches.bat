@echo off
rem ====================================================================
rem Gitブランチ削除ツール（対話型）
rem リモートブランチとローカルブランチを数字で選択して削除
rem ====================================================================

setlocal enabledelayedexpansion

rem ====================================================================
rem メイン処理
rem ====================================================================

:MAIN_MENU
cls
echo ========================================
echo Gitブランチ削除ツール
echo ========================================
echo.

rem Gitリポジトリのパスに移動（初回のみ）
if not defined GIT_PROJECT_INITIALIZED (
    set GIT_PROJECT_PATH=C:\Users\%username%\source\Git\project

    if not exist "!GIT_PROJECT_PATH!" (
        echo [エラー] Gitプロジェクトフォルダが見つかりません。
        echo パス: !GIT_PROJECT_PATH!
        pause
        exit /b 1
    )

    echo Gitプロジェクトパス: !GIT_PROJECT_PATH!
    cd /d "!GIT_PROJECT_PATH!"
    echo.

    rem Gitリポジトリかどうか確認
    git rev-parse --git-dir >nul 2>&1
    if errorlevel 1 (
        echo [エラー] このフォルダはGitリポジトリではありません。
        pause
        exit /b 1
    )

    set GIT_PROJECT_INITIALIZED=1
)

echo [1] リモートブランチを削除
echo [2] ローカルブランチを削除
echo [3] リモート＆ローカル両方を削除
echo [4] 終了
echo.
set /p CHOICE="選択してください (1-4): "

if "%CHOICE%"=="1" goto DELETE_REMOTE
if "%CHOICE%"=="2" goto DELETE_LOCAL
if "%CHOICE%"=="3" goto DELETE_BOTH
if "%CHOICE%"=="4" goto END
echo 無効な選択です。
pause
goto MAIN_MENU

rem ====================================================================
rem リモートブランチ削除
rem ====================================================================
:DELETE_REMOTE
cls
echo ========================================
echo リモートブランチ削除
echo ========================================
echo.

rem リモート名を取得
for /f "delims=" %%r in ('git remote') do set REMOTE_NAME=%%r

if not defined REMOTE_NAME (
    echo [エラー] リモートリポジトリが設定されていません。
    pause
    goto MAIN_MENU
)

echo リモート名: %REMOTE_NAME%
echo.
echo リモートブランチ一覧を取得中...
echo.

rem リモートブランチ一覧を取得（origin/HEAD を除外）
set INDEX=0
set BRANCH_COUNT=0

rem 一時ファイルにブランチ一覧を保存
set TEMP_FILE=%TEMP%\git_remote_branches_%RANDOM%.txt
git branch -r | findstr /v "HEAD" > "%TEMP_FILE%"

rem ブランチをカウントして表示
for /f "usebackq delims=" %%b in ("%TEMP_FILE%") do (
    set /a INDEX+=1
    set "BRANCH[!INDEX!]=%%b"

    rem ブランチ名から空白を削除
    set "DISPLAY_BRANCH=%%b"
    set "DISPLAY_BRANCH=!DISPLAY_BRANCH:  =!"
    set "DISPLAY_BRANCH=!DISPLAY_BRANCH: =!"

    echo [!INDEX!] !DISPLAY_BRANCH!
    set /a BRANCH_COUNT+=1
)

del "%TEMP_FILE%"

if %BRANCH_COUNT% EQU 0 (
    echo リモートブランチが見つかりません。
    pause
    goto MAIN_MENU
)

echo.
echo [0] キャンセル
echo.
set /p BRANCH_NUM="削除するブランチ番号を入力 (1-%BRANCH_COUNT%, 0=キャンセル): "

if "%BRANCH_NUM%"=="0" goto MAIN_MENU

rem 入力チェック
if %BRANCH_NUM% LSS 1 goto INVALID_REMOTE
if %BRANCH_NUM% GTR %BRANCH_COUNT% goto INVALID_REMOTE

rem 選択されたブランチを取得
set "SELECTED_BRANCH=!BRANCH[%BRANCH_NUM%]!"
rem 空白とリモート名を削除
set "SELECTED_BRANCH=!SELECTED_BRANCH:  =!"
set "SELECTED_BRANCH=!SELECTED_BRANCH: =!"
set "SELECTED_BRANCH=!SELECTED_BRANCH:%REMOTE_NAME%/=!"

echo.
echo ========================================
echo 選択されたブランチ: %REMOTE_NAME%/!SELECTED_BRANCH!
echo ========================================
echo.

rem main/master/develop の保護
if "!SELECTED_BRANCH!"=="main" goto PROTECTED_BRANCH
if "!SELECTED_BRANCH!"=="master" goto PROTECTED_BRANCH
if "!SELECTED_BRANCH!"=="develop" goto PROTECTED_BRANCH

choice /M "このリモートブランチを削除しますか"
if errorlevel 2 goto MAIN_MENU

echo.
echo リモートブランチを削除中...
git push %REMOTE_NAME% --delete !SELECTED_BRANCH!

if errorlevel 1 (
    echo.
    echo [エラー] リモートブランチの削除に失敗しました。
    pause
    goto MAIN_MENU
)

echo.
echo リモートブランチを削除しました: %REMOTE_NAME%/!SELECTED_BRANCH!
pause
goto MAIN_MENU

:INVALID_REMOTE
echo 無効な番号です。
pause
goto DELETE_REMOTE

:PROTECTED_BRANCH
echo.
echo [警告] main/master/develop ブランチは保護されています。
echo このツールでは削除できません。
pause
goto MAIN_MENU

rem ====================================================================
rem ローカルブランチ削除
rem ====================================================================
:DELETE_LOCAL
cls
echo ========================================
echo ローカルブランチ削除
echo ========================================
echo.

rem 現在のブランチを取得
for /f "delims=" %%b in ('git branch --show-current') do set CURRENT_BRANCH=%%b

echo 現在のブランチ: %CURRENT_BRANCH%
echo.
echo ローカルブランチ一覧:
echo.

rem ローカルブランチ一覧を取得
set INDEX=0
set BRANCH_COUNT=0

rem 一時ファイルにブランチ一覧を保存
set TEMP_FILE=%TEMP%\git_local_branches_%RANDOM%.txt
git branch > "%TEMP_FILE%"

rem ブランチをカウントして表示
for /f "usebackq delims=" %%b in ("%TEMP_FILE%") do (
    set "BRANCH_LINE=%%b"

    rem * を削除してブランチ名を取得
    set "BRANCH_NAME=!BRANCH_LINE:* =!"
    set "BRANCH_NAME=!BRANCH_NAME:  =!"
    set "BRANCH_NAME=!BRANCH_NAME: =!"

    rem 現在のブランチでない場合のみ表示
    if not "!BRANCH_NAME!"=="%CURRENT_BRANCH%" (
        set /a INDEX+=1
        set "BRANCH[!INDEX!]=!BRANCH_NAME!"
        echo [!INDEX!] !BRANCH_NAME!
        set /a BRANCH_COUNT+=1
    )
)

del "%TEMP_FILE%"

if %BRANCH_COUNT% EQU 0 (
    echo 削除可能なローカルブランチがありません。
    echo （現在のブランチ以外のブランチがありません）
    pause
    goto MAIN_MENU
)

echo.
echo [0] キャンセル
echo.
set /p BRANCH_NUM="削除するブランチ番号を入力 (1-%BRANCH_COUNT%, 0=キャンセル): "

if "%BRANCH_NUM%"=="0" goto MAIN_MENU

rem 入力チェック
if %BRANCH_NUM% LSS 1 goto INVALID_LOCAL
if %BRANCH_NUM% GTR %BRANCH_COUNT% goto INVALID_LOCAL

rem 選択されたブランチを取得
set "SELECTED_BRANCH=!BRANCH[%BRANCH_NUM%]!"

echo.
echo ========================================
echo 選択されたブランチ: !SELECTED_BRANCH!
echo ========================================
echo.

rem main/master/develop の保護
if "!SELECTED_BRANCH!"=="main" goto PROTECTED_BRANCH
if "!SELECTED_BRANCH!"=="master" goto PROTECTED_BRANCH
if "!SELECTED_BRANCH!"=="develop" goto PROTECTED_BRANCH

rem マージ確認
echo このブランチの削除方法を選択してください：
echo [1] 通常の削除 (マージ済みブランチのみ)
echo [2] 強制削除 (マージされていなくても削除)
echo [0] キャンセル
echo.
set /p DELETE_MODE="選択 (1-2, 0=キャンセル): "

if "%DELETE_MODE%"=="0" goto MAIN_MENU
if "%DELETE_MODE%"=="1" goto DELETE_LOCAL_NORMAL
if "%DELETE_MODE%"=="2" goto DELETE_LOCAL_FORCE

echo 無効な選択です。
pause
goto DELETE_LOCAL

:DELETE_LOCAL_NORMAL
choice /M "ローカルブランチを削除しますか"
if errorlevel 2 goto MAIN_MENU

echo.
echo ローカルブランチを削除中...
git branch -d !SELECTED_BRANCH!

if errorlevel 1 (
    echo.
    echo [エラー] ローカルブランチの削除に失敗しました。
    echo このブランチはマージされていない可能性があります。
    echo 強制削除する場合は、メニューから「強制削除」を選択してください。
    pause
    goto MAIN_MENU
)

echo.
echo ローカルブランチを削除しました: !SELECTED_BRANCH!
pause
goto MAIN_MENU

:DELETE_LOCAL_FORCE
echo.
echo [警告] 強制削除を選択しています。
echo マージされていない変更は失われます。
echo.
choice /M "本当に強制削除しますか"
if errorlevel 2 goto MAIN_MENU

echo.
echo ローカルブランチを強制削除中...
git branch -D !SELECTED_BRANCH!

if errorlevel 1 (
    echo.
    echo [エラー] ローカルブランチの削除に失敗しました。
    pause
    goto MAIN_MENU
)

echo.
echo ローカルブランチを強制削除しました: !SELECTED_BRANCH!
pause
goto MAIN_MENU

:INVALID_LOCAL
echo 無効な番号です。
pause
goto DELETE_LOCAL

rem ====================================================================
rem リモート＆ローカル両方削除
rem ====================================================================
:DELETE_BOTH
cls
echo ========================================
echo リモート＆ローカルブランチ両方削除
echo ========================================
echo.

rem リモート名を取得
for /f "delims=" %%r in ('git remote') do set REMOTE_NAME=%%r

if not defined REMOTE_NAME (
    echo [エラー] リモートリポジトリが設定されていません。
    pause
    goto MAIN_MENU
)

rem 現在のブランチを取得
for /f "delims=" %%b in ('git branch --show-current') do set CURRENT_BRANCH=%%b

echo リモート名: %REMOTE_NAME%
echo 現在のブランチ: %CURRENT_BRANCH%
echo.
echo 共通するブランチを検索中...
echo.

rem ローカルブランチとリモートブランチの共通ブランチを検索
set INDEX=0
set BRANCH_COUNT=0

rem 一時ファイル
set TEMP_LOCAL=%TEMP%\git_local_%RANDOM%.txt
set TEMP_REMOTE=%TEMP%\git_remote_%RANDOM%.txt

rem ローカルブランチ一覧
git branch | sed "s/\* //" | sed "s/  //" > "%TEMP_LOCAL%"

rem リモートブランチ一覧（origin/ を除去）
git branch -r | findstr /v "HEAD" | sed "s/  %REMOTE_NAME%\///" | sed "s/  //" > "%TEMP_REMOTE%"

rem 共通ブランチを検索
for /f "usebackq delims=" %%l in ("%TEMP_LOCAL%") do (
    set "LOCAL_BR=%%l"

    rem 現在のブランチと保護ブランチを除外
    if not "!LOCAL_BR!"=="%CURRENT_BRANCH%" (
        if not "!LOCAL_BR!"=="main" (
            if not "!LOCAL_BR!"=="master" (
                if not "!LOCAL_BR!"=="develop" (
                    rem リモートに同名ブランチが存在するか確認
                    findstr /x "!LOCAL_BR!" "%TEMP_REMOTE%" >nul
                    if !errorlevel! EQU 0 (
                        set /a INDEX+=1
                        set "BRANCH[!INDEX!]=!LOCAL_BR!"
                        echo [!INDEX!] !LOCAL_BR! ^(ローカル＆リモート^)
                        set /a BRANCH_COUNT+=1
                    )
                )
            )
        )
    )
)

del "%TEMP_LOCAL%" "%TEMP_REMOTE%"

if %BRANCH_COUNT% EQU 0 (
    echo 削除可能な共通ブランチがありません。
    pause
    goto MAIN_MENU
)

echo.
echo [0] キャンセル
echo.
set /p BRANCH_NUM="削除するブランチ番号を入力 (1-%BRANCH_COUNT%, 0=キャンセル): "

if "%BRANCH_NUM%"=="0" goto MAIN_MENU

rem 入力チェック
if %BRANCH_NUM% LSS 1 goto INVALID_BOTH
if %BRANCH_NUM% GTR %BRANCH_COUNT% goto INVALID_BOTH

rem 選択されたブランチを取得
set "SELECTED_BRANCH=!BRANCH[%BRANCH_NUM%]!"

echo.
echo ========================================
echo 選択されたブランチ: !SELECTED_BRANCH!
echo ========================================
echo リモート: %REMOTE_NAME%/!SELECTED_BRANCH!
echo ローカル: !SELECTED_BRANCH!
echo.

echo このブランチの削除方法を選択してください：
echo [1] 通常の削除 (ローカルはマージ済みのみ)
echo [2] 強制削除 (ローカルはマージされていなくても削除)
echo [0] キャンセル
echo.
set /p DELETE_MODE="選択 (1-2, 0=キャンセル): "

if "%DELETE_MODE%"=="0" goto MAIN_MENU
if "%DELETE_MODE%"=="1" goto DELETE_BOTH_NORMAL
if "%DELETE_MODE%"=="2" goto DELETE_BOTH_FORCE

echo 無効な選択です。
pause
goto DELETE_BOTH

:DELETE_BOTH_NORMAL
choice /M "リモート＆ローカルブランチを削除しますか"
if errorlevel 2 goto MAIN_MENU

echo.
echo リモートブランチを削除中...
git push %REMOTE_NAME% --delete !SELECTED_BRANCH!

if errorlevel 1 (
    echo [エラー] リモートブランチの削除に失敗しました。
    pause
    goto MAIN_MENU
)

echo ローカルブランチを削除中...
git branch -d !SELECTED_BRANCH!

if errorlevel 1 (
    echo [エラー] ローカルブランチの削除に失敗しました。
    echo リモートブランチは削除されましたが、ローカルブランチはマージされていない可能性があります。
    pause
    goto MAIN_MENU
)

echo.
echo リモート＆ローカルブランチを削除しました: !SELECTED_BRANCH!
pause
goto MAIN_MENU

:DELETE_BOTH_FORCE
echo.
echo [警告] ローカルブランチを強制削除します。
echo マージされていない変更は失われます。
echo.
choice /M "本当にリモート＆ローカルブランチを削除しますか"
if errorlevel 2 goto MAIN_MENU

echo.
echo リモートブランチを削除中...
git push %REMOTE_NAME% --delete !SELECTED_BRANCH!

if errorlevel 1 (
    echo [エラー] リモートブランチの削除に失敗しました。
    pause
    goto MAIN_MENU
)

echo ローカルブランチを強制削除中...
git branch -D !SELECTED_BRANCH!

if errorlevel 1 (
    echo [エラー] ローカルブランチの削除に失敗しました。
    pause
    goto MAIN_MENU
)

echo.
echo リモート＆ローカルブランチを削除しました: !SELECTED_BRANCH!
pause
goto MAIN_MENU

:INVALID_BOTH
echo 無効な番号です。
pause
goto DELETE_BOTH

:END
endlocal
exit /b 0
