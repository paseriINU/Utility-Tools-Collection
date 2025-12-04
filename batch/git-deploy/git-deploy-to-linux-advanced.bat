@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul

REM ==============================================================================
REM Git Deploy to Linux - Advanced Version
REM Git変更ファイルをLinuxサーバーに転送（環境選択・拡張子フィルタ対応）
REM ==============================================================================
REM
REM 機能:
REM   1. 複数環境から転送先を選択
REM   2. 変更ファイルのみ OR すべてのファイルを選択
REM   3. 拡張子フィルタ (.c .pc .h)
REM   4. 削除されたファイルは自動除外
REM   5. 全部転送 or 個別選択
REM   6. Linux側でディレクトリ自動作成・パーミッション設定
REM
REM 必要な環境:
REM   - Git がインストールされていること
REM   - SCP/SSH コマンドが利用可能であること
REM   - SSH公開鍵認証が設定されていること（推奨）
REM
REM ==============================================================================

REM ============================================================
REM 設定 - ここを編集してください
REM ============================================================

REM SSH接続情報
set SSH_USER=youruser
set SSH_HOST=linux-server

REM Gitリポジトリのローカルパス
set LOCAL_REPO=C:\path\to\local\repo

REM 共通グループ
set COMMON_GROUP=common_group

REM 転送対象の拡張子（スペース区切り）
set TARGET_EXTENSIONS=.c .pc .h

REM 環境設定（追加・変更可能）
REM 環境名|転送先パス|オーナー の形式
set ENV_COUNT=3
set ENV_1=tst1t|/path/to/tst1t/|tzy_tst13
set ENV_2=tst2t|/path/to/tst2t/|tzy_tst23
set ENV_3=tst3t|/path/to/tst3t/|tzy_tst33

REM ============================================================

echo.
echo ================================================================
echo   Git Deploy to Linux - Advanced Version
echo ================================================================
echo.

REM Gitリポジトリチェック
cd /d "%LOCAL_REPO%"
if not exist ".git" (
    echo [エラー] Gitリポジトリではありません: %LOCAL_REPO%
    pause
    exit /b 1
)

echo [情報] Gitリポジトリ: %LOCAL_REPO%
echo.

REM ============================================================
REM 環境選択
REM ============================================================
echo ================================================================
echo 転送先環境を選択してください
echo ================================================================
echo.

for /l %%i in (1,1,%ENV_COUNT%) do (
    set ENV_LINE=!ENV_%%i!
    for /f "tokens=1 delims=|" %%a in ("!ENV_LINE!") do (
        echo %%i. %%a
    )
)

echo.
set /p ENV_CHOICE="番号を入力 (1-%ENV_COUNT%): "

REM 入力チェック
set VALID_CHOICE=0
for /l %%i in (1,1,%ENV_COUNT%) do (
    if "%ENV_CHOICE%"=="%%i" set VALID_CHOICE=1
)

if "%VALID_CHOICE%"=="0" (
    echo [エラー] 無効な選択です
    pause
    exit /b 1
)

REM 選択された環境の情報を取得
set ENV_LINE=!ENV_%ENV_CHOICE%!
for /f "tokens=1,2,3 delims=|" %%a in ("!ENV_LINE!") do (
    set ENV_NAME=%%a
    set DEST_DIR=%%b
    set OWNER=%%c
)

echo.
echo [選択] 環境: %ENV_NAME%
echo [情報] 転送先: %SSH_USER%@%SSH_HOST%:%DEST_DIR%
echo [情報] オーナー: %OWNER%:%COMMON_GROUP%
echo.

REM ============================================================
REM 転送モード選択
REM ============================================================
echo ================================================================
echo 転送するファイルを選択
echo ================================================================
echo.
echo 1. 変更されたファイルのみ (git status)
echo 2. すべてのファイル
echo.
set /p MODE="番号を入力 (1-2): "

if "%MODE%"=="1" (
    set MODE_NAME=変更ファイルのみ
) else if "%MODE%"=="2" (
    set MODE_NAME=すべてのファイル
) else (
    echo [エラー] 無効な選択です
    pause
    exit /b 1
)

echo.
echo [選択] モード: %MODE_NAME%
echo [情報] 対象拡張子: %TARGET_EXTENSIONS%
echo.
pause

REM ============================================================
REM ファイルリスト取得（メモリに保存）
REM ============================================================
echo.
echo [実行] ファイルリストを取得中...
echo.

REM 一時ファイルを使用してファイルリストを保存
set TEMP_FILE_LIST=%TEMP%\git_deploy_files_%RANDOM%.txt
if exist "%TEMP_FILE_LIST%" del "%TEMP_FILE_LIST%"

set FILE_COUNT=0

if "%MODE%"=="1" (
    REM 変更されたファイルのみ
    echo [情報] Git status から変更ファイルを取得中...

    for /f "tokens=1,2*" %%a in ('git status --short 2^>nul') do (
        set STATUS=%%a
        set FILEPATH=%%b

        REM 削除されたファイル (D で始まる) を除外
        echo !STATUS! | findstr /r "^D" >nul
        if errorlevel 1 (
            REM 拡張子チェック
            set FILE_EXT=%%~xb
            set IS_TARGET=0

            for %%e in (%TARGET_EXTENSIONS%) do (
                if /i "!FILE_EXT!"=="%%e" set IS_TARGET=1
            )

            if "!IS_TARGET!"=="1" (
                echo %%b>> "%TEMP_FILE_LIST%"
                set /a FILE_COUNT+=1
            )
        )
    )
) else (
    REM すべてのファイル
    echo [情報] リポジトリ内のすべての対象ファイルを取得中...

    for /r %%f in (*) do (
        set FULLPATH=%%f
        set FILE_EXT=%%~xf
        set IS_TARGET=0

        for %%e in (%TARGET_EXTENSIONS%) do (
            if /i "!FILE_EXT!"=="%%e" set IS_TARGET=1
        )

        if "!IS_TARGET!"=="1" (
            set RELPATH=!FULLPATH:%LOCAL_REPO%\=!
            echo !RELPATH!>> "%TEMP_FILE_LIST%"
            set /a FILE_COUNT+=1
        )
    )
)

if %FILE_COUNT%==0 (
    echo.
    echo [情報] 転送対象のファイルがありません
    if exist "%TEMP_FILE_LIST%" del "%TEMP_FILE_LIST%"
    pause
    exit /b 0
)

echo.
echo [成功] %FILE_COUNT% 個のファイルが見つかりました
echo.

REM ============================================================
REM ファイルリスト表示
REM ============================================================
echo ================================================================
echo 転送予定のファイル一覧
echo ================================================================
echo.

set INDEX=1
for /f "usebackq delims=" %%f in ("%TEMP_FILE_LIST%") do (
    echo !INDEX!. %%f
    set /a INDEX+=1
)

echo.

REM ============================================================
REM 転送確認
REM ============================================================
echo これらのファイルを転送しますか？
echo.
echo   [A] すべて転送
echo   [I] 個別に選択
echo   [C] キャンセル
echo.

:ASK_TRANSFER_MODE
set /p TRANSFER_MODE="選択してください (A/I/C): "
set TRANSFER_MODE=%TRANSFER_MODE:a=A%
set TRANSFER_MODE=%TRANSFER_MODE:i=I%
set TRANSFER_MODE=%TRANSFER_MODE:c=C%

if "%TRANSFER_MODE%"=="C" (
    echo.
    echo [キャンセル] 転送を中止しました
    if exist "%TEMP_FILE_LIST%" del "%TEMP_FILE_LIST%"
    pause
    exit /b 0
)

if "%TRANSFER_MODE%" neq "A" if "%TRANSFER_MODE%" neq "I" (
    echo [エラー] A, I, C のいずれかを入力してください
    goto ASK_TRANSFER_MODE
)

echo.

REM ============================================================
REM 個別選択
REM ============================================================
if "%TRANSFER_MODE%"=="I" (
    echo ================================================================
    echo 個別ファイル選択
    echo ================================================================
    echo.

    set TEMP_SELECTED=%TEMP%\git_deploy_selected_%RANDOM%.txt
    if exist "%TEMP_SELECTED%" del "%TEMP_SELECTED%"

    set SELECTED_COUNT=0

    for /f "usebackq delims=" %%f in ("%TEMP_FILE_LIST%") do (
        :ASK_FILE
        set /p FILE_CONFIRM="転送: %%f (y/n): "
        set FILE_CONFIRM=!FILE_CONFIRM:Y=y!
        set FILE_CONFIRM=!FILE_CONFIRM:N=n!

        if "!FILE_CONFIRM!"=="y" (
            echo %%f>> "%TEMP_SELECTED%"
            set /a SELECTED_COUNT+=1
        ) else if "!FILE_CONFIRM!"=="n" (
            REM スキップ
        ) else (
            echo [エラー] y または n を入力してください
            goto ASK_FILE
        )
    )

    echo.

    if !SELECTED_COUNT!==0 (
        echo [情報] 転送するファイルが選択されませんでした
        if exist "%TEMP_FILE_LIST%" del "%TEMP_FILE_LIST%"
        if exist "%TEMP_SELECTED%" del "%TEMP_SELECTED%"
        pause
        exit /b 0
    )

    echo [選択] !SELECTED_COUNT! 個のファイルを転送します
    echo.

    REM 選択されたファイルリストに置き換え
    del "%TEMP_FILE_LIST%"
    move "%TEMP_SELECTED%" "%TEMP_FILE_LIST%" >nul
    set FILE_COUNT=!SELECTED_COUNT!
) else (
    echo [選択] すべてのファイルを転送します
    echo.
)

pause

REM ============================================================
REM ファイル転送
REM ============================================================
echo.
echo ================================================================
echo ファイル転送開始
echo ================================================================
echo.

set SUCCESS_COUNT=0
set FAILED_COUNT=0
set TEMP_FAILED=%TEMP%\git_deploy_failed_%RANDOM%.txt
if exist "%TEMP_FAILED%" del "%TEMP_FAILED%"

for /f "usebackq delims=" %%f in ("%TEMP_FILE_LIST%") do (
    REM Windowsのパス区切り(\)をLinux形式(/)に変換
    set FILEPATH=%%f
    set LINUX_PATH=!FILEPATH:\=/!

    echo [転送] %%f

    REM Linux側で親ディレクトリを作成
    ssh %SSH_USER%@%SSH_HOST% "mkdir -p '%DEST_DIR%$(dirname '!LINUX_PATH!')' && chmod 777 '%DEST_DIR%$(dirname '!LINUX_PATH!')' && chown %OWNER%:%COMMON_GROUP% '%DEST_DIR%$(dirname '!LINUX_PATH!')'" 2>nul

    if errorlevel 1 (
        echo   [警告] ディレクトリ作成に失敗しました（既に存在する可能性があります）
    )

    REM ファイルを転送
    scp "%%f" %SSH_USER%@%SSH_HOST%:%DEST_DIR%!LINUX_PATH! 2>nul

    if errorlevel 1 (
        echo   [✗] 失敗
        echo %%f>> "%TEMP_FAILED%"
        set /a FAILED_COUNT+=1
    ) else (
        REM パーミッションと所有者を設定
        ssh %SSH_USER%@%SSH_HOST% "chmod 777 '%DEST_DIR%!LINUX_PATH!' && chown %OWNER%:%COMMON_GROUP% '%DEST_DIR%!LINUX_PATH!'" 2>nul

        if errorlevel 1 (
            echo   [✓] 転送成功（パーミッション設定失敗）
        ) else (
            echo   [✓] 成功
        )
        set /a SUCCESS_COUNT+=1
    )
    echo.
)

REM ============================================================
REM 結果サマリー
REM ============================================================
echo ================================================================
echo 転送結果
echo ================================================================
echo.
echo   成功: %SUCCESS_COUNT% ファイル
echo   失敗: %FAILED_COUNT% ファイル
echo.

if %FAILED_COUNT% gtr 0 (
    echo [失敗したファイル]
    for /f "usebackq delims=" %%f in ("%TEMP_FAILED%") do (
        echo   - %%f
    )
    echo.
)

REM 一時ファイル削除
if exist "%TEMP_FILE_LIST%" del "%TEMP_FILE_LIST%"
if exist "%TEMP_FAILED%" del "%TEMP_FAILED%"

if %FAILED_COUNT%==0 (
    echo すべてのファイル転送が完了しました！
) else (
    echo 一部のファイル転送に失敗しました
)

echo.
pause
