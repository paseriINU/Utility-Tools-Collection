@echo off
rem ====================================================================
rem リモートWindowsサーバ上でバッチファイルをCMDから実行するスクリプト
rem （設定ファイル対応版）
rem ====================================================================

setlocal enabledelayedexpansion

rem ====================================================================
rem 設定ファイルの読み込み
rem ====================================================================

set CONFIG_FILE=%~dp0config.ini

if not exist "%CONFIG_FILE%" (
    echo [エラー] 設定ファイルが見つかりません: %CONFIG_FILE%
    echo.
    echo config.ini を作成してください。サンプル：
    echo.
    echo [Server]
    echo REMOTE_SERVER=192.168.1.100
    echo REMOTE_USER=Administrator
    echo REMOTE_BATCH_PATH=C:\Scripts\target_script.bat
    echo.
    pause
    exit /b 1
)

rem 設定ファイルから値を読み込む
for /f "usebackq tokens=1,* delims==" %%a in ("%CONFIG_FILE%") do (
    set "%%a=%%b"
)

rem 必須項目のチェック
if not defined REMOTE_SERVER (
    echo [エラー] REMOTE_SERVER が設定されていません
    exit /b 1
)
if not defined REMOTE_USER (
    echo [エラー] REMOTE_USER が設定されていません
    exit /b 1
)
if not defined REMOTE_BATCH_PATH (
    echo [エラー] REMOTE_BATCH_PATH が設定されていません
    exit /b 1
)

rem オプション項目のデフォルト値
if not defined AUTO_DELETE set AUTO_DELETE=1
if not defined TASK_NAME set TASK_NAME=RemoteExec_%RANDOM%

rem ====================================================================
rem メイン処理
rem ====================================================================

echo ========================================
echo リモートバッチ実行ツール（設定ファイル版）
echo ========================================
echo.
echo リモートサーバ: %REMOTE_SERVER%
echo 実行ユーザー  : %REMOTE_USER%
echo 実行ファイル  : %REMOTE_BATCH_PATH%
echo タスク名      : %TASK_NAME%
echo.

rem パスワード入力（設定ファイルにパスワードがない場合）
if not defined REMOTE_PASSWORD (
    echo リモートサーバのパスワードを入力してください：
    set /p REMOTE_PASSWORD=
    echo.
)

echo タスクを作成中...

rem リモートサーバにタスクを作成
schtasks /Create ^
    /S %REMOTE_SERVER% ^
    /U %REMOTE_USER% ^
    /P %REMOTE_PASSWORD% ^
    /TN %TASK_NAME% ^
    /TR "%REMOTE_BATCH_PATH%" ^
    /SC ONCE ^
    /ST 00:00 ^
    /RU SYSTEM ^
    /F

if errorlevel 1 (
    echo.
    echo [エラー] タスクの作成に失敗しました。
    goto :ERROR_EXIT
)

echo タスク作成成功
echo.
echo タスクを実行中...

rem タスクを即座に実行
schtasks /Run ^
    /S %REMOTE_SERVER% ^
    /U %REMOTE_USER% ^
    /P %REMOTE_PASSWORD% ^
    /TN %TASK_NAME%

if errorlevel 1 (
    echo.
    echo [エラー] タスクの実行に失敗しました。
    goto :CLEANUP
)

echo タスク実行開始
echo.
echo 実行状態を確認中（5秒待機）...
timeout /t 5 /nobreak >nul

rem タスクの状態を確認
schtasks /Query ^
    /S %REMOTE_SERVER% ^
    /U %REMOTE_USER% ^
    /P %REMOTE_PASSWORD% ^
    /TN %TASK_NAME% ^
    /FO LIST

echo.
echo 処理が完了しました。

:CLEANUP
if "%AUTO_DELETE%"=="1" (
    echo.
    echo タスクを削除中...
    timeout /t 3 /nobreak >nul

    schtasks /Delete ^
        /S %REMOTE_SERVER% ^
        /U %REMOTE_USER% ^
        /P %REMOTE_PASSWORD% ^
        /TN %TASK_NAME% ^
        /F >nul 2>&1

    if errorlevel 1 (
        echo [警告] タスクの削除に失敗しました。
    ) else (
        echo タスク削除完了
    )
)

goto :END

:ERROR_EXIT
echo.
echo 処理を中断しました。
exit /b 1

:END
endlocal
exit /b 0
