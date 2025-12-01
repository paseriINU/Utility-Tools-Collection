@echo off
rem ====================================================================
rem リモートWindowsサーバ上でバッチファイルをPowerShell Remotingで実行
rem （設定ファイル対応版）
rem ====================================================================

setlocal enabledelayedexpansion

rem ====================================================================
rem 設定ファイルの読み込み
rem ====================================================================

set CONFIG_FILE=%~dp0config_ps.ini

if not exist "%CONFIG_FILE%" (
    echo [エラー] 設定ファイルが見つかりません: %CONFIG_FILE%
    echo.
    echo config_ps.ini を作成してください。サンプル：
    echo.
    echo [Server]
    echo REMOTE_SERVER=192.168.1.100
    echo REMOTE_USER=Administrator
    echo REMOTE_BATCH_PATH=C:\Scripts\target_script.bat
    echo BATCH_ARGUMENTS=
    echo OUTPUT_LOG=remote_exec_output.log
    echo USE_SSL=0
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
if not defined OUTPUT_LOG set OUTPUT_LOG=%~dp0remote_exec_output.log
if not defined USE_SSL set USE_SSL=0

rem ====================================================================
rem メイン処理
rem ====================================================================

echo ========================================
echo PowerShell Remoting - リモートバッチ実行
echo （設定ファイル版）
echo ========================================
echo.
echo リモートサーバ: %REMOTE_SERVER%
echo 実行ユーザー  : %REMOTE_USER%
echo 実行ファイル  : %REMOTE_BATCH_PATH%
if defined BATCH_ARGUMENTS (
    echo 引数          : %BATCH_ARGUMENTS%
)
echo 出力ログ      : %OUTPUT_LOG%
if "%USE_SSL%"=="1" (
    echo プロトコル    : HTTPS ^(ポート 5986^)
) else (
    echo プロトコル    : HTTP ^(ポート 5985^)
)
echo.

rem PowerShellスクリプトのパス
set PS_SCRIPT=%~dp0Invoke-RemoteBatch.ps1

if not exist "%PS_SCRIPT%" (
    echo [エラー] PowerShellスクリプトが見つかりません: %PS_SCRIPT%
    goto :ERROR_EXIT
)

rem PowerShellが利用可能か確認
powershell -Command "Write-Host 'PowerShell確認OK'" >nul 2>&1
if errorlevel 1 (
    echo [エラー] PowerShellが利用できません。
    goto :ERROR_EXIT
)

rem PowerShellスクリプトを実行
set PS_PARAMS=-ComputerName "%REMOTE_SERVER%" -UserName "%REMOTE_USER%" -BatchPath "%REMOTE_BATCH_PATH%"

if defined BATCH_ARGUMENTS (
    set PS_PARAMS=!PS_PARAMS! -Arguments "%BATCH_ARGUMENTS%"
)

if defined OUTPUT_LOG (
    set PS_PARAMS=!PS_PARAMS! -OutputLog "%OUTPUT_LOG%"
)

if "%USE_SSL%"=="1" (
    set PS_PARAMS=!PS_PARAMS! -UseSSL
)

powershell -ExecutionPolicy Bypass -File "%PS_SCRIPT%" !PS_PARAMS!

if errorlevel 1 (
    goto :ERROR_EXIT
)

echo.
echo 処理が完了しました。
goto :END

:ERROR_EXIT
echo.
echo 処理を中断しました。
exit /b 1

:END
endlocal
exit /b 0
