@echo off
rem ====================================================================
rem リモートWindowsサーバ上でバッチファイルをPowerShell Remotingで実行
rem PowerShellスクリプト（Invoke-RemoteBatch.ps1）のラッパー
rem ====================================================================

setlocal enabledelayedexpansion

rem ====================================================================
rem 設定項目（必要に応じて編集してください）
rem ====================================================================

rem リモートサーバのコンピュータ名またはIPアドレス
set REMOTE_SERVER=192.168.1.100

rem リモートサーバの管理者ユーザー名
set REMOTE_USER=Administrator

rem リモートサーバで実行するバッチファイルのフルパス
set REMOTE_BATCH_PATH=C:\Scripts\target_script.bat

rem バッチファイルに渡す引数（オプション）
set BATCH_ARGUMENTS=

rem 実行結果を保存するローカルファイル（オプション）
set OUTPUT_LOG=%~dp0remote_exec_output.log

rem HTTPS使用（1=使用する, 0=使用しない）
set USE_SSL=0

rem ====================================================================
rem メイン処理
rem ====================================================================

echo ========================================
echo PowerShell Remoting - リモートバッチ実行
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

rem 詳細ログを有効化する場合は -Verbose を追加
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
