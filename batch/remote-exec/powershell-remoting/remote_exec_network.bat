@echo off
rem ====================================================================
rem ネットワークパス上のPowerShellスクリプトを実行
rem バッチをローカルに、.ps1をサーバ上に配置する場合に使用
rem ====================================================================

setlocal enabledelayedexpansion

rem ====================================================================
rem 設定項目（必要に応じて編集してください）
rem ====================================================================

rem ネットワークパス上のPowerShellスクリプト
set NETWORK_PS_SCRIPT=\\192.168.1.100\Share\Scripts\Invoke-RemoteBatch.ps1

rem リモートサーバのコンピュータ名またはIPアドレス
set REMOTE_SERVER=192.168.1.100

rem リモートサーバの管理者ユーザー名
set REMOTE_USER=Administrator

rem リモートサーバで実行するバッチファイルのフルパス
set REMOTE_BATCH_PATH=C:\Scripts\target_script.bat

rem バッチファイルに渡す引数（オプション）
set BATCH_ARGUMENTS=

rem 実行結果を保存するローカルファイル（オプション）
set OUTPUT_LOG=%USERPROFILE%\Desktop\remote_exec_output.log

rem HTTPS使用（1=使用する, 0=使用しない）
set USE_SSL=0

rem ====================================================================
rem メイン処理
rem ====================================================================

echo ========================================
echo リモートバッチ実行（ネットワークパス版）
echo ========================================
echo.
echo スクリプト    : %NETWORK_PS_SCRIPT%
echo リモートサーバ: %REMOTE_SERVER%
echo 実行ユーザー  : %REMOTE_USER%
echo 実行ファイル  : %REMOTE_BATCH_PATH%
echo 出力ログ      : %OUTPUT_LOG%
echo.

rem PowerShellスクリプトの存在確認
if not exist "%NETWORK_PS_SCRIPT%" (
    echo [エラー] PowerShellスクリプトが見つかりません: %NETWORK_PS_SCRIPT%
    echo.
    echo ネットワークパスにアクセスできるか確認してください。
    pause
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

rem ネットワークパスのスクリプトを実行
powershell -ExecutionPolicy Bypass -File "%NETWORK_PS_SCRIPT%" !PS_PARAMS!

if errorlevel 1 (
    goto :ERROR_EXIT
)

echo.
echo 処理が完了しました。
goto :END

:ERROR_EXIT
echo.
echo 処理を中断しました。
pause
exit /b 1

:END
endlocal
exit /b 0
