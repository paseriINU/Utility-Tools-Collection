@echo off
rem ====================================================================
rem JP1ジョブネット起動ツール（REST API版 - バッチラッパー）
rem ====================================================================

setlocal

cls
echo ========================================
echo JP1ジョブネット起動ツール
echo （REST API版）
echo ========================================
echo.

rem ====================================================================
rem 設定項目（ここを編集してください）
rem ====================================================================

set JP1_HOST=192.168.1.100
set JP1_PORT=22250
set JP1_USER=jp1admin
set JOBNET_PATH=/main_unit/jobgroup1/daily_batch
set USE_SSL=false

rem ====================================================================
rem PowerShellスクリプト実行
rem ====================================================================

if "%USE_SSL%"=="true" (
    set SSL_FLAG=-UseSSL
) else (
    set SSL_FLAG=
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Start-JP1Job.ps1" ^
    -JP1Host "%JP1_HOST%" ^
    -JP1Port %JP1_PORT% ^
    -JP1User "%JP1_USER%" ^
    -JobnetPath "%JOBNET_PATH%" ^
    %SSL_FLAG%

if errorlevel 1 (
    echo.
    echo [エラー] 処理に失敗しました。
    pause
    exit /b 1
)

echo.
pause
endlocal
exit /b 0
