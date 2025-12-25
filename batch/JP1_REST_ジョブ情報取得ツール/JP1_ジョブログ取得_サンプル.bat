@echo off
title JP1 ジョブログ取得サンプル
setlocal

rem ============================================================================
rem JP1 ジョブログ取得サンプル
rem
rem 説明:
rem   JP1_REST_ジョブ情報取得ツール.bat を呼び出し、
rem   取得したログをファイルに保存します。
rem
rem 使い方:
rem   1. 下記の UNIT_PATH を取得したいジョブネットのパスに変更
rem   2. 下記の OUTPUT_FILE を出力先ファイルパスに変更
rem   3. このバッチをダブルクリックで実行
rem ============================================================================

rem === ここを編集してください ===
set "UNIT_PATH=/JobGroup/Jobnet"
set "OUTPUT_FILE=%~dp0joblog_output.txt"

rem スクリプトのディレクトリを取得
set "SCRIPT_DIR=%~dp0"

echo.
echo ================================================================
echo   JP1 ジョブログ取得サンプル
echo ================================================================
echo.
echo   対象: %UNIT_PATH%
echo   出力: %OUTPUT_FILE%
echo.
echo ログを取得中...

rem JP1_REST_ジョブ情報取得ツール.bat を呼び出し、結果を直接ファイルに保存
call "%SCRIPT_DIR%JP1_REST_ジョブ情報取得ツール.bat" "%UNIT_PATH%" > "%OUTPUT_FILE%" 2>&1
set "EXIT_CODE=%ERRORLEVEL%"

rem エラーコードでチェック
if %EXIT_CODE% neq 0 (
    echo.
    echo [エラー] ログの取得に失敗しました（終了コード: %EXIT_CODE%）
    echo.
    del "%OUTPUT_FILE%" >nul 2>&1
    pause
    exit /b %EXIT_CODE%
)

rem 結果が空かチェック
for %%A in ("%OUTPUT_FILE%") do set "FILE_SIZE=%%~zA"
if "%FILE_SIZE%"=="0" (
    echo.
    echo [警告] 取得結果が空です
    echo.
    del "%OUTPUT_FILE%" >nul 2>&1
    pause
    exit /b 1
)

echo.
echo ================================================================
echo   取得完了 - ファイルに保存しました
echo ================================================================
echo.
echo 出力ファイル: %OUTPUT_FILE%
echo.

pause
exit /b 0
