@echo off
title JP1 ジョブログ取得サンプル
setlocal enabledelayedexpansion

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

rem JP1_REST_ジョブ情報取得ツール.bat を呼び出し
echo ログを取得中...

rem 一時ファイルに結果を保存
set "TEMP_FILE=%TEMP%\jp1_log_%RANDOM%.txt"

call "%SCRIPT_DIR%JP1_REST_ジョブ情報取得ツール.bat" "%UNIT_PATH%" > "%TEMP_FILE%" 2>&1
set "EXIT_CODE=%ERRORLEVEL%"

rem エラーチェック
if %EXIT_CODE% neq 0 (
    echo.
    echo [エラー] ログの取得に失敗しました（終了コード: %EXIT_CODE%）
    echo.
    if exist "%TEMP_FILE%" (
        echo エラー内容:
        type "%TEMP_FILE%"
        del "%TEMP_FILE%" >nul 2>&1
    )
    echo.
    pause
    exit /b %EXIT_CODE%
)

rem 結果が空かチェック
for %%A in ("%TEMP_FILE%") do set "FILE_SIZE=%%~zA"
if "%FILE_SIZE%"=="0" (
    echo.
    echo [警告] 取得結果が空です
    echo.
    del "%TEMP_FILE%" >nul 2>&1
    pause
    exit /b 1
)

rem ERRORで始まる行があるかチェック
findstr /B "ERROR:" "%TEMP_FILE%" >nul 2>&1
if %ERRORLEVEL% equ 0 (
    echo.
    echo [エラー] API呼び出しでエラーが発生しました
    echo.
    type "%TEMP_FILE%"
    echo.
    del "%TEMP_FILE%" >nul 2>&1
    pause
    exit /b 1
)

rem 出力ファイルに保存（UTF-8からShift-JISに変換）
powershell -NoProfile -Command "Get-Content '%TEMP_FILE%' -Encoding UTF8 | Set-Content '%OUTPUT_FILE%' -Encoding Default"

echo.
echo ================================================================
echo   取得完了 - ファイルに保存しました
echo ================================================================
echo.
echo 出力ファイル: %OUTPUT_FILE%
echo.

rem 一時ファイルを削除
del "%TEMP_FILE%" >nul 2>&1

pause
exit /b 0
