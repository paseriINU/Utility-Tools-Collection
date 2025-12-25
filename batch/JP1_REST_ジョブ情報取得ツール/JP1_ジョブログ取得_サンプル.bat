<# :
@echo off
chcp 65001 >nul
title JP1 ジョブログ取得サンプル
setlocal enabledelayedexpansion

rem ============================================================================
rem JP1 ジョブログ取得サンプル
rem
rem 説明:
rem   JP1_REST_ジョブ情報取得ツール.bat を呼び出し、
rem   取得したログをクリップボードにコピーします。
rem
rem 使い方:
rem   1. 下記の UNIT_PATH を編集
rem   2. このファイルをダブルクリックで実行
rem ============================================================================

rem 取得対象のユニットパス（ここを編集）
set "UNIT_PATH=/JobGroup/Jobnet"

rem スクリプトのディレクトリを取得
set "SCRIPT_DIR=%~dp0"

echo.
echo ================================================================
echo   JP1 ジョブログ取得サンプル
echo ================================================================
echo.
echo   対象: %UNIT_PATH%
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

rem クリップボードにコピー
type "%TEMP_FILE%" | clip

echo.
echo ================================================================
echo   取得完了 - クリップボードにコピーしました
echo ================================================================
echo.
echo 取得内容:
echo ----------------------------------------
type "%TEMP_FILE%"
echo ----------------------------------------
echo.

rem 一時ファイルを削除
del "%TEMP_FILE%" >nul 2>&1

pause
exit /b 0
: #>
