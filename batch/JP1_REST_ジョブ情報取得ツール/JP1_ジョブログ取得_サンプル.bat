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
rem
rem 終了コード:
rem   0: 正常終了
rem   1: 引数エラー
rem   2: API接続エラー（ユニット一覧取得）
rem   3: 5MB超過エラー（結果切り捨て）
rem   4: 詳細取得エラー
rem ============================================================================

rem === ここを編集してください ===
set "UNIT_PATH=/JobGroup/Jobnet"

rem 出力ファイル名（日付_時間形式）
set "DATETIME=%date:~0,4%%date:~5,2%%date:~8,2%_%time:~0,2%%time:~3,2%%time:~6,2%"
set "DATETIME=%DATETIME: =0%"
set "OUTPUT_FILE=%~dp0%DATETIME%.txt"

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

rem エラーコード別のハンドリング
if %EXIT_CODE% equ 0 goto :SUCCESS
if %EXIT_CODE% equ 1 goto :ERR_ARGUMENT
if %EXIT_CODE% equ 2 goto :ERR_API_CONNECTION
if %EXIT_CODE% equ 3 goto :ERR_5MB_EXCEEDED
if %EXIT_CODE% equ 4 goto :ERR_DETAIL_FETCH
goto :ERR_UNKNOWN

:ERR_ARGUMENT
echo.
echo [エラー] 引数エラー（ユニットパスが指定されていません）
goto :ERROR_EXIT

:ERR_API_CONNECTION
echo.
echo [エラー] API接続エラー（ユニット一覧の取得に失敗しました）
echo          - Web Consoleが起動しているか確認してください
echo          - 接続設定（ホスト名・ポート）を確認してください
echo          - 認証情報（ユーザー名・パスワード）を確認してください
goto :ERROR_EXIT

:ERR_5MB_EXCEEDED
echo.
echo [エラー] 5MB超過エラー（実行結果が大きすぎて切り捨てられました）
echo          - 対象ユニットの出力サイズを確認してください
goto :ERROR_EXIT

:ERR_DETAIL_FETCH
echo.
echo [エラー] 詳細取得エラー（実行結果詳細の取得に失敗しました）
echo          - ユニットパスが正しいか確認してください
echo          - 実行履歴が存在するか確認してください
goto :ERROR_EXIT

:ERR_UNKNOWN
echo.
echo [エラー] 不明なエラー（終了コード: %EXIT_CODE%）
goto :ERROR_EXIT

:ERROR_EXIT
echo.
del "%OUTPUT_FILE%" >nul 2>&1
pause
exit /b %EXIT_CODE%

:SUCCESS
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

rem メモ帳で開く
start notepad "%OUTPUT_FILE%"

exit /b 0
