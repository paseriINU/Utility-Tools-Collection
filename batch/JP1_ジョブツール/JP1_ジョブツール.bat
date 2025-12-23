@echo off
rem ============================================================================
rem JP1 ジョブツール
rem
rem 概要:
rem   JP1/AJS3のジョブネット起動とログ取得を行うツールです。
rem   ジョブネット起動 → 完了待ち → ログ取得を連続実行します。
rem
rem 使い方:
rem   1. 下記の「設定セクション」を編集
rem   2. JP1/AJS3がインストールされているサーバでこのファイルを実行
rem
rem 作成日: 2025-12-23
rem ============================================================================

chcp 932 >nul
title JP1 ジョブツール
setlocal enabledelayedexpansion

rem ============================================================================
rem ■ 設定セクション（ここを編集してください）
rem ============================================================================

rem スケジューラーサービス名（通常は AJSROOT1）
set "SCHEDULER_SERVICE=AJSROOT1"

rem 完了待ちのタイムアウト（秒）。0の場合は無制限
set "WAIT_TIMEOUT=0"

rem 状態確認の間隔（秒）
set "POLLING_INTERVAL=10"

rem ログ出力先フォルダ（バッチファイルと同じ場所に出力する場合は %~dp0 のまま）
set "OUTPUT_DIR=%~dp0"

rem ----------------------------------------------------------------------------
rem 選択肢1: TEST
rem ----------------------------------------------------------------------------
set "MENU1_NAME=TEST"
set "MENU1_JOBNET=/main_unit/jobgroup1/test_batch"
set "MENU1_JOB1=/main_unit/jobgroup1/test_batch/job1"
set "MENU1_JOB2=/main_unit/jobgroup1/test_batch/job2"

rem ----------------------------------------------------------------------------
rem 選択肢2: TEST2
rem ----------------------------------------------------------------------------
set "MENU2_NAME=TEST2"
set "MENU2_JOBNET=/main_unit/jobgroup2/test2_batch"
set "MENU2_JOB1=/main_unit/jobgroup2/test2_batch/job1"
set "MENU2_JOB2=/main_unit/jobgroup2/test2_batch/job2"

rem ----------------------------------------------------------------------------
rem 選択肢3: TEST3
rem ----------------------------------------------------------------------------
set "MENU3_NAME=TEST3"
set "MENU3_JOBNET=/main_unit/jobgroup3/test3_batch"
set "MENU3_JOB1=/main_unit/jobgroup3/test3_batch/job1"
set "MENU3_JOB2=/main_unit/jobgroup3/test3_batch/job2"

rem ============================================================================
rem ■ メイン処理（以下は編集不要）
rem ============================================================================

echo.
echo ================================================================
echo   JP1 ジョブツール
echo ================================================================
echo.

rem ----------------------------------------------------------------------------
rem メニュー表示
rem ----------------------------------------------------------------------------
echo 実行するジョブを選択してください:
echo.
echo   1. %MENU1_NAME%
echo   2. %MENU2_NAME%
echo   3. %MENU3_NAME%
echo.
echo   0. キャンセル
echo.
set /p "MENU_CHOICE=選択 (0-3): "

if "%MENU_CHOICE%"=="0" (
    echo 処理をキャンセルしました。
    goto :NORMAL_EXIT
)

if "%MENU_CHOICE%"=="1" (
    set "JOBNET_PATH=%MENU1_JOBNET%"
    set "JOB_PATH1=%MENU1_JOB1%"
    set "JOB_PATH2=%MENU1_JOB2%"
    set "SELECTED_NAME=%MENU1_NAME%"
    goto :START_JOB
)

if "%MENU_CHOICE%"=="2" (
    set "JOBNET_PATH=%MENU2_JOBNET%"
    set "JOB_PATH1=%MENU2_JOB1%"
    set "JOB_PATH2=%MENU2_JOB2%"
    set "SELECTED_NAME=%MENU2_NAME%"
    goto :START_JOB
)

if "%MENU_CHOICE%"=="3" (
    set "JOBNET_PATH=%MENU3_JOBNET%"
    set "JOB_PATH1=%MENU3_JOB1%"
    set "JOB_PATH2=%MENU3_JOB2%"
    set "SELECTED_NAME=%MENU3_NAME%"
    goto :START_JOB
)

echo [エラー] 無効な選択です
goto :ERROR_EXIT

:START_JOB
echo.
echo ================================================================
echo 選択: %SELECTED_NAME%
echo ================================================================
echo.
echo 設定情報:
echo   スケジューラー   : %SCHEDULER_SERVICE%
echo   ジョブネットパス : %JOBNET_PATH%
echo   ジョブパス1      : %JOB_PATH1%
echo   ジョブパス2      : %JOB_PATH2%
if "%WAIT_TIMEOUT%"=="0" (
    echo   タイムアウト     : 無制限
) else (
    echo   タイムアウト     : %WAIT_TIMEOUT%秒
)
echo.

rem ============================================================================
rem ジョブネット起動
rem ============================================================================
echo ================================================================
echo ジョブネット起動中...
echo ================================================================
echo.

rem ajsentry実行前に現在の最新実行登録番号を取得（比較用）
set "BEFORE_EXEC_REG_NUM="
for /f "delims=" %%A in ('ajsshow -F %SCHEDULER_SERVICE% -g 1 -i "%%ll" "%JOBNET_PATH%" 2^>^&1') do (
    if not defined BEFORE_EXEC_REG_NUM set "BEFORE_EXEC_REG_NUM=%%A"
)

rem ajsentry実行（-n: 即時実行, -w: 完了待ち）
rem -wオプションによりジョブネット終了まで待機し、終了コードで結果を判定
echo コマンド実行中: ajsentry -F %SCHEDULER_SERVICE% -n -w %JOBNET_PATH%
echo.
ajsentry -F %SCHEDULER_SERVICE% -n -w %JOBNET_PATH%
set "AJSENTRY_EXITCODE=%ERRORLEVEL%"
echo.

rem 実行登録番号を取得（ajsentry後の最新世代）
set "EXEC_REG_NUM="
for /f "delims=" %%A in ('ajsshow -F %SCHEDULER_SERVICE% -g 1 -i "%%ll" "%JOBNET_PATH%" 2^>^&1') do (
    if not defined EXEC_REG_NUM set "EXEC_REG_NUM=%%A"
)

rem 実行登録番号が変わったことを確認（自分が起動したジョブであることを保証）
if "!EXEC_REG_NUM!"=="!BEFORE_EXEC_REG_NUM!" (
    echo [エラー] 実行登録番号が変化していません。ジョブが実行されませんでした。
    goto :ERROR_EXIT
)
echo   実行登録番号: !EXEC_REG_NUM!
echo.

rem ajsentry終了後、ajsshowで1回だけ結果を取得
rem （ajsentryの戻り値はコマンド実行成否であり、ジョブネット結果ではない）
set "JOB_STATUS="
for /f "delims=" %%i in ('ajsshow -F %SCHEDULER_SERVICE% -B !EXEC_REG_NUM! -i "%%CC" "%JOBNET_PATH%" 2^>^&1') do (
    if not defined JOB_STATUS set "JOB_STATUS=%%i"
)

echo ================================================================
echo ジョブネット状態: !JOB_STATUS!
echo.

rem 状態判定
echo !JOB_STATUS! | findstr /i "正常終了" >nul
if !ERRORLEVEL!==0 (
    echo [OK] ジョブネットが正常終了しました
) else (
    echo !JOB_STATUS! | findstr /i "異常終了 強制終了 中断" >nul
    if !ERRORLEVEL!==0 (
        echo [NG] ジョブネットが異常終了しました
    ) else (
        echo [--] ジョブネット状態: !JOB_STATUS!
    )
)
echo ================================================================
echo.

rem ============================================================================
rem ジョブログ取得（2つのジョブ）
rem ============================================================================
echo ================================================================
echo ジョブログを取得中...
echo ================================================================
echo.

rem ----------------------------------------------------------------------------
rem ジョブ1のログ取得
rem ----------------------------------------------------------------------------
echo [ジョブ1] %JOB_PATH1%
echo.

rem まず実終了コード（%%RR）を取得してエラー判定（実行登録番号で特定）
set "RETURN_CODE1="
for /f "delims=" %%A in ('ajsshow -F %SCHEDULER_SERVICE% -B !EXEC_REG_NUM! -i "%%RR" "%JOB_PATH1%" 2^>^&1') do (
    if not defined RETURN_CODE1 set "RETURN_CODE1=%%A"
)

rem 数値かどうかでエラー判定（数値以外ならエラー）
set "RETURN_CODE1_NUM="
for /f "delims=0123456789" %%A in ("!RETURN_CODE1!") do set "RETURN_CODE1_NUM=%%A"
if defined RETURN_CODE1_NUM (
    echo [エラー] ジョブ1の情報取得に失敗しました
    echo   エラー: !RETURN_CODE1!
    goto :GET_JOB2
)

echo   実終了コード: !RETURN_CODE1!
if not "!RETURN_CODE1!"=="0" (
    echo   [警告] ジョブ1は異常終了しています
)

rem 標準出力ファイルパスを取得（実行登録番号で特定）
set "LOG_FILE_PATH1="
for /f "delims=" %%A in ('ajsshow -F %SCHEDULER_SERVICE% -B !EXEC_REG_NUM! -i "%%so" "%JOB_PATH1%" 2^>^&1') do (
    if not defined LOG_FILE_PATH1 set "LOG_FILE_PATH1=%%A"
)

echo   標準出力ファイル: !LOG_FILE_PATH1!

rem サブルーチンでファイル処理（特殊文字対策）
call :PROCESS_LOG1
goto :AFTER_LOG1

:PROCESS_LOG1
if not exist "%LOG_FILE_PATH1%" (
    echo [情報] ジョブ1の標準出力ファイルが存在しません
    exit /b
)
echo.
echo ----------------------------------------------------------------
echo [ジョブ1] %JOB_PATH1%
echo ----------------------------------------------------------------
type "%LOG_FILE_PATH1%"
echo.

rem ログをファイルに出力
set "OUTPUT_FILE1=%OUTPUT_DIR%job1_log.txt"
copy "%LOG_FILE_PATH1%" "%OUTPUT_FILE1%" >nul
echo [OK] ジョブ1のログを出力しました: %OUTPUT_FILE1%
echo.

rem ======================================================================
rem ■ ここに入れたい処理を記述してください
rem ======================================================================


rem ======================================================================

exit /b

:AFTER_LOG1
echo.

:GET_JOB2
rem ----------------------------------------------------------------------------
rem ジョブ2のログ取得
rem ----------------------------------------------------------------------------
echo [ジョブ2] %JOB_PATH2%
echo.

rem まず実終了コード（%%RR）を取得してエラー判定（実行登録番号で特定）
set "RETURN_CODE2="
for /f "delims=" %%A in ('ajsshow -F %SCHEDULER_SERVICE% -B !EXEC_REG_NUM! -i "%%RR" "%JOB_PATH2%" 2^>^&1') do (
    if not defined RETURN_CODE2 set "RETURN_CODE2=%%A"
)

rem 数値かどうかでエラー判定（数値以外ならエラー）
set "RETURN_CODE2_NUM="
for /f "delims=0123456789" %%A in ("!RETURN_CODE2!") do set "RETURN_CODE2_NUM=%%A"
if defined RETURN_CODE2_NUM (
    echo [エラー] ジョブ2の情報取得に失敗しました
    echo   エラー: !RETURN_CODE2!
    goto :LOG_DONE
)

echo   実終了コード: !RETURN_CODE2!
if not "!RETURN_CODE2!"=="0" (
    echo   [警告] ジョブ2は異常終了しています
)

rem 標準出力ファイルパスを取得（実行登録番号で特定）
set "LOG_FILE_PATH2="
for /f "delims=" %%A in ('ajsshow -F %SCHEDULER_SERVICE% -B !EXEC_REG_NUM! -i "%%so" "%JOB_PATH2%" 2^>^&1') do (
    if not defined LOG_FILE_PATH2 set "LOG_FILE_PATH2=%%A"
)

echo   標準出力ファイル: !LOG_FILE_PATH2!

rem サブルーチンでファイル処理（特殊文字対策）
call :PROCESS_LOG2
goto :AFTER_LOG2

:PROCESS_LOG2
if not exist "%LOG_FILE_PATH2%" (
    echo [情報] ジョブ2の標準出力ファイルが存在しません
    exit /b
)
echo.
echo ----------------------------------------------------------------
echo [ジョブ2] %JOB_PATH2%
echo ----------------------------------------------------------------
type "%LOG_FILE_PATH2%"
echo.

rem ログをファイルに出力
set "OUTPUT_FILE2=%OUTPUT_DIR%job2_log.txt"
copy "%LOG_FILE_PATH2%" "%OUTPUT_FILE2%" >nul
echo [OK] ジョブ2のログを出力しました: %OUTPUT_FILE2%
echo.

rem ======================================================================
rem ■ ここに入れたい処理を記述してください
rem ======================================================================


rem ======================================================================

exit /b

:AFTER_LOG2
echo.

:LOG_DONE

rem ============================================================================
rem サマリー表示
rem ============================================================================
echo ================================================================
echo 処理サマリー
echo ================================================================
echo.
echo   選択         : %SELECTED_NAME%
echo   ジョブネット : %JOBNET_PATH%
echo   ジョブ1      : %JOB_PATH1%
echo   ジョブ2      : %JOB_PATH2%
echo   起動結果     : 成功
echo   実行結果     : 正常終了
echo   ログ取得     : 成功（ファイル出力済み）
echo.

:NORMAL_EXIT
echo 処理が完了しました。
echo.
pause
exit /b 0

:ERROR_EXIT
echo.
echo 追加の確認事項:
echo   - JP1/AJS3サービスが起動しているか
echo   - ジョブネットパス、ジョブパスが正しいか
echo.
pause
exit /b 1
