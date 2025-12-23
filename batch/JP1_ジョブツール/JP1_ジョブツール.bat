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
echo.
echo   0. キャンセル
echo.
set /p "MENU_CHOICE=選択 (0-2): "

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

rem ajsentry実行
echo コマンド実行中: ajsentry -F %SCHEDULER_SERVICE% %JOBNET_PATH%
echo.
ajsentry -F %SCHEDULER_SERVICE% %JOBNET_PATH%
set "AJSENTRY_EXITCODE=%ERRORLEVEL%"
echo.

if %AJSENTRY_EXITCODE% neq 0 (
    echo [エラー] ジョブネットの起動に失敗しました
    echo   終了コード: %AJSENTRY_EXITCODE%
    goto :ERROR_EXIT
)

echo [OK] ジョブネットの起動に成功しました
echo.

rem ============================================================================
rem ジョブ完了待ち
rem ============================================================================
echo ================================================================
echo ジョブ完了待機中...
echo ================================================================
echo.

set "ELAPSED_SECONDS=0"
set "JOB_STATUS=unknown"

:WAIT_LOOP
rem タイムアウトチェック
if not "%WAIT_TIMEOUT%"=="0" (
    if %ELAPSED_SECONDS% geq %WAIT_TIMEOUT% (
        echo.
        echo [タイムアウト] 完了待ちがタイムアウトしました
        set "JOB_STATUS=timeout"
        goto :WAIT_DONE
    )
)

rem ajsshowでステータス（%CC）を取得して状態確認
rem %CC = 状態コード（E:正常終了, A:異常終了, K:強制終了, r:実行中 など）
set "WAIT_STATUS="
for /f "delims=" %%i in ('ajsshow -F %SCHEDULER_SERVICE% -g 1 -i "%%CC" "%JOBNET_PATH%" 2^>^&1') do (
    if not defined WAIT_STATUS set "WAIT_STATUS=%%i"
)

rem 出力がない場合はエラー
if not defined WAIT_STATUS (
    echo.
    echo [エラー] ステータスを取得できませんでした
    set "JOB_STATUS=error"
    goto :WAIT_DONE
)

rem KAVS エラーチェック
echo !WAIT_STATUS! | findstr /r "^KAVS.*-E" >nul
if %ERRORLEVEL%==0 (
    echo.
    echo [エラー] ajsshowでエラーが発生しました
    echo   エラー: !WAIT_STATUS!
    set "JOB_STATUS=error"
    goto :WAIT_DONE
)

rem 状態判定（E:正常終了、A/K:異常終了、その他:実行中）
if "!WAIT_STATUS!"=="E" (
    set "JOB_STATUS=normal"
    goto :WAIT_DONE
)
if "!WAIT_STATUS!"=="A" (
    set "JOB_STATUS=abnormal"
    goto :WAIT_DONE
)
if "!WAIT_STATUS!"=="K" (
    set "JOB_STATUS=abnormal"
    goto :WAIT_DONE
)

rem その他は実行中として待機継続
set /a "MINUTES=ELAPSED_SECONDS/60"
set /a "SECS=ELAPSED_SECONDS%%60"
echo   状態: !WAIT_STATUS! ^(経過時間: !MINUTES!分!SECS!秒^)

:WAIT_CONTINUE
ping -n %POLLING_INTERVAL% 127.0.0.1 >nul
set /a "ELAPSED_SECONDS+=POLLING_INTERVAL"
goto :WAIT_LOOP

:WAIT_DONE
echo.
echo ================================================================
if "%JOB_STATUS%"=="normal" (
    echo [OK] ジョブネットが正常終了しました
) else if "%JOB_STATUS%"=="abnormal" (
    echo [NG] ジョブネットが異常終了しました
    goto :ERROR_EXIT
) else if "%JOB_STATUS%"=="timeout" (
    echo [NG] 完了待ちがタイムアウトしました
    goto :ERROR_EXIT
) else if "%JOB_STATUS%"=="not_found" (
    echo [NG] ジョブネットが見つかりません
    goto :ERROR_EXIT
) else if "%JOB_STATUS%"=="error" (
    echo [NG] コマンド実行エラー
    goto :ERROR_EXIT
) else (
    echo [--] ジョブネット状態: %JOB_STATUS%
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

rem まず実終了コード（%%RR）を取得してエラー判定
set "RETURN_CODE1="
for /f "delims=" %%A in ('ajsshow -F %SCHEDULER_SERVICE% -g 1 -i "%%RR" "%JOB_PATH1%" 2^>^&1') do (
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

rem 標準出力ファイルパスを取得
set "LOG_FILE_PATH1="
for /f "delims=" %%A in ('ajsshow -F %SCHEDULER_SERVICE% -g 1 -i "%%so" "%JOB_PATH1%" 2^>^&1') do (
    if not defined LOG_FILE_PATH1 set "LOG_FILE_PATH1=%%A"
)
set "LOG_FILE_PATH1=!LOG_FILE_PATH1:'=!"

echo   標準出力ファイル: !LOG_FILE_PATH1!

if defined LOG_FILE_PATH1 (
    if exist "!LOG_FILE_PATH1!" (
        echo.
        echo ----------------------------------------------------------------
        echo [ジョブ1] %JOB_PATH1%
        echo ----------------------------------------------------------------
        type "!LOG_FILE_PATH1!"
        echo.

        rem ログをファイルに出力
        set "OUTPUT_FILE1=%OUTPUT_DIR%job1_log.txt"
        copy "!LOG_FILE_PATH1!" "!OUTPUT_FILE1!" >nul
        echo [OK] ジョブ1のログを出力しました: !OUTPUT_FILE1!
        echo.

        rem ======================================================================
        rem ■ ここに入れたい処理を記述してください
        rem ======================================================================


        rem ======================================================================

    ) else (
        echo [情報] ジョブ1の標準出力ファイルが存在しません
    )
) else (
    echo [情報] ジョブ1の標準出力ファイルパスを取得できませんでした
)
echo.

:GET_JOB2
rem ----------------------------------------------------------------------------
rem ジョブ2のログ取得
rem ----------------------------------------------------------------------------
echo [ジョブ2] %JOB_PATH2%
echo.

rem まず実終了コード（%%RR）を取得してエラー判定
set "RETURN_CODE2="
for /f "delims=" %%A in ('ajsshow -F %SCHEDULER_SERVICE% -g 1 -i "%%RR" "%JOB_PATH2%" 2^>^&1') do (
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

rem 標準出力ファイルパスを取得
set "LOG_FILE_PATH2="
for /f "delims=" %%A in ('ajsshow -F %SCHEDULER_SERVICE% -g 1 -i "%%so" "%JOB_PATH2%" 2^>^&1') do (
    if not defined LOG_FILE_PATH2 set "LOG_FILE_PATH2=%%A"
)
set "LOG_FILE_PATH2=!LOG_FILE_PATH2:'=!"

echo   標準出力ファイル: !LOG_FILE_PATH2!

if defined LOG_FILE_PATH2 (
    if exist "!LOG_FILE_PATH2!" (
        echo.
        echo ----------------------------------------------------------------
        echo [ジョブ2] %JOB_PATH2%
        echo ----------------------------------------------------------------
        type "!LOG_FILE_PATH2!"
        echo.

        rem ログをファイルに出力
        set "OUTPUT_FILE2=%OUTPUT_DIR%job2_log.txt"
        copy "!LOG_FILE_PATH2!" "!OUTPUT_FILE2!" >nul
        echo [OK] ジョブ2のログを出力しました: !OUTPUT_FILE2!
    ) else (
        echo [情報] ジョブ2の標準出力ファイルが存在しません
    )
) else (
    echo [情報] ジョブ2の標準出力ファイルパスを取得できませんでした
)
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
