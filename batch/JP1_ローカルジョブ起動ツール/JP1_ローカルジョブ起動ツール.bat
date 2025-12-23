@echo off
rem ============================================================================
rem JP1 ローカルジョブ起動ツール
rem
rem 概要:
rem   JP1/AJS3のジョブネットをローカルで起動するバッチファイルです。
rem
rem 使い方:
rem   1. 下記の「設定セクション」を編集
rem   2. JP1/AJS3がインストールされているサーバでこのファイルを実行
rem
rem 作成日: 2025-12-23
rem ============================================================================

chcp 932 >nul
title JP1 ローカルジョブ起動ツール
setlocal enabledelayedexpansion

rem ============================================================================
rem ■ 設定セクション（ここを編集してください）
rem ============================================================================

rem スケジューラーサービス名（通常は AJSROOT1）
set "SCHEDULER_SERVICE=AJSROOT1"

rem 起動するジョブネットのフルパス
set "JOBNET_PATH=/main_unit/jobgroup1/daily_batch"

rem ジョブ完了を待つ場合は 1、起動のみの場合は 0
set "WAIT_FOR_COMPLETION=1"

rem 完了待ちのタイムアウト（秒）。0の場合は無制限
set "WAIT_TIMEOUT=0"

rem 状態確認の間隔（秒）
set "POLLING_INTERVAL=10"

rem ============================================================================
rem ■ メイン処理（以下は編集不要）
rem ============================================================================

echo.
echo ================================================================
echo   JP1 ローカルジョブ起動ツール
echo ================================================================
echo.

rem 設定情報の表示
echo 設定情報:
echo   スケジューラー   : %SCHEDULER_SERVICE%
echo   ジョブネットパス : %JOBNET_PATH%
if "%WAIT_FOR_COMPLETION%"=="1" (
    echo   完了待ち         : 有効
    if "%WAIT_TIMEOUT%"=="0" (
        echo   タイムアウト     : 無制限
    ) else (
        echo   タイムアウト     : %WAIT_TIMEOUT%秒
    )
) else (
    echo   完了待ち         : 無効
)
echo.

rem ============================================================================
rem ajsentry実行（ジョブネット起動）
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

rem ajsentry実行（-n: 即時実行）
echo コマンド実行中: ajsentry -F %SCHEDULER_SERVICE% -n %JOBNET_PATH%
echo.
ajsentry -F %SCHEDULER_SERVICE% -n %JOBNET_PATH%
set "AJSENTRY_EXITCODE=%ERRORLEVEL%"
echo.

if %AJSENTRY_EXITCODE% neq 0 (
    echo [エラー] ジョブネットの起動に失敗しました
    echo   終了コード: %AJSENTRY_EXITCODE%
    goto :error_exit
)

echo [OK] ジョブネットの起動に成功しました

rem 実行登録番号を取得（ajsentry後の最新世代）
set "EXEC_REG_NUM="
for /f "delims=" %%A in ('ajsshow -F %SCHEDULER_SERVICE% -g 1 -i "%%ll" "%JOBNET_PATH%" 2^>^&1') do (
    if not defined EXEC_REG_NUM set "EXEC_REG_NUM=%%A"
)

rem 実行登録番号が変わったことを確認（自分が起動したジョブであることを保証）
if "!EXEC_REG_NUM!"=="!BEFORE_EXEC_REG_NUM!" (
    echo [警告] 実行登録番号が変化していません。既存のジョブを追跡します。
)
echo   実行登録番号: !EXEC_REG_NUM!
echo.

rem ============================================================================
rem 完了待ち処理
rem ============================================================================
if "%WAIT_FOR_COMPLETION%"=="0" goto :show_details

echo ================================================================
echo ジョブ完了待機中...
echo ================================================================
echo.

set "ELAPSED_SECONDS=0"
set "JOB_STATUS=unknown"

:wait_loop
rem タイムアウトチェック
if not "%WAIT_TIMEOUT%"=="0" (
    if %ELAPSED_SECONDS% geq %WAIT_TIMEOUT% (
        echo.
        echo [タイムアウト] 完了待ちがタイムアウトしました
        echo   タイムアウト時間: %WAIT_TIMEOUT%秒
        set "JOB_STATUS=timeout"
        goto :wait_done
    )
)

rem ajsshowでステータス（%CC）を取得して状態確認（実行登録番号で特定）
rem %CC = 状態（日本語文字列: 正常終了, 異常終了, 強制終了, 実行中 など）
set "WAIT_STATUS="
for /f "delims=" %%i in ('ajsshow -F %SCHEDULER_SERVICE% -B !EXEC_REG_NUM! -i "%%CC" "%JOBNET_PATH%" 2^>^&1') do (
    if not defined WAIT_STATUS set "WAIT_STATUS=%%i"
)

rem 出力がない場合はエラー
if not defined WAIT_STATUS (
    echo.
    echo [エラー] ステータスを取得できませんでした
    set "JOB_STATUS=error"
    goto :wait_done
)

rem KAVS エラーチェック
echo !WAIT_STATUS! | findstr /r "^KAVS.*-E" >nul
if %ERRORLEVEL%==0 (
    echo.
    echo [エラー] ajsshowでエラーが発生しました
    echo   エラー: !WAIT_STATUS!
    set "JOB_STATUS=error"
    goto :wait_done
)

rem 状態判定（日本語文字列で判定）
echo !WAIT_STATUS! | findstr /i "正常終了" >nul
if !ERRORLEVEL!==0 (
    set "JOB_STATUS=normal"
    goto :wait_done
)
echo !WAIT_STATUS! | findstr /i "異常終了 強制終了 中断" >nul
if !ERRORLEVEL!==0 (
    set "JOB_STATUS=abnormal"
    goto :wait_done
)

rem その他は実行中として待機継続
set /a "MINUTES=ELAPSED_SECONDS/60"
set /a "SECS=ELAPSED_SECONDS%%60"
echo   状態: !WAIT_STATUS! ^(経過時間: !MINUTES!分!SECS!秒^)

:wait_continue
rem 待機
ping -n %POLLING_INTERVAL% 127.0.0.1 >nul
set /a "ELAPSED_SECONDS+=POLLING_INTERVAL"
goto :wait_loop

:wait_done
echo.
echo ================================================================
if "%JOB_STATUS%"=="normal" (
    echo [OK] ジョブネットが正常終了しました
) else if "%JOB_STATUS%"=="abnormal" (
    echo [NG] ジョブネットが異常終了しました
) else if "%JOB_STATUS%"=="timeout" (
    echo [NG] 完了待ちがタイムアウトしました
) else if "%JOB_STATUS%"=="not_found" (
    echo [NG] ジョブネットが見つかりません
) else if "%JOB_STATUS%"=="error" (
    echo [NG] コマンド実行エラー
) else (
    echo [--] ジョブネット状態: %JOB_STATUS%
)
echo ================================================================
echo.

rem ============================================================================
rem 詳細情報の取得
rem ============================================================================
:show_details
echo ================================================================
echo ジョブ詳細情報を取得中...
echo ================================================================
echo.

echo 詳細情報 (ajsshow -E -B !EXEC_REG_NUM!):
echo ----------------------------------------------------------------
ajsshow -F %SCHEDULER_SERVICE% -B !EXEC_REG_NUM! -E %JOBNET_PATH%
echo ----------------------------------------------------------------
echo.

rem ============================================================================
rem サマリー表示
rem ============================================================================
:summary
echo ================================================================
echo 処理サマリー
echo ================================================================
echo.
echo   ジョブネット: %JOBNET_PATH%
echo   起動結果    : 成功

if "%WAIT_FOR_COMPLETION%"=="1" (
    if "%JOB_STATUS%"=="normal" (
        echo   実行結果    : 正常終了
    ) else if "%JOB_STATUS%"=="abnormal" (
        echo   実行結果    : 異常終了
    ) else if "%JOB_STATUS%"=="timeout" (
        echo   実行結果    : タイムアウト
    ) else if "%JOB_STATUS%"=="error" (
        echo   実行結果    : コマンドエラー
    ) else (
        echo   実行結果    : %JOB_STATUS%
    )
)
echo.

rem 終了判定
if "%JOB_STATUS%"=="normal" goto :normal_exit
if "%JOB_STATUS%"=="unknown" goto :normal_exit
if "%JOB_STATUS%"=="error" goto :error_exit
if "%WAIT_FOR_COMPLETION%"=="0" goto :normal_exit
goto :error_exit

:normal_exit
echo 処理が完了しました。
echo.
pause
exit /b 0

:error_exit
echo.
echo 追加の確認事項:
echo   - JP1/AJS3サービスが起動しているか
echo   - ジョブネットパスが正しいか
echo.
pause
exit /b 1
