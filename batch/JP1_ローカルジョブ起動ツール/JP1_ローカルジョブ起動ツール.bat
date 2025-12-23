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

rem ajsentry実行
echo コマンド実行中: ajsentry -F %SCHEDULER_SERVICE% %JOBNET_PATH%
echo.
ajsentry -F %SCHEDULER_SERVICE% %JOBNET_PATH%
set "AJSENTRY_EXITCODE=%ERRORLEVEL%"
echo.

if %AJSENTRY_EXITCODE% neq 0 (
    echo [エラー] ジョブネットの起動に失敗しました
    echo   終了コード: %AJSENTRY_EXITCODE%
    goto :error_exit
)

echo [OK] ジョブネットの起動に成功しました
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

rem ajsstatusで状態確認
set "STATUS_LINE="
for /f "delims=" %%i in ('ajsstatus -F %SCHEDULER_SERVICE% %JOBNET_PATH% 2^>^&1') do (
    set "STATUS_LINE=%%i"
)

rem 状態を判定
echo !STATUS_LINE! | findstr /i "ended abnormally abnormal end abend killed interrupted failed" >nul
if %ERRORLEVEL%==0 (
    set "JOB_STATUS=abnormal"
    goto :wait_done
)

echo !STATUS_LINE! | findstr /i "ended normally normal end completed" >nul
if %ERRORLEVEL%==0 (
    set "JOB_STATUS=normal"
    goto :wait_done
)

echo !STATUS_LINE! | findstr /i "now running running wait queued executing" >nul
if %ERRORLEVEL%==0 (
    set /a "MINUTES=ELAPSED_SECONDS/60"
    set /a "SECS=ELAPSED_SECONDS%%60"
    echo   状態: 実行中... ^(経過時間: !MINUTES!分!SECS!秒^)
    goto :wait_continue
)

echo !STATUS_LINE! | findstr /i "not registered not found does not exist" >nul
if %ERRORLEVEL%==0 (
    set "JOB_STATUS=not_found"
    goto :wait_done
)

rem その他の状態
set /a "MINUTES=ELAPSED_SECONDS/60"
set /a "SECS=ELAPSED_SECONDS%%60"
echo   状態: 確認中... ^(経過時間: !MINUTES!分!SECS!秒^) - !STATUS_LINE!

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

echo 詳細情報 (ajsshow -E):
echo ----------------------------------------------------------------
ajsshow -F %SCHEDULER_SERVICE% -E %JOBNET_PATH%
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
    ) else (
        echo   実行結果    : %JOB_STATUS%
    )
)
echo.

rem 終了判定
if "%JOB_STATUS%"=="normal" goto :normal_exit
if "%JOB_STATUS%"=="unknown" goto :normal_exit
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
