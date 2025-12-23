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
rem 1: ajsentry -n -w で完了まで待機
rem 0: ajsentry -n で即時起動のみ（完了を待たない）
set "WAIT_FOR_COMPLETION=1"

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
    echo   完了待ち         : 有効（ajsentry -w）
) else (
    echo   完了待ち         : 無効（即時起動のみ）
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

rem ajsentry実行
rem -n: 即時実行登録
rem -w: 完了待ち（WAIT_FOR_COMPLETION=1の場合のみ）
if "%WAIT_FOR_COMPLETION%"=="1" (
    echo コマンド実行中: ajsentry -F %SCHEDULER_SERVICE% -n -w %JOBNET_PATH%
    echo.
    ajsentry -F %SCHEDULER_SERVICE% -n -w %JOBNET_PATH%
) else (
    echo コマンド実行中: ajsentry -F %SCHEDULER_SERVICE% -n %JOBNET_PATH%
    echo.
    ajsentry -F %SCHEDULER_SERVICE% -n %JOBNET_PATH%
)
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
    goto :error_exit
)
echo   実行登録番号: !EXEC_REG_NUM!
echo.

rem 結果判定
echo ================================================================
if "%WAIT_FOR_COMPLETION%"=="1" (
    rem ajsentry -w の終了コードでジョブネットの結果を判定
    if %AJSENTRY_EXITCODE% equ 0 (
        echo [OK] ジョブネットが正常終了しました
    ) else (
        echo [NG] ジョブネットが異常終了しました
        echo   終了コード: %AJSENTRY_EXITCODE%
    )
) else (
    rem 即時起動のみの場合はajsentryの起動成否を判定
    if %AJSENTRY_EXITCODE% equ 0 (
        echo [OK] ジョブネットの起動に成功しました
    ) else (
        echo [NG] ジョブネットの起動に失敗しました
        echo   終了コード: %AJSENTRY_EXITCODE%
    )
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
