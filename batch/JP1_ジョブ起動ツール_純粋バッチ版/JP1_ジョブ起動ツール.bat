@echo off
rem ============================================================================
rem JP1 ジョブ起動ツール（純粋バッチ版）
rem
rem 概要:
rem   JP1/AJS3のジョブネットをローカルで起動するバッチファイルです。
rem   PowerShellを使用せず、純粋なバッチコマンドのみで動作します。
rem
rem 使い方:
rem   1. 下記の「設定セクション」を編集
rem   2. JP1/AJS3がインストールされているサーバでこのファイルを実行
rem
rem 作成日: 2025-12-23
rem ============================================================================

chcp 65001 >nul
title JP1 ジョブ起動ツール（純粋バッチ版）
setlocal enabledelayedexpansion

rem ============================================================================
rem ■ 設定セクション（ここを編集してください）
rem ============================================================================

rem JP1/AJS3コマンドのパス（インストールディレクトリ）
set "JP1_BIN=C:\Program Files\HITACHI\JP1AJS3\bin"

rem スケジューラーサービス名（通常は AJSROOT1）
set "SCHEDULER_SERVICE=AJSROOT1"

rem JP1ユーザー名（空の場合は実行時に入力）
set "JP1_USER=jp1admin"

rem JP1パスワード（空の場合は実行時に入力、セキュリティ上空欄推奨）
set "JP1_PASSWORD="

rem 起動するジョブネットのフルパス
set "JOBNET_PATH=/main_unit/jobgroup1/daily_batch"

rem ジョブ完了を待つ場合は 1、起動のみの場合は 0
set "WAIT_FOR_COMPLETION=1"

rem 完了待ちのタイムアウト（秒）。0の場合は無制限
set "WAIT_TIMEOUT=3600"

rem 状態確認の間隔（秒）
set "POLLING_INTERVAL=10"

rem ホスト名（ローカル実行の場合は localhost または空欄）
set "JP1_HOST=localhost"

rem ============================================================================
rem ■ メイン処理（以下は編集不要）
rem ============================================================================

echo.
echo ================================================================
echo   JP1 ジョブ起動ツール（純粋バッチ版）
echo ================================================================
echo.

rem コマンドパスの確認
if not exist "%JP1_BIN%\ajsentry.exe" (
    echo [エラー] ajsentry.exe が見つかりません
    echo   パス: %JP1_BIN%\ajsentry.exe
    echo.
    echo JP1/AJS3のインストールディレクトリを確認してください。
    echo.
    goto :error_exit
)

rem 設定情報の表示
echo 設定情報:
echo   JP1コマンドパス  : %JP1_BIN%
echo   スケジューラー   : %SCHEDULER_SERVICE%
echo   JP1ユーザー      : %JP1_USER%
echo   ジョブネットパス : %JOBNET_PATH%
echo   ホスト           : %JP1_HOST%
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

rem JP1ユーザー名の入力（空の場合）
if "%JP1_USER%"=="" (
    set /p "JP1_USER=JP1ユーザー名を入力してください: "
    if "!JP1_USER!"=="" (
        echo [エラー] JP1ユーザー名が入力されていません
        goto :error_exit
    )
)

rem JP1パスワードの入力（空の場合）
if "%JP1_PASSWORD%"=="" (
    echo [注意] JP1パスワードが設定されていません。
    set /p "JP1_PASSWORD=JP1パスワードを入力してください: "
    if "!JP1_PASSWORD!"=="" (
        echo [エラー] JP1パスワードが入力されていません
        goto :error_exit
    )
    echo.
)

rem 実行確認
echo ジョブネットを起動しますか？ (y/n)
set /p "CONFIRM="
if /i not "%CONFIRM%"=="y" (
    echo 処理をキャンセルしました。
    goto :normal_exit
)
echo.

rem ============================================================================
rem ajsentry実行（ジョブネット起動）
rem ============================================================================
echo ================================================================
echo ジョブネット起動中...
echo ================================================================
echo.

rem ajsentryコマンドの構築
set "AJSENTRY_CMD="%JP1_BIN%\ajsentry.exe""
if not "%JP1_HOST%"=="" set "AJSENTRY_CMD=%AJSENTRY_CMD% -h %JP1_HOST%"
set "AJSENTRY_CMD=%AJSENTRY_CMD% -u %JP1_USER% -p %JP1_PASSWORD% -F %SCHEDULER_SERVICE% %JOBNET_PATH%"

rem 一時ファイルの準備
set "TEMP_OUTPUT=%TEMP%\jp1_output_%RANDOM%.txt"

rem ajsentry実行
echo コマンド実行中: ajsentry ...
call %AJSENTRY_CMD% > "%TEMP_OUTPUT%" 2>&1
set "AJSENTRY_EXITCODE=%ERRORLEVEL%"

rem 結果表示
echo.
echo ----------------------------------------------------------------
echo ajsentry出力:
type "%TEMP_OUTPUT%"
echo ----------------------------------------------------------------
echo.

if %AJSENTRY_EXITCODE% neq 0 (
    echo [エラー] ジョブネットの起動に失敗しました
    echo   終了コード: %AJSENTRY_EXITCODE%
    del "%TEMP_OUTPUT%" 2>nul
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

set "START_TIME=%TIME%"
set "ELAPSED_SECONDS=0"
set "JOB_FINISHED=0"
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
set "AJSSTATUS_CMD="%JP1_BIN%\ajsstatus.exe""
if not "%JP1_HOST%"=="" set "AJSSTATUS_CMD=%AJSSTATUS_CMD% -h %JP1_HOST%"
set "AJSSTATUS_CMD=%AJSSTATUS_CMD% -u %JP1_USER% -p %JP1_PASSWORD% -F %SCHEDULER_SERVICE% %JOBNET_PATH%"

call %AJSSTATUS_CMD% > "%TEMP_OUTPUT%" 2>&1
set "AJSSTATUS_EXITCODE=%ERRORLEVEL%"

rem 状態判定
set "STATUS_LINE="
for /f "delims=" %%i in ('type "%TEMP_OUTPUT%"') do (
    set "STATUS_LINE=%%i"
)

rem 状態を小文字に変換して判定（簡易版）
echo !STATUS_LINE! | findstr /i "ended abnormally abnormal end abend killed interrupted failed" >nul
if %ERRORLEVEL%==0 (
    set "JOB_STATUS=abnormal"
    set "JOB_FINISHED=1"
    goto :wait_done
)

echo !STATUS_LINE! | findstr /i "ended normally normal end completed" >nul
if %ERRORLEVEL%==0 (
    set "JOB_STATUS=normal"
    set "JOB_FINISHED=1"
    goto :wait_done
)

echo !STATUS_LINE! | findstr /i "now running running wait queued executing" >nul
if %ERRORLEVEL%==0 (
    rem 実行中 - 継続して待機
    set /a "MINUTES=ELAPSED_SECONDS/60"
    set /a "SECS=ELAPSED_SECONDS%%60"
    echo   状態: 実行中... ^(経過時間: !MINUTES!分!SECS!秒^)
    goto :wait_continue
)

echo !STATUS_LINE! | findstr /i "not registered not found does not exist" >nul
if %ERRORLEVEL%==0 (
    set "JOB_STATUS=not_found"
    set "JOB_FINISHED=1"
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

if not exist "%JP1_BIN%\ajsshow.exe" (
    echo [情報] ajsshowコマンドが見つかりません
    echo   パス: %JP1_BIN%\ajsshow.exe
    goto :summary
)

set "AJSSHOW_CMD="%JP1_BIN%\ajsshow.exe""
if not "%JP1_HOST%"=="" set "AJSSHOW_CMD=%AJSSHOW_CMD% -h %JP1_HOST%"
set "AJSSHOW_CMD=%AJSSHOW_CMD% -u %JP1_USER% -p %JP1_PASSWORD% -F %SCHEDULER_SERVICE% -E %JOBNET_PATH%"

call %AJSSHOW_CMD% > "%TEMP_OUTPUT%" 2>&1

echo 詳細情報 (ajsshow -E):
echo ----------------------------------------------------------------
type "%TEMP_OUTPUT%"
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
echo   サーバ      : %JP1_HOST%
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

rem 一時ファイルの削除
del "%TEMP_OUTPUT%" 2>nul

rem 終了判定
if "%JOB_STATUS%"=="normal" goto :normal_exit
if "%JOB_STATUS%"=="unknown" goto :normal_exit
if "%WAIT_FOR_COMPLETION%"=="0" goto :normal_exit
goto :error_exit

:normal_exit
echo 処理が完了しました。
echo.
del "%TEMP_OUTPUT%" 2>nul
pause
exit /b 0

:error_exit
echo.
echo 追加の確認事項:
echo   - JP1/AJS3サービスが起動しているか
echo   - JP1ユーザー名、パスワードが正しいか
echo   - ジョブネットパスが正しいか
echo   - ajsentryコマンドのパスが正しいか
echo.
del "%TEMP_OUTPUT%" 2>nul
pause
exit /b 1
