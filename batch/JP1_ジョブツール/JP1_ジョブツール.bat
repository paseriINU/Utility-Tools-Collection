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

rem JP1/AJS3コマンドのパス（インストールディレクトリ）
set "JP1_BIN=C:\Program Files (x86)\HITACHI\JP1AJS3\bin"

rem スケジューラーサービス名（通常は AJSROOT1）
set "SCHEDULER_SERVICE=AJSROOT1"

rem JP1ユーザー名（空の場合は実行時に入力）
set "JP1_USER=jp1admin"

rem JP1パスワード（空の場合は実行時に入力、セキュリティ上空欄推奨）
set "JP1_PASSWORD="

rem 起動するジョブネットのフルパス
set "JOBNET_PATH=/main_unit/jobgroup1/daily_batch"

rem ログ取得対象のジョブのフルパス（ジョブネット内のジョブを指定）
rem ※ジョブネットではなく、その中の個別ジョブを指定
set "JOB_PATH=/main_unit/jobgroup1/daily_batch/job1"

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
echo   JP1 ジョブツール
echo ================================================================
echo.

rem コマンドパスの確認
if not exist "%JP1_BIN%\ajsentry.exe" (
    echo [エラー] ajsentry.exe が見つかりません
    echo   パス: %JP1_BIN%\ajsentry.exe
    goto :ERROR_EXIT
)
if not exist "%JP1_BIN%\ajsshow.exe" (
    echo [エラー] ajsshow.exe が見つかりません
    echo   パス: %JP1_BIN%\ajsshow.exe
    goto :ERROR_EXIT
)

echo 設定情報:
echo   JP1コマンドパス  : %JP1_BIN%
echo   スケジューラー   : %SCHEDULER_SERVICE%
echo   JP1ユーザー      : %JP1_USER%
echo   ジョブネットパス : %JOBNET_PATH%
echo   ジョブパス       : %JOB_PATH%
echo   タイムアウト     : %WAIT_TIMEOUT%秒
echo.

rem 認証情報の確認
if "%JP1_USER%"=="" (
    set /p "JP1_USER=JP1ユーザー名を入力してください: "
    if "!JP1_USER!"=="" (
        echo [エラー] JP1ユーザー名が入力されていません
        goto :ERROR_EXIT
    )
)

if "%JP1_PASSWORD%"=="" (
    echo [注意] JP1パスワードが設定されていません。
    set /p "JP1_PASSWORD=JP1パスワードを入力してください: "
    if "!JP1_PASSWORD!"=="" (
        echo [エラー] JP1パスワードが入力されていません
        goto :ERROR_EXIT
    )
    echo.
)

rem 実行確認
echo ジョブネットを起動し、完了後にログを取得しますか？ (y/n)
set /p "CONFIRM="
if /i not "%CONFIRM%"=="y" (
    echo 処理をキャンセルしました。
    goto :NORMAL_EXIT
)

echo.

rem ============================================================================
rem ジョブネット起動
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

rem ajsstatusで状態確認
set "AJSSTATUS_CMD="%JP1_BIN%\ajsstatus.exe""
if not "%JP1_HOST%"=="" set "AJSSTATUS_CMD=%AJSSTATUS_CMD% -h %JP1_HOST%"
set "AJSSTATUS_CMD=%AJSSTATUS_CMD% -u %JP1_USER% -p %JP1_PASSWORD% -F %SCHEDULER_SERVICE% %JOBNET_PATH%"

call %AJSSTATUS_CMD% > "%TEMP_OUTPUT%" 2>&1

rem 状態判定
set "STATUS_LINE="
for /f "delims=" %%i in ('type "%TEMP_OUTPUT%"') do (
    set "STATUS_LINE=%%i"
)

echo !STATUS_LINE! | findstr /i "ended abnormally abnormal end abend killed interrupted failed" >nul
if %ERRORLEVEL%==0 (
    set "JOB_STATUS=abnormal"
    goto :WAIT_DONE
)

echo !STATUS_LINE! | findstr /i "ended normally normal end completed" >nul
if %ERRORLEVEL%==0 (
    set "JOB_STATUS=normal"
    goto :WAIT_DONE
)

echo !STATUS_LINE! | findstr /i "now running running wait queued executing" >nul
if %ERRORLEVEL%==0 (
    set /a "MINUTES=ELAPSED_SECONDS/60"
    set /a "SECS=ELAPSED_SECONDS%%60"
    echo   状態: 実行中... ^(経過時間: !MINUTES!分!SECS!秒^)
    goto :WAIT_CONTINUE
)

echo !STATUS_LINE! | findstr /i "not registered not found does not exist" >nul
if %ERRORLEVEL%==0 (
    set "JOB_STATUS=not_found"
    goto :WAIT_DONE
)

set /a "MINUTES=ELAPSED_SECONDS/60"
set /a "SECS=ELAPSED_SECONDS%%60"
echo   状態: 確認中... ^(経過時間: !MINUTES!分!SECS!秒^)

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
    del "%TEMP_OUTPUT%" 2>nul
    goto :ERROR_EXIT
) else if "%JOB_STATUS%"=="timeout" (
    echo [NG] 完了待ちがタイムアウトしました
    del "%TEMP_OUTPUT%" 2>nul
    goto :ERROR_EXIT
) else if "%JOB_STATUS%"=="not_found" (
    echo [NG] ジョブネットが見つかりません
    del "%TEMP_OUTPUT%" 2>nul
    goto :ERROR_EXIT
) else (
    echo [--] ジョブネット状態: %JOB_STATUS%
)
echo ================================================================
echo.

del "%TEMP_OUTPUT%" 2>nul

rem ============================================================================
rem ジョブログ取得
rem ============================================================================
echo ================================================================
echo ジョブログを取得中...
echo ================================================================
echo.

rem 一時ファイル作成
set "TEMP_AJSSHOW=%TEMP%\jp1_ajsshow_%RANDOM%.txt"

rem ajsshowコマンド実行（標準出力ファイルパスを取得）
echo 実行コマンド: ajsshow -F %SCHEDULER_SERVICE% -g 1 -i '%%so' "%JOB_PATH%"
echo.

"%JP1_BIN%\ajsshow.exe" -F %SCHEDULER_SERVICE% -g 1 -i '%%so' "%JOB_PATH%" > "%TEMP_AJSSHOW%" 2>&1
set "AJSSHOW_EXITCODE=%ERRORLEVEL%"

echo ajsshow結果:
type "%TEMP_AJSSHOW%"
echo.

if not %AJSSHOW_EXITCODE%==0 (
    echo [エラー] ジョブ情報の取得に失敗しました（終了コード: %AJSSHOW_EXITCODE%）
    echo.
    echo 以下を確認してください:
    echo   - ジョブパスが正しいか: %JOB_PATH%
    echo   - ジョブが実行済みか
    echo   - ジョブネットではなくジョブを指定しているか
    del "%TEMP_AJSSHOW%" 2>nul
    goto :ERROR_EXIT
)

rem 標準出力ファイルパスを抽出（シングルクォートを除去）
set "LOG_FILE_PATH="
for /f "usebackq delims=" %%A in ("%TEMP_AJSSHOW%") do (
    if not defined LOG_FILE_PATH set "LOG_FILE_PATH=%%A"
)
del "%TEMP_AJSSHOW%" 2>nul

rem シングルクォートを除去
set "LOG_FILE_PATH=%LOG_FILE_PATH:'=%"

echo [情報] 標準出力ファイル: %LOG_FILE_PATH%

if not defined LOG_FILE_PATH (
    echo [エラー] 標準出力ファイルパスを取得できませんでした
    goto :ERROR_EXIT
)

echo.

rem ファイル存在チェック
if not exist "%LOG_FILE_PATH%" (
    echo [エラー] 標準出力ファイルが存在しません: %LOG_FILE_PATH%
    goto :ERROR_EXIT
)

rem ファイルサイズチェック
for %%F in ("%LOG_FILE_PATH%") do set "FILE_SIZE=%%~zF"
if "%FILE_SIZE%"=="0" (
    echo [情報] スプールは空です
    goto :NORMAL_EXIT
)

rem コンソールに出力
echo ================================================================
echo 取得したスプール内容:
echo ================================================================
echo.
type "%LOG_FILE_PATH%"
echo.

rem クリップボードにコピー
type "%LOG_FILE_PATH%" | clip

echo ================================================================
echo [OK] スプール内容をクリップボードにコピーしました
echo ================================================================
echo.

rem ============================================================================
rem サマリー表示
rem ============================================================================
echo ================================================================
echo 処理サマリー
echo ================================================================
echo.
echo   ジョブネット : %JOBNET_PATH%
echo   ジョブ       : %JOB_PATH%
echo   起動結果     : 成功
echo   実行結果     : 正常終了
echo   ログ取得     : 成功（クリップボードにコピー済み）
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
echo   - JP1ユーザー名、パスワードが正しいか
echo   - ジョブネットパス、ジョブパスが正しいか
echo.
pause
exit /b 1
