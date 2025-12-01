@echo off
rem ====================================================================
rem JP1ジョブネット起動ツール（ajsentryコマンド版）
rem ====================================================================

setlocal enabledelayedexpansion

cls
echo ========================================
echo JP1ジョブネット起動ツール
echo （ajsentryコマンド版）
echo ========================================
echo.

rem ====================================================================
rem 設定項目（ここを編集してください）
rem ====================================================================

rem JP1/AJS3のホスト名またはIPアドレス
set JP1_HOST=192.168.1.100

rem JP1ユーザー名
set JP1_USER=jp1admin

rem JP1パスワード（空の場合は実行時に入力を求めます）
set JP1_PASSWORD=

rem 起動するジョブネットのフルパス
rem 例: /main_unit/jobgroup1/daily_batch
set JOBNET_PATH=/main_unit/jobgroup1/daily_batch

rem ajsentryコマンドのパス（環境変数PATHに含まれている場合は"ajsentry"でOK）
set AJSENTRY_CMD=ajsentry

rem ====================================================================
rem メイン処理
rem ====================================================================

echo JP1ホスト      : %JP1_HOST%
echo JP1ユーザー    : %JP1_USER%
echo ジョブネットパス: %JOBNET_PATH%
echo.

rem パスワードが設定されていない場合は入力を求める
if "%JP1_PASSWORD%"=="" (
    echo [注意] パスワードが設定されていません。
    set /p JP1_PASSWORD="JP1パスワードを入力してください: "
    echo.
)

rem ajsentryコマンドの存在確認
where %AJSENTRY_CMD% >nul 2>&1
if errorlevel 1 (
    echo [エラー] ajsentryコマンドが見つかりません。
    echo.
    echo 以下を確認してください：
    echo - JP1/AJS3 - View または JP1/AJS3 - Manager がインストールされているか
    echo - 環境変数PATHにajsentryのパスが含まれているか
    echo - AJSENTRY_CMDの設定が正しいか
    echo.
    pause
    exit /b 1
)

echo ジョブネットを起動しますか？
choice /M "実行する場合はYを押してください"
if errorlevel 2 (
    echo 処理をキャンセルしました。
    pause
    exit /b 0
)

echo.
echo ========================================
echo ジョブネット起動中...
echo ========================================
echo.

rem ajsentryコマンドを実行
rem -h: ホスト名
rem -u: ユーザー名
rem -p: パスワード
rem -F: ジョブネットのフルパス
%AJSENTRY_CMD% -h %JP1_HOST% -u %JP1_USER% -p %JP1_PASSWORD% -F "%JOBNET_PATH%"

if errorlevel 1 (
    echo.
    echo [エラー] ジョブネットの起動に失敗しました。
    echo.
    echo エラーコード: %errorlevel%
    echo.
    echo 以下を確認してください：
    echo - JP1ホスト名、ユーザー名、パスワードが正しいか
    echo - ジョブネットパスが正しいか
    echo - ネットワーク接続が正常か
    echo - JP1/AJS3サービスが起動しているか
    echo.
    pause
    exit /b 1
)

echo.
echo ========================================
echo ジョブネットの起動に成功しました
echo ========================================
echo.
echo ジョブネット: %JOBNET_PATH%
echo ホスト      : %JP1_HOST%
echo.

pause
endlocal
exit /b 0
