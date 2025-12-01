@echo off
rem ====================================================================
rem JP1ジョブネット起動ツール（設定ファイル版）
rem ====================================================================

setlocal enabledelayedexpansion

cls
echo ========================================
echo JP1ジョブネット起動ツール
echo （設定ファイル版）
echo ========================================
echo.

rem ====================================================================
rem 設定ファイルの読み込み
rem ====================================================================

set CONFIG_FILE=%~dp0config.ini

if not exist "%CONFIG_FILE%" (
    echo [エラー] 設定ファイルが見つかりません: %CONFIG_FILE%
    echo.
    echo config.ini を作成してください。サンプル：
    echo.
    echo [JP1]
    echo JP1_HOST=192.168.1.100
    echo JP1_USER=jp1admin
    echo JP1_PASSWORD=
    echo JOBNET_PATH=/main_unit/jobgroup1/daily_batch
    echo AJSENTRY_CMD=ajsentry
    echo.
    pause
    exit /b 1
)

rem 設定ファイルから値を読み込む
for /f "usebackq tokens=1,* delims==" %%a in ("%CONFIG_FILE%") do (
    set "%%a=%%b"
)

rem 必須項目のチェック
if not defined JP1_HOST (
    echo [エラー] JP1_HOSTが設定されていません。
    pause
    exit /b 1
)

if not defined JP1_USER (
    echo [エラー] JP1_USERが設定されていません。
    pause
    exit /b 1
)

if not defined JOBNET_PATH (
    echo [エラー] JOBNET_PATHが設定されていません。
    pause
    exit /b 1
)

if not defined AJSENTRY_CMD set AJSENTRY_CMD=ajsentry

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
    echo - config.iniのAJSENTRY_CMDの設定が正しいか
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
