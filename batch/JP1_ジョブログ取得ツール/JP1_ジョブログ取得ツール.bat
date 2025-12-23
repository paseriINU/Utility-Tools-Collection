@echo off
chcp 932 >nul
title JP1 ジョブログ取得ツール
setlocal enabledelayedexpansion

rem ==============================================================================
rem ■ JP1ジョブログ取得ツール
rem
rem ■ 説明
rem    JP1/AJS3の指定されたジョブの標準出力（スプール）を取得し、
rem    テキストファイルに出力します。
rem    ajsshowで標準出力ファイルパスを取得し、直接読み取ります。
rem    ※ PCジョブ・UNIXジョブ用（QUEUEジョブには対応していません）
rem
rem ■ 使い方
rem    1. 下記の「設定セクション」を編集
rem    2. このファイルをダブルクリックで実行
rem    3. 取得したログがテキストファイルに出力されます
rem
rem ■ 注意
rem    このファイルはShift-JIS（CP932）で保存してください
rem ==============================================================================

rem ==============================================================================
rem ■ 設定セクション（ここを編集してください）
rem ==============================================================================

rem スケジューラーサービス名（デフォルト: AJSROOT1）
set SCHEDULER_SERVICE=AJSROOT1

rem 取得対象のジョブのフルパス（ジョブネット内のジョブを指定）
rem 例: /main_unit/jobgroup1/daily_batch/job1
set JOB_PATH=/main_unit/jobgroup1/daily_batch/job1

rem ログ出力先フォルダ（バッチファイルと同じ場所に出力する場合は %~dp0 のまま）
set OUTPUT_DIR=%~dp0

rem ==============================================================================
rem ■ メイン処理（以下は編集不要）
rem ==============================================================================

echo.
echo ================================================================
echo   JP1 ジョブログ取得ツール
echo ================================================================
echo.

echo 設定内容:
echo   スケジューラーサービス: %SCHEDULER_SERVICE%
echo   ジョブパス            : %JOB_PATH%
echo.

rem ========================================
rem 標準出力ファイルパスの取得（ajsshow -i %so）
rem ========================================
echo ========================================
echo 標準出力ファイルパスを取得中...
echo ========================================
echo.

rem ajsshowコマンド実行（標準出力ファイルパスを取得）
rem フォーマット: %so → 標準出力ファイル名
rem 公式ドキュメント: https://itpfdoc.hitachi.co.jp/manuals/3021/30213L4920/AJSO0131.HTM
echo 実行コマンド: ajsshow -F %SCHEDULER_SERVICE% -g 1 -i '%%so' "%JOB_PATH%"
echo.

rem 標準出力ファイルパスを直接取得（一時ファイル不要）
set LOG_FILE_PATH=
for /f "delims=" %%A in ('ajsshow -F %SCHEDULER_SERVICE% -g 1 -i "%%so" "%JOB_PATH%" 2^>^&1') do (
    if not defined LOG_FILE_PATH set LOG_FILE_PATH=%%A
)

echo ajsshow結果: !LOG_FILE_PATH!
echo.

if not defined LOG_FILE_PATH (
    echo [エラー] ジョブ情報の取得に失敗しました
    echo.
    echo 以下を確認してください:
    echo   - ジョブパスが正しいか: %JOB_PATH%
    echo   - ジョブが実行済みか（少なくとも1回実行されている必要があります）
    echo   - JP1ユーザーに権限があるか
    echo   - スケジューラサービス名が正しいか: %SCHEDULER_SERVICE%
    goto :ERROR_EXIT
)

rem シングルクォートを除去
set LOG_FILE_PATH=%LOG_FILE_PATH:'=%

echo [情報] 標準出力ファイル: %LOG_FILE_PATH%

if not defined LOG_FILE_PATH (
    echo [エラー] 標準出力ファイルパスを取得できませんでした
    goto :ERROR_EXIT
)

echo.

rem ========================================
rem スプール取得・ファイル出力
rem ========================================
echo ========================================
echo スプールを取得中...
echo ========================================
echo.

rem ファイル存在チェック
if not exist "%LOG_FILE_PATH%" (
    echo [エラー] 標準出力ファイルが存在しません: %LOG_FILE_PATH%
    echo.
    echo 以下を確認してください:
    echo   - ジョブが実行済みか
    echo   - スプールが保存されているか（保存期間の設定を確認）
    goto :ERROR_EXIT
)

rem ファイルサイズチェック
for %%F in ("%LOG_FILE_PATH%") do set FILE_SIZE=%%~zF
if "%FILE_SIZE%"=="0" (
    echo ========================================
    echo [情報] スプールは空です
    echo ========================================
    goto :SUCCESS_EXIT
)

rem コンソールに出力
echo ========================================
echo 取得したスプール内容:
echo ========================================
echo.
type "%LOG_FILE_PATH%"
echo.

rem ファイルに出力
set OUTPUT_FILE=%OUTPUT_DIR%joblog.txt
copy "%LOG_FILE_PATH%" "%OUTPUT_FILE%" >nul

echo ========================================
echo [OK] スプール内容をファイルに出力しました
echo   出力先: %OUTPUT_FILE%
echo ========================================

:SUCCESS_EXIT
echo.
pause
exit /b 0

:ERROR_EXIT
echo.
pause
exit /b 1
