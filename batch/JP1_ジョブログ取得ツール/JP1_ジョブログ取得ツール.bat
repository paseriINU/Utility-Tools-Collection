@echo off
chcp 932 >nul
title JP1 ジョブログ取得ツール
setlocal enabledelayedexpansion

rem ==============================================================================
rem ■ JP1ジョブログ取得ツール
rem
rem ■ 説明
rem    JP1/AJS3の指定されたジョブの標準出力（スプール）を取得し、
rem    クリップボードにコピーします。
rem    ajsshowで標準出力ファイルパスを取得し、直接読み取ります。
rem    ※ PCジョブ・UNIXジョブ用（QUEUEジョブには対応していません）
rem
rem ■ 使い方
rem    1. 下記の「設定セクション」を編集
rem    2. このファイルをダブルクリックで実行
rem    3. 取得したログがクリップボードにコピーされます
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

rem 一時ファイル作成
set TEMP_AJSSHOW=%TEMP%\jp1_ajsshow_%RANDOM%.txt

rem ajsshowコマンド実行（標準出力ファイルパスを取得）
rem フォーマット: %so → 標準出力ファイル名
rem 公式ドキュメント: https://itpfdoc.hitachi.co.jp/manuals/3021/30213L4920/AJSO0131.HTM
echo 実行コマンド: ajsshow -F %SCHEDULER_SERVICE% -g 1 -i '%%so' "%JOB_PATH%"
echo.

ajsshow -F %SCHEDULER_SERVICE% -g 1 -i '%%so' "%JOB_PATH%" > "%TEMP_AJSSHOW%" 2>&1
set AJSSHOW_EXITCODE=%ERRORLEVEL%

echo ajsshow結果:
type "%TEMP_AJSSHOW%"
echo.

if not %AJSSHOW_EXITCODE%==0 (
    echo [エラー] ジョブ情報の取得に失敗しました（終了コード: %AJSSHOW_EXITCODE%）
    echo.
    echo 以下を確認してください:
    echo   - ジョブパスが正しいか: %JOB_PATH%
    echo   - ジョブが実行済みか（少なくとも1回実行されている必要があります）
    echo   - JP1ユーザーに権限があるか
    echo   - スケジューラサービス名が正しいか: %SCHEDULER_SERVICE%
    del "%TEMP_AJSSHOW%" 2>nul
    goto :ERROR_EXIT
)

rem 標準出力ファイルパスを抽出（シングルクォートを除去）
set LOG_FILE_PATH=
for /f "usebackq delims=" %%A in ("%TEMP_AJSSHOW%") do (
    if not defined LOG_FILE_PATH set LOG_FILE_PATH=%%A
)
del "%TEMP_AJSSHOW%" 2>nul

rem シングルクォートを除去
set LOG_FILE_PATH=%LOG_FILE_PATH:'=%

echo [情報] 標準出力ファイル: %LOG_FILE_PATH%

if not defined LOG_FILE_PATH (
    echo [エラー] 標準出力ファイルパスを取得できませんでした
    goto :ERROR_EXIT
)

echo.

rem ========================================
rem スプール取得（ファイル直接読み取り）
rem ========================================
echo ========================================
echo スプールを取得中...
echo ========================================
echo.

set SPOOL_FILE=%TEMP%\jp1_spool_%RANDOM%.txt

echo ファイルを読み取り中: %LOG_FILE_PATH%

if exist "%LOG_FILE_PATH%" (
    copy "%LOG_FILE_PATH%" "%SPOOL_FILE%" >nul 2>&1
    if exist "%SPOOL_FILE%" (
        echo [OK] 標準出力ファイルを読み取りました
    ) else (
        echo [エラー] ファイルのコピーに失敗しました
        goto :ERROR_EXIT
    )
) else (
    echo [エラー] 標準出力ファイルが存在しません: %LOG_FILE_PATH%
    echo.
    echo 以下を確認してください:
    echo   - ジョブが実行済みか
    echo   - スプールが保存されているか（保存期間の設定を確認）
    goto :ERROR_EXIT
)

echo.

rem 結果確認
if not exist "%SPOOL_FILE%" (
    echo ========================================
    echo [エラー] スプールを取得できませんでした
    echo ========================================
    goto :ERROR_EXIT
)

rem ファイルサイズチェック
for %%F in ("%SPOOL_FILE%") do set FILE_SIZE=%%~zF
if "%FILE_SIZE%"=="0" (
    echo ========================================
    echo [情報] スプールは空です
    echo ========================================
    del "%SPOOL_FILE%" 2>nul
    goto :SUCCESS_EXIT
)

rem コンソールに出力
echo ========================================
echo 取得したスプール内容:
echo ========================================
echo.
type "%SPOOL_FILE%"
echo.

rem クリップボードにコピー
clip < "%SPOOL_FILE%"

echo ========================================
echo [OK] スプール内容をクリップボードにコピーしました
echo ========================================

rem 一時ファイル削除
del "%SPOOL_FILE%" 2>nul

:SUCCESS_EXIT
echo.
pause
exit /b 0

:ERROR_EXIT
echo.
pause
exit /b 1
