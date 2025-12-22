@echo off
chcp 932 >nul
title JP1 ジョブログ取得ツール（バッチ版）
setlocal enabledelayedexpansion

rem ==============================================================================
rem ■ JP1ジョブログ取得ツール（純粋バッチ版）
rem
rem ■ 説明
rem    JP1/AJS3の指定されたジョブの標準出力（スプール）を取得し、
rem    クリップボードにコピーします。
rem    jpqjobgetコマンドを使用してスプールを取得します。
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
echo   JP1 ジョブログ取得ツール（バッチ版）
echo ================================================================
echo.

echo 設定内容:
echo   スケジューラーサービス: %SCHEDULER_SERVICE%
echo   ジョブパス            : %JOB_PATH%
echo.

rem ========================================
rem 実行登録番号の取得（ajsshow -i）
rem ========================================
echo ========================================
echo 実行登録番号を取得中...
echo ========================================
echo.

rem 一時ファイル作成
set TEMP_AJSSHOW=%TEMP%\jp1_ajsshow_%RANDOM%.txt

rem ajsshowコマンド実行（-g 1 -i で実行登録番号を取得）
rem フォーマット: %ll（2バイト版）→ 実行登録番号を出力（jpqjobgetの-nオプションで使用）
rem 公式ドキュメント: https://itpfdoc.hitachi.co.jp/manuals/3021/30213L4920/AJSO0131.HTM
echo 実行コマンド: ajsshow -F %SCHEDULER_SERVICE% -g 1 -i '%%ll' "%JOB_PATH%"
echo.

ajsshow -F %SCHEDULER_SERVICE% -g 1 -i '%%ll' "%JOB_PATH%" > "%TEMP_AJSSHOW%" 2>&1
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

rem 実行登録番号を抽出（出力をそのまま取得）
set /p EXEC_REG_NO=<"%TEMP_AJSSHOW%"

del "%TEMP_AJSSHOW%" 2>nul

if not defined EXEC_REG_NO (
    echo [エラー] 実行登録番号を取得できませんでした
    echo.
    echo ajsshow -i の出力から実行登録番号を特定できませんでした。
    echo ジョブが実行されていることを確認してください。
    goto :ERROR_EXIT
)

echo [OK] 実行登録番号: %EXEC_REG_NO%
echo.

rem ========================================
rem スプール取得（jpqjobget -s: 標準出力）
rem ========================================
echo ========================================
echo スプールを取得中...
echo ========================================
echo.

set SPOOL_FILE=%TEMP%\jp1_spool_%RANDOM%.txt

echo 実行コマンド: jpqjobget -F %SCHEDULER_SERVICE% -n %EXEC_REG_NO% -s "%JOB_PATH%"
jpqjobget -F %SCHEDULER_SERVICE% -n %EXEC_REG_NO% -s "%JOB_PATH%" > "%SPOOL_FILE%" 2>&1
set JPQJOBGET_EXITCODE=%ERRORLEVEL%

echo jpqjobget終了コード: %JPQJOBGET_EXITCODE%

:SHOW_RESULT
echo.

if not exist "%SPOOL_FILE%" (
    echo ========================================
    echo [エラー] スプールを取得できませんでした
    echo ========================================
    echo.
    echo 以下を確認してください:
    echo   - ジョブパスが正しいか: %JOB_PATH%
    echo   - ジョブが実行済みか
    echo   - JP1ユーザーに権限があるか
    echo   - スプールが保存されているか
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
