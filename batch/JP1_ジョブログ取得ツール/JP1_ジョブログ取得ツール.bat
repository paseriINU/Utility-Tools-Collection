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
rem    まずログファイルを直接読み取り、失敗時はjpqjobgetを使用します。
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
rem ジョブ情報の取得（ajsshow -i）
rem ========================================
echo ========================================
echo ジョブ情報を取得中...
echo ========================================
echo.

rem 一時ファイル作成
set TEMP_AJSSHOW=%TEMP%\jp1_ajsshow_%RANDOM%.txt

rem ajsshowコマンド実行（-g 1 -i で実行登録番号と標準エラーファイルパスを取得）
rem フォーマット: %ll %rr → 実行登録番号 標準エラーファイルパス
rem 公式ドキュメント: https://itpfdoc.hitachi.co.jp/manuals/3021/30213L4920/AJSO0131.HTM
echo 実行コマンド: ajsshow -F %SCHEDULER_SERVICE% -g 1 -i '%%ll %%rr' "%JOB_PATH%"
echo.

ajsshow -F %SCHEDULER_SERVICE% -g 1 -i '%%ll %%rr' "%JOB_PATH%" > "%TEMP_AJSSHOW%" 2>&1
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

rem 実行登録番号とログファイルパスを抽出
set EXEC_REG_NO=
set LOG_FILE_PATH=

rem 出力を解析（形式: 実行登録番号 ファイルパス）
for /f "tokens=1,*" %%A in ('type "%TEMP_AJSSHOW%"') do (
    if not defined EXEC_REG_NO (
        set EXEC_REG_NO=%%A
        set LOG_FILE_PATH=%%B
    )
)

del "%TEMP_AJSSHOW%" 2>nul

echo [情報] 実行登録番号: %EXEC_REG_NO%
echo [情報] ログファイル: %LOG_FILE_PATH%
echo.

if not defined EXEC_REG_NO (
    echo [エラー] 実行登録番号を取得できませんでした
    goto :ERROR_EXIT
)

rem ========================================
rem スプール取得（まずファイル直接読み取りを試行）
rem ========================================
echo ========================================
echo スプールを取得中...
echo ========================================
echo.

set SPOOL_FILE=%TEMP%\jp1_spool_%RANDOM%.txt
set LOG_CONTENT_FOUND=0

rem 方法1: ログファイルを直接読み取る
if defined LOG_FILE_PATH (
    echo [試行1] ログファイルを直接読み取り: %LOG_FILE_PATH%
    if exist "%LOG_FILE_PATH%" (
        copy "%LOG_FILE_PATH%" "%SPOOL_FILE%" >nul 2>&1
        if exist "%SPOOL_FILE%" (
            for %%F in ("%SPOOL_FILE%") do (
                if %%~zF GTR 0 (
                    echo [OK] ログファイルを直接読み取りました
                    set LOG_CONTENT_FOUND=1
                )
            )
        )
    ) else (
        echo [情報] ファイルが存在しません
    )
)

rem 方法2: ファイルが読めなかった場合、jpqjobgetを試行
if !LOG_CONTENT_FOUND!==0 (
    echo [試行2] jpqjobgetでスプールを取得...
    echo 実行コマンド: jpqjobget -F %SCHEDULER_SERVICE% -n %EXEC_REG_NO% -s "%JOB_PATH%"
    jpqjobget -F %SCHEDULER_SERVICE% -n %EXEC_REG_NO% -s "%JOB_PATH%" > "%SPOOL_FILE%" 2>&1
    set JPQJOBGET_EXITCODE=!ERRORLEVEL!
    echo jpqjobget終了コード: !JPQJOBGET_EXITCODE!

    rem エラーチェック
    findstr /i "無効 KAVS エラー" "%SPOOL_FILE%" >nul 2>&1
    if !ERRORLEVEL!==0 (
        echo.
        echo [警告] jpqjobgetでエラーが発生しました
        type "%SPOOL_FILE%"
        del "%SPOOL_FILE%" 2>nul
    ) else (
        for %%F in ("%SPOOL_FILE%") do (
            if %%~zF GTR 0 (
                echo [OK] jpqjobgetでスプールを取得しました
                set LOG_CONTENT_FOUND=1
            )
        )
    )
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
