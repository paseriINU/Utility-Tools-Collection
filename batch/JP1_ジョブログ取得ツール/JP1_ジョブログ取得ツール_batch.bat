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

rem JP1ユーザー名（空の場合は現在のログインユーザーで実行）
set JP1_USER=

rem JP1パスワード（空の場合は実行時に入力、JP1_USERが空の場合は不要）
set JP1_PASSWORD=

rem 取得するスプールの種類（stdout=標準出力、stderr=標準エラー出力、both=両方）
set SPOOL_TYPE=stdout

rem ==============================================================================
rem ■ メイン処理（以下は編集不要）
rem ==============================================================================

echo.
echo ================================================================
echo   JP1 ジョブログ取得ツール（バッチ版）
echo ================================================================
echo.

rem コマンドパス検索
set AJSSHOW_PATH=
set JPQJOBGET_PATH=

rem 検索パスリスト
for %%P in (
    "C:\Program Files (x86)\Hitachi\JP1AJS2\bin"
    "C:\Program Files\Hitachi\JP1AJS2\bin"
    "C:\Program Files (x86)\HITACHI\JP1AJS3\bin"
    "C:\Program Files\HITACHI\JP1AJS3\bin"
) do (
    if exist "%%~P\ajsshow.exe" if not defined AJSSHOW_PATH set "AJSSHOW_PATH=%%~P\ajsshow.exe"
    if exist "%%~P\jpqjobget.exe" if not defined JPQJOBGET_PATH set "JPQJOBGET_PATH=%%~P\jpqjobget.exe"
)

if not defined AJSSHOW_PATH (
    echo [エラー] ajsshowコマンドが見つかりません
    echo.
    echo 以下のパスを確認してください:
    echo   - C:\Program Files ^(x86^)\Hitachi\JP1AJS2\bin\ajsshow.exe
    echo   - C:\Program Files\Hitachi\JP1AJS2\bin\ajsshow.exe
    echo   - C:\Program Files ^(x86^)\HITACHI\JP1AJS3\bin\ajsshow.exe
    echo   - C:\Program Files\HITACHI\JP1AJS3\bin\ajsshow.exe
    goto :ERROR_EXIT
)

if not defined JPQJOBGET_PATH (
    echo [エラー] jpqjobgetコマンドが見つかりません
    echo.
    echo 以下のパスを確認してください:
    echo   - C:\Program Files ^(x86^)\Hitachi\JP1AJS2\bin\jpqjobget.exe
    echo   - C:\Program Files\Hitachi\JP1AJS2\bin\jpqjobget.exe
    echo   - C:\Program Files ^(x86^)\HITACHI\JP1AJS3\bin\jpqjobget.exe
    echo   - C:\Program Files\HITACHI\JP1AJS3\bin\jpqjobget.exe
    goto :ERROR_EXIT
)

echo コマンドパス:
echo   ajsshow  : %AJSSHOW_PATH%
echo   jpqjobget: %JPQJOBGET_PATH%
echo.

echo 設定内容:
echo   スケジューラーサービス: %SCHEDULER_SERVICE%
echo   ジョブパス            : %JOB_PATH%
echo   スプール種類          : %SPOOL_TYPE%
echo.

rem JP1認証パラメータの構築
set AUTH_PARAMS=
if not "%JP1_USER%"=="" (
    if "%JP1_PASSWORD%"=="" (
        echo [注意] JP1パスワードが設定されていません。
        set /p JP1_PASSWORD=JP1パスワードを入力してください:
        echo.
    )
    set AUTH_PARAMS=-u %JP1_USER% -p %JP1_PASSWORD%
)

rem ========================================
rem ジョブ情報の取得（ajsshow）
rem ========================================
echo ========================================
echo ジョブ情報を取得中...
echo ========================================
echo.

rem 一時ファイル作成
set TEMP_AJSSHOW=%TEMP%\jp1_ajsshow_%RANDOM%.txt

rem ajsshowコマンド実行
rem %## = ジョブ名、%I = ジョブ番号
set AJSSHOW_CMD="%AJSSHOW_PATH%" %AUTH_PARAMS% -F %SCHEDULER_SERVICE% -g 1 -i "%%## %%I" "%JOB_PATH%"
echo 実行コマンド: %AJSSHOW_CMD%

%AJSSHOW_CMD% > "%TEMP_AJSSHOW%" 2>&1
set AJSSHOW_EXITCODE=%ERRORLEVEL%

echo ajsshow結果:
type "%TEMP_AJSSHOW%"
echo.

if not %AJSSHOW_EXITCODE%==0 (
    echo [エラー] ジョブ情報の取得に失敗しました（終了コード: %AJSSHOW_EXITCODE%）
    echo.
    echo 以下を確認してください:
    echo   - ジョブパスが正しいか: %JOB_PATH%
    echo   - ジョブが実行済みか
    echo   - JP1ユーザーに権限があるか
    del "%TEMP_AJSSHOW%" 2>nul
    goto :ERROR_EXIT
)

rem ジョブ番号を抽出（最後の数字を取得）
set JOB_NO=
for /f "tokens=*" %%L in ('type "%TEMP_AJSSHOW%"') do (
    set LINE=%%L
    rem 行末の数字を抽出
    for %%N in (!LINE!) do (
        echo %%N | findstr /r "^[0-9][0-9]*$" >nul
        if !ERRORLEVEL!==0 set JOB_NO=%%N
    )
)

del "%TEMP_AJSSHOW%" 2>nul

if not defined JOB_NO (
    echo [エラー] ジョブ番号を取得できませんでした
    goto :ERROR_EXIT
)

echo [OK] ジョブ番号: %JOB_NO%
echo.

rem ========================================
rem スプール取得（jpqjobget）
rem ========================================
echo ========================================
echo スプールを取得中...
echo ========================================
echo.

set SPOOL_CONTENT=
set SPOOL_FILE=%TEMP%\jp1_spool_%RANDOM%.txt

rem スプール種類に応じてオプションを設定
if /i "%SPOOL_TYPE%"=="stdout" (
    set SPOOL_OPTIONS=-oso
) else if /i "%SPOOL_TYPE%"=="stderr" (
    set SPOOL_OPTIONS=-ose
) else if /i "%SPOOL_TYPE%"=="both" (
    set SPOOL_OPTIONS=-oso -ose
) else (
    set SPOOL_OPTIONS=-oso
)

rem stdoutの取得
if /i "%SPOOL_TYPE%"=="stdout" goto :GET_STDOUT
if /i "%SPOOL_TYPE%"=="both" goto :GET_BOTH
if /i "%SPOOL_TYPE%"=="stderr" goto :GET_STDERR
goto :GET_STDOUT

:GET_STDOUT
echo   取得中: stdout ...
set STDOUT_FILE=%TEMP%\jp1_stdout_%RANDOM%.txt
"%JPQJOBGET_PATH%" -j %JOB_NO% -oso "%STDOUT_FILE%" >nul 2>&1
set JPQJOBGET_EXITCODE=%ERRORLEVEL%

if %JPQJOBGET_EXITCODE%==0 if exist "%STDOUT_FILE%" (
    echo   [OK] stdout を取得しました
    set SPOOL_FILE=%STDOUT_FILE%
) else (
    echo   [警告] stdout の取得に失敗しました（終了コード: %JPQJOBGET_EXITCODE%）
)
goto :SHOW_RESULT

:GET_STDERR
echo   取得中: stderr ...
set STDERR_FILE=%TEMP%\jp1_stderr_%RANDOM%.txt
"%JPQJOBGET_PATH%" -j %JOB_NO% -ose "%STDERR_FILE%" >nul 2>&1
set JPQJOBGET_EXITCODE=%ERRORLEVEL%

if %JPQJOBGET_EXITCODE%==0 if exist "%STDERR_FILE%" (
    echo   [OK] stderr を取得しました
    set SPOOL_FILE=%STDERR_FILE%
) else (
    echo   [警告] stderr の取得に失敗しました（終了コード: %JPQJOBGET_EXITCODE%）
)
goto :SHOW_RESULT

:GET_BOTH
set COMBINED_FILE=%TEMP%\jp1_combined_%RANDOM%.txt

echo   取得中: stderr ...
set STDERR_FILE=%TEMP%\jp1_stderr_%RANDOM%.txt
"%JPQJOBGET_PATH%" -j %JOB_NO% -ose "%STDERR_FILE%" >nul 2>&1
if %ERRORLEVEL%==0 if exist "%STDERR_FILE%" (
    echo   [OK] stderr を取得しました
    echo ===== STDERR ===== > "%COMBINED_FILE%"
    type "%STDERR_FILE%" >> "%COMBINED_FILE%"
    echo. >> "%COMBINED_FILE%"
    del "%STDERR_FILE%" 2>nul
) else (
    echo   [情報] stderr は空です
)

echo   取得中: stdout ...
set STDOUT_FILE=%TEMP%\jp1_stdout_%RANDOM%.txt
"%JPQJOBGET_PATH%" -j %JOB_NO% -oso "%STDOUT_FILE%" >nul 2>&1
if %ERRORLEVEL%==0 if exist "%STDOUT_FILE%" (
    echo   [OK] stdout を取得しました
    echo ===== STDOUT ===== >> "%COMBINED_FILE%"
    type "%STDOUT_FILE%" >> "%COMBINED_FILE%"
    del "%STDOUT_FILE%" 2>nul
) else (
    echo   [情報] stdout は空です
)

set SPOOL_FILE=%COMBINED_FILE%
goto :SHOW_RESULT

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
