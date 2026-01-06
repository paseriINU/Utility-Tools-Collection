@echo off
title JP1 ジョブログ取得サンプル
setlocal enabledelayedexpansion

rem UNCパス対応（PushD/PopDで自動マッピング）
pushd "%~dp0"

rem ============================================================================
rem JP1 ジョブログ取得サンプル
rem
rem 説明:
rem   JP1_REST_ジョブ情報取得ツール.bat を呼び出し、
rem   取得したログをファイルに保存します。
rem
rem 使い方:
rem   1. 下記の UNIT_PATH を取得したいジョブのパスに変更
rem      例: /JobGroup/Jobnet/Job1
rem   2. このバッチをダブルクリックで実行
rem
rem 終了コード（実行順）:
rem   0: 正常終了
rem   1: 引数エラー
rem   2: ユニット未検出（STEP 1）
rem   3: ユニット種別エラー（STEP 1）
rem   4: 実行世代なし（STEP 2）
rem   5: 5MB超過エラー（STEP 3）
rem   6: 詳細取得エラー（STEP 3）
rem   7: ジョブ開始日時取得エラー（START_TIMEが空）
rem   9: API接続エラー（各STEP）
rem ============================================================================

rem === ここを編集してください（ジョブのパスを指定）===
set "UNIT_PATH=/JobGroup/Jobnet/Job1"

rem === スクロール位置の設定 ===
rem メモ帳で開いた後、指定した文字列を検索してその位置にスクロールします
rem 空欄の場合はスクロールせずにファイル先頭を表示します
rem 例: "エラー", "ERROR", "RC=", "異常終了" など
set "SCROLL_TO_TEXT="

rem === ファイル名の設定 ===
rem ファイル名形式: 【ジョブ実行結果】【{日時}実行分】{ジョブネット名}_{コメント}.txt
rem 例: 【ジョブ実行結果】【20251010_163520実行分】Jobnet_テスト.txt
rem
rem ジョブネットコメントが取得できない場合に使用するデフォルト値
set "DEFAULT_COMMENT=NoComment"

rem 一時ファイル名（後でジョブ開始日時を使用してリネーム）
set "TEMP_FILE=%~dp0temp_output.txt"

rem スクリプトのディレクトリを取得
set "SCRIPT_DIR=%~dp0"

echo.
echo ================================================================
echo   JP1 ジョブログ取得サンプル
echo ================================================================
echo.
echo   対象: %UNIT_PATH%
echo.
echo ログを取得中...

rem JP1_REST_ジョブ情報取得ツール.bat を呼び出し、結果を一時ファイルに保存
call "%SCRIPT_DIR%JP1_REST_ジョブ情報取得ツール.bat" "%UNIT_PATH%" > "%TEMP_FILE%" 2>&1
set "EXIT_CODE=%ERRORLEVEL%"

rem エラーコード別のハンドリング（実行順）
if %EXIT_CODE% equ 0 goto :SUCCESS
if %EXIT_CODE% equ 1 goto :ERR_ARGUMENT
if %EXIT_CODE% equ 2 goto :ERR_UNIT_NOT_FOUND
if %EXIT_CODE% equ 3 goto :ERR_UNIT_TYPE
if %EXIT_CODE% equ 4 goto :ERR_NO_GENERATION
if %EXIT_CODE% equ 5 goto :ERR_5MB_EXCEEDED
if %EXIT_CODE% equ 6 goto :ERR_DETAIL_FETCH
if %EXIT_CODE% equ 9 goto :ERR_API_CONNECTION
goto :ERR_UNKNOWN

:ERR_ARGUMENT
echo.
echo [エラー] 引数エラー（ユニットパスが指定されていません）
goto :ERROR_EXIT

:ERR_UNIT_NOT_FOUND
echo.
echo [エラー] ユニット未検出（指定したユニットが存在しません）
echo          - ユニットパスが正しいか確認してください: %UNIT_PATH%
echo          - スケジューラーサービス名が正しいか確認してください
goto :ERROR_EXIT

:ERR_UNIT_TYPE
echo.
echo [エラー] ユニット種別エラー（指定したユニットがジョブではありません）
echo          - 指定したパスはジョブネットまたはジョブグループの可能性があります
echo          - ジョブのフルパスを指定してください
goto :ERROR_EXIT

:ERR_NO_GENERATION
echo.
echo [エラー] 実行世代なし（実行履歴が存在しません）
echo          - ジョブが一度も実行されていない可能性があります
echo          - 世代指定（generation）の設定を確認してください
goto :ERROR_EXIT

:ERR_5MB_EXCEEDED
echo.
echo [エラー] 5MB超過エラー（実行結果が大きすぎて切り捨てられました）
echo          - 対象ユニットの出力サイズを確認してください
goto :ERROR_EXIT

:ERR_DETAIL_FETCH
echo.
echo [エラー] 詳細取得エラー（実行結果詳細の取得に失敗しました）
echo          - execIDが正しいか確認してください
goto :ERROR_EXIT

:ERR_API_CONNECTION
echo.
echo [エラー] API接続エラー（Web Consoleへの接続に失敗しました）
echo          - Web Consoleが起動しているか確認してください
echo          - 接続設定（ホスト名・ポート）を確認してください
echo          - 認証情報（ユーザー名・パスワード）を確認してください
goto :ERROR_EXIT

:ERR_UNKNOWN
echo.
echo [エラー] 不明なエラー（終了コード: %EXIT_CODE%）
goto :ERROR_EXIT

:ERROR_EXIT
echo.
del "%TEMP_FILE%" >nul 2>&1
pause
popd
exit /b %EXIT_CODE%

:SUCCESS
rem 結果が空かチェック
for %%A in ("%TEMP_FILE%") do set "FILE_SIZE=%%~zA"
if "%FILE_SIZE%"=="0" (
    echo.
    echo [警告] 取得結果が空です
    echo.
    del "%TEMP_FILE%" >nul 2>&1
    pause
    popd
    exit /b 1
)

rem 一時ファイルから START_TIME: 行を取得してジョブ開始日時を抽出
set "JOB_START_TIME="
for /f "tokens=1,* delims=:" %%a in ('type "%TEMP_FILE%" 2^>nul') do (
    if "%%a"=="START_TIME" (
        set "JOB_START_TIME=%%b"
        goto :GOT_START_TIME
    )
)
:GOT_START_TIME

rem 一時ファイルから JOBNET_NAME: 行を取得してジョブネット名を抽出
set "JOBNET_NAME="
for /f "tokens=1,* delims=:" %%a in ('type "%TEMP_FILE%" 2^>nul') do (
    if "%%a"=="JOBNET_NAME" (
        set "JOBNET_NAME=%%b"
        goto :GOT_JOBNET_NAME
    )
)
:GOT_JOBNET_NAME

rem 一時ファイルから JOBNET_COMMENT: 行を取得してジョブネットコメントを抽出
set "JOBNET_COMMENT="
for /f "tokens=1,* delims=:" %%a in ('type "%TEMP_FILE%" 2^>nul') do (
    if "%%a"=="JOBNET_COMMENT" (
        set "JOBNET_COMMENT=%%b"
        goto :GOT_JOBNET_COMMENT
    )
)
:GOT_JOBNET_COMMENT

rem ジョブネットコメントが空の場合はデフォルト値を使用
if "%JOBNET_COMMENT%"=="" set "JOBNET_COMMENT=%DEFAULT_COMMENT%"

rem ジョブネット名が空の場合はデフォルト値を使用
if "%JOBNET_NAME%"=="" set "JOBNET_NAME=UnknownJobnet"

rem ジョブ開始日時が取得できなかった場合はエラー
if "%JOB_START_TIME%"=="" (
    echo.
    echo [エラー] ジョブ開始日時が取得できませんでした
    echo          - ジョブが実行中の可能性があります
    echo          - 世代指定（generation）の設定を確認してください
    del "%TEMP_FILE%" >nul 2>&1
    pause
    popd
    exit /b 7
)

rem ジョブ開始日時が2日以上前かチェック
set "OLD_DATA_WARNING="
for /f %%D in ('powershell -NoProfile -Command "$dt=[DateTime]::ParseExact('%JOB_START_TIME%','yyyyMMdd_HHmmss',$null); if(((Get-Date)-$dt).TotalDays -ge 2){'OLD'}"') do set "OLD_DATA_WARNING=%%D"

if "%OLD_DATA_WARNING%"=="OLD" (
    echo.
    echo ================================================================
    echo   [警告] ジョブ開始日時が2日以上前です
    echo ================================================================
    echo.
    echo   ジョブ開始日時: %JOB_START_TIME%
    echo.
    echo   意図した世代のログか確認してください。
    echo   続行する場合は任意のキーを押してください...
    echo.
    pause >nul
)

rem 出力ファイル名を新形式で作成
rem 形式: 【ジョブ実行結果】【{日時}実行分】{ジョブネット名}_{コメント}.txt
set "OUTPUT_FILE=%~dp0【ジョブ実行結果】【%JOB_START_TIME%実行分】%JOBNET_NAME%_%JOBNET_COMMENT%.txt"

rem START_TIME:行、JOBNET_NAME:行、JOBNET_COMMENT:行を除いた実行結果詳細を出力ファイルに保存
(for /f "usebackq tokens=* delims=" %%L in ("%TEMP_FILE%") do (
    set "LINE=%%L"
    setlocal enabledelayedexpansion
    set "SKIP="
    if "!LINE:~0,11!"=="START_TIME:" set "SKIP=1"
    if "!LINE:~0,12!"=="JOBNET_NAME:" set "SKIP=1"
    if "!LINE:~0,15!"=="JOBNET_COMMENT:" set "SKIP=1"
    if not defined SKIP echo !LINE!
    endlocal
)) > "%OUTPUT_FILE%"

rem 一時ファイルを削除
del "%TEMP_FILE%" >nul 2>&1

echo.
echo ================================================================
echo   取得完了 - ファイルに保存しました
echo ================================================================
echo.
echo ジョブネット名: %JOBNET_NAME%
echo コメント:       %JOBNET_COMMENT%
echo ジョブ開始日時: %JOB_START_TIME%
echo 出力ファイル:   %OUTPUT_FILE%
echo.

rem メモ帳で開く
start notepad "%OUTPUT_FILE%"

rem 検索文字列が指定されている場合、その位置にスクロール
if not "%SCROLL_TO_TEXT%"=="" (
    echo スクロール位置: %SCROLL_TO_TEXT%
    echo.
    rem PowerShellでSendKeysを使ってメモ帳の検索機能を操作
    rem 注意: ウィンドウ名はファイル名で検索（日本語/英語環境両対応）
    powershell -NoProfile -Command ^
        "$searchText = '%SCROLL_TO_TEXT%';" ^
        "Start-Sleep -Milliseconds 800;" ^
        "$wshell = New-Object -ComObject WScript.Shell;" ^
        "$activated = $wshell.AppActivate('メモ帳');" ^
        "if (-not $activated) { $activated = $wshell.AppActivate('Notepad') };" ^
        "if ($activated) {" ^
        "  Start-Sleep -Milliseconds 100;" ^
        "  $wshell.SendKeys('^f');" ^
        "  Start-Sleep -Milliseconds 300;" ^
        "  $wshell.SendKeys($searchText);" ^
        "  Start-Sleep -Milliseconds 100;" ^
        "  $wshell.SendKeys('{ENTER}');" ^
        "  Start-Sleep -Milliseconds 100;" ^
        "  $wshell.SendKeys('{ESCAPE}');" ^
        "}"
)

popd
exit /b 0
