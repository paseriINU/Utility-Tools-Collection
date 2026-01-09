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
rem 比較モード:
rem   UNIT_PATH_2 を設定すると、2つのジョブの実行日時を比較して
rem   新しい方のログのみを取得します。
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
rem   8: 比較モードで両方のジョブ取得に失敗
rem   9: API接続エラー（各STEP）
rem ============================================================================

rem === ここを編集してください（ジョブのパスを指定）===
set "UNIT_PATH=/JobGroup/Jobnet/Job1"

rem === 比較用ジョブパス（オプション）===
rem 2つ目のジョブパスを指定すると、両方の実行日時を比較して
rem 新しい方のログのみを取得します。不要な場合は空欄のまま。
set "UNIT_PATH_2="

rem === スクロール位置の設定 ===
rem メモ帳で開いた後、指定した文字列を含む行に自動でジャンプします
rem 事前に行番号を特定し、Ctrl+Gで移動するため検索窓は開きません
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
set "TEMP_FILE_2=%~dp0temp_output_2.txt"

rem スクリプトのディレクトリを取得
set "SCRIPT_DIR=%~dp0"

echo.
echo ================================================================
echo   JP1 ジョブログ取得サンプル
echo ================================================================
echo.

rem 比較モードかどうかを判定
if "%UNIT_PATH_2%"=="" (
    echo   対象: %UNIT_PATH%
    echo.
    echo ログを取得中...
    goto :SINGLE_MODE
) else (
    echo   対象1: %UNIT_PATH%
    echo   対象2: %UNIT_PATH_2%
    echo.
    echo   [比較モード] 新しい方のログを取得します
    echo.
    echo ログを取得中...
    goto :COMPARE_MODE
)

rem ============================================================================
rem 単一モード（従来通り）
rem ============================================================================
:SINGLE_MODE

rem JP1_REST_ジョブ情報取得ツール.bat を呼び出し、結果を一時ファイルに保存
call "%SCRIPT_DIR%JP1_REST_ジョブ情報取得ツール.bat" "%UNIT_PATH%" > "%TEMP_FILE%" 2>&1
set "EXIT_CODE=%ERRORLEVEL%"

rem エラーコード別のハンドリング（実行順）
if %EXIT_CODE% equ 0 goto :SUCCESS
set "ERROR_TARGET=%UNIT_PATH%"
goto :HANDLE_ERROR

rem ============================================================================
rem 比較モード（2つのジョブを比較）
rem ============================================================================
:COMPARE_MODE

rem 1つ目のジョブを取得
echo.
echo   [1/2] %UNIT_PATH% を取得中...
call "%SCRIPT_DIR%JP1_REST_ジョブ情報取得ツール.bat" "%UNIT_PATH%" > "%TEMP_FILE%" 2>&1
set "EXIT_CODE_1=%ERRORLEVEL%"

rem 2つ目のジョブを取得
echo   [2/2] %UNIT_PATH_2% を取得中...
call "%SCRIPT_DIR%JP1_REST_ジョブ情報取得ツール.bat" "%UNIT_PATH_2%" > "%TEMP_FILE_2%" 2>&1
set "EXIT_CODE_2=%ERRORLEVEL%"

rem 両方失敗した場合
if %EXIT_CODE_1% neq 0 if %EXIT_CODE_2% neq 0 (
    echo.
    echo [エラー] 両方のジョブ取得に失敗しました
    echo   対象1: %UNIT_PATH% （終了コード: %EXIT_CODE_1%）
    echo   対象2: %UNIT_PATH_2% （終了コード: %EXIT_CODE_2%）
    del "%TEMP_FILE%" >nul 2>&1
    del "%TEMP_FILE_2%" >nul 2>&1
    pause
    popd
    exit /b 8
)

rem 1つ目だけ失敗した場合、2つ目を使用
if %EXIT_CODE_1% neq 0 (
    echo.
    echo   [情報] 対象1の取得に失敗（終了コード: %EXIT_CODE_1%）、対象2を使用します
    del "%TEMP_FILE%" >nul 2>&1
    move "%TEMP_FILE_2%" "%TEMP_FILE%" >nul
    set "SELECTED_PATH=%UNIT_PATH_2%"
    goto :SUCCESS
)

rem 2つ目だけ失敗した場合、1つ目を使用
if %EXIT_CODE_2% neq 0 (
    echo.
    echo   [情報] 対象2の取得に失敗（終了コード: %EXIT_CODE_2%）、対象1を使用します
    del "%TEMP_FILE_2%" >nul 2>&1
    set "SELECTED_PATH=%UNIT_PATH%"
    goto :SUCCESS
)

rem 両方成功した場合、START_TIMEを比較
set "START_TIME_1="
for /f "tokens=1,* delims=:" %%a in ('type "%TEMP_FILE%" 2^>nul') do (
    if "%%a"=="START_TIME" (
        set "START_TIME_1=%%b"
        goto :GOT_ST1
    )
)
:GOT_ST1

set "START_TIME_2="
for /f "tokens=1,* delims=:" %%a in ('type "%TEMP_FILE_2%" 2^>nul') do (
    if "%%a"=="START_TIME" (
        set "START_TIME_2=%%b"
        goto :GOT_ST2
    )
)
:GOT_ST2

echo.
echo   対象1 開始日時: %START_TIME_1%
echo   対象2 開始日時: %START_TIME_2%

rem 日時を比較（文字列比較でOK: yyyyMMdd_HHmmss形式）
if "%START_TIME_1%" geq "%START_TIME_2%" (
    echo.
    echo   → 対象1 の方が新しい（または同じ）ため、対象1 のログを取得します
    del "%TEMP_FILE_2%" >nul 2>&1
    set "SELECTED_PATH=%UNIT_PATH%"
) else (
    echo.
    echo   → 対象2 の方が新しいため、対象2 のログを取得します
    del "%TEMP_FILE%" >nul 2>&1
    move "%TEMP_FILE_2%" "%TEMP_FILE%" >nul
    set "SELECTED_PATH=%UNIT_PATH_2%"
)

goto :SUCCESS

rem ============================================================================
rem エラーハンドリング
rem ============================================================================
:HANDLE_ERROR
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
echo          - ユニットパスが正しいか確認してください: %ERROR_TARGET%
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
del "%TEMP_FILE_2%" >nul 2>&1
pause
popd
exit /b %EXIT_CODE%

rem ============================================================================
rem 成功時の処理
rem ============================================================================
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

rem 一時ファイルから RUNNING_WARNING: 行を取得して実行中警告を抽出
set "RUNNING_WARNING="
for /f "tokens=1,* delims=:" %%a in ('type "%TEMP_FILE%" 2^>nul') do (
    if "%%a"=="RUNNING_WARNING" (
        set "RUNNING_WARNING=%%b"
        goto :GOT_RUNNING_WARNING
    )
)
:GOT_RUNNING_WARNING

rem 実行中警告がある場合、ユーザーに表示して確認を求める
if not "%RUNNING_WARNING%"=="" (
    echo.
    echo ================================================================
    echo   [警告] 実行中のジョブが検出されました
    echo ================================================================
    echo.
    echo   %RUNNING_WARNING%
    echo.
    echo   取得されるログは「直前に終了した世代」のものです。
    echo   現在実行中のジョブのログではありません。
    echo.
    echo   続行する場合は任意のキーを押してください...
    echo.
    pause >nul
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

rem メタデータ行を除いた実行結果詳細を出力ファイルに保存
rem 除外対象: RUNNING_WARNING:, START_TIME:, JOBNET_NAME:, JOBNET_COMMENT:
(for /f "usebackq tokens=* delims=" %%L in ("%TEMP_FILE%") do (
    set "LINE=%%L"
    setlocal enabledelayedexpansion
    set "SKIP="
    if "!LINE:~0,16!"=="RUNNING_WARNING:" set "SKIP=1"
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
if defined SELECTED_PATH (
    echo 選択されたジョブ: %SELECTED_PATH%
)
echo ジョブネット名: %JOBNET_NAME%
echo コメント:       %JOBNET_COMMENT%
echo ジョブ開始日時: %JOB_START_TIME%
echo 出力ファイル:   %OUTPUT_FILE%
echo.

rem 検索文字列が指定されている場合、行番号を事前に特定
set "SCROLL_LINE_NUM="
if not "%SCROLL_TO_TEXT%"=="" (
    echo スクロール位置: %SCROLL_TO_TEXT%
    rem PowerShellで行番号を特定
    for /f %%L in ('powershell -NoProfile -Command ^
        "$searchText = [regex]::Escape('%SCROLL_TO_TEXT%');" ^
        "$i = 0;" ^
        "Get-Content -Path '%OUTPUT_FILE%' -Encoding UTF8 | ForEach-Object { $i++; if ($_ -match $searchText) { Write-Output $i; break } }"') do set "SCROLL_LINE_NUM=%%L"
)

rem メモ帳で開く
start notepad "%OUTPUT_FILE%"

rem 行番号が特定できた場合、その行にジャンプ（Ctrl+Gで行へ移動）
if defined SCROLL_LINE_NUM (
    echo ジャンプ先行番号: %SCROLL_LINE_NUM%
    echo.
    rem PowerShellでCtrl+Gを送信して行へ移動
    rem ※検索ダイアログではなく「行へ移動」ダイアログを使用（移動後自動で閉じる）
    powershell -NoProfile -Command ^
        "$lineNum = '%SCROLL_LINE_NUM%';" ^
        "Start-Sleep -Milliseconds 600;" ^
        "$wshell = New-Object -ComObject WScript.Shell;" ^
        "$activated = $wshell.AppActivate('メモ帳');" ^
        "if (-not $activated) { $activated = $wshell.AppActivate('Notepad') };" ^
        "if ($activated) {" ^
        "  Start-Sleep -Milliseconds 100;" ^
        "  $wshell.SendKeys('^g');" ^
        "  Start-Sleep -Milliseconds 200;" ^
        "  $wshell.SendKeys($lineNum);" ^
        "  Start-Sleep -Milliseconds 100;" ^
        "  $wshell.SendKeys('{ENTER}');" ^
        "}"
) else (
    if not "%SCROLL_TO_TEXT%"=="" (
        echo [情報] 指定した文字列がファイル内に見つかりませんでした
    )
    echo.
)

popd
exit /b 0
