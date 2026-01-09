@echo off
title JP1 ジョブツール サンプル
setlocal enabledelayedexpansion

rem UNCパス対応（PushD/PopDで自動マッピング）
pushd "%~dp0"

rem ============================================================================
rem JP1 ジョブツール サンプル
rem
rem 説明:
rem   JP1_REST_ジョブ実行ログ取得ツール.bat または JP1_REST_ジョブ情報取得ツール.bat を呼び出し、
rem   ジョブの実行結果をファイルに保存します。
rem
rem 機能:
rem   1. 実行モード: ジョブを即時実行してログを取得
rem   2. 取得モード: 実行せずに既存の結果を取得
rem
rem 使い方:
rem   1. 下記の UNIT_PATH を実行したいジョブのパスに変更
rem      例: /JobGroup/Jobnet/Job1
rem   2. 下記の MODE を設定（EXEC=実行, GET=取得）
rem   3. このバッチをダブルクリックで実行
rem
rem 終了コード（実行順）:
rem   0: 正常終了
rem   1: 引数エラー
rem   2: ユニット未検出（STEP 1）
rem   3: ユニット種別エラー（STEP 1）
rem   4: ルートジョブネット特定エラー / 実行世代なし
rem   5: 即時実行登録エラー / 5MB超過エラー
rem   6: タイムアウト / 詳細取得エラー
rem   7: 5MB超過エラー（EXEC版のみ）
rem   8: 詳細取得エラー（EXEC版） / 比較モードで両方取得失敗（GET版）
rem   9: API接続エラー（各STEP）
rem   10: ジョブ開始日時取得エラー / 比較モードで実行中検出（GET版）
rem   11: 実行中のジョブが検出された（待機タイムアウト、GET版）
rem ============================================================================

rem === ここを編集してください ===

rem モード設定
rem   EXEC: ジョブを即時実行してログを取得
rem   GET:  実行せずに既存の結果を取得
set "MODE=GET"

rem ジョブパス（必須）
rem 例: /JobGroup/Jobnet/Job1
set "UNIT_PATH=/JobGroup/Jobnet/Job1"

rem 比較用ジョブパス（オプション、GETモードのみ）
rem 2つのジョブを比較して新しい方を取得する場合に設定
rem 空欄の場合は比較なし
set "UNIT_PATH_2="

rem === スクロール位置の設定 ===
rem メモ帳で開いた後、指定した文字列を含む行に自動でジャンプします
rem 事前に行番号を特定し、Ctrl+Gで移動するため検索窓は開きません
rem 空欄の場合はスクロールせずにファイル先頭を表示します
rem 例: "エラー", "ERROR", "RC=", "異常終了" など
set "SCROLL_TO_TEXT="

rem === Excel貼り付け設定 ===
rem ログをExcelファイルに貼り付ける場合は以下を設定してください
rem 空欄の場合はExcel貼り付けを行いません
rem
rem Excelファイル名（このバッチと同じフォルダに配置）
set "EXCEL_FILE_NAME="
rem 例: set "EXCEL_FILE_NAME=ログ貼り付け用.xlsx"

rem 貼り付け先シート名
set "EXCEL_SHEET_NAME=Sheet1"

rem 貼り付け先セル位置
set "EXCEL_PASTE_CELL=A1"

rem === ファイル名の設定 ===
rem ファイル名形式: 【ジョブ実行結果】【{日時}実行分】【{終了状態}】{ジョブネット名}_{コメント}.txt
rem 例: 【ジョブ実行結果】【20251010_163520実行分】【正常終了】Jobnet_テスト.txt
rem
rem ジョブネットコメントが取得できない場合に使用するデフォルト値
set "DEFAULT_COMMENT=NoComment"

rem 一時ファイル名（後でジョブ開始日時を使用してリネーム）
set "TEMP_FILE=%~dp0temp_output.txt"

rem スクリプトのディレクトリを取得
set "SCRIPT_DIR=%~dp0"

echo.
echo ================================================================
echo   JP1 ジョブツール サンプル
echo ================================================================
echo.
echo   対象: %UNIT_PATH%
if not "%UNIT_PATH_2%"=="" echo   比較: %UNIT_PATH_2%
echo   モード: %MODE%
echo.

rem モードに応じてツールを選択
rem ※標準出力のみファイルにリダイレクト（標準エラー出力は待機メッセージ等をコンソールに表示）
if /i "%MODE%"=="EXEC" (
    echo ジョブを実行中...
    echo （完了まで時間がかかる場合があります）
    echo.
    call "%SCRIPT_DIR%JP1_REST_ジョブ実行ログ取得ツール.bat" "%UNIT_PATH%" > "%TEMP_FILE%"
) else if /i "%MODE%"=="GET" (
    echo ログを取得中...
    echo.
    if "%UNIT_PATH_2%"=="" (
        call "%SCRIPT_DIR%JP1_REST_ジョブ情報取得ツール.bat" "%UNIT_PATH%" > "%TEMP_FILE%"
    ) else (
        call "%SCRIPT_DIR%JP1_REST_ジョブ情報取得ツール.bat" "%UNIT_PATH%" "%UNIT_PATH_2%" > "%TEMP_FILE%"
    )
) else (
    echo [エラー] 無効なモードです: %MODE%
    echo          EXEC または GET を指定してください
    pause
    popd
    exit /b 1
)

set "EXIT_CODE=%ERRORLEVEL%"

rem エラーコード別のハンドリング（実行順）
if %EXIT_CODE% equ 0 goto :SUCCESS
if %EXIT_CODE% equ 1 goto :ERR_ARGUMENT
if %EXIT_CODE% equ 2 goto :ERR_UNIT_NOT_FOUND
if %EXIT_CODE% equ 3 goto :ERR_UNIT_TYPE
if %EXIT_CODE% equ 4 goto :ERR_ROOT_JOBNET_OR_NO_GEN
if %EXIT_CODE% equ 5 goto :ERR_EXEC_REGISTER_OR_5MB
if %EXIT_CODE% equ 6 goto :ERR_TIMEOUT_OR_DETAIL
if %EXIT_CODE% equ 7 goto :ERR_5MB_EXCEEDED
if %EXIT_CODE% equ 8 goto :ERR_DETAIL_OR_COMPARE
if %EXIT_CODE% equ 9 goto :ERR_API_CONNECTION
if %EXIT_CODE% equ 10 goto :ERR_RUNNING_JOB
if %EXIT_CODE% equ 11 goto :ERR_RUNNING_DETECTED
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

:ERR_ROOT_JOBNET_OR_NO_GEN
echo.
if /i "%MODE%"=="EXEC" (
    echo [エラー] ルートジョブネット特定エラー
    echo          - ジョブの定義情報からルートジョブネットが特定できませんでした
) else (
    echo [エラー] 実行世代なし
    echo          - 指定したジョブに実行履歴が存在しません
)
goto :ERROR_EXIT

:ERR_EXEC_REGISTER_OR_5MB
echo.
if /i "%MODE%"=="EXEC" (
    echo [エラー] 即時実行登録エラー
    echo          - ジョブネットの即時実行に失敗しました
    echo          - ユーザーに実行権限があるか確認してください
    echo          - ジョブネットが既に実行中でないか確認してください
) else (
    echo [エラー] 5MB超過エラー（実行結果が大きすぎて切り捨てられました）
    echo          - 対象ユニットの出力サイズを確認してください
)
goto :ERROR_EXIT

:ERR_TIMEOUT_OR_DETAIL
echo.
if /i "%MODE%"=="EXEC" (
    echo [エラー] タイムアウト（ジョブが完了しませんでした）
    echo          - ジョブの実行に時間がかかりすぎています
    echo          - タイムアウト設定を確認してください
) else (
    echo [エラー] 詳細取得エラー（実行結果詳細の取得に失敗しました）
    echo          - execIDが正しいか確認してください
)
goto :ERROR_EXIT

:ERR_5MB_EXCEEDED
echo.
echo [エラー] 5MB超過エラー（実行結果が大きすぎて切り捨てられました）
echo          - 対象ユニットの出力サイズを確認してください
goto :ERROR_EXIT

:ERR_DETAIL_OR_COMPARE
echo.
if /i "%MODE%"=="EXEC" (
    echo [エラー] 詳細取得エラー（実行結果詳細の取得に失敗しました）
    echo          - execIDが正しいか確認してください
) else (
    echo [エラー] 比較モードで両方のジョブ取得に失敗しました
    echo          - 両方のジョブパスを確認してください
)
goto :ERROR_EXIT

:ERR_API_CONNECTION
echo.
echo [エラー] API接続エラー（Web Consoleへの接続に失敗しました）
echo          - Web Consoleが起動しているか確認してください
echo          - 接続設定（ホスト名・ポート）を確認してください
echo          - 認証情報（ユーザー名・パスワード）を確認してください
goto :ERROR_EXIT

:ERR_RUNNING_JOB
echo.
echo [エラー] 比較モードで実行中のジョブが検出されました
echo          - 実行中のジョブがあるため、処理を中断しました
echo          - ジョブの完了後に再実行してください
rem 一時ファイルの内容を表示（エラーメッセージ確認用）
type "%TEMP_FILE%"
goto :ERROR_EXIT

:ERR_RUNNING_DETECTED
echo.
echo [エラー] 実行中のジョブが検出されました（待機タイムアウト）
echo.
rem 一時ファイルからRUNNING_ERROR:, RUNNING_JOB:, WAIT_TIMEOUT:を表示
for /f "tokens=1,* delims=:" %%a in ('type "%TEMP_FILE%" 2^>nul') do (
    if "%%a"=="RUNNING_ERROR" echo          %%b
    if "%%a"=="RUNNING_JOB" echo          %%b
    if "%%a"=="WAIT_TIMEOUT" echo          %%b
)
echo.
echo          最大待機時間を超えてもジョブが終了しなかったため、ログ取得を中止しました。
echo          ジョブの終了を待ってから再度実行してください。
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

rem 一時ファイルから END_STATUS: 行を取得して終了状態（日本語）を抽出
set "END_STATUS="
for /f "tokens=1,* delims=:" %%a in ('type "%TEMP_FILE%" 2^>nul') do (
    if "%%a"=="END_STATUS" (
        set "END_STATUS=%%b"
        goto :GOT_END_STATUS
    )
)
:GOT_END_STATUS

rem 一時ファイルから JOBNET_COMMENT: 行を取得してジョブネットコメントを抽出
set "JOBNET_COMMENT="
for /f "tokens=1,* delims=:" %%a in ('type "%TEMP_FILE%" 2^>nul') do (
    if "%%a"=="JOBNET_COMMENT" (
        set "JOBNET_COMMENT=%%b"
        goto :GOT_JOBNET_COMMENT
    )
)
:GOT_JOBNET_COMMENT

rem 一時ファイルから JOB_STATUS: 行を取得してジョブ終了ステータスを抽出（EXEC版のみ）
set "JOB_STATUS="
for /f "tokens=1,* delims=:" %%a in ('type "%TEMP_FILE%" 2^>nul') do (
    if "%%a"=="JOB_STATUS" (
        set "JOB_STATUS=%%b"
        goto :GOT_JOB_STATUS
    )
)
:GOT_JOB_STATUS

rem 比較結果を表示（GETモードで比較モードの場合）
set "SELECTED_PATH="
set "SELECTED_TIME="
set "REJECTED_PATH="
set "REJECTED_TIME="
for /f "tokens=1,* delims=:" %%a in ('type "%TEMP_FILE%" 2^>nul') do (
    if "%%a"=="SELECTED_PATH" set "SELECTED_PATH=%%b"
    if "%%a"=="SELECTED_TIME" set "SELECTED_TIME=%%b"
    if "%%a"=="REJECTED_PATH" set "REJECTED_PATH=%%b"
    if "%%a"=="REJECTED_TIME" set "REJECTED_TIME=%%b"
)

if not "%SELECTED_PATH%"=="" (
    echo.
    echo ================================================================
    echo   比較結果
    echo ================================================================
    echo.
    echo   [選択] %SELECTED_PATH%
    echo          開始日時: %SELECTED_TIME%
    echo.
    echo   [除外] %REJECTED_PATH%
    echo          開始日時: %REJECTED_TIME%
    echo.
)

rem ジョブネットコメントが空の場合はデフォルト値を使用
if "%JOBNET_COMMENT%"=="" set "JOBNET_COMMENT=%DEFAULT_COMMENT%"

rem ジョブネット名が空の場合はデフォルト値を使用
if "%JOBNET_NAME%"=="" set "JOBNET_NAME=UnknownJobnet"

rem 終了状態が空の場合はデフォルト値を使用
if "%END_STATUS%"=="" set "END_STATUS=不明"

rem ジョブ開始日時が取得できなかった場合はエラー
if "%JOB_START_TIME%"=="" (
    echo.
    echo [エラー] ジョブ開始日時が取得できませんでした
    echo          - ジョブが開始されなかった可能性があります
    del "%TEMP_FILE%" >nul 2>&1
    pause
    popd
    exit /b 10
)

rem 出力ファイル名を新形式で作成
rem 形式: 【ジョブ実行結果】【{日時}実行分】【{終了状態}】{ジョブネット名}_{コメント}.txt
set "OUTPUT_FILE=%~dp0【ジョブ実行結果】【%JOB_START_TIME%実行分】【%END_STATUS%】%JOBNET_NAME%_%JOBNET_COMMENT%.txt"

rem メタデータ行を除いた実行結果詳細を出力ファイルに保存
rem 除外対象: START_TIME:, END_STATUS:, JOBNET_NAME:, JOBNET_COMMENT:, JOB_STATUS:, SELECTED_*, REJECTED_*
rem ※空行を保持するためPowerShellで処理（for /fは空行をスキップしてしまうため）
powershell -NoProfile -Command ^
    "$excludePatterns = @('^START_TIME:', '^END_STATUS:', '^JOBNET_NAME:', '^JOBNET_COMMENT:', '^JOB_STATUS:', '^SELECTED_PATH:', '^SELECTED_TIME:', '^REJECTED_PATH:', '^REJECTED_TIME:');" ^
    "$content = Get-Content -Path '%TEMP_FILE%' -Encoding Default;" ^
    "$filtered = $content | Where-Object { $line = $_; -not ($excludePatterns | Where-Object { $line -match $_ }) };" ^
    "$filtered | Out-File -FilePath '%OUTPUT_FILE%' -Encoding Default"

rem 一時ファイルを削除
del "%TEMP_FILE%" >nul 2>&1

echo.
echo ================================================================
echo   実行完了 - ファイルに保存しました
echo ================================================================
echo.
echo ジョブネット名: %JOBNET_NAME%
echo コメント:       %JOBNET_COMMENT%
echo ジョブ開始日時: %JOB_START_TIME%
echo 終了状態:       %END_STATUS%
echo 出力ファイル:   %OUTPUT_FILE%
echo.

rem ジョブ終了状態によってメッセージを変える（EXEC版でJOB_STATUSがある場合）
if not "%JOB_STATUS%"=="" (
    if "%JOB_STATUS%"=="NORMAL" (
        echo [OK] ジョブは正常終了しました
    ) else if "%JOB_STATUS%"=="ABNORMAL" (
        echo [NG] ジョブは異常終了しました
    ) else if "%JOB_STATUS%"=="WARNING" (
        echo [警告] ジョブは警告終了しました
    ) else (
        echo [情報] ジョブ終了状態: %END_STATUS%
    )
) else (
    rem GET版の場合はEND_STATUSで判定
    if "%END_STATUS%"=="正常終了" (
        echo [OK] ジョブは正常終了しました
    ) else if "%END_STATUS%"=="異常検出終了" (
        echo [NG] ジョブは異常終了しました
    ) else if "%END_STATUS%"=="警告検出終了" (
        echo [警告] ジョブは警告終了しました
    ) else (
        echo [情報] ジョブ終了状態: %END_STATUS%
    )
)
echo.

rem ======================================================================
rem Excel貼り付け処理（EXCEL_FILE_NAMEが設定されている場合のみ）
rem ======================================================================
if not "%EXCEL_FILE_NAME%"=="" (
    set "EXCEL_PATH=%~dp0%EXCEL_FILE_NAME%"
    if exist "!EXCEL_PATH!" (
        echo Excelにログを貼り付け中...
        powershell -NoProfile -ExecutionPolicy Bypass -Command ^
            "$logFile = '%OUTPUT_FILE%';" ^
            "$excelPath = '!EXCEL_PATH!';" ^
            "$sheetName = '%EXCEL_SHEET_NAME%';" ^
            "$pasteCell = '%EXCEL_PASTE_CELL%';" ^
            "try {" ^
            "  $logContent = Get-Content $logFile -Encoding Default -Raw;" ^
            "  $excel = New-Object -ComObject Excel.Application;" ^
            "  $excel.Visible = $true;" ^
            "  $workbook = $excel.Workbooks.Open($excelPath);" ^
            "  $sheet = $workbook.Worksheets.Item($sheetName);" ^
            "  $sheet.Range($pasteCell).Value2 = $logContent;" ^
            "  $workbook.Save();" ^
            "  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) | Out-Null;" ^
            "  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null;" ^
            "  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null;" ^
            "  Write-Host '[OK] Excelにログを貼り付けました:' $sheetName $pasteCell;" ^
            "} catch {" ^
            "  Write-Host '[エラー] Excel貼り付けに失敗しました:' $_.Exception.Message;" ^
            "}"
        echo.
    ) else (
        echo [警告] Excelファイルが見つかりません: !EXCEL_PATH!
        echo.
    )
)

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
