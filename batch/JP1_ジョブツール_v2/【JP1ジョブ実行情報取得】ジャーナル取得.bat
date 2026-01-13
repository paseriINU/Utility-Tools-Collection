@echo off
title JP1 ジャーナル実行情報取得
setlocal enabledelayedexpansion

rem ============================================================================
rem JP1 ジャーナル実行情報取得
rem
rem 説明:
rem   JP1ジョブ実行情報取得.bat を呼び出し、複数ジョブの実行情報（標準出力/エラー）を
rem   取得します。取得したログは、JP1ジョブ実行情報取得.bat内の$jobExcelMappingに
rem   基づいて対応するExcelファイルに貼り付けられます。
rem
rem 使い方:
rem   1. 下記の JOBNET_PATH_1〜3 をジョブネットのパスに設定
rem   2. 各ジョブネット配下のジョブ名を JOB_x_1, JOB_x_2 に設定
rem   3. JP1ジョブ実行情報取得.bat の $jobExcelMapping でExcelファイルを設定
rem   4. このバッチをダブルクリックで実行
rem
rem 構成:
rem   ジョブネット1: JOBNET_PATH_1 + JOB_1_1, JOB_1_2
rem   ジョブネット2: JOBNET_PATH_2 + JOB_2_1, JOB_2_2
rem   ジョブネット3: JOBNET_PATH_3 + JOB_3_1, JOB_3_2
rem ============================================================================

rem === ここを編集してください ===

rem --- ジョブネット1 ---
set "JOBNET_PATH_1=/グループ/ネット1"
set "JOB_1_1=ジョブA"
set "JOB_1_2=ジョブB"

rem --- ジョブネット2（不要な場合はJOBNET_PATH_2を空欄にする） ---
set "JOBNET_PATH_2="
set "JOB_2_1=ジョブC"
set "JOB_2_2=ジョブD"

rem --- ジョブネット3（不要な場合はJOBNET_PATH_3を空欄にする） ---
set "JOBNET_PATH_3="
set "JOB_3_1=ジョブE"
set "JOB_3_2=ジョブF"

rem 出力オプション
rem   /NOTEPAD  - メモ帳で開く
rem   /EXCEL    - Excelに貼り付け（$jobExcelMappingで自動選択）
rem   /WINMERGE - WinMergeで比較
set "JP1_OUTPUT_MODE=/EXCEL"

rem スクロール位置の設定（/NOTEPAD モード時のみ有効）
rem メモ帳で開いた後、指定した文字列を含む行に自動でジャンプします
rem 空欄の場合はスクロールせずにファイル先頭を表示します
set "JP1_SCROLL_TO_TEXT="

rem ===================================

rem UNCパス対応
pushd "%~dp0"

echo.
echo ================================================================
echo   JP1 ジャーナル実行情報取得
echo ================================================================
echo.

rem 有効なジョブネット数をカウント
set "TOTAL_JOBNETS=0"
if not "%JOBNET_PATH_1%"=="" set /a TOTAL_JOBNETS+=1
if not "%JOBNET_PATH_2%"=="" set /a TOTAL_JOBNETS+=1
if not "%JOBNET_PATH_3%"=="" set /a TOTAL_JOBNETS+=1

set "CURRENT_JOBNET=0"
set "EXITCODE_1_1=0"
set "EXITCODE_1_2=0"
set "EXITCODE_2_1=0"
set "EXITCODE_2_2=0"
set "EXITCODE_3_1=0"
set "EXITCODE_3_2=0"

rem --- ジョブネット1の処理 ---
if not "%JOBNET_PATH_1%"=="" (
    set /a CURRENT_JOBNET+=1
    echo [!CURRENT_JOBNET!/%TOTAL_JOBNETS%] ジョブネット1の実行情報を取得中...
    echo       ジョブネット: %JOBNET_PATH_1%
    echo.

    rem ジョブ1-1
    set "UNIT_PATH=%JOBNET_PATH_1%/%JOB_1_1%"
    echo   [1/2] %JOB_1_1% を取得中...
    set "JP1_UNIT_PATH=!UNIT_PATH!"
    call "JP1ジョブ実行情報取得.bat" "!UNIT_PATH!"
    set "EXITCODE_1_1=!ERRORLEVEL!"
    if !EXITCODE_1_1! neq 0 (
        echo         [エラー] 取得失敗（終了コード: !EXITCODE_1_1!）
    ) else (
        echo         [完了] 取得成功
    )

    rem ジョブ1-2
    set "UNIT_PATH=%JOBNET_PATH_1%/%JOB_1_2%"
    echo   [2/2] %JOB_1_2% を取得中...
    set "JP1_UNIT_PATH=!UNIT_PATH!"
    call "JP1ジョブ実行情報取得.bat" "!UNIT_PATH!"
    set "EXITCODE_1_2=!ERRORLEVEL!"
    if !EXITCODE_1_2! neq 0 (
        echo         [エラー] 取得失敗（終了コード: !EXITCODE_1_2!）
    ) else (
        echo         [完了] 取得成功
    )
    echo.
)

rem --- ジョブネット2の処理 ---
if not "%JOBNET_PATH_2%"=="" (
    set /a CURRENT_JOBNET+=1
    echo [!CURRENT_JOBNET!/%TOTAL_JOBNETS%] ジョブネット2の実行情報を取得中...
    echo       ジョブネット: %JOBNET_PATH_2%
    echo.

    rem ジョブ2-1
    set "UNIT_PATH=%JOBNET_PATH_2%/%JOB_2_1%"
    echo   [1/2] %JOB_2_1% を取得中...
    set "JP1_UNIT_PATH=!UNIT_PATH!"
    call "JP1ジョブ実行情報取得.bat" "!UNIT_PATH!"
    set "EXITCODE_2_1=!ERRORLEVEL!"
    if !EXITCODE_2_1! neq 0 (
        echo         [エラー] 取得失敗（終了コード: !EXITCODE_2_1!）
    ) else (
        echo         [完了] 取得成功
    )

    rem ジョブ2-2
    set "UNIT_PATH=%JOBNET_PATH_2%/%JOB_2_2%"
    echo   [2/2] %JOB_2_2% を取得中...
    set "JP1_UNIT_PATH=!UNIT_PATH!"
    call "JP1ジョブ実行情報取得.bat" "!UNIT_PATH!"
    set "EXITCODE_2_2=!ERRORLEVEL!"
    if !EXITCODE_2_2! neq 0 (
        echo         [エラー] 取得失敗（終了コード: !EXITCODE_2_2!）
    ) else (
        echo         [完了] 取得成功
    )
    echo.
)

rem --- ジョブネット3の処理 ---
if not "%JOBNET_PATH_3%"=="" (
    set /a CURRENT_JOBNET+=1
    echo [!CURRENT_JOBNET!/%TOTAL_JOBNETS%] ジョブネット3の実行情報を取得中...
    echo       ジョブネット: %JOBNET_PATH_3%
    echo.

    rem ジョブ3-1
    set "UNIT_PATH=%JOBNET_PATH_3%/%JOB_3_1%"
    echo   [1/2] %JOB_3_1% を取得中...
    set "JP1_UNIT_PATH=!UNIT_PATH!"
    call "JP1ジョブ実行情報取得.bat" "!UNIT_PATH!"
    set "EXITCODE_3_1=!ERRORLEVEL!"
    if !EXITCODE_3_1! neq 0 (
        echo         [エラー] 取得失敗（終了コード: !EXITCODE_3_1!）
    ) else (
        echo         [完了] 取得成功
    )

    rem ジョブ3-2
    set "UNIT_PATH=%JOBNET_PATH_3%/%JOB_3_2%"
    echo   [2/2] %JOB_3_2% を取得中...
    set "JP1_UNIT_PATH=!UNIT_PATH!"
    call "JP1ジョブ実行情報取得.bat" "!UNIT_PATH!"
    set "EXITCODE_3_2=!ERRORLEVEL!"
    if !EXITCODE_3_2! neq 0 (
        echo         [エラー] 取得失敗（終了コード: !EXITCODE_3_2!）
    ) else (
        echo         [完了] 取得成功
    )
    echo.
)

popd

rem 結果サマリー
echo ================================================================
echo   処理結果
echo ================================================================
if not "%JOBNET_PATH_1%"=="" (
    echo   ジョブネット1: %JOBNET_PATH_1%
    echo     - %JOB_1_1%: & if %EXITCODE_1_1% equ 0 (echo 成功) else (echo 失敗 [コード:%EXITCODE_1_1%])
    echo     - %JOB_1_2%: & if %EXITCODE_1_2% equ 0 (echo 成功) else (echo 失敗 [コード:%EXITCODE_1_2%])
)
if not "%JOBNET_PATH_2%"=="" (
    echo   ジョブネット2: %JOBNET_PATH_2%
    echo     - %JOB_2_1%: & if %EXITCODE_2_1% equ 0 (echo 成功) else (echo 失敗 [コード:%EXITCODE_2_1%])
    echo     - %JOB_2_2%: & if %EXITCODE_2_2% equ 0 (echo 成功) else (echo 失敗 [コード:%EXITCODE_2_2%])
)
if not "%JOBNET_PATH_3%"=="" (
    echo   ジョブネット3: %JOBNET_PATH_3%
    echo     - %JOB_3_1%: & if %EXITCODE_3_1% equ 0 (echo 成功) else (echo 失敗 [コード:%EXITCODE_3_1%])
    echo     - %JOB_3_2%: & if %EXITCODE_3_2% equ 0 (echo 成功) else (echo 失敗 [コード:%EXITCODE_3_2%])
)
echo ================================================================
echo.

pause
exit /b 0
