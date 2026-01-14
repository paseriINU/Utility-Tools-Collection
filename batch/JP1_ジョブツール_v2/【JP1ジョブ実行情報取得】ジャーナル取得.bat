@echo off
title %~n0
setlocal enabledelayedexpansion

rem ============================================================================
rem JP1 ジャーナル実行情報取得
rem
rem 説明:
rem   JP1ジョブ実行情報取得.bat を呼び出し、複数ジョブの実行情報（標準出力/エラー）を
rem   取得します。取得したログは、JP1ジョブ実行情報取得.bat内の$jobExcelMappingに
rem   基づいて対応するExcelファイルに貼り付けられます。
rem
rem   ★ 同一ジョブネット内の2ジョブは1回の実行で処理されます（ジョブネットは1回だけ実行）
rem
rem 使い方:
rem   1. 下記の JOBNET_PATH_1〜3 をジョブネットのパスに設定
rem   2. 各ジョブネット配下のジョブ名を JOB_x_1, JOB_x_2 に設定
rem   3. JP1ジョブ実行情報取得.bat の $jobExcelMapping でExcelファイルを設定
rem   4. このバッチをダブルクリックで実行
rem
rem 構成:
rem   ジョブネット1: JOBNET_PATH_1 + JOB_1_1, JOB_1_2 → 1回の実行で2ジョブ処理
rem   ジョブネット2: JOBNET_PATH_2 + JOB_2_1, JOB_2_2 → 1回の実行で2ジョブ処理
rem   ジョブネット3: JOBNET_PATH_3 + JOB_3_1, JOB_3_2 → 1回の実行で2ジョブ処理
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

rem 連続呼び出し時は個別の待機をスキップ（最後に1回だけ待機する）
set "JP1_SKIP_FINAL_WAIT=1"

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
set "EXITCODE_1=0"
set "EXITCODE_2=0"
set "EXITCODE_3=0"

rem --- ジョブネット1の処理（2ジョブを1回の実行で処理） ---
if not "%JOBNET_PATH_1%"=="" (
    set /a CURRENT_JOBNET+=1
    echo [!CURRENT_JOBNET!/%TOTAL_JOBNETS%] ジョブネット1の実行情報を取得中...
    echo       ジョブネット: %JOBNET_PATH_1%
    echo       ジョブ1: %JOB_1_1%
    echo       ジョブ2: %JOB_1_2%
    echo.

    rem 2つのジョブパスを引数で渡す（1回の実行で2ジョブ処理）
    call "JP1ジョブ実行情報取得.bat" "%JOBNET_PATH_1%/%JOB_1_1%" "%JOBNET_PATH_1%/%JOB_1_2%"
    set "EXITCODE_1=!ERRORLEVEL!"

    if !EXITCODE_1! neq 0 (
        echo [エラー] ジョブネット1の取得に失敗しました（終了コード: !EXITCODE_1!）
    ) else (
        echo [完了] ジョブネット1の取得が完了しました
    )
    echo.
)

rem --- ジョブネット2の処理（2ジョブを1回の実行で処理） ---
if not "%JOBNET_PATH_2%"=="" (
    set /a CURRENT_JOBNET+=1
    echo [!CURRENT_JOBNET!/%TOTAL_JOBNETS%] ジョブネット2の実行情報を取得中...
    echo       ジョブネット: %JOBNET_PATH_2%
    echo       ジョブ1: %JOB_2_1%
    echo       ジョブ2: %JOB_2_2%
    echo.

    rem 2つのジョブパスを引数で渡す（1回の実行で2ジョブ処理）
    call "JP1ジョブ実行情報取得.bat" "%JOBNET_PATH_2%/%JOB_2_1%" "%JOBNET_PATH_2%/%JOB_2_2%"
    set "EXITCODE_2=!ERRORLEVEL!"

    if !EXITCODE_2! neq 0 (
        echo [エラー] ジョブネット2の取得に失敗しました（終了コード: !EXITCODE_2!）
    ) else (
        echo [完了] ジョブネット2の取得が完了しました
    )
    echo.
)

rem --- ジョブネット3の処理（2ジョブを1回の実行で処理） ---
if not "%JOBNET_PATH_3%"=="" (
    set /a CURRENT_JOBNET+=1
    echo [!CURRENT_JOBNET!/%TOTAL_JOBNETS%] ジョブネット3の実行情報を取得中...
    echo       ジョブネット: %JOBNET_PATH_3%
    echo       ジョブ1: %JOB_3_1%
    echo       ジョブ2: %JOB_3_2%
    echo.

    rem 2つのジョブパスを引数で渡す（1回の実行で2ジョブ処理）
    call "JP1ジョブ実行情報取得.bat" "%JOBNET_PATH_3%/%JOB_3_1%" "%JOBNET_PATH_3%/%JOB_3_2%"
    set "EXITCODE_3=!ERRORLEVEL!"

    if !EXITCODE_3! neq 0 (
        echo [エラー] ジョブネット3の取得に失敗しました（終了コード: !EXITCODE_3!）
    ) else (
        echo [完了] ジョブネット3の取得が完了しました
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
    echo     - %JOB_1_1%, %JOB_1_2%
    if %EXITCODE_1% equ 0 (echo     結果: 成功) else (echo     結果: 失敗 [コード:%EXITCODE_1%])
)
if not "%JOBNET_PATH_2%"=="" (
    echo   ジョブネット2: %JOBNET_PATH_2%
    echo     - %JOB_2_1%, %JOB_2_2%
    if %EXITCODE_2% equ 0 (echo     結果: 成功) else (echo     結果: 失敗 [コード:%EXITCODE_2%])
)
if not "%JOBNET_PATH_3%"=="" (
    echo   ジョブネット3: %JOBNET_PATH_3%
    echo     - %JOB_3_1%, %JOB_3_2%
    if %EXITCODE_3% equ 0 (echo     結果: 成功) else (echo     結果: 失敗 [コード:%EXITCODE_3%])
)
echo ================================================================
echo.

rem メモ帳/Excelでファイルを開く時間を確保するため10秒待機
echo このウィンドウは10秒後に閉じます...
timeout /t 10 /nobreak >nul
exit /b 0
