@echo off
title %~n0
setlocal enabledelayedexpansion

rem ============================================================================
rem JP1 ジャーナル取得
rem
rem 説明:
rem   JP1ジョブ情報取得.bat を呼び出し、複数ジョブのログを取得します。
rem   各ジョブのログは、JP1ジョブ情報取得.bat内の$jobExcelMappingに基づいて
rem   対応するExcelファイルに貼り付けられます。
rem
rem 使い方:
rem   1. 下記の UNIT_PATH_1〜6 を取得したいジョブのパスに変更（不要な場合は空欄）
rem   2. JP1ジョブ情報取得.bat の $jobExcelMapping でExcelファイルを設定
rem   3. このバッチをダブルクリックで実行
rem ============================================================================

rem === ここを編集してください ===

rem ジョブパス1
set "UNIT_PATH_1=/グループ/ネット/ジョブ1"

rem ジョブパス2
set "UNIT_PATH_2=/グループ/ネット/ジョブ2"

rem ジョブパス3（不要な場合は空欄にする）
set "UNIT_PATH_3="

rem ジョブパス4（不要な場合は空欄にする）
set "UNIT_PATH_4="

rem ジョブパス5（不要な場合は空欄にする）
set "UNIT_PATH_5="

rem ジョブパス6（不要な場合は空欄にする）
set "UNIT_PATH_6="

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
echo   JP1 ジャーナル取得
echo ================================================================
echo.

rem 有効なジョブ数をカウント
set "TOTAL_JOBS=0"
if not "%UNIT_PATH_1%"=="" set /a TOTAL_JOBS+=1
if not "%UNIT_PATH_2%"=="" set /a TOTAL_JOBS+=1
if not "%UNIT_PATH_3%"=="" set /a TOTAL_JOBS+=1
if not "%UNIT_PATH_4%"=="" set /a TOTAL_JOBS+=1
if not "%UNIT_PATH_5%"=="" set /a TOTAL_JOBS+=1
if not "%UNIT_PATH_6%"=="" set /a TOTAL_JOBS+=1

set "CURRENT_JOB=0"
set "EXITCODE1=0"
set "EXITCODE2=0"
set "EXITCODE3=0"
set "EXITCODE4=0"
set "EXITCODE5=0"
set "EXITCODE6=0"

rem --- ジョブ1の処理 ---
if not "%UNIT_PATH_1%"=="" (
    set /a CURRENT_JOB+=1
    echo [!CURRENT_JOB!/%TOTAL_JOBS%] ジョブ1のログを取得中...
    echo       パス: %UNIT_PATH_1%
    echo.

    set "JP1_UNIT_PATH=%UNIT_PATH_1%"
    call "JP1ジョブ情報取得.bat" "%UNIT_PATH_1%"
    set "EXITCODE1=!ERRORLEVEL!"

    if !EXITCODE1! neq 0 (
        echo [エラー] ジョブ1の取得に失敗しました（終了コード: !EXITCODE1!）
    ) else (
        echo [完了] ジョブ1の取得が完了しました
    )
    echo.
)

rem --- ジョブ2の処理 ---
if not "%UNIT_PATH_2%"=="" (
    set /a CURRENT_JOB+=1
    echo [!CURRENT_JOB!/%TOTAL_JOBS%] ジョブ2のログを取得中...
    echo       パス: %UNIT_PATH_2%
    echo.

    set "JP1_UNIT_PATH=%UNIT_PATH_2%"
    call "JP1ジョブ情報取得.bat" "%UNIT_PATH_2%"
    set "EXITCODE2=!ERRORLEVEL!"

    if !EXITCODE2! neq 0 (
        echo [エラー] ジョブ2の取得に失敗しました（終了コード: !EXITCODE2!）
    ) else (
        echo [完了] ジョブ2の取得が完了しました
    )
    echo.
)

rem --- ジョブ3の処理 ---
if not "%UNIT_PATH_3%"=="" (
    set /a CURRENT_JOB+=1
    echo [!CURRENT_JOB!/%TOTAL_JOBS%] ジョブ3のログを取得中...
    echo       パス: %UNIT_PATH_3%
    echo.

    set "JP1_UNIT_PATH=%UNIT_PATH_3%"
    call "JP1ジョブ情報取得.bat" "%UNIT_PATH_3%"
    set "EXITCODE3=!ERRORLEVEL!"

    if !EXITCODE3! neq 0 (
        echo [エラー] ジョブ3の取得に失敗しました（終了コード: !EXITCODE3!）
    ) else (
        echo [完了] ジョブ3の取得が完了しました
    )
    echo.
)

rem --- ジョブ4の処理 ---
if not "%UNIT_PATH_4%"=="" (
    set /a CURRENT_JOB+=1
    echo [!CURRENT_JOB!/%TOTAL_JOBS%] ジョブ4のログを取得中...
    echo       パス: %UNIT_PATH_4%
    echo.

    set "JP1_UNIT_PATH=%UNIT_PATH_4%"
    call "JP1ジョブ情報取得.bat" "%UNIT_PATH_4%"
    set "EXITCODE4=!ERRORLEVEL!"

    if !EXITCODE4! neq 0 (
        echo [エラー] ジョブ4の取得に失敗しました（終了コード: !EXITCODE4!）
    ) else (
        echo [完了] ジョブ4の取得が完了しました
    )
    echo.
)

rem --- ジョブ5の処理 ---
if not "%UNIT_PATH_5%"=="" (
    set /a CURRENT_JOB+=1
    echo [!CURRENT_JOB!/%TOTAL_JOBS%] ジョブ5のログを取得中...
    echo       パス: %UNIT_PATH_5%
    echo.

    set "JP1_UNIT_PATH=%UNIT_PATH_5%"
    call "JP1ジョブ情報取得.bat" "%UNIT_PATH_5%"
    set "EXITCODE5=!ERRORLEVEL!"

    if !EXITCODE5! neq 0 (
        echo [エラー] ジョブ5の取得に失敗しました（終了コード: !EXITCODE5!）
    ) else (
        echo [完了] ジョブ5の取得が完了しました
    )
    echo.
)

rem --- ジョブ6の処理 ---
if not "%UNIT_PATH_6%"=="" (
    set /a CURRENT_JOB+=1
    echo [!CURRENT_JOB!/%TOTAL_JOBS%] ジョブ6のログを取得中...
    echo       パス: %UNIT_PATH_6%
    echo.

    set "JP1_UNIT_PATH=%UNIT_PATH_6%"
    call "JP1ジョブ情報取得.bat" "%UNIT_PATH_6%"
    set "EXITCODE6=!ERRORLEVEL!"

    if !EXITCODE6! neq 0 (
        echo [エラー] ジョブ6の取得に失敗しました（終了コード: !EXITCODE6!）
    ) else (
        echo [完了] ジョブ6の取得が完了しました
    )
    echo.
)

popd

rem 結果サマリー
echo ================================================================
echo   処理結果
echo ================================================================
if not "%UNIT_PATH_1%"=="" (
    echo   ジョブ1: %UNIT_PATH_1%
    if %EXITCODE1% equ 0 (echo          結果: 成功) else (echo          結果: 失敗 [コード:%EXITCODE1%])
)
if not "%UNIT_PATH_2%"=="" (
    echo   ジョブ2: %UNIT_PATH_2%
    if %EXITCODE2% equ 0 (echo          結果: 成功) else (echo          結果: 失敗 [コード:%EXITCODE2%])
)
if not "%UNIT_PATH_3%"=="" (
    echo   ジョブ3: %UNIT_PATH_3%
    if %EXITCODE3% equ 0 (echo          結果: 成功) else (echo          結果: 失敗 [コード:%EXITCODE3%])
)
if not "%UNIT_PATH_4%"=="" (
    echo   ジョブ4: %UNIT_PATH_4%
    if %EXITCODE4% equ 0 (echo          結果: 成功) else (echo          結果: 失敗 [コード:%EXITCODE4%])
)
if not "%UNIT_PATH_5%"=="" (
    echo   ジョブ5: %UNIT_PATH_5%
    if %EXITCODE5% equ 0 (echo          結果: 成功) else (echo          結果: 失敗 [コード:%EXITCODE5%])
)
if not "%UNIT_PATH_6%"=="" (
    echo   ジョブ6: %UNIT_PATH_6%
    if %EXITCODE6% equ 0 (echo          結果: 成功) else (echo          結果: 失敗 [コード:%EXITCODE6%])
)
echo ================================================================
echo.

rem メモ帳/Excelでファイルを開く時間を確保するため10秒待機
echo このウィンドウは10秒後に閉じます...
timeout /t 10 /nobreak >nul
exit /b 0
