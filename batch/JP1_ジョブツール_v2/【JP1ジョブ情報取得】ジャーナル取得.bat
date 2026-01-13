@echo off
title JP1 ジャーナル取得（複数ジョブ対応）
setlocal

rem ============================================================================
rem JP1 ジャーナル取得（複数ジョブ対応）
rem
rem 説明:
rem   JP1ジョブ情報取得.bat を呼び出し、複数ジョブのログを取得します。
rem   各ジョブのログは、JP1ジョブ情報取得.bat内の$jobExcelMappingに基づいて
rem   対応するExcelファイルに貼り付けられます。
rem
rem 使い方:
rem   1. 下記の UNIT_PATH_1, UNIT_PATH_2 を取得したいジョブのパスに変更
rem   2. JP1ジョブ情報取得.bat の $jobExcelMapping でExcelファイルを設定
rem   3. このバッチをダブルクリックで実行
rem ============================================================================

rem === ここを編集してください ===

rem ジョブパス1（必須）
rem 例: /JobGroup/Jobnet/Job1
set "UNIT_PATH_1=/JobGroup/Jobnet/Job1"

rem ジョブパス2（必須）
rem 例: /JobGroup/Jobnet/Job2
set "UNIT_PATH_2=/JobGroup/Jobnet/Job2"

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
echo   JP1 ジャーナル取得（複数ジョブ対応）
echo ================================================================
echo.

rem --- ジョブ1の処理 ---
echo [1/2] ジョブ1のログを取得中...
echo       パス: %UNIT_PATH_1%
echo.

set "JP1_UNIT_PATH=%UNIT_PATH_1%"
call "JP1ジョブ情報取得.bat" "%UNIT_PATH_1%"
set "EXITCODE1=%ERRORLEVEL%"

if %EXITCODE1% neq 0 (
    echo [エラー] ジョブ1の取得に失敗しました（終了コード: %EXITCODE1%）
) else (
    echo [完了] ジョブ1の取得が完了しました
)
echo.

rem --- ジョブ2の処理 ---
echo [2/2] ジョブ2のログを取得中...
echo       パス: %UNIT_PATH_2%
echo.

set "JP1_UNIT_PATH=%UNIT_PATH_2%"
call "JP1ジョブ情報取得.bat" "%UNIT_PATH_2%"
set "EXITCODE2=%ERRORLEVEL%"

if %EXITCODE2% neq 0 (
    echo [エラー] ジョブ2の取得に失敗しました（終了コード: %EXITCODE2%）
) else (
    echo [完了] ジョブ2の取得が完了しました
)
echo.

popd

rem 結果サマリー
echo ================================================================
echo   処理結果
echo ================================================================
echo   ジョブ1: %UNIT_PATH_1%
if %EXITCODE1% equ 0 (echo          結果: 成功) else (echo          結果: 失敗 [コード:%EXITCODE1%])
echo   ジョブ2: %UNIT_PATH_2%
if %EXITCODE2% equ 0 (echo          結果: 成功) else (echo          結果: 失敗 [コード:%EXITCODE2%])
echo ================================================================
echo.

pause
exit /b 0
