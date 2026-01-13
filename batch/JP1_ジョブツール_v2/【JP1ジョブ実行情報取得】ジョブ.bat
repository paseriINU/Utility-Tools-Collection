@echo off
title JP1 ジョブ実行
setlocal

rem ============================================================================
rem JP1 ジョブ実行
rem
rem 説明:
rem   JP1ジョブ実行情報取得.bat を呼び出し、ジョブを即時実行してログを取得します。
rem
rem 使い方:
rem   1. 下記の UNIT_PATH を実行したいジョブのパスに変更
rem   2. 下記の JP1_OUTPUT_MODE を設定
rem   3. /EXCEL使用時は JP1ジョブ実行情報取得.bat の $jobExcelMapping を設定
rem   4. このバッチをダブルクリックで実行
rem ============================================================================

rem === ここを編集してください ===

rem ジョブパス（必須）
rem 例: /JobGroup/Jobnet/Job1
set "UNIT_PATH=/JobGroup/Jobnet/Job1"

rem 出力オプション
rem   /NOTEPAD  - メモ帳で開く
rem   /EXCEL    - Excelに貼り付け（JP1ジョブ実行情報取得.batの$jobExcelMappingで設定）
rem   /WINMERGE - WinMergeで比較
set "JP1_OUTPUT_MODE=/NOTEPAD"

rem スクロール位置の設定（/NOTEPAD モード時のみ有効）
rem メモ帳で開いた後、指定した文字列を含む行に自動でジャンプします
rem 空欄の場合はスクロールせずにファイル先頭を表示します
set "JP1_SCROLL_TO_TEXT="

rem ===================================

rem UNCパス対応
pushd "%~dp0"

rem メインツールを呼び出し
call "JP1ジョブ実行情報取得.bat" "%UNIT_PATH%"

set "EXITCODE=%ERRORLEVEL%"

popd

rem 終了コードを表示
if %EXITCODE% neq 0 (
    echo.
    echo 終了コード: %EXITCODE%
    pause
)

exit /b %EXITCODE%
