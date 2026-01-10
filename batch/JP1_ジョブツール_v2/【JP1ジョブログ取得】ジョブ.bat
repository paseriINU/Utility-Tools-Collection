@echo off
title JP1 標準ログ
setlocal

rem ============================================================================
rem JP1 標準ログ
rem
rem 説明:
rem   JP1ジョブログ取得.bat を呼び出し、ジョブのログを取得します。
rem
rem 使い方:
rem   1. 下記の UNIT_PATH を取得したいジョブのパスに変更
rem   2. 下記の OUTPUT_MODE を設定
rem   3. このバッチをダブルクリックで実行
rem ============================================================================

rem === ここを編集してください ===

rem ジョブパス（必須）
rem 例: /JobGroup/Jobnet/Job1
set "UNIT_PATH=/JobGroup/Jobnet/Job1"

rem 比較用ジョブパス（オプション）
rem 2つのジョブを比較して新しい方を取得する場合に設定
set "UNIT_PATH_2="

rem 出力オプション
rem   /NOTEPAD  - メモ帳で開く
rem   /EXCEL    - Excelに貼り付け
rem   /WINMERGE - WinMergeで比較
set "JP1_OUTPUT_MODE=/NOTEPAD"

rem スクロール位置の設定（/NOTEPAD モード時のみ有効）
rem メモ帳で開いた後、指定した文字列を含む行に自動でジャンプします
rem 事前に行番号を特定し、Ctrl+Gで移動するため検索窓は開きません
rem 空欄の場合はスクロールせずにファイル先頭を表示します
rem 例: "エラー", "ERROR", "RC=", "異常終了" など
set "JP1_SCROLL_TO_TEXT="

rem Excel貼り付け設定（/EXCEL モード時のみ有効）
rem Excelファイルパス（相対パスまたはフルパス）
rem 例: "ログ貼り付け用.xlsx" または "C:\Users\Documents\ログ.xlsx"
set "EXCEL_FILE_NAME="

rem 貼り付け先シート名
rem 例: "Sheet1", "ログ貼り付け" など
set "EXCEL_SHEET_NAME="

rem 貼り付け先セル位置
rem 例: "A1", "B2" など
set "EXCEL_PASTE_CELL="

rem ===================================

rem UNCパス対応
pushd "%~dp0"

rem メインツールを呼び出し
call "JP1ジョブログ取得.bat" "%UNIT_PATH%" "%UNIT_PATH_2%"

set "EXITCODE=%ERRORLEVEL%"

popd

rem 終了コードを表示
if %EXITCODE% neq 0 (
    echo.
    echo 終了コード: %EXITCODE%
    pause
)

exit /b %EXITCODE%
