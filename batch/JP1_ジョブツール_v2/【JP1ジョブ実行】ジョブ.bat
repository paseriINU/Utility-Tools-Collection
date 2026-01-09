@echo off
title JP1　ジョブ実行
setlocal

rem ============================================================================
rem JP1　ジョブ実行
rem
rem 説明:
rem   JP1ジョブ実行.bat を呼び出し、ジョブを即時実行してログを取得します。
rem
rem 使い方:
rem   1. 下記の UNIT_PATH を実行したいジョブのパスに変更
rem   2. 下記の OUTPUT_MODE を設定
rem   3. このバッチをダブルクリックで実行
rem ============================================================================

rem === ここを編集してください ===

rem ジョブパス（必須）
rem 例: /JobGroup/Jobnet/Job1
set "UNIT_PATH=/JobGroup/Jobnet/Job1"

rem 出力オプション
rem   /LOG      - ログファイル出力のみ（デフォルト）
rem   /NOTEPAD  - メモ帳で開く
rem   /EXCEL    - Excelに貼り付け
rem   /WINMERGE - WinMergeで比較
set "OUTPUT_MODE=/NOTEPAD"

rem === Excel貼り付け設定（OUTPUT_MODE=/EXCEL の場合のみ使用）===
rem Excelファイル名（このバッチと同じフォルダに配置）
rem 空欄の場合はExcel貼り付けを行いません
set "EXCEL_FILE_NAME="
rem 例: set "EXCEL_FILE_NAME=ログ貼り付け用.xlsx"

rem 貼り付け先シート名
set "EXCEL_SHEET_NAME=Sheet1"

rem 貼り付け先セル位置
set "EXCEL_PASTE_CELL=A1"

rem ===================================

rem UNCパス対応
pushd "%~dp0"

rem メインツールを呼び出し
call "JP1ジョブ実行.bat" "%UNIT_PATH%" "%OUTPUT_MODE%"

set "EXITCODE=%ERRORLEVEL%"

popd

rem 終了コードを表示
if %EXITCODE% neq 0 (
    echo.
    echo 終了コード: %EXITCODE%
    pause
)

exit /b %EXITCODE%
