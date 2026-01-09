@echo off
title JP1　標準ログ
setlocal

rem ============================================================================
rem JP1　標準ログ
rem
rem 説明:
rem   JP1ジョブ実行.bat または JP1ジョブログ取得.bat を呼び出し、
rem   ジョブの実行結果を取得します。
rem
rem 機能:
rem   1. EXECモード: ジョブを即時実行してログを取得
rem   2. GETモード:  実行せずに既存の結果を取得
rem
rem 使い方:
rem   1. 下記の UNIT_PATH を実行したいジョブのパスに変更
rem   2. 下記の MODE を設定（EXEC=実行, GET=取得）
rem   3. 下記の OUTPUT_MODE を設定
rem   4. このバッチをダブルクリックで実行
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
set "UNIT_PATH_2="

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

rem モードに応じてメインツールを呼び出し
if /i "%MODE%"=="EXEC" (
    call "JP1ジョブ実行.bat" "%UNIT_PATH%" "%OUTPUT_MODE%"
) else (
    call "JP1ジョブログ取得.bat" "%UNIT_PATH%" "%UNIT_PATH_2%" "%OUTPUT_MODE%"
)

set "EXITCODE=%ERRORLEVEL%"

popd

rem 終了コードを表示
if %EXITCODE% neq 0 (
    echo.
    echo 終了コード: %EXITCODE%
    pause
)

exit /b %EXITCODE%
