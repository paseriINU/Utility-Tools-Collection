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
rem   /LOG      - ログファイル出力のみ（デフォルト）
rem   /NOTEPAD  - メモ帳で開く
rem   /EXCEL    - Excelに貼り付け
rem   /WINMERGE - WinMergeで比較
set "OUTPUT_MODE=/NOTEPAD"

rem ===================================

rem UNCパス対応
pushd "%~dp0"

rem メインツールを呼び出し
call "JP1ジョブログ取得.bat" "%UNIT_PATH%" "%UNIT_PATH_2%" "%OUTPUT_MODE%"

set "EXITCODE=%ERRORLEVEL%"

popd

rem 終了コードを表示
if %EXITCODE% neq 0 (
    echo.
    echo 終了コード: %EXITCODE%
    pause
)

exit /b %EXITCODE%
