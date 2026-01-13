@echo off
title JP1 ジャーナル実行情報取得
setlocal

rem ============================================================================
rem JP1 ジャーナル実行情報取得
rem
rem 説明:
rem   JP1ジョブ実行情報取得.bat を呼び出し、ジョブの実行情報（標準出力/エラー）を
rem   取得します。取得したログは、JP1ジョブ実行情報取得.bat内の$jobExcelMappingに
rem   基づいて対応するExcelファイルに貼り付けられます。
rem
rem 使い方:
rem   1. 下記の UNIT_PATH を取得したいジョブのパスに変更
rem   2. JP1ジョブ実行情報取得.bat の $jobExcelMapping でExcelファイルを設定
rem   3. このバッチをダブルクリックで実行
rem ============================================================================

rem === ここを編集してください ===

rem ジョブパス（必須）
rem 例: /JobGroup/Jobnet/Job1
set "UNIT_PATH=/JobGroup/Jobnet/Job1"

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

rem --- ジョブの処理 ---
echo ジョブの実行情報を取得中...
echo パス: %UNIT_PATH%
echo.

set "JP1_UNIT_PATH=%UNIT_PATH%"
call "JP1ジョブ実行情報取得.bat" "%UNIT_PATH%"
set "EXITCODE=%ERRORLEVEL%"

if %EXITCODE% neq 0 (
    echo [エラー] ジョブの取得に失敗しました（終了コード: %EXITCODE%）
) else (
    echo [完了] ジョブの取得が完了しました
)
echo.

popd

rem 結果サマリー
echo ================================================================
echo   処理結果
echo ================================================================
echo   ジョブ: %UNIT_PATH%
if %EXITCODE% equ 0 (echo   結果: 成功) else (echo   結果: 失敗 [コード:%EXITCODE%])
echo ================================================================
echo.

pause
exit /b 0
