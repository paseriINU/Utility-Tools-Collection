@echo off
rem ====================================================================
rem Gitブランチ間の差分ファイルを抽出してフォルダ構造を保ったままコピー
rem ====================================================================

setlocal enabledelayedexpansion

rem ====================================================================
rem 設定項目（必要に応じて編集してください）
rem ====================================================================

rem 比較元ブランチ（基準）
set BASE_BRANCH=main

rem 比較先ブランチ（差分を取得したいブランチ）
set TARGET_BRANCH=develop

rem 出力先フォルダ（相対パスまたは絶対パス）
set OUTPUT_DIR=diff_output

rem 削除されたファイルも含めるか（1=含める, 0=含めない）
set INCLUDE_DELETED=0

rem ====================================================================
rem メイン処理
rem ====================================================================

echo ========================================
echo Git差分ファイル抽出ツール
echo ========================================
echo.

rem Gitリポジトリのパスに移動
set GIT_PROJECT_PATH=C:\Users\%username%\source\Git\project

if not exist "%GIT_PROJECT_PATH%" (
    echo [エラー] Gitプロジェクトフォルダが見つかりません。
    echo パス: %GIT_PROJECT_PATH%
    echo.
    echo フォルダが存在するか確認してください。
    pause
    exit /b 1
)

echo Gitプロジェクトパス: %GIT_PROJECT_PATH%
cd /d "%GIT_PROJECT_PATH%"
echo.

rem Gitリポジトリかどうか確認
git rev-parse --git-dir >nul 2>&1
if errorlevel 1 (
    echo [エラー] このフォルダはGitリポジトリではありません。
    echo パス: %GIT_PROJECT_PATH%
    pause
    exit /b 1
)

rem リポジトリのルートディレクトリを取得
for /f "delims=" %%i in ('git rev-parse --show-toplevel') do set REPO_ROOT=%%i
rem Windowsパス形式に変換（/をバックスラッシュに）
set REPO_ROOT=%REPO_ROOT:/=\%

echo リポジトリルート: %REPO_ROOT%
echo 比較元ブランチ  : %BASE_BRANCH%
echo 比較先ブランチ  : %TARGET_BRANCH%
echo 出力先フォルダ  : %OUTPUT_DIR%
echo.

rem ブランチの存在確認
git rev-parse --verify %BASE_BRANCH% >nul 2>&1
if errorlevel 1 (
    echo [エラー] ブランチ '%BASE_BRANCH%' が見つかりません。
    pause
    exit /b 1
)

git rev-parse --verify %TARGET_BRANCH% >nul 2>&1
if errorlevel 1 (
    echo [エラー] ブランチ '%TARGET_BRANCH%' が見つかりません。
    pause
    exit /b 1
)

rem 出力先フォルダを作成（既存の場合は確認）
if exist "%OUTPUT_DIR%" (
    echo [警告] 出力先フォルダ '%OUTPUT_DIR%' は既に存在します。
    choice /M "上書きしますか"
    if errorlevel 2 (
        echo 処理を中止しました。
        pause
        exit /b 0
    )
    echo 既存のフォルダをクリア中...
    rd /s /q "%OUTPUT_DIR%"
)

mkdir "%OUTPUT_DIR%"

rem 差分ファイルリストを取得
echo 差分ファイルを検出中...
echo.

rem 一時ファイルに差分リストを保存
set TEMP_FILE=%TEMP%\git_diff_files_%RANDOM%.txt

if "%INCLUDE_DELETED%"=="1" (
    rem 削除されたファイルも含める
    git diff --name-only %BASE_BRANCH%...%TARGET_BRANCH% > "%TEMP_FILE%"
) else (
    rem 削除されたファイルを除外（追加・変更のみ）
    git diff --name-only --diff-filter=ACMR %BASE_BRANCH%...%TARGET_BRANCH% > "%TEMP_FILE%"
)

rem ファイル数をカウント
set FILE_COUNT=0
for /f %%i in ('type "%TEMP_FILE%" ^| find /c /v ""') do set FILE_COUNT=%%i

if %FILE_COUNT% EQU 0 (
    echo [情報] 差分ファイルが見つかりませんでした。
    echo 2つのブランチは同じ内容です。
    del "%TEMP_FILE%"
    pause
    exit /b 0
)

echo 検出された差分ファイル数: %FILE_COUNT% 個
echo.
echo ファイルをコピー中...
echo.

rem 各ファイルをコピー
set COPY_COUNT=0
set ERROR_COUNT=0

for /f "usebackq delims=" %%f in ("%TEMP_FILE%") do (
    set "FILE_PATH=%%f"

    rem Unixスタイルのパスをバックスラッシュに変換
    set "FILE_PATH=!FILE_PATH:/=\!"

    rem フルパス
    set "SOURCE_FILE=%REPO_ROOT%\!FILE_PATH!"
    set "DEST_FILE=%OUTPUT_DIR%\!FILE_PATH!"

    rem ファイルの存在確認（削除されたファイルはスキップ）
    if exist "!SOURCE_FILE!" (
        rem コピー先のディレクトリを作成
        for %%d in ("!DEST_FILE!") do set "DEST_DIR=%%~dpd"
        if not exist "!DEST_DIR!" mkdir "!DEST_DIR!"

        rem ファイルをコピー
        copy /y "!SOURCE_FILE!" "!DEST_FILE!" >nul 2>&1

        if errorlevel 1 (
            echo [エラー] !FILE_PATH!
            set /a ERROR_COUNT+=1
        ) else (
            echo [コピー] !FILE_PATH!
            set /a COPY_COUNT+=1
        )
    ) else (
        echo [削除済] !FILE_PATH! ^(スキップ^)
    )
)

rem 一時ファイルを削除
del "%TEMP_FILE%"

echo.
echo ========================================
echo 処理完了
echo ========================================
echo.
echo コピーしたファイル数: %COPY_COUNT% 個
if %ERROR_COUNT% GTR 0 (
    echo エラー: %ERROR_COUNT% 個
)
echo 出力先: %OUTPUT_DIR%
echo.

rem 出力先フォルダを開く
choice /M "出力先フォルダを開きますか"
if not errorlevel 2 (
    explorer "%OUTPUT_DIR%"
)

endlocal
exit /b 0
