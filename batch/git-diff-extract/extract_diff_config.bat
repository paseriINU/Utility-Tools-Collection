@echo off
rem ====================================================================
rem Gitブランチ間の差分ファイルを抽出（設定ファイル版）
rem ====================================================================

setlocal enabledelayedexpansion

rem ====================================================================
rem 設定ファイルの読み込み
rem ====================================================================

set CONFIG_FILE=%~dp0config.ini

if not exist "%CONFIG_FILE%" (
    echo [エラー] 設定ファイルが見つかりません: %CONFIG_FILE%
    echo.
    echo config.ini を作成してください。サンプル：
    echo.
    echo [Branches]
    echo BASE_BRANCH=main
    echo TARGET_BRANCH=develop
    echo.
    echo [Output]
    echo OUTPUT_DIR=diff_output
    echo INCLUDE_DELETED=0
    echo.
    pause
    exit /b 1
)

rem 設定ファイルから値を読み込む
for /f "usebackq tokens=1,* delims==" %%a in ("%CONFIG_FILE%") do (
    set "%%a=%%b"
)

rem 必須項目のチェック
if not defined BASE_BRANCH set BASE_BRANCH=main
if not defined TARGET_BRANCH set TARGET_BRANCH=develop
if not defined OUTPUT_DIR set OUTPUT_DIR=diff_output
if not defined INCLUDE_DELETED set INCLUDE_DELETED=0

rem ====================================================================
rem メイン処理
rem ====================================================================

echo ========================================
echo Git差分ファイル抽出ツール（設定ファイル版）
echo ========================================
echo.

rem Gitリポジトリのパスに移動
set GIT_PROJECT_PATH=C:\Users\%username%\source\Git\project

if not exist "%GIT_PROJECT_PATH%" (
    echo [エラー] Gitプロジェクトフォルダが見つかりません。
    echo パス: %GIT_PROJECT_PATH%
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
    pause
    exit /b 1
)

rem リポジトリのルートディレクトリを取得
for /f "delims=" %%i in ('git rev-parse --show-toplevel') do set REPO_ROOT=%%i
set REPO_ROOT=%REPO_ROOT:/=\%

echo リポジトリルート: %REPO_ROOT%
echo 比較元ブランチ  : %BASE_BRANCH%
echo 比較先ブランチ  : %TARGET_BRANCH%
echo 出力先フォルダ  : %OUTPUT_DIR%
echo 削除ファイル含む: %INCLUDE_DELETED%
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

rem 出力先フォルダを作成
if exist "%OUTPUT_DIR%" (
    echo [警告] 出力先フォルダは既に存在します。
    choice /M "上書きしますか"
    if errorlevel 2 (
        echo 処理を中止しました。
        pause
        exit /b 0
    )
    rd /s /q "%OUTPUT_DIR%"
)

mkdir "%OUTPUT_DIR%"

rem 差分ファイルリストを取得
echo 差分ファイルを検出中...
echo.

set TEMP_FILE=%TEMP%\git_diff_files_%RANDOM%.txt

if "%INCLUDE_DELETED%"=="1" (
    git diff --name-only %BASE_BRANCH%...%TARGET_BRANCH% > "%TEMP_FILE%"
) else (
    git diff --name-only --diff-filter=ACMR %BASE_BRANCH%...%TARGET_BRANCH% > "%TEMP_FILE%"
)

rem ファイル数をカウント
set FILE_COUNT=0
for /f %%i in ('type "%TEMP_FILE%" ^| find /c /v ""') do set FILE_COUNT=%%i

if %FILE_COUNT% EQU 0 (
    echo [情報] 差分ファイルが見つかりませんでした。
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
    set "FILE_PATH=!FILE_PATH:/=\!"
    set "SOURCE_FILE=%REPO_ROOT%\!FILE_PATH!"
    set "DEST_FILE=%OUTPUT_DIR%\!FILE_PATH!"

    if exist "!SOURCE_FILE!" (
        for %%d in ("!DEST_FILE!") do set "DEST_DIR=%%~dpd"
        if not exist "!DEST_DIR!" mkdir "!DEST_DIR!"

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

choice /M "出力先フォルダを開きますか"
if not errorlevel 2 (
    explorer "%OUTPUT_DIR%"
)

endlocal
exit /b 0
