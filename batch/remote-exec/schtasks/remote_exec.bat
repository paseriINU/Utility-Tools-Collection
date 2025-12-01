@echo off
rem ====================================================================
rem リモートWindowsサーバ上でバッチファイルをCMDから実行するスクリプト
rem タスクスケジューラ（schtasks）を使用してリモート実行を実現
rem ====================================================================

setlocal enabledelayedexpansion

rem ====================================================================
rem 設定項目（必要に応じて編集してください）
rem ====================================================================

rem リモートサーバのコンピュータ名またはIPアドレス
set REMOTE_SERVER=192.168.1.100

rem リモートサーバの管理者ユーザー名（ドメイン\ユーザー名 または ユーザー名）
set REMOTE_USER=Administrator

rem リモートサーバで実行するバッチファイルのフルパス
set REMOTE_BATCH_PATH=C:\Scripts\target_script.bat

rem 作成する一時タスク名（ランダムな名前を推奨）
set TASK_NAME=RemoteExec_%RANDOM%

rem タスク実行後に自動削除するか（1=削除する, 0=削除しない）
set AUTO_DELETE=1

rem ====================================================================
rem メイン処理
rem ====================================================================

echo ========================================
echo リモートバッチ実行ツール
echo ========================================
echo.
echo リモートサーバ: %REMOTE_SERVER%
echo 実行ユーザー  : %REMOTE_USER%
echo 実行ファイル  : %REMOTE_BATCH_PATH%
echo タスク名      : %TASK_NAME%
echo.

rem パスワード入力（セキュリティのため画面に表示されません）
echo リモートサーバのパスワードを入力してください：
set /p REMOTE_PASSWORD=

echo.
echo タスクを作成中...

rem リモートサーバにタスクを作成
schtasks /Create ^
    /S %REMOTE_SERVER% ^
    /U %REMOTE_USER% ^
    /P %REMOTE_PASSWORD% ^
    /TN %TASK_NAME% ^
    /TR "%REMOTE_BATCH_PATH%" ^
    /SC ONCE ^
    /ST 00:00 ^
    /RU SYSTEM ^
    /F

if errorlevel 1 (
    echo.
    echo [エラー] タスクの作成に失敗しました。
    echo - サーバ名、ユーザー名、パスワードが正しいか確認してください
    echo - リモートサーバへのネットワーク接続を確認してください
    echo - 管理者権限があるか確認してください
    goto :ERROR_EXIT
)

echo タスク作成成功
echo.
echo タスクを実行中...

rem タスクを即座に実行
schtasks /Run ^
    /S %REMOTE_SERVER% ^
    /U %REMOTE_USER% ^
    /P %REMOTE_PASSWORD% ^
    /TN %TASK_NAME%

if errorlevel 1 (
    echo.
    echo [エラー] タスクの実行に失敗しました。
    goto :CLEANUP
)

echo タスク実行開始
echo.
echo 実行状態を確認中（5秒待機）...
timeout /t 5 /nobreak >nul

rem タスクの状態を確認
schtasks /Query ^
    /S %REMOTE_SERVER% ^
    /U %REMOTE_USER% ^
    /P %REMOTE_PASSWORD% ^
    /TN %TASK_NAME% ^
    /FO LIST

echo.
echo ========================================
echo 注意: タスクはバックグラウンドで実行されます。
echo       実行結果を確認するには、リモートサーバの
echo       ログファイルやタスクスケジューラを確認してください。
echo ========================================

:CLEANUP
if "%AUTO_DELETE%"=="1" (
    echo.
    echo タスクを削除中...

    rem 削除前に少し待機（タスクが完了するまで）
    timeout /t 3 /nobreak >nul

    schtasks /Delete ^
        /S %REMOTE_SERVER% ^
        /U %REMOTE_USER% ^
        /P %REMOTE_PASSWORD% ^
        /TN %TASK_NAME% ^
        /F >nul 2>&1

    if errorlevel 1 (
        echo [警告] タスクの削除に失敗しました。手動で削除してください。
    ) else (
        echo タスク削除完了
    )
)

echo.
echo 処理が完了しました。
goto :END

:ERROR_EXIT
echo.
echo 処理を中断しました。
exit /b 1

:END
endlocal
exit /b 0
