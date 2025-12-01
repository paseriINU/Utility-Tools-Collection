@echo off
rem ====================================================================
rem リモートWindowsサーバ上でバッチファイルをWinRMで実行するスクリプト
rem PowerShellのInvoke-Commandを使用してリモート実行し、結果を取得
rem ====================================================================

setlocal enabledelayedexpansion

rem ====================================================================
rem 設定項目（必要に応じて編集してください）
rem ====================================================================

rem リモートサーバのコンピュータ名またはIPアドレス
set REMOTE_SERVER=192.168.1.100

rem リモートサーバの管理者ユーザー名
set REMOTE_USER=Administrator

rem リモートサーバで実行するバッチファイルのフルパス
set REMOTE_BATCH_PATH=C:\Scripts\target_script.bat

rem 実行結果を保存するローカルファイル（オプション）
set OUTPUT_LOG=%~dp0remote_exec_output.log

rem ====================================================================
rem メイン処理
rem ====================================================================

echo ========================================
echo リモートバッチ実行ツール (WinRM版)
echo ========================================
echo.
echo リモートサーバ: %REMOTE_SERVER%
echo 実行ユーザー  : %REMOTE_USER%
echo 実行ファイル  : %REMOTE_BATCH_PATH%
echo 出力ログ      : %OUTPUT_LOG%
echo.

rem PowerShellが利用可能か確認
powershell -Command "Write-Host 'PowerShell確認OK'" >nul 2>&1
if errorlevel 1 (
    echo [エラー] PowerShellが利用できません。
    goto :ERROR_EXIT
)

echo パスワードを入力してください：
set /p REMOTE_PASSWORD=

echo.
echo リモートサーバに接続中...
echo.

rem PowerShellでリモート実行
powershell -ExecutionPolicy Bypass -Command ^
    "$password = ConvertTo-SecureString '%REMOTE_PASSWORD%' -AsPlainText -Force; ^
     $credential = New-Object System.Management.Automation.PSCredential('%REMOTE_USER%', $password); ^
     try { ^
         Write-Host '接続確認中...' -ForegroundColor Cyan; ^
         $session = New-PSSession -ComputerName '%REMOTE_SERVER%' -Credential $credential -ErrorAction Stop; ^
         Write-Host '接続成功' -ForegroundColor Green; ^
         Write-Host ''; ^
         Write-Host '========================================' -ForegroundColor Yellow; ^
         Write-Host 'バッチファイル実行結果：' -ForegroundColor Yellow; ^
         Write-Host '========================================' -ForegroundColor Yellow; ^
         Write-Host ''; ^
         $result = Invoke-Command -Session $session -ScriptBlock { ^
             cmd.exe /c '%REMOTE_BATCH_PATH%' 2>&1 ^
         }; ^
         $result | Out-String; ^
         Write-Host ''; ^
         Write-Host '========================================' -ForegroundColor Yellow; ^
         Write-Host '実行完了' -ForegroundColor Green; ^
         Write-Host '========================================' -ForegroundColor Yellow; ^
         if ('%OUTPUT_LOG%' -ne '') { ^
             $result | Out-File -FilePath '%OUTPUT_LOG%' -Encoding UTF8; ^
             Write-Host ''; ^
             Write-Host '結果をログファイルに保存しました: %OUTPUT_LOG%' -ForegroundColor Cyan; ^
         }; ^
         Remove-PSSession -Session $session; ^
         exit 0; ^
     } catch { ^
         Write-Host ''; ^
         Write-Host '[エラー] リモート実行に失敗しました' -ForegroundColor Red; ^
         Write-Host $_.Exception.Message -ForegroundColor Red; ^
         Write-Host ''; ^
         Write-Host 'トラブルシューティング：' -ForegroundColor Yellow; ^
         Write-Host '1. リモートサーバでWinRMが有効になっているか確認' -ForegroundColor Gray; ^
         Write-Host '   winrm quickconfig' -ForegroundColor Gray; ^
         Write-Host '2. ファイアウォールでポート5985(HTTP)/5986(HTTPS)が開いているか確認' -ForegroundColor Gray; ^
         Write-Host '3. ユーザー名とパスワードが正しいか確認' -ForegroundColor Gray; ^
         Write-Host '4. リモートサーバがTrustedHostsに登録されているか確認' -ForegroundColor Gray; ^
         Write-Host '   winrm get winrm/config/client' -ForegroundColor Gray; ^
         exit 1; ^
     }"

if errorlevel 1 (
    goto :ERROR_EXIT
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
