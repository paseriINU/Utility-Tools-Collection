@echo off
:: UTF-8モードに設定
chcp 65001 >nul

echo ========================================
echo Invoke-RemoteBatch.ps1 テスト実行
echo ========================================
echo.
echo PowerShellスクリプトの構文チェック中...
echo.

:: PowerShellスクリプトの構文チェック
powershell -NoProfile -ExecutionPolicy Bypass -Command "& { $ErrorActionPreference = 'Stop'; try { $null = [System.Management.Automation.PSParser]::Tokenize((Get-Content '%~dp0Invoke-RemoteBatch.ps1' -Raw), [ref]$null); Write-Host '✓ 構文チェックOK' -ForegroundColor Green } catch { Write-Host '[エラー] 構文エラーが見つかりました:' -ForegroundColor Red; Write-Host $_.Exception.Message -ForegroundColor Yellow; exit 1 } }"

if errorlevel 1 (
    echo.
    echo 構文エラーがあるため、実行できません。
    pause
    exit /b 1
)

echo.
echo ========================================
echo スクリプトのヘルプを表示
echo ========================================
echo.

powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-Help '%~dp0Invoke-RemoteBatch.ps1' -Full"

echo.
echo ========================================
echo 実際に実行する場合は、以下のコマンドを使用:
echo ========================================
echo.
echo powershell -NoProfile -ExecutionPolicy Bypass -NoExit -File "%~dp0Invoke-RemoteBatch.ps1" -ComputerName "サーバ名" -UserName "ユーザー名" -BatchPath "C:\path\to\batch.bat"
echo.
echo (-NoExit オプションでウィンドウを閉じずにエラーを確認可能)
echo.
pause
