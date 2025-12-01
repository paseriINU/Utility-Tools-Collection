@echo off
rem ====================================================================
rem JP1ジョブネット起動ツール（リモート実行版）
rem PowerShell Remotingを使用してリモートサーバでajsentryを実行
rem ====================================================================

setlocal enabledelayedexpansion

cls
echo ========================================
echo JP1ジョブネット起動ツール
echo （リモート実行版）
echo ========================================
echo.

rem ====================================================================
rem 設定項目（ここを編集してください）
rem ====================================================================

rem JP1/AJS3が稼働しているリモートサーバ
set JP1_SERVER=192.168.1.100

rem リモートサーバのユーザー名（Windowsログインユーザー）
set REMOTE_USER=Administrator

rem JP1ユーザー名
set JP1_USER=jp1admin

rem JP1パスワード（空の場合は実行時に入力を求めます）
set JP1_PASSWORD=

rem 起動するジョブネットのフルパス
set JOBNET_PATH=/main_unit/jobgroup1/daily_batch

rem ajsentryコマンドのパス（リモートサーバ上）
set AJSENTRY_PATH=C:\Program Files\HITACHI\JP1AJS3\bin\ajsentry.exe

rem ====================================================================
rem メイン処理
rem ====================================================================

echo JP1サーバ      : %JP1_SERVER%
echo リモートユーザー: %REMOTE_USER%
echo JP1ユーザー    : %JP1_USER%
echo ジョブネットパス: %JOBNET_PATH%
echo.

rem JP1パスワードが設定されていない場合は入力を求める
if "%JP1_PASSWORD%"=="" (
    echo [注意] JP1パスワードが設定されていません。
    set /p JP1_PASSWORD="JP1パスワードを入力してください: "
    echo.
)

echo ジョブネットを起動しますか？
choice /M "実行する場合はYを押してください"
if errorlevel 2 (
    echo 処理をキャンセルしました。
    pause
    exit /b 0
)

echo.
echo ========================================
echo リモート接続してジョブネット起動中...
echo ========================================
echo.
echo リモートサーバの認証情報を入力してください。
echo.

rem PowerShell Remotingでリモート実行
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
"$credential = Get-Credential -UserName '%REMOTE_USER%' -Message 'JP1サーバ(%JP1_SERVER%)の認証情報を入力'; ^
if ($null -eq $credential) { ^
    Write-Host '[エラー] 認証情報の入力がキャンセルされました。' -ForegroundColor Red; ^
    exit 1; ^
}; ^
try { ^
    Write-Host 'リモートサーバに接続中...' -ForegroundColor Cyan; ^
    $session = New-PSSession -ComputerName '%JP1_SERVER%' -Credential $credential -ErrorAction Stop; ^
    Write-Host 'ajsentryコマンドを実行中...' -ForegroundColor Cyan; ^
    $result = Invoke-Command -Session $session -ScriptBlock { ^
        param($ajsPath, $jp1User, $jp1Pass, $jobnetPath); ^
        $output = & $ajsPath -h localhost -u $jp1User -p $jp1Pass -F $jobnetPath 2>&1; ^
        return @{ExitCode = $LASTEXITCODE; Output = $output}; ^
    } -ArgumentList '%AJSENTRY_PATH%', '%JP1_USER%', '%JP1_PASSWORD%', '%JOBNET_PATH%'; ^
    Remove-PSSession -Session $session; ^
    Write-Host ''; ^
    Write-Host '========================================' -ForegroundColor Cyan; ^
    if ($result.ExitCode -eq 0) { ^
        Write-Host 'ジョブネットの起動に成功しました' -ForegroundColor Green; ^
    } else { ^
        Write-Host 'ジョブネットの起動に失敗しました' -ForegroundColor Red; ^
        Write-Host 'エラーコード:' $result.ExitCode -ForegroundColor Red; ^
    }; ^
    Write-Host '========================================' -ForegroundColor Cyan; ^
    Write-Host ''; ^
    Write-Host '実行結果:'; ^
    Write-Host $result.Output; ^
    exit $result.ExitCode; ^
} catch { ^
    Write-Host ''; ^
    Write-Host '[エラー] リモート実行に失敗しました。' -ForegroundColor Red; ^
    Write-Host $_.Exception.Message -ForegroundColor Red; ^
    Write-Host ''; ^
    Write-Host '以下を確認してください：' -ForegroundColor Yellow; ^
    Write-Host '- リモートサーバのWinRMサービスが有効か' -ForegroundColor Yellow; ^
    Write-Host '- PowerShell Remotingが有効か（Enable-PSRemoting）' -ForegroundColor Yellow; ^
    Write-Host '- ファイアウォールで5985/5986ポートが開いているか' -ForegroundColor Yellow; ^
    Write-Host '- ネットワーク接続が正常か' -ForegroundColor Yellow; ^
    exit 1; ^
}"

set EXEC_RESULT=%errorlevel%

echo.
if %EXEC_RESULT% EQU 0 (
    echo ジョブネット: %JOBNET_PATH%
    echo サーバ      : %JP1_SERVER%
) else (
    echo.
    echo 追加の確認事項：
    echo - ajsentryのパスが正しいか: %AJSENTRY_PATH%
    echo - JP1ユーザー名、パスワードが正しいか
    echo - ジョブネットパスが正しいか
    echo - JP1/AJS3サービスが起動しているか
)
echo.

pause
endlocal
exit /b %EXEC_RESULT%
