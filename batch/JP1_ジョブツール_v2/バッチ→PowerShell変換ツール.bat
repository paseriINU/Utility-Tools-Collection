<# :
@echo off
setlocal
chcp 65001 >nul
title バッチ→PowerShell変換ツール

rem ============================================================================
rem ■ バッチファイル部分（PowerShellを起動するためのラッパー）
rem ============================================================================
rem
rem 【このツールの目的】
rem   バッチファイル（.bat）のコードをPowerShell（.ps1）形式に変換します。
rem   変換対象のバッチファイルをこのツールにドラッグ&ドロップしてください。
rem
rem 【使い方】
rem   1. 変換したいバッチファイルをこのツールにドラッグ&ドロップ
rem   2. 同じフォルダに .ps1 ファイルが生成されます
rem
rem 【注意事項】
rem   - 完全な自動変換は困難なため、変換後のコードは手動で確認・修正が必要です
rem   - goto/ラベル構造は関数形式に変換を試みますが、複雑な場合は手動調整が必要です
rem
rem 【対応できない項目（手動変換が必要）】
rem   - goto/ラベル構造 → while/関数に書き換え必要
rem   - 遅延環境変数展開（!変数!形式）
rem   - 複雑な for /f オプション（tokens,delims,skip,usebackq の組み合わせ）
rem   - 引数を持つ call :ラベル
rem   - shift コマンド
rem   - choice コマンド
rem   - 複雑なリダイレクト（2>file, 1>out 2>err, < input）
rem   - システム管理コマンド（net, sc, reg, wmic）
rem   - 複数行のif/forブロック（複雑な括弧対応）
rem   - 複合条件（if defined VAR if exist file.txt ...）
rem
rem ============================================================================

rem 引数チェック（ドラッグ&ドロップされたファイル）
if "%~1"=="" (
    echo.
    echo [エラー] 変換対象のバッチファイルが指定されていません。
    echo          バッチファイルをこのツールにドラッグ＆ドロップしてください。
    echo.
    pause
    exit /b 1
)

rem 拡張子チェック
if /i not "%~x1"==".bat" (
    if /i not "%~x1"==".cmd" (
        echo.
        echo [エラー] バッチファイル（.bat または .cmd）を指定してください。
        echo          指定されたファイル: %~nx1
        echo.
        pause
        exit /b 1
    )
)

rem ファイル存在チェック
if not exist "%~1" (
    echo.
    echo [エラー] ファイルが見つかりません: %~1
    echo.
    pause
    exit /b 1
)

rem 環境変数に入力ファイルパスを設定
set "INPUT_FILE=%~f1"
set "INPUT_DIR=%~dp1"
set "INPUT_NAME=%~n1"

rem PowerShell実行
powershell -NoProfile -ExecutionPolicy Bypass -Command "$inputFile=$env:INPUT_FILE; $inputDir=$env:INPUT_DIR; $inputName=$env:INPUT_NAME; iex ((gc '%~f0' -Encoding UTF8) -join \"`n\")"
set EXITCODE=%ERRORLEVEL%

pause
exit /b %EXITCODE%
: #>

# ==============================================================================
# ■ バッチ→PowerShell変換ツール
# ==============================================================================
#
# 【このツールの目的】
# バッチファイル（.bat / .cmd）のコードをPowerShellスクリプト（.ps1）形式に
# 変換するツールです。
#
# 【変換対象コマンド】
# - 基本コマンド: echo, set, if, for, rem, pause, exit
# - ファイル操作: dir, copy, xcopy, del, rd, md, move, ren, type
# - 検索: find, findstr
# - 変数: %VAR%, %~dp0, %~n0, %errorlevel% など
# - リダイレクト: >, >>, 2>&1, |
# - 制御: goto, call, ラベル
#
# 【注意事項】
# - 完全な自動変換は困難です。変換後のコードは必ず確認・修正してください
# - 複雑なgoto/ラベル構造は手動での調整が必要な場合があります
#
# 【対応できない項目（手動変換が必要）】
# - goto/ラベル構造 → while/関数に書き換え必要
# - 遅延環境変数展開（!変数!形式）
# - 複雑な for /f オプション（tokens,delims,skip,usebackq の組み合わせ）
# - 引数を持つ call :ラベル
# - shift コマンド
# - choice コマンド
# - 複雑なリダイレクト（2>file, 1>out 2>err, < input）
# - システム管理コマンド（net, sc, reg, wmic）
# - 複数行のif/forブロック（複雑な括弧対応）
# - 複合条件（if defined VAR if exist file.txt ...）
#
# ==============================================================================

# タイトル表示
Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  バッチ→PowerShell変換ツール" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

# 入力ファイル情報を表示
Write-Host "[入力ファイル] $inputFile" -ForegroundColor Yellow
Write-Host ""

# ==============================================================================
# ■ 変換ルール定義
# ==============================================================================

# 変換ルールを定義（順序が重要：より具体的なパターンを先に）
$conversionRules = @(
    # ----------------------------------------------------------------------
    # コメント
    # ----------------------------------------------------------------------
    @{ Pattern = '^rem\s+(.*)$'; Replacement = '# $1'; Description = 'rem → #' }
    @{ Pattern = '^::\s*(.*)$'; Replacement = '# $1'; Description = ':: → #' }
    @{ Pattern = '^REM\s+(.*)$'; Replacement = '# $1'; Description = 'REM → #' }

    # ----------------------------------------------------------------------
    # echo関連
    # ----------------------------------------------------------------------
    @{ Pattern = '^@echo\s+off\s*$'; Replacement = '# @echo off (PowerShellでは不要)'; Description = '@echo off' }
    @{ Pattern = '^echo\s+off\s*$'; Replacement = '# echo off (PowerShellでは不要)'; Description = 'echo off' }
    @{ Pattern = '^echo\s+on\s*$'; Replacement = '# echo on (PowerShellでは不要)'; Description = 'echo on' }
    @{ Pattern = '^echo\.\s*$'; Replacement = 'Write-Host ""'; Description = 'echo. → Write-Host ""' }
    @{ Pattern = '^echo\s+(.+)$'; Replacement = 'Write-Host "$1"'; Description = 'echo → Write-Host' }

    # ----------------------------------------------------------------------
    # 変数設定（set）
    # ----------------------------------------------------------------------
    @{ Pattern = '^set\s+/p\s+"?(\w+)=([^"]*)"?\s*$'; Replacement = '$$1 = Read-Host "$2"'; Description = 'set /p → Read-Host' }
    @{ Pattern = '^set\s+/p\s+(\w+)=(.*)$'; Replacement = '$$1 = Read-Host "$2"'; Description = 'set /p → Read-Host' }
    @{ Pattern = '^set\s+/a\s+"?(\w+)=([^"]*)"?\s*$'; Replacement = '$$1 = $2'; Description = 'set /a → 算術演算' }
    @{ Pattern = '^set\s+/a\s+(\w+)=(.*)$'; Replacement = '$$1 = $2'; Description = 'set /a → 算術演算' }
    @{ Pattern = '^set\s+"(\w+)=([^"]*)"$'; Replacement = '$$1 = "$2"'; Description = 'set "VAR=value"' }
    @{ Pattern = '^set\s+(\w+)=(.*)$'; Replacement = '$$1 = "$2"'; Description = 'set VAR=value' }
    @{ Pattern = '^setlocal\s*$'; Replacement = '# setlocal (PowerShellでは不要)'; Description = 'setlocal' }
    @{ Pattern = '^endlocal\s*$'; Replacement = '# endlocal (PowerShellでは不要)'; Description = 'endlocal' }

    # ----------------------------------------------------------------------
    # if文
    # ----------------------------------------------------------------------
    @{ Pattern = '^if\s+exist\s+"([^"]+)"\s+\('; Replacement = 'if (Test-Path "$1") {'; Description = 'if exist → Test-Path' }
    @{ Pattern = '^if\s+exist\s+(\S+)\s+\('; Replacement = 'if (Test-Path "$1") {'; Description = 'if exist → Test-Path' }
    @{ Pattern = '^if\s+exist\s+"([^"]+)"\s+(.+)$'; Replacement = 'if (Test-Path "$1") { $2 }'; Description = 'if exist → Test-Path' }
    @{ Pattern = '^if\s+exist\s+(\S+)\s+(.+)$'; Replacement = 'if (Test-Path "$1") { $2 }'; Description = 'if exist → Test-Path' }
    @{ Pattern = '^if\s+not\s+exist\s+"([^"]+)"\s+\('; Replacement = 'if (-not (Test-Path "$1")) {'; Description = 'if not exist' }
    @{ Pattern = '^if\s+not\s+exist\s+(\S+)\s+\('; Replacement = 'if (-not (Test-Path "$1")) {'; Description = 'if not exist' }
    @{ Pattern = '^if\s+not\s+exist\s+"([^"]+)"\s+(.+)$'; Replacement = 'if (-not (Test-Path "$1")) { $2 }'; Description = 'if not exist' }
    @{ Pattern = '^if\s+not\s+exist\s+(\S+)\s+(.+)$'; Replacement = 'if (-not (Test-Path "$1")) { $2 }'; Description = 'if not exist' }
    @{ Pattern = '^if\s+"([^"]+)"=="([^"]+)"\s+\('; Replacement = 'if ("$1" -eq "$2") {'; Description = 'if "A"=="B"' }
    @{ Pattern = '^if\s+"([^"]+)"=="([^"]+)"\s+(.+)$'; Replacement = 'if ("$1" -eq "$2") { $3 }'; Description = 'if "A"=="B"' }
    @{ Pattern = '^if\s+not\s+"([^"]+)"=="([^"]+)"\s+\('; Replacement = 'if ("$1" -ne "$2") {'; Description = 'if not "A"=="B"' }
    @{ Pattern = '^if\s+errorlevel\s+(\d+)\s+\('; Replacement = 'if ($LASTEXITCODE -ge $1) {'; Description = 'if errorlevel' }
    @{ Pattern = '^if\s+errorlevel\s+(\d+)\s+(.+)$'; Replacement = 'if ($LASTEXITCODE -ge $1) { $2 }'; Description = 'if errorlevel' }
    @{ Pattern = '^if\s+/i\s+"([^"]+)"=="([^"]+)"'; Replacement = 'if ("$1" -ieq "$2")'; Description = 'if /i (大文字小文字無視)' }
    @{ Pattern = '^\)\s*else\s*\('; Replacement = '} else {'; Description = ') else (' }
    @{ Pattern = '^\)\s*else\s+'; Replacement = '} else { '; Description = ') else' }
    @{ Pattern = '^\)\s*$'; Replacement = '}'; Description = ')' }

    # ----------------------------------------------------------------------
    # for文
    # ----------------------------------------------------------------------
    @{ Pattern = '^for\s+/f\s+"tokens=([^"]+)"\s+%%(\w)\s+in\s+\(''([^'']+)''\)\s+do\s+\('; Replacement = 'foreach ($line in (& $2)) { # tokens=$1'; Description = 'for /f tokens' }
    @{ Pattern = '^for\s+/f\s+%%(\w)\s+in\s+\(''([^'']+)''\)\s+do\s+\('; Replacement = 'foreach ($$1 in (& $2)) {'; Description = 'for /f command' }
    @{ Pattern = '^for\s+/f\s+"([^"]+)"\s+%%(\w)\s+in\s+\("([^"]+)"\)\s+do\s+\('; Replacement = 'foreach ($$2 in (Get-Content "$3")) { # $1'; Description = 'for /f file' }
    @{ Pattern = '^for\s+/r\s+"?([^"]*)"?\s+%%(\w)\s+in\s+\(([^)]+)\)\s+do\s+\('; Replacement = 'foreach ($$2 in (Get-ChildItem -Path "$1" -Filter "$3" -Recurse)) {'; Description = 'for /r 再帰' }
    @{ Pattern = '^for\s+/d\s+%%(\w)\s+in\s+\(([^)]+)\)\s+do\s+\('; Replacement = 'foreach ($$1 in (Get-ChildItem -Path "$2" -Directory)) {'; Description = 'for /d ディレクトリ' }
    @{ Pattern = '^for\s+%%(\w)\s+in\s+\(([^)]+)\)\s+do\s+\('; Replacement = 'foreach ($$1 in @($2)) {'; Description = 'for %%i in () do' }
    @{ Pattern = '^for\s+/l\s+%%(\w)\s+in\s+\((\d+),(\d+),(\d+)\)\s+do\s+\('; Replacement = 'for ($$1 = $2; $$1 -le $4; $$1 += $3) {'; Description = 'for /l ループ' }

    # ----------------------------------------------------------------------
    # ファイル操作コマンド
    # ----------------------------------------------------------------------
    @{ Pattern = '^dir\s+/b\s+"?([^"]*)"?\s*$'; Replacement = 'Get-ChildItem -Name "$1"'; Description = 'dir /b → Get-ChildItem -Name' }
    @{ Pattern = '^dir\s+/s\s+"?([^"]*)"?\s*$'; Replacement = 'Get-ChildItem -Recurse "$1"'; Description = 'dir /s → Get-ChildItem -Recurse' }
    @{ Pattern = '^dir\s+"?([^"]*)"?\s*$'; Replacement = 'Get-ChildItem "$1"'; Description = 'dir → Get-ChildItem' }
    @{ Pattern = '^dir\s*$'; Replacement = 'Get-ChildItem'; Description = 'dir → Get-ChildItem' }
    @{ Pattern = '^copy\s+/y\s+"?([^"]*)"?\s+"?([^"]*)"?'; Replacement = 'Copy-Item -Path "$1" -Destination "$2" -Force'; Description = 'copy /y → Copy-Item -Force' }
    @{ Pattern = '^copy\s+"?([^"]*)"?\s+"?([^"]*)"?'; Replacement = 'Copy-Item -Path "$1" -Destination "$2"'; Description = 'copy → Copy-Item' }
    @{ Pattern = '^xcopy\s+/s\s+/e\s+/y\s+"?([^"]*)"?\s+"?([^"]*)"?'; Replacement = 'Copy-Item -Path "$1" -Destination "$2" -Recurse -Force'; Description = 'xcopy → Copy-Item -Recurse' }
    @{ Pattern = '^xcopy\s+"?([^"]*)"?\s+"?([^"]*)"?'; Replacement = 'Copy-Item -Path "$1" -Destination "$2" -Recurse'; Description = 'xcopy → Copy-Item -Recurse' }
    @{ Pattern = '^del\s+/q\s+"?([^"]*)"?'; Replacement = 'Remove-Item -Path "$1" -Force'; Description = 'del /q → Remove-Item -Force' }
    @{ Pattern = '^del\s+"?([^"]*)"?'; Replacement = 'Remove-Item -Path "$1"'; Description = 'del → Remove-Item' }
    @{ Pattern = '^rd\s+/s\s+/q\s+"?([^"]*)"?'; Replacement = 'Remove-Item -Path "$1" -Recurse -Force'; Description = 'rd /s /q → Remove-Item -Recurse' }
    @{ Pattern = '^rmdir\s+/s\s+/q\s+"?([^"]*)"?'; Replacement = 'Remove-Item -Path "$1" -Recurse -Force'; Description = 'rmdir /s /q → Remove-Item -Recurse' }
    @{ Pattern = '^rd\s+"?([^"]*)"?'; Replacement = 'Remove-Item -Path "$1"'; Description = 'rd → Remove-Item' }
    @{ Pattern = '^md\s+"?([^"]*)"?'; Replacement = 'New-Item -ItemType Directory -Path "$1" -Force | Out-Null'; Description = 'md → New-Item Directory' }
    @{ Pattern = '^mkdir\s+"?([^"]*)"?'; Replacement = 'New-Item -ItemType Directory -Path "$1" -Force | Out-Null'; Description = 'mkdir → New-Item Directory' }
    @{ Pattern = '^move\s+/y\s+"?([^"]*)"?\s+"?([^"]*)"?'; Replacement = 'Move-Item -Path "$1" -Destination "$2" -Force'; Description = 'move /y → Move-Item -Force' }
    @{ Pattern = '^move\s+"?([^"]*)"?\s+"?([^"]*)"?'; Replacement = 'Move-Item -Path "$1" -Destination "$2"'; Description = 'move → Move-Item' }
    @{ Pattern = '^ren\s+"?([^"]*)"?\s+"?([^"]*)"?'; Replacement = 'Rename-Item -Path "$1" -NewName "$2"'; Description = 'ren → Rename-Item' }
    @{ Pattern = '^type\s+"?([^"]*)"?'; Replacement = 'Get-Content -Path "$1"'; Description = 'type → Get-Content' }

    # ----------------------------------------------------------------------
    # 検索コマンド
    # ----------------------------------------------------------------------
    @{ Pattern = '^find\s+"([^"]+)"\s+"?([^"]*)"?'; Replacement = 'Select-String -Pattern "$1" -Path "$2"'; Description = 'find → Select-String' }
    @{ Pattern = '^findstr\s+/i\s+"([^"]+)"\s+"?([^"]*)"?'; Replacement = 'Select-String -Pattern "$1" -Path "$2" -CaseSensitive:$false'; Description = 'findstr /i → Select-String' }
    @{ Pattern = '^findstr\s+"([^"]+)"\s+"?([^"]*)"?'; Replacement = 'Select-String -Pattern "$1" -Path "$2"'; Description = 'findstr → Select-String' }

    # ----------------------------------------------------------------------
    # その他のコマンド
    # ----------------------------------------------------------------------
    @{ Pattern = '^pause\s*$'; Replacement = 'Read-Host "続行するには Enter キーを押してください"'; Description = 'pause → Read-Host' }
    @{ Pattern = '^exit\s+/b\s+(\d+)\s*$'; Replacement = 'exit $1'; Description = 'exit /b N → exit N' }
    @{ Pattern = '^exit\s+/b\s+%(\w+)%\s*$'; Replacement = 'exit $$1'; Description = 'exit /b %VAR%' }
    @{ Pattern = '^exit\s+/b\s*$'; Replacement = 'exit 0'; Description = 'exit /b → exit' }
    @{ Pattern = '^exit\s*$'; Replacement = 'exit'; Description = 'exit' }
    @{ Pattern = '^cd\s+/d\s+"?([^"]*)"?'; Replacement = 'Set-Location -Path "$1"'; Description = 'cd /d → Set-Location' }
    @{ Pattern = '^cd\s+"?([^"]*)"?'; Replacement = 'Set-Location -Path "$1"'; Description = 'cd → Set-Location' }
    @{ Pattern = '^pushd\s+"?([^"]*)"?'; Replacement = 'Push-Location -Path "$1"'; Description = 'pushd → Push-Location' }
    @{ Pattern = '^popd\s*$'; Replacement = 'Pop-Location'; Description = 'popd → Pop-Location' }
    @{ Pattern = '^cls\s*$'; Replacement = 'Clear-Host'; Description = 'cls → Clear-Host' }
    @{ Pattern = '^title\s+(.+)$'; Replacement = '$Host.UI.RawUI.WindowTitle = "$1"'; Description = 'title → WindowTitle' }
    @{ Pattern = '^chcp\s+(\d+)\s*>nul'; Replacement = '# chcp $1 (PowerShellでは [Console]::OutputEncoding を使用)'; Description = 'chcp' }
    @{ Pattern = '^chcp\s+(\d+)'; Replacement = '# chcp $1 (PowerShellでは [Console]::OutputEncoding を使用)'; Description = 'chcp' }
    @{ Pattern = '^call\s+:(\w+)'; Replacement = '$1'; Description = 'call :label → 関数呼び出し' }
    @{ Pattern = '^call\s+"?([^"]*)"?'; Replacement = '& "$1"'; Description = 'call → &' }
    @{ Pattern = '^start\s+/wait\s+"?([^"]*)"?'; Replacement = 'Start-Process -FilePath "$1" -Wait'; Description = 'start /wait → Start-Process -Wait' }
    @{ Pattern = '^start\s+"?([^"]*)"?'; Replacement = 'Start-Process -FilePath "$1"'; Description = 'start → Start-Process' }
    @{ Pattern = '^timeout\s+/t\s+(\d+)'; Replacement = 'Start-Sleep -Seconds $1'; Description = 'timeout → Start-Sleep' }

    # ----------------------------------------------------------------------
    # goto/ラベル（警告付きで変換）
    # ----------------------------------------------------------------------
    @{ Pattern = '^goto\s+:?eof\s*$'; Replacement = 'return  # goto :eof'; Description = 'goto :eof → return' }
    @{ Pattern = '^goto\s+:?(\w+)\s*$'; Replacement = '# TODO: goto $1 → 関数呼び出しに書き換えが必要'; Description = 'goto :label' }
    @{ Pattern = '^:(\w+)\s*$'; Replacement = '# --- ラベル: $1 --- (関数化を検討: function $1 { ... })'; Description = ':label' }
)

# ==============================================================================
# ■ 変数置換ルール
# ==============================================================================

function Convert-Variables {
    param([string]$line)

    # バッチ変数をPowerShell変数に変換

    # 特殊変数
    $line = $line -replace '%~dp0', '$PSScriptRoot\'
    $line = $line -replace '%~d0', '(Split-Path $PSScriptRoot -Qualifier)'
    $line = $line -replace '%~p0', '(Split-Path $PSScriptRoot -NoQualifier)'
    $line = $line -replace '%~f0', '$MyInvocation.MyCommand.Path'
    $line = $line -replace '%~n0', '[System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Path)'
    $line = $line -replace '%~x0', '[System.IO.Path]::GetExtension($MyInvocation.MyCommand.Path)'
    $line = $line -replace '%~nx0', '[System.IO.Path]::GetFileName($MyInvocation.MyCommand.Path)'

    # 引数（%1, %2, ... → $args[0], $args[1], ...）
    $line = $line -replace '%~f(\d)', '$args[$1]'
    $line = $line -replace '%~dp(\d)', '(Split-Path $args[$1] -Parent)'
    $line = $line -replace '%~n(\d)', '[System.IO.Path]::GetFileNameWithoutExtension($args[$1])'
    $line = $line -replace '%~x(\d)', '[System.IO.Path]::GetExtension($args[$1])'
    $line = $line -replace '%~(\d)', '$args[$1]'
    $line = $line -replace '%(\d)', '$args[$1]'
    $line = $line -replace '%\*', '$args'

    # 環境変数
    $line = $line -replace '%errorlevel%', '$LASTEXITCODE'
    $line = $line -replace '%ERRORLEVEL%', '$LASTEXITCODE'
    $line = $line -replace '%cd%', '$PWD'
    $line = $line -replace '%CD%', '$PWD'
    $line = $line -replace '%date%', '(Get-Date -Format "yyyy/MM/dd")'
    $line = $line -replace '%DATE%', '(Get-Date -Format "yyyy/MM/dd")'
    $line = $line -replace '%time%', '(Get-Date -Format "HH:mm:ss")'
    $line = $line -replace '%TIME%', '(Get-Date -Format "HH:mm:ss")'
    $line = $line -replace '%random%', '(Get-Random -Maximum 32767)'
    $line = $line -replace '%RANDOM%', '(Get-Random -Maximum 32767)'

    # ユーザー定義変数 %VAR% → $VAR または $env:VAR
    # ここでは $VAR に変換（環境変数の場合は手動で $env:VAR に変更が必要）
    $line = $line -replace '%(\w+)%', '$$1'

    # for文内の %%i → $i
    $line = $line -replace '%%(\w)', '$$1'

    return $line
}

# ==============================================================================
# ■ リダイレクト変換
# ==============================================================================

function Convert-Redirects {
    param([string]$line)

    # リダイレクトを変換
    # 注意: 複雑なリダイレクトは手動調整が必要

    # 2>&1 はそのまま使用可能
    # >> → | Out-File -Append
    # > → | Out-File
    # ただし、コマンドの構造を変更する必要があるため、コメントで注記

    if ($line -match '>\s*nul\s*2>&1' -or $line -match '2>&1\s*>\s*nul') {
        $line = $line -replace '\s*>\s*nul\s*2>&1', ' | Out-Null'
        $line = $line -replace '\s*2>&1\s*>\s*nul', ' | Out-Null'
    } elseif ($line -match '>\s*nul') {
        $line = $line -replace '\s*>\s*nul', ' | Out-Null'
    } elseif ($line -match '>>\s*"?([^">\s]+)"?') {
        $line = $line -replace '\s*>>\s*"?([^">\s]+)"?', ' | Out-File -FilePath "$1" -Append -Encoding Default'
    } elseif ($line -match '>\s*"?([^">\s]+)"?') {
        $line = $line -replace '\s*>\s*"?([^">\s]+)"?', ' | Out-File -FilePath "$1" -Encoding Default'
    }

    return $line
}

# ==============================================================================
# ■ メイン変換処理
# ==============================================================================

Write-Host "[処理開始] 変換を開始します..." -ForegroundColor Green
Write-Host ""

# 入力ファイルを読み込み
try {
    $inputLines = Get-Content -Path $inputFile -Encoding Default
    Write-Host "[読み込み] $($inputLines.Count) 行を読み込みました" -ForegroundColor Gray
} catch {
    Write-Host "[エラー] ファイルの読み込みに失敗しました: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# 変換結果を格納する配列
$outputLines = @()

# ヘッダーコメントを追加
$outputLines += "# =============================================================================="
$outputLines += "# このファイルは「バッチ→PowerShell変換ツール」で自動生成されました"
$outputLines += "# 元ファイル: $inputFile"
$outputLines += "# 変換日時: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
$outputLines += "# "
$outputLines += "# 【重要】自動変換には限界があります。以下の点を確認してください："
$outputLines += "#   - 変数名が正しく変換されているか"
$outputLines += "#   - goto/ラベル構造が正しく関数化されているか"
$outputLines += "#   - リダイレクトが正しく変換されているか"
$outputLines += "#   - 文字エンコーディングが適切か"
$outputLines += "# =============================================================================="
$outputLines += ""

# 変換統計
$convertedCount = 0
$unchangedCount = 0
$warningCount = 0

# 各行を変換
foreach ($line in $inputLines) {
    $originalLine = $line
    $converted = $false
    $warning = $false

    # 空行はそのまま
    if ($line.Trim() -eq "") {
        $outputLines += ""
        continue
    }

    # 行頭の空白を保持
    $indent = ""
    if ($line -match '^(\s+)') {
        $indent = $Matches[1]
        $line = $line.TrimStart()
    }

    # @記号で始まる行の処理（エコー抑制）
    $suppressEcho = $false
    if ($line -match '^@') {
        $line = $line.Substring(1)
        $suppressEcho = $true
    }

    # 変換ルールを適用
    foreach ($rule in $conversionRules) {
        if ($line -match $rule.Pattern) {
            $line = $line -replace $rule.Pattern, $rule.Replacement
            $converted = $true

            # goto/ラベル関連は警告
            if ($rule.Description -match 'goto|label|ラベル') {
                $warning = $true
            }
            break
        }
    }

    # 変数置換
    $line = Convert-Variables -line $line

    # リダイレクト変換
    $line = Convert-Redirects -line $line

    # 統計更新
    if ($converted) {
        $convertedCount++
        if ($warning) { $warningCount++ }
    } else {
        $unchangedCount++
    }

    # 結果を追加
    $outputLines += "$indent$line"
}

# 出力ファイルパス
$outputFile = Join-Path $inputDir "$inputName.ps1"

# ファイルに書き込み
try {
    $outputLines | Out-File -FilePath $outputFile -Encoding UTF8
    Write-Host ""
    Write-Host "[変換完了]" -ForegroundColor Green
    Write-Host "  出力ファイル: $outputFile" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "[変換統計]" -ForegroundColor Cyan
    Write-Host "  変換された行: $convertedCount 行"
    Write-Host "  未変換の行:   $unchangedCount 行"
    if ($warningCount -gt 0) {
        Write-Host "  要確認の行:   $warningCount 行 (goto/ラベル関連)" -ForegroundColor Yellow
    }
    Write-Host ""
    Write-Host "[注意] 変換後のコードは必ず確認・テストしてください。" -ForegroundColor Yellow
} catch {
    Write-Host "[エラー] ファイルの書き込みに失敗しました: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

exit 0
