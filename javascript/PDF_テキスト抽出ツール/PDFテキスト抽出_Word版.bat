<# :
@echo off
chcp 65001 >nul
title PDF テキスト抽出ツール（Word版）
setlocal

powershell -NoProfile -ExecutionPolicy Bypass -Command "$scriptDir=('%~dp0' -replace '\\$',''); iex ((gc '%~f0' -Encoding UTF8) -join \"`n\")"
set EXITCODE=%ERRORLEVEL%

pause
exit /b %EXITCODE%
: #>

#==============================================================================
# PDF テキスト抽出ツール（Word版）
#==============================================================================
#
# 機能:
#   - 2つのPDFをWordで開いてテキスト抽出
#   - 一時ファイルに保存してWinMergeで比較
#
# 必要な環境:
#   - Microsoft Word 2013 以降
#   - WinMerge
#
#==============================================================================

#region 設定
#==============================================================================

# WinMergeのパス（環境に合わせて変更してください）
$WINMERGE_PATH = "C:\Program Files\WinMerge\WinMergeU.exe"

# 一時ファイルの保存先（空の場合はシステムの一時フォルダ）
$TEMP_FOLDER = ""

#==============================================================================
#endregion

# UTF-8出力設定
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ヘッダー表示
Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  PDF テキスト抽出ツール（Word版）" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Word でPDFを開いてテキスト抽出 → WinMergeで比較" -ForegroundColor Gray
Write-Host ""

#region 環境チェック
Write-Host "[チェック] 環境確認中..." -ForegroundColor Yellow

# WinMergeチェック
if (-not (Test-Path $WINMERGE_PATH)) {
    # 32bit版も確認
    $winmerge32 = "C:\Program Files (x86)\WinMerge\WinMergeU.exe"
    if (Test-Path $winmerge32) {
        $WINMERGE_PATH = $winmerge32
    } else {
        Write-Host ""
        Write-Host "[エラー] WinMergeが見つかりません" -ForegroundColor Red
        Write-Host "  確認したパス:" -ForegroundColor Gray
        Write-Host "    - C:\Program Files\WinMerge\WinMergeU.exe" -ForegroundColor Gray
        Write-Host "    - C:\Program Files (x86)\WinMerge\WinMergeU.exe" -ForegroundColor Gray
        Write-Host ""
        Write-Host "  WinMergeをインストールしてください:" -ForegroundColor Yellow
        Write-Host "    https://winmerge.org/" -ForegroundColor Cyan
        exit 1
    }
}
Write-Host "  [OK] WinMerge: $WINMERGE_PATH" -ForegroundColor Green

# Wordチェック
try {
    $wordApp = New-Object -ComObject Word.Application -ErrorAction Stop
    $wordVersion = $wordApp.Version
    $wordApp.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordApp) | Out-Null
    Write-Host "  [OK] Word: バージョン $wordVersion" -ForegroundColor Green

    # Word 2013 (15.0) 以降かチェック
    if ([double]$wordVersion -lt 15.0) {
        Write-Host ""
        Write-Host "[警告] Word 2013以降が推奨です（現在: $wordVersion）" -ForegroundColor Yellow
        Write-Host "  PDFを開けない可能性があります" -ForegroundColor Yellow
    }
} catch {
    Write-Host ""
    Write-Host "[エラー] Microsoft Wordが見つかりません" -ForegroundColor Red
    Write-Host "  Word 2013以降がインストールされている必要があります" -ForegroundColor Yellow
    exit 1
}

Write-Host ""
#endregion

#region ファイル選択ダイアログ
Add-Type -AssemblyName System.Windows.Forms

function Select-PdfFile {
    param([string]$Title)

    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = $Title
    $dialog.Filter = "PDFファイル (*.pdf)|*.pdf"
    $dialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.FileName
    }
    return $null
}

Write-Host "[入力] PDFファイルを選択してください" -ForegroundColor Yellow
Write-Host ""

# 旧PDF選択
Write-Host "  旧PDF（比較元）を選択..." -ForegroundColor Cyan
$pdf1Path = Select-PdfFile -Title "旧PDF（比較元）を選択"
if (-not $pdf1Path) {
    Write-Host "[キャンセル] 処理を中止しました" -ForegroundColor Yellow
    exit 0
}
$pdf1Name = [System.IO.Path]::GetFileNameWithoutExtension($pdf1Path)
Write-Host "    選択: $pdf1Path" -ForegroundColor Green

# 新PDF選択
Write-Host "  新PDF（比較先）を選択..." -ForegroundColor Cyan
$pdf2Path = Select-PdfFile -Title "新PDF（比較先）を選択"
if (-not $pdf2Path) {
    Write-Host "[キャンセル] 処理を中止しました" -ForegroundColor Yellow
    exit 0
}
$pdf2Name = [System.IO.Path]::GetFileNameWithoutExtension($pdf2Path)
Write-Host "    選択: $pdf2Path" -ForegroundColor Green

Write-Host ""
#endregion

#region テキスト抽出
Write-Host "[処理] PDFからテキストを抽出中..." -ForegroundColor Yellow
Write-Host ""

# 一時フォルダ設定
if ($TEMP_FOLDER -eq "" -or -not (Test-Path $TEMP_FOLDER)) {
    $TEMP_FOLDER = $env:TEMP
}

# 出力ファイルパス
$txt1Path = Join-Path $TEMP_FOLDER "${pdf1Name}_旧.txt"
$txt2Path = Join-Path $TEMP_FOLDER "${pdf2Name}_新.txt"

# Word起動
$wordApp = $null
try {
    Write-Host "  Wordを起動中..." -ForegroundColor Gray
    $wordApp = New-Object -ComObject Word.Application
    $wordApp.Visible = $false
    $wordApp.DisplayAlerts = 0  # wdAlertsNone

    # 旧PDF処理
    Write-Host "  旧PDFを処理中: $pdf1Name" -ForegroundColor Cyan
    try {
        $doc1 = $wordApp.Documents.Open($pdf1Path, $false, $true)  # ReadOnly

        # テキストを取得
        $text1 = $doc1.Content.Text

        # BOM付きUTF-8で保存
        $utf8WithBom = New-Object System.Text.UTF8Encoding($true)
        [System.IO.File]::WriteAllText($txt1Path, $text1, $utf8WithBom)

        $doc1.Close($false)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc1) | Out-Null

        Write-Host "    [OK] 保存: $txt1Path" -ForegroundColor Green
    } catch {
        Write-Host "    [エラー] $($_.Exception.Message)" -ForegroundColor Red
        throw
    }

    # 新PDF処理
    Write-Host "  新PDFを処理中: $pdf2Name" -ForegroundColor Cyan
    try {
        $doc2 = $wordApp.Documents.Open($pdf2Path, $false, $true)  # ReadOnly

        # テキストを取得
        $text2 = $doc2.Content.Text

        # BOM付きUTF-8で保存
        [System.IO.File]::WriteAllText($txt2Path, $text2, $utf8WithBom)

        $doc2.Close($false)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc2) | Out-Null

        Write-Host "    [OK] 保存: $txt2Path" -ForegroundColor Green
    } catch {
        Write-Host "    [エラー] $($_.Exception.Message)" -ForegroundColor Red
        throw
    }

} catch {
    Write-Host ""
    Write-Host "[エラー] PDF処理に失敗しました" -ForegroundColor Red
    Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "考えられる原因:" -ForegroundColor Yellow
    Write-Host "  - PDFが暗号化されている" -ForegroundColor Gray
    Write-Host "  - PDFが破損している" -ForegroundColor Gray
    Write-Host "  - Word のバージョンが古い" -ForegroundColor Gray
    exit 1
} finally {
    if ($wordApp) {
        $wordApp.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordApp) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        Write-Host "  Wordを終了しました" -ForegroundColor Gray
    }
}

Write-Host ""
#endregion

#region WinMerge起動
Write-Host "[起動] WinMergeを起動します..." -ForegroundColor Yellow

try {
    Start-Process -FilePath $WINMERGE_PATH -ArgumentList "`"$txt1Path`"", "`"$txt2Path`""
    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Green
    Write-Host "  WinMergeを起動しました" -ForegroundColor Green
    Write-Host "================================================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "  一時ファイル:" -ForegroundColor Gray
    Write-Host "    旧: $txt1Path" -ForegroundColor Gray
    Write-Host "    新: $txt2Path" -ForegroundColor Gray
    Write-Host ""
} catch {
    Write-Host "[エラー] WinMergeの起動に失敗しました" -ForegroundColor Red
    Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
#endregion
