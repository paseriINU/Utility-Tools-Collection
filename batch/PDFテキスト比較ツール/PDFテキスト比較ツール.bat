<# :
@echo off
chcp 65001 >nul
start "" /b powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -STA -Command "iex ((gc '%~f0' -Encoding UTF8) -join \"`n\")"
exit /b
: #>

#==============================================================================
# PDF テキスト比較ツール（GUI版）
#==============================================================================
#
# 機能:
#   - ドラッグ&ドロップで2つのPDFを選択
#   - WordでPDFを開いてテキスト抽出
#   - WinMergeで比較表示
#
# 必要な環境:
#   - Microsoft Word 2013 以降
#   - WinMerge
#
#==============================================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#region 設定
# WinMergeのパス
$WINMERGE_PATHS = @(
    "C:\Program Files\WinMerge\WinMergeU.exe",
    "C:\Program Files (x86)\WinMerge\WinMergeU.exe"
)
#endregion

#region WinMerge検索
$winmergePath = $null
foreach ($path in $WINMERGE_PATHS) {
    if (Test-Path $path) {
        $winmergePath = $path
        break
    }
}

if (-not $winmergePath) {
    [System.Windows.Forms.MessageBox]::Show(
        "WinMergeが見つかりません。`n`nインストールしてください:`nhttps://winmerge.org/",
        "エラー",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    exit 1
}
#endregion

#region フォーム作成
$form = New-Object System.Windows.Forms.Form
$form.Text = "PDF テキスト比較ツール"
$form.Size = New-Object System.Drawing.Size(500, 420)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.Font = New-Object System.Drawing.Font("Meiryo UI", 9)
$form.AllowDrop = $true

# タイトルラベル
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "PDFファイルをドラッグ&ドロップしてください"
$titleLabel.Location = New-Object System.Drawing.Point(20, 15)
$titleLabel.Size = New-Object System.Drawing.Size(450, 25)
$titleLabel.Font = New-Object System.Drawing.Font("Meiryo UI", 11, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($titleLabel)

# 説明ラベル
$descLabel = New-Object System.Windows.Forms.Label
$descLabel.Text = "※ PDFをWordで開いてテキスト抽出し、WinMergeで比較します"
$descLabel.Location = New-Object System.Drawing.Point(20, 42)
$descLabel.Size = New-Object System.Drawing.Size(450, 18)
$descLabel.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
$form.Controls.Add($descLabel)

# 旧PDFラベル
$oldLabel = New-Object System.Windows.Forms.Label
$oldLabel.Text = "旧PDF（比較元）:"
$oldLabel.Location = New-Object System.Drawing.Point(20, 68)
$oldLabel.Size = New-Object System.Drawing.Size(120, 20)
$form.Controls.Add($oldLabel)

# 旧PDFドロップゾーン
$oldDropZone = New-Object System.Windows.Forms.Panel
$oldDropZone.Location = New-Object System.Drawing.Point(20, 88)
$oldDropZone.Size = New-Object System.Drawing.Size(440, 80)
$oldDropZone.BorderStyle = "FixedSingle"
$oldDropZone.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
$oldDropZone.AllowDrop = $true
$form.Controls.Add($oldDropZone)

$oldDropLabel = New-Object System.Windows.Forms.Label
$oldDropLabel.Text = "ここにPDFをドロップ`nまたはクリックして選択"
$oldDropLabel.TextAlign = "MiddleCenter"
$oldDropLabel.Dock = "Fill"
$oldDropLabel.ForeColor = [System.Drawing.Color]::Gray
$oldDropZone.Controls.Add($oldDropLabel)

# 新PDFラベル
$newLabel = New-Object System.Windows.Forms.Label
$newLabel.Text = "新PDF（比較先）:"
$newLabel.Location = New-Object System.Drawing.Point(20, 178)
$newLabel.Size = New-Object System.Drawing.Size(120, 20)
$form.Controls.Add($newLabel)

# 新PDFドロップゾーン
$newDropZone = New-Object System.Windows.Forms.Panel
$newDropZone.Location = New-Object System.Drawing.Point(20, 198)
$newDropZone.Size = New-Object System.Drawing.Size(440, 80)
$newDropZone.BorderStyle = "FixedSingle"
$newDropZone.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
$newDropZone.AllowDrop = $true
$form.Controls.Add($newDropZone)

$newDropLabel = New-Object System.Windows.Forms.Label
$newDropLabel.Text = "ここにPDFをドロップ`nまたはクリックして選択"
$newDropLabel.TextAlign = "MiddleCenter"
$newDropLabel.Dock = "Fill"
$newDropLabel.ForeColor = [System.Drawing.Color]::Gray
$newDropZone.Controls.Add($newDropLabel)

# 比較実行ボタン
$compareButton = New-Object System.Windows.Forms.Button
$compareButton.Text = "比較実行"
$compareButton.Location = New-Object System.Drawing.Point(180, 298)
$compareButton.Size = New-Object System.Drawing.Size(120, 35)
$compareButton.Enabled = $false
$compareButton.Font = New-Object System.Drawing.Font("Meiryo UI", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($compareButton)

# ステータスラベル
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = ""
$statusLabel.Location = New-Object System.Drawing.Point(20, 343)
$statusLabel.Size = New-Object System.Drawing.Size(450, 20)
$statusLabel.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($statusLabel)

# グローバル変数
$script:oldPdfPath = $null
$script:newPdfPath = $null
#endregion

#region 関数定義
function Update-ButtonState {
    if ($script:oldPdfPath -and $script:newPdfPath) {
        $compareButton.Enabled = $true
        $statusLabel.Text = "準備完了 - 「比較実行」をクリックしてください"
        $statusLabel.ForeColor = [System.Drawing.Color]::Green
    } else {
        $compareButton.Enabled = $false
        $statusLabel.Text = ""
    }
}

function Set-OldPdf($path) {
    if ($path -and $path.ToLower().EndsWith(".pdf")) {
        $script:oldPdfPath = $path
        $fileName = [System.IO.Path]::GetFileName($path)
        $oldDropLabel.Text = $fileName
        $oldDropLabel.ForeColor = [System.Drawing.Color]::Black
        $oldDropZone.BackColor = [System.Drawing.Color]::FromArgb(200, 230, 200)
        Update-ButtonState
    }
}

function Set-NewPdf($path) {
    if ($path -and $path.ToLower().EndsWith(".pdf")) {
        $script:newPdfPath = $path
        $fileName = [System.IO.Path]::GetFileName($path)
        $newDropLabel.Text = $fileName
        $newDropLabel.ForeColor = [System.Drawing.Color]::Black
        $newDropZone.BackColor = [System.Drawing.Color]::FromArgb(200, 230, 200)
        Update-ButtonState
    }
}

function Select-PdfFile($title) {
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = $title
    $dialog.Filter = "PDFファイル (*.pdf)|*.pdf"
    $dialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.FileName
    }
    return $null
}

function Extract-TextFromPdf($pdfPath) {
    $wordApp = $null
    $doc = $null
    $text = ""

    try {
        $wordApp = New-Object -ComObject Word.Application
        $wordApp.Visible = $false
        $wordApp.DisplayAlerts = 0

        $doc = $wordApp.Documents.Open($pdfPath, $false, $true)

        # Word定数
        $wdGoToPage = 1
        $wdGoToAbsolute = 1
        $wdStatisticPages = 2
        $wdHeaderFooterPrimary = 1

        # ページ数を取得
        $pageCount = $doc.ComputeStatistics($wdStatisticPages)
        $allTexts = @()

        for ($page = 1; $page -le $pageCount; $page++) {
            # ページの開始位置を取得
            $pageStart = $doc.GoTo($wdGoToPage, $wdGoToAbsolute, $page)
            $startPos = $pageStart.Start

            # ページの終了位置を取得
            if ($page -lt $pageCount) {
                $nextPage = $doc.GoTo($wdGoToPage, $wdGoToAbsolute, $page + 1)
                $endPos = $nextPage.Start
            } else {
                $endPos = $doc.Content.End
            }

            # そのページに対応するセクションのヘッダーを取得
            $sectionIndex = $pageStart.Information(10)  # wdActiveEndSectionNumber = 10
            if ($sectionIndex -ge 1 -and $sectionIndex -le $doc.Sections.Count) {
                try {
                    $section = $doc.Sections.Item($sectionIndex)
                    $header = $section.Headers.Item($wdHeaderFooterPrimary)
                    if ($header.Exists) {
                        $headerText = $header.Range.Text.Trim()
                        if ($headerText) {
                            $allTexts += $headerText
                        }
                    }
                } catch { }
            }

            # ページ範囲のテキストを取得
            $range = $doc.Range($startPos, $endPos)
            $bodyText = $range.Text.Trim()
            if ($bodyText) {
                $allTexts += $bodyText
            }

            # そのページに対応するセクションのフッターを取得
            if ($sectionIndex -ge 1 -and $sectionIndex -le $doc.Sections.Count) {
                try {
                    $section = $doc.Sections.Item($sectionIndex)
                    $footer = $section.Footers.Item($wdHeaderFooterPrimary)
                    if ($footer.Exists) {
                        $footerText = $footer.Range.Text.Trim()
                        if ($footerText) {
                            $allTexts += $footerText
                        }
                    }
                } catch { }
            }
        }

        $text = $allTexts -join "`r`n"
        $doc.Close($false)
    } catch {
        throw "PDFの読み込みに失敗しました: $($_.Exception.Message)"
    } finally {
        if ($doc) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
        }
        if ($wordApp) {
            $wordApp.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordApp) | Out-Null
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }

    # テキストのクリーニング処理
    # 1. Word特有のセクション区切り文字（垂直タブ等）を改行に置換
    $text = $text -replace "`v", "`r`n"
    # 2. フォームフィード文字を改行に置換
    $text = $text -replace "`f", "`r`n"
    # 3. ベル文字等の制御文字を除去（タブ、改行、復帰は保持）
    $text = $text -replace '[^\x09\x0A\x0D\x20-\x7E\u0080-\uFFFF]', ''
    # 4. 改行コードをCRLFに統一
    $text = $text -replace "`r`n", "`n"
    $text = $text -replace "`r", "`n"
    $text = $text -replace "`n", "`r`n"
    # 5. 連続する空行を2行までに制限
    $text = $text -replace "(`r`n){3,}", "`r`n`r`n"

    return $text
}

function Start-Comparison {
    $statusLabel.Text = "処理中..."
    $statusLabel.ForeColor = [System.Drawing.Color]::Blue
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $compareButton.Enabled = $false
    $form.Refresh()

    try {
        # ファイル名取得
        $oldName = [System.IO.Path]::GetFileNameWithoutExtension($script:oldPdfPath)
        $newName = [System.IO.Path]::GetFileNameWithoutExtension($script:newPdfPath)

        # 一時ファイルパス
        $tempFolder = $env:TEMP
        $oldTxtPath = Join-Path $tempFolder "${oldName}_旧.txt"
        $newTxtPath = Join-Path $tempFolder "${newName}_新.txt"

        # 旧PDF処理
        $statusLabel.Text = "旧PDFを処理中..."
        $form.Refresh()
        $oldText = Extract-TextFromPdf $script:oldPdfPath
        $utf8WithBom = New-Object System.Text.UTF8Encoding($true)
        [System.IO.File]::WriteAllText($oldTxtPath, $oldText, $utf8WithBom)

        # 新PDF処理
        $statusLabel.Text = "新PDFを処理中..."
        $form.Refresh()
        $newText = Extract-TextFromPdf $script:newPdfPath
        [System.IO.File]::WriteAllText($newTxtPath, $newText, $utf8WithBom)

        # WinMerge起動
        $statusLabel.Text = "WinMergeを起動中..."
        $form.Refresh()
        Start-Process -FilePath $winmergePath -ArgumentList "`"$oldTxtPath`"", "`"$newTxtPath`""

        $statusLabel.Text = "完了 - WinMergeを起動しました"
        $statusLabel.ForeColor = [System.Drawing.Color]::Green

    } catch {
        $statusLabel.Text = "エラー: $($_.Exception.Message)"
        $statusLabel.ForeColor = [System.Drawing.Color]::Red
        [System.Windows.Forms.MessageBox]::Show(
            "処理中にエラーが発生しました。`n`n$($_.Exception.Message)",
            "エラー",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        Update-ButtonState
    }
}
#endregion

#region イベントハンドラ
# 旧PDFドロップゾーン - ドラッグエンター
$oldDropZone.Add_DragEnter({
    param($sender, $e)
    if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $files = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
        if ($files[0].ToLower().EndsWith(".pdf")) {
            $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy
            $sender.BackColor = [System.Drawing.Color]::FromArgb(200, 200, 255)
        }
    }
})

# 旧PDFドロップゾーン - ドラッグリーブ
$oldDropZone.Add_DragLeave({
    param($sender, $e)
    if (-not $script:oldPdfPath) {
        $sender.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
    } else {
        $sender.BackColor = [System.Drawing.Color]::FromArgb(200, 230, 200)
    }
})

# 旧PDFドロップゾーン - ドロップ
$oldDropZone.Add_DragDrop({
    param($sender, $e)
    $files = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
    Set-OldPdf $files[0]
})

# 旧PDFドロップゾーン - クリック
$oldDropZone.Add_Click({
    $path = Select-PdfFile "旧PDF（比較元）を選択"
    if ($path) { Set-OldPdf $path }
})
$oldDropLabel.Add_Click({
    $path = Select-PdfFile "旧PDF（比較元）を選択"
    if ($path) { Set-OldPdf $path }
})

# 新PDFドロップゾーン - ドラッグエンター
$newDropZone.Add_DragEnter({
    param($sender, $e)
    if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $files = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
        if ($files[0].ToLower().EndsWith(".pdf")) {
            $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy
            $sender.BackColor = [System.Drawing.Color]::FromArgb(200, 200, 255)
        }
    }
})

# 新PDFドロップゾーン - ドラッグリーブ
$newDropZone.Add_DragLeave({
    param($sender, $e)
    if (-not $script:newPdfPath) {
        $sender.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
    } else {
        $sender.BackColor = [System.Drawing.Color]::FromArgb(200, 230, 200)
    }
})

# 新PDFドロップゾーン - ドロップ
$newDropZone.Add_DragDrop({
    param($sender, $e)
    $files = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
    Set-NewPdf $files[0]
})

# 新PDFドロップゾーン - クリック
$newDropZone.Add_Click({
    $path = Select-PdfFile "新PDF（比較先）を選択"
    if ($path) { Set-NewPdf $path }
})
$newDropLabel.Add_Click({
    $path = Select-PdfFile "新PDF（比較先）を選択"
    if ($path) { Set-NewPdf $path }
})

# 比較実行ボタン
$compareButton.Add_Click({
    Start-Comparison
})
#endregion

# フォーム表示
[void]$form.ShowDialog()
