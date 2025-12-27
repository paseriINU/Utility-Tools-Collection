<# :
@echo off
chcp 65001 >nul
title PDF内部構造調査ツール
setlocal

rem 引数チェック
if "%~1"=="" (
    echo.
    echo ================================================================
    echo   PDF内部構造調査ツール
    echo ================================================================
    echo.
    echo   使い方: このバッチファイルにPDFファイルをドラッグ＆ドロップ
    echo.
    echo   または: PDF内部構造調査ツール.bat "PDFファイルのパス"
    echo.
    pause
    exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -Command "$pdfPath='%~1'; iex ((gc '%~f0' -Encoding UTF8) -join \"`n\")"
pause
exit /b 0
: #>

# ================================================================
#   PDF内部構造調査ツール
#   PDFのしおり（/Title）のエンコーディングを調査します
# ================================================================

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  PDF内部構造調査ツール" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

# ファイル存在チェック
if (-not (Test-Path $pdfPath)) {
    Write-Host "[エラー] ファイルが見つかりません: $pdfPath" -ForegroundColor Red
    exit 1
}

Write-Host "対象ファイル: $pdfPath" -ForegroundColor Yellow
Write-Host ""

# PDFファイルを読み込み
Write-Host "PDFファイルを読み込み中..." -ForegroundColor Gray
try {
    $bytes = [System.IO.File]::ReadAllBytes($pdfPath)
    $content = [System.Text.Encoding]::GetEncoding("iso-8859-1").GetString($bytes)
    Write-Host "読み込み完了: $($bytes.Length) バイト" -ForegroundColor Green
} catch {
    Write-Host "[エラー] ファイル読み込み失敗: $_" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  1. /Title パターン検索結果" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

# /Title を検索
$titleMatches = [regex]::Matches($content, '/Title\s*[\(<][^\r\n]{0,200}')
Write-Host "/Title 検出数: $($titleMatches.Count) 件" -ForegroundColor Yellow
Write-Host ""

if ($titleMatches.Count -gt 0) {
    $count = 0
    foreach ($match in $titleMatches) {
        $count++
        if ($count -gt 10) {
            Write-Host "... 以下省略 (残り $($titleMatches.Count - 10) 件)" -ForegroundColor Gray
            break
        }

        $value = $match.Value
        Write-Host "[$count] $value" -ForegroundColor White

        # 16進数ダンプ
        $hexDump = ""
        $displayLen = [Math]::Min($value.Length, 50)
        for ($i = 0; $i -lt $displayLen; $i++) {
            $hexDump += "{0:X2} " -f [int][char]$value[$i]
        }
        Write-Host "    HEX: $hexDump" -ForegroundColor DarkGray
        Write-Host ""
    }
} else {
    Write-Host "/Title が見つかりませんでした" -ForegroundColor Red
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  2. BOM (FF FE / FE FF) パターン検索" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

# BOMパターンを検索（UTF-16LE: FF FE, UTF-16BE: FE FF）
$bomCount = 0
$bomSamples = @()

for ($i = 0; $i -lt $bytes.Length - 20; $i++) {
    $isBOM = $false
    $bomType = ""

    if ($bytes[$i] -eq 0xFF -and $bytes[$i+1] -eq 0xFE) {
        $isBOM = $true
        $bomType = "UTF-16LE"
    } elseif ($bytes[$i] -eq 0xFE -and $bytes[$i+1] -eq 0xFF) {
        $isBOM = $true
        $bomType = "UTF-16BE"
    }

    if ($isBOM) {
        $bomCount++
        if ($bomSamples.Count -lt 5) {
            # BOM後の20バイトを取得
            $sampleBytes = $bytes[($i)..([Math]::Min($i+21, $bytes.Length-1))]
            $hexDump = ($sampleBytes | ForEach-Object { "{0:X2}" -f $_ }) -join " "

            # UTF-16としてデコード試行
            $dataBytes = $bytes[($i+2)..([Math]::Min($i+41, $bytes.Length-1))]
            try {
                if ($bomType -eq "UTF-16LE") {
                    $decoded = [System.Text.Encoding]::Unicode.GetString($dataBytes)
                } else {
                    $decoded = [System.Text.Encoding]::BigEndianUnicode.GetString($dataBytes)
                }
                # 制御文字を除去
                $decoded = $decoded -replace '[\x00-\x1F]', ''
                $decoded = $decoded.Substring(0, [Math]::Min(30, $decoded.Length))
            } catch {
                $decoded = "(デコード失敗)"
            }

            $bomSamples += @{
                Position = $i
                Type = $bomType
                Hex = $hexDump
                Decoded = $decoded
            }
        }
    }
}

Write-Host "BOM検出数: $bomCount 件" -ForegroundColor Yellow
Write-Host ""

if ($bomSamples.Count -gt 0) {
    $sampleNum = 0
    foreach ($sample in $bomSamples) {
        $sampleNum++
        Write-Host "[$sampleNum] 位置: $($sample.Position), タイプ: $($sample.Type)" -ForegroundColor White
        Write-Host "    HEX: $($sample.Hex)" -ForegroundColor DarkGray
        Write-Host "    デコード: $($sample.Decoded)" -ForegroundColor Green
        Write-Host ""
    }
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  3. ObjStm (Object Streams) 詳細分析" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

# ObjStmオブジェクトを検索してストリームヘッダーを分析
$objStmPattern = '(\d+)\s+0\s+obj\s*<<([^>]*?/Type\s*/ObjStm[^>]*?)>>\s*stream\r?\n'
$objStmMatches = [regex]::Matches($content, $objStmPattern)
Write-Host "ObjStm検出数: $($objStmMatches.Count) 件" -ForegroundColor Yellow
Write-Host ""

if ($objStmMatches.Count -gt 0) {
    Write-Host "  -> しおりがObject Streams内に格納されている可能性があります" -ForegroundColor Yellow
    Write-Host ""

    # ストリームヘッダーの統計
    $headerStats = @{}
    $filterStats = @{}
    $sampleCount = 0

    foreach ($match in $objStmMatches) {
        $objNum = $match.Groups[1].Value
        $dictContent = $match.Groups[2].Value
        $streamStartPos = $match.Index + $match.Length

        # フィルター検出
        $filterMatch = [regex]::Match($dictContent, '/Filter\s*(/\w+|\[.*?\])')
        $filterType = if ($filterMatch.Success) { $filterMatch.Groups[1].Value } else { "なし" }

        if (-not $filterStats.ContainsKey($filterType)) {
            $filterStats[$filterType] = 0
        }
        $filterStats[$filterType]++

        # ストリームの先頭バイトを取得
        if ($streamStartPos -lt $bytes.Length - 10) {
            $headerBytes = $bytes[$streamStartPos..($streamStartPos + 9)]
            $headerHex = ($headerBytes[0..1] | ForEach-Object { "{0:X2}" -f $_ }) -join " "

            if (-not $headerStats.ContainsKey($headerHex)) {
                $headerStats[$headerHex] = 0
            }
            $headerStats[$headerHex]++

            # サンプル表示（最初の5件）
            if ($sampleCount -lt 5) {
                $sampleCount++
                $fullHeaderHex = ($headerBytes | ForEach-Object { "{0:X2}" -f $_ }) -join " "
                Write-Host "[$sampleCount] オブジェクト $objNum" -ForegroundColor White
                Write-Host "    フィルター: $filterType" -ForegroundColor Gray
                Write-Host "    ストリームヘッダー: $fullHeaderHex" -ForegroundColor Gray

                # ヘッダーの解釈
                $interpretation = switch ($headerHex) {
                    "78 9C" { "zlib (デフォルト圧縮)" }
                    "78 DA" { "zlib (最高圧縮)" }
                    "78 01" { "zlib (無圧縮)" }
                    "78 5E" { "zlib (低圧縮)" }
                    "1F 8B" { "gzip" }
                    default { "不明な形式" }
                }
                Write-Host "    解釈: $interpretation" -ForegroundColor $(if ($interpretation -eq "不明な形式") { "Red" } else { "Green" })
                Write-Host ""
            }
        }
    }

    # 統計サマリー
    Write-Host "--- ストリームヘッダー統計 ---" -ForegroundColor Cyan
    foreach ($key in $headerStats.Keys | Sort-Object) {
        $interpretation = switch ($key) {
            "78 9C" { "(zlib デフォルト)" }
            "78 DA" { "(zlib 最高)" }
            "78 01" { "(zlib 無圧縮)" }
            "78 5E" { "(zlib 低圧縮)" }
            "1F 8B" { "(gzip)" }
            default { "(不明)" }
        }
        Write-Host "  $key $interpretation : $($headerStats[$key]) 件" -ForegroundColor $(if ($interpretation -eq "(不明)") { "Red" } else { "White" })
    }

    Write-Host ""
    Write-Host "--- フィルター統計 ---" -ForegroundColor Cyan
    foreach ($key in $filterStats.Keys | Sort-Object) {
        Write-Host "  $key : $($filterStats[$key]) 件" -ForegroundColor White
    }
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  4. Outlines (しおり構造) 検索" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

$outlinesMatch = [regex]::Match($content, '/Type\s*/Outlines')
$firstMatch = [regex]::Match($content, '/First\s+\d+\s+\d+\s+R')
$countMatch = [regex]::Match($content, '/Count\s+-?\d+')

Write-Host "/Type /Outlines: $(if ($outlinesMatch.Success) { '検出' } else { '未検出' })" -ForegroundColor $(if ($outlinesMatch.Success) { 'Green' } else { 'Red' })
Write-Host "/First (子しおり参照): $(if ($firstMatch.Success) { $firstMatch.Value } else { '未検出' })" -ForegroundColor $(if ($firstMatch.Success) { 'Green' } else { 'Red' })
Write-Host "/Count (しおり数): $(if ($countMatch.Success) { $countMatch.Value } else { '未検出' })" -ForegroundColor $(if ($countMatch.Success) { 'Green' } else { 'Red' })

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  5. 日本語テキスト「第1部」検索" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

# 各エンコーディングで「第1部」を検索
$searchPatterns = @(
    @{ Name = "UTF-16BE"; Pattern = [byte[]]@(0x7B, 0x2C, 0x00, 0x31, 0x90, 0xE8) },  # 第1部
    @{ Name = "UTF-16LE"; Pattern = [byte[]]@(0x2C, 0x7B, 0x31, 0x00, 0xE8, 0x90) },  # 第1部
    @{ Name = "UTF-8"; Pattern = [byte[]]@(0xE7, 0xAC, 0xAC, 0x31, 0xE9, 0x83, 0xA8) },  # 第1部
    @{ Name = "Shift-JIS"; Pattern = [byte[]]@(0x91, 0xE6, 0x31, 0x95, 0x94) }  # 第1部
)

foreach ($sp in $searchPatterns) {
    $found = $false
    $foundPos = -1

    for ($i = 0; $i -lt $bytes.Length - $sp.Pattern.Length; $i++) {
        $match = $true
        for ($j = 0; $j -lt $sp.Pattern.Length; $j++) {
            if ($bytes[$i + $j] -ne $sp.Pattern[$j]) {
                $match = $false
                break
            }
        }
        if ($match) {
            $found = $true
            $foundPos = $i
            break
        }
    }

    if ($found) {
        Write-Host "$($sp.Name): 検出 (位置: $foundPos)" -ForegroundColor Green
        # 周辺のバイトを表示
        $contextStart = [Math]::Max(0, $foundPos - 5)
        $contextEnd = [Math]::Min($bytes.Length - 1, $foundPos + 20)
        $contextBytes = $bytes[$contextStart..$contextEnd]
        $hexContext = ($contextBytes | ForEach-Object { "{0:X2}" -f $_ }) -join " "
        Write-Host "    周辺HEX: $hexContext" -ForegroundColor DarkGray
    } else {
        Write-Host "$($sp.Name): 未検出" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  7. Adobe JavaScript形式しおり検索" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

# JavaScriptアクション形式のしおりを検索
$jsActionMatches = [regex]::Matches($content, '/S\s*/JavaScript\s*/JS\s*\([^)]{0,200}')
Write-Host "JavaScript形式アクション検出数: $($jsActionMatches.Count) 件" -ForegroundColor Yellow

if ($jsActionMatches.Count -gt 0) {
    Write-Host "  -> Adobe JavaScriptで作成されたしおりの可能性があります" -ForegroundColor Yellow
    Write-Host ""

    $jsCount = 0
    foreach ($match in $jsActionMatches) {
        $jsCount++
        if ($jsCount -gt 5) {
            Write-Host "... 以下省略 (残り $($jsActionMatches.Count - 5) 件)" -ForegroundColor Gray
            break
        }
        Write-Host "[$jsCount] $($match.Value.Substring(0, [Math]::Min(80, $match.Value.Length)))..." -ForegroundColor White
    }
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  8. XRef/XRefStm 検索" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

# XRef (通常の相互参照テーブル)
$xrefMatches = [regex]::Matches($content, '\bxref\b')
Write-Host "xref (通常形式) 検出数: $($xrefMatches.Count) 件" -ForegroundColor $(if ($xrefMatches.Count -gt 0) { 'Green' } else { 'Gray' })

# XRefStm (圧縮相互参照ストリーム)
$xrefStmMatches = [regex]::Matches($content, '/Type\s*/XRef')
Write-Host "XRefStm (圧縮形式) 検出数: $($xrefStmMatches.Count) 件" -ForegroundColor $(if ($xrefStmMatches.Count -gt 0) { 'Yellow' } else { 'Gray' })

if ($xrefStmMatches.Count -gt 0 -and $xrefMatches.Count -eq 0) {
    Write-Host ""
    Write-Host "  [重要] このPDFは圧縮XRefのみを使用しています" -ForegroundColor Red
    Write-Host "  -> オブジェクトの多くがObjStm内に格納されている可能性が高いです" -ForegroundColor Red
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Green
Write-Host "  調査完了" -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Green
Write-Host ""
Write-Host "上記の結果をスクリーンショットで共有してください。" -ForegroundColor Yellow
Write-Host ""
