<# :
@echo off
chcp 65001 >nul
title ファイル拡張子解析ツール
setlocal

rem UNCパス対応（PushD/PopDで自動マッピング）
pushd "%~dp0"

powershell -NoProfile -ExecutionPolicy Bypass -Command "$scriptDir=('%~dp0' -replace '\\$',''); try { iex ((gc '%~f0' -Encoding UTF8) -join \"`n\") } finally { Set-Location C:\ }"
set EXITCODE=%ERRORLEVEL%

popd

pause
exit /b %EXITCODE%
: #>

#==============================================================================
# ファイル拡張子解析ツール
#   指定フォルダとそのサブフォルダ内のファイル拡張子を一覧表示します
#==============================================================================

#region タイトル表示
Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  ファイル拡張子解析ツール" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "指定フォルダ内の拡張子の種類とファイル数を解析します。" -ForegroundColor White
Write-Host ""
#endregion

#region フォルダ選択
function Select-FolderDialog {
    param (
        [string]$Description = "フォルダを選択してください"
    )

    Add-Type -AssemblyName System.Windows.Forms
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = $Description
    $folderBrowser.ShowNewFolderButton = $false

    $result = $folderBrowser.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $folderBrowser.SelectedPath
    }
    return $null
}

Write-Host "フォルダの指定方法を選択してください:" -ForegroundColor Yellow
Write-Host " 1. フォルダ選択ダイアログを使用"
Write-Host " 2. パスを直接入力"
Write-Host ""
Write-Host " 0. キャンセル"
Write-Host ""
$choice = Read-Host "選択 (0-2)"

$targetPath = $null

switch ($choice) {
    "1" {
        Write-Host ""
        Write-Host "フォルダ選択ダイアログを開いています..." -ForegroundColor Gray
        $targetPath = Select-FolderDialog -Description "解析対象のフォルダを選択してください"
    }
    "2" {
        Write-Host ""
        $targetPath = Read-Host "フォルダパスを入力してください"
        $targetPath = $targetPath.Trim('"', "'", " ")
    }
    "0" {
        Write-Host ""
        Write-Host "キャンセルしました。" -ForegroundColor Yellow
        exit 0
    }
    default {
        Write-Host ""
        Write-Host "[エラー] 無効な選択です。" -ForegroundColor Red
        exit 1
    }
}

if ([string]::IsNullOrWhiteSpace($targetPath)) {
    Write-Host ""
    Write-Host "[エラー] フォルダが選択されませんでした。" -ForegroundColor Red
    exit 1
}

if (-not (Test-Path -Path $targetPath -PathType Container)) {
    Write-Host ""
    Write-Host "[エラー] 指定されたパスが存在しないか、フォルダではありません: $targetPath" -ForegroundColor Red
    exit 1
}
#endregion

#region 拡張子解析
Write-Host ""
Write-Host "----------------------------------------------------------------" -ForegroundColor Gray
Write-Host "解析対象: $targetPath" -ForegroundColor White
Write-Host "----------------------------------------------------------------" -ForegroundColor Gray
Write-Host ""
Write-Host "ファイルを検索しています..." -ForegroundColor Gray

try {
    # サブフォルダを含めてすべてのファイルを取得
    $files = Get-ChildItem -Path $targetPath -Recurse -File -ErrorAction SilentlyContinue

    if ($null -eq $files -or $files.Count -eq 0) {
        Write-Host ""
        Write-Host "[情報] ファイルが見つかりませんでした。" -ForegroundColor Yellow
        exit 0
    }

    $totalFiles = $files.Count
    Write-Host "検出ファイル数: $totalFiles 件" -ForegroundColor Green
    Write-Host ""

    # 拡張子ごとにグループ化
    $extensionGroups = $files | Group-Object {
        if ([string]::IsNullOrWhiteSpace($_.Extension)) {
            "(拡張子なし)"
        } else {
            $_.Extension.ToLower()
        }
    } | Sort-Object Count -Descending

    # 結果表示
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host "  拡張子一覧 (ファイル数降順)" -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""

    # ヘッダー
    $header = "{0,-20} {1,10} {2,10}" -f "拡張子", "ファイル数", "割合"
    Write-Host $header -ForegroundColor White
    Write-Host ("-" * 42) -ForegroundColor Gray

    foreach ($group in $extensionGroups) {
        $percentage = [math]::Round(($group.Count / $totalFiles) * 100, 1)
        $row = "{0,-20} {1,10} {2,9}%" -f $group.Name, $group.Count, $percentage

        # 割合に応じて色分け
        if ($percentage -ge 30) {
            Write-Host $row -ForegroundColor Yellow
        } elseif ($percentage -ge 10) {
            Write-Host $row -ForegroundColor Green
        } else {
            Write-Host $row -ForegroundColor White
        }
    }

    Write-Host ("-" * 42) -ForegroundColor Gray
    $totalRow = "{0,-20} {1,10} {2,9}%" -f "合計", $totalFiles, "100.0"
    Write-Host $totalRow -ForegroundColor Cyan

    # サマリー
    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host "  サマリー" -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  総ファイル数    : $totalFiles 件" -ForegroundColor White
    Write-Host "  拡張子の種類数  : $($extensionGroups.Count) 種類" -ForegroundColor White

    # フォルダ数をカウント
    $folders = Get-ChildItem -Path $targetPath -Recurse -Directory -ErrorAction SilentlyContinue
    $folderCount = if ($null -ne $folders) { $folders.Count } else { 0 }
    Write-Host "  サブフォルダ数  : $folderCount 件" -ForegroundColor White

    # 上位3拡張子を表示
    if ($extensionGroups.Count -ge 3) {
        Write-Host ""
        Write-Host "  上位3拡張子:" -ForegroundColor Yellow
        for ($i = 0; $i -lt 3; $i++) {
            $ext = $extensionGroups[$i]
            $pct = [math]::Round(($ext.Count / $totalFiles) * 100, 1)
            Write-Host "    $($i + 1). $($ext.Name) ($($ext.Count) 件, $pct%)" -ForegroundColor White
        }
    }

    Write-Host ""

    #region 詳細表示オプション
    Write-Host "詳細情報を表示しますか?" -ForegroundColor Yellow
    Write-Host " 1. 拡張子別ファイル一覧を表示"
    Write-Host " 2. フォルダ別に拡張子を表示"
    Write-Host " 3. CSVファイルにエクスポート"
    Write-Host ""
    Write-Host " 0. 終了"
    Write-Host ""
    $detailChoice = Read-Host "選択 (0-3)"

    switch ($detailChoice) {
        "1" {
            Write-Host ""
            Write-Host "表示する拡張子を入力してください (例: .txt, .pdf):" -ForegroundColor Yellow
            $extFilter = Read-Host "拡張子"
            $extFilter = $extFilter.Trim().ToLower()

            if (-not $extFilter.StartsWith(".") -and $extFilter -ne "(拡張子なし)") {
                $extFilter = "." + $extFilter
            }

            $filteredFiles = $files | Where-Object {
                if ($extFilter -eq "(拡張子なし)") {
                    [string]::IsNullOrWhiteSpace($_.Extension)
                } else {
                    $_.Extension.ToLower() -eq $extFilter
                }
            }

            if ($filteredFiles.Count -eq 0) {
                Write-Host ""
                Write-Host "[情報] 該当するファイルがありません。" -ForegroundColor Yellow
            } else {
                Write-Host ""
                Write-Host "================================================================" -ForegroundColor Cyan
                Write-Host "  $extFilter ファイル一覧 ($($filteredFiles.Count) 件)" -ForegroundColor Cyan
                Write-Host "================================================================" -ForegroundColor Cyan
                Write-Host ""

                foreach ($file in ($filteredFiles | Sort-Object FullName)) {
                    # 相対パスで表示
                    $relativePath = $file.FullName.Substring($targetPath.Length).TrimStart('\', '/')
                    Write-Host "  $relativePath" -ForegroundColor White
                }
            }
        }
        "2" {
            # フォルダ別に拡張子を表示
            Write-Host ""
            Write-Host "================================================================" -ForegroundColor Cyan
            Write-Host "  フォルダ別 拡張子一覧" -ForegroundColor Cyan
            Write-Host "================================================================" -ForegroundColor Cyan
            Write-Host ""

            # ルートフォルダのファイルを処理
            $rootFiles = $files | Where-Object { $_.DirectoryName -eq $targetPath }
            if ($rootFiles.Count -gt 0) {
                $rootExtensions = ($rootFiles | ForEach-Object {
                    if ([string]::IsNullOrWhiteSpace($_.Extension)) { "(拡張子なし)" } else { $_.Extension.ToLower() }
                } | Sort-Object -Unique) -join ", "
                Write-Host "[ルート]" -ForegroundColor Yellow
                Write-Host "  拡張子: $rootExtensions" -ForegroundColor White
                Write-Host "  ファイル数: $($rootFiles.Count)" -ForegroundColor Gray
                Write-Host ""
            }

            # サブフォルダごとにグループ化
            $folderGroups = $files | Group-Object DirectoryName | Where-Object { $_.Name -ne $targetPath } | Sort-Object Name

            foreach ($folder in $folderGroups) {
                # 相対パスで表示
                $relativeFolderPath = $folder.Name.Substring($targetPath.Length).TrimStart('\', '/')

                # このフォルダ内の拡張子を取得
                $folderExtensions = ($folder.Group | ForEach-Object {
                    if ([string]::IsNullOrWhiteSpace($_.Extension)) { "(拡張子なし)" } else { $_.Extension.ToLower() }
                } | Sort-Object -Unique) -join ", "

                Write-Host "[$relativeFolderPath]" -ForegroundColor Yellow
                Write-Host "  拡張子: $folderExtensions" -ForegroundColor White
                Write-Host "  ファイル数: $($folder.Count)" -ForegroundColor Gray
                Write-Host ""
            }
        }
        "3" {
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $csvPath = Join-Path -Path $scriptDir -ChildPath "拡張子解析結果_$timestamp.csv"

            # CSV出力用データ作成
            $csvData = @()
            foreach ($group in $extensionGroups) {
                $percentage = [math]::Round(($group.Count / $totalFiles) * 100, 1)
                $csvData += [PSCustomObject]@{
                    "拡張子" = $group.Name
                    "ファイル数" = $group.Count
                    "割合(%)" = $percentage
                }
            }

            # 合計行を追加
            $csvData += [PSCustomObject]@{
                "拡張子" = "合計"
                "ファイル数" = $totalFiles
                "割合(%)" = 100.0
            }

            # BOM付きUTF-8でCSV出力
            $csvData | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

            Write-Host ""
            Write-Host "[成功] CSVファイルを出力しました:" -ForegroundColor Green
            Write-Host "  $csvPath" -ForegroundColor White
        }
        "0" {
            Write-Host ""
            Write-Host "終了します。" -ForegroundColor Gray
        }
        default {
            Write-Host ""
            Write-Host "終了します。" -ForegroundColor Gray
        }
    }
    #endregion

} catch {
    Write-Host ""
    Write-Host "[エラー] 解析中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
#endregion

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  処理完了" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
