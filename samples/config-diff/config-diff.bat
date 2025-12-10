<# :
@echo off
chcp 65001 >nul
title 設定ファイル比較ツール
setlocal

rem UNCパス対応（PushD/PopDで自動マッピング）
pushd "%~dp0"

powershell -NoProfile -ExecutionPolicy Bypass -Command "$scriptDir=('%~dp0' -replace '\\$',''); try { iex ((gc '%~f0' -Encoding UTF8) -join \"`n\") } finally { Set-Location C:\ }"
set EXITCODE=%ERRORLEVEL%

popd

pause
exit /b %EXITCODE%
: #>

<#
.SYNOPSIS
    環境間の設定ファイル比較ツール

.DESCRIPTION
    2つのフォルダ内の設定ファイルを比較し、差分を検出します。
    対応形式: .ini, .json, .xml, .properties, .conf, .cfg, .yaml, .yml

.NOTES
    作成日: 2025-12-10
    バージョン: 1.0
#>

# ==============================================================================
# ■ 設定セクション（ここを編集してください）
# ==============================================================================

$Config = @{
    # 比較元フォルダ（本番環境など）
    SourceFolder = "C:\Config\Production"

    # 比較先フォルダ（開発環境など）
    TargetFolder = "C:\Config\Development"

    # 比較対象の拡張子（空の場合はすべて）
    Extensions = @(".ini", ".json", ".xml", ".properties", ".conf", ".cfg", ".yaml", ".yml")

    # 無視するファイル名パターン（正規表現）
    IgnorePatterns = @("\.bak$", "\.backup$", "~$")

    # 結果をファイル出力するか
    ExportResult = $true

    # 出力先フォルダ（空の場合はデスクトップ）
    OutputFolder = ""
}

# ==============================================================================
# ■ メイン処理（以下は編集不要）
# ==============================================================================

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ヘッダー表示
Write-Host ""
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host "  設定ファイル比較ツール" -ForegroundColor Cyan
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host ""

#region フォルダ存在確認
if (-not (Test-Path $Config.SourceFolder)) {
    Write-Host "[エラー] 比較元フォルダが見つかりません: $($Config.SourceFolder)" -ForegroundColor Red
    exit 1
}

if (-not (Test-Path $Config.TargetFolder)) {
    Write-Host "[エラー] 比較先フォルダが見つかりません: $($Config.TargetFolder)" -ForegroundColor Red
    exit 1
}

Write-Host "比較元: $($Config.SourceFolder)" -ForegroundColor White
Write-Host "比較先: $($Config.TargetFolder)" -ForegroundColor White
Write-Host ""
#endregion

#region ファイル収集
Write-Host "ファイルを収集中..." -ForegroundColor Cyan

# 比較元ファイル取得
$sourceFiles = Get-ChildItem -Path $Config.SourceFolder -Recurse -File | Where-Object {
    $file = $_
    # 拡張子フィルタ
    $extMatch = ($Config.Extensions.Count -eq 0) -or ($Config.Extensions -contains $file.Extension.ToLower())
    # 無視パターンチェック
    $ignored = $false
    foreach ($pattern in $Config.IgnorePatterns) {
        if ($file.Name -match $pattern) {
            $ignored = $true
            break
        }
    }
    $extMatch -and (-not $ignored)
}

# 比較先ファイル取得
$targetFiles = Get-ChildItem -Path $Config.TargetFolder -Recurse -File | Where-Object {
    $file = $_
    $extMatch = ($Config.Extensions.Count -eq 0) -or ($Config.Extensions -contains $file.Extension.ToLower())
    $ignored = $false
    foreach ($pattern in $Config.IgnorePatterns) {
        if ($file.Name -match $pattern) {
            $ignored = $true
            break
        }
    }
    $extMatch -and (-not $ignored)
}

# 相対パスでハッシュテーブル作成
$sourceDict = @{}
foreach ($file in $sourceFiles) {
    $relativePath = $file.FullName.Substring($Config.SourceFolder.Length).TrimStart('\')
    $sourceDict[$relativePath] = $file
}

$targetDict = @{}
foreach ($file in $targetFiles) {
    $relativePath = $file.FullName.Substring($Config.TargetFolder.Length).TrimStart('\')
    $targetDict[$relativePath] = $file
}

Write-Host "比較元: $($sourceDict.Count) ファイル" -ForegroundColor Green
Write-Host "比較先: $($targetDict.Count) ファイル" -ForegroundColor Green
Write-Host ""
#endregion

#region 比較処理
Write-Host "比較中..." -ForegroundColor Cyan
Write-Host ""

$results = @{
    OnlyInSource = @()    # 比較元のみ
    OnlyInTarget = @()    # 比較先のみ
    Different = @()       # 両方にあるが内容が異なる
    Identical = @()       # 同一
}

# 比較元ファイルをチェック
foreach ($relativePath in $sourceDict.Keys) {
    $sourceFile = $sourceDict[$relativePath]

    if ($targetDict.ContainsKey($relativePath)) {
        # 両方に存在 → 内容比較
        $targetFile = $targetDict[$relativePath]

        try {
            $sourceHash = (Get-FileHash -Path $sourceFile.FullName -Algorithm MD5).Hash
            $targetHash = (Get-FileHash -Path $targetFile.FullName -Algorithm MD5).Hash

            if ($sourceHash -eq $targetHash) {
                $results.Identical += $relativePath
            } else {
                $results.Different += [PSCustomObject]@{
                    Path = $relativePath
                    SourceFile = $sourceFile.FullName
                    TargetFile = $targetFile.FullName
                }
            }
        } catch {
            Write-Host "[警告] 比較エラー: $relativePath - $($_.Exception.Message)" -ForegroundColor Yellow
        }
    } else {
        # 比較元のみ
        $results.OnlyInSource += $relativePath
    }
}

# 比較先のみのファイルをチェック
foreach ($relativePath in $targetDict.Keys) {
    if (-not $sourceDict.ContainsKey($relativePath)) {
        $results.OnlyInTarget += $relativePath
    }
}
#endregion

#region 結果表示
Write-Host "========================================================================" -ForegroundColor Yellow
Write-Host " 比較結果" -ForegroundColor Yellow
Write-Host "========================================================================" -ForegroundColor Yellow
Write-Host ""

# サマリー
Write-Host "同一ファイル      : $($results.Identical.Count) 件" -ForegroundColor Green
Write-Host "差分あり          : $($results.Different.Count) 件" -ForegroundColor Yellow
Write-Host "比較元のみ        : $($results.OnlyInSource.Count) 件" -ForegroundColor Cyan
Write-Host "比較先のみ        : $($results.OnlyInTarget.Count) 件" -ForegroundColor Magenta
Write-Host ""

# 差分詳細
if ($results.Different.Count -gt 0) {
    Write-Host "--- 差分のあるファイル ---" -ForegroundColor Yellow
    foreach ($item in $results.Different) {
        Write-Host "  [差分] $($item.Path)" -ForegroundColor Yellow
    }
    Write-Host ""
}

if ($results.OnlyInSource.Count -gt 0) {
    Write-Host "--- 比較元のみに存在 ---" -ForegroundColor Cyan
    foreach ($path in $results.OnlyInSource) {
        Write-Host "  [比較元のみ] $path" -ForegroundColor Cyan
    }
    Write-Host ""
}

if ($results.OnlyInTarget.Count -gt 0) {
    Write-Host "--- 比較先のみに存在 ---" -ForegroundColor Magenta
    foreach ($path in $results.OnlyInTarget) {
        Write-Host "  [比較先のみ] $path" -ForegroundColor Magenta
    }
    Write-Host ""
}
#endregion

#region 結果出力
if ($Config.ExportResult -and ($results.Different.Count -gt 0 -or $results.OnlyInSource.Count -gt 0 -or $results.OnlyInTarget.Count -gt 0)) {
    $outputFolder = if ($Config.OutputFolder -ne "") { $Config.OutputFolder } else { "$env:USERPROFILE\Desktop" }
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $outputFile = Join-Path $outputFolder "config-diff_$timestamp.txt"

    $output = @()
    $output += "設定ファイル比較結果"
    $output += "===================="
    $output += "実行日時: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $output += "比較元: $($Config.SourceFolder)"
    $output += "比較先: $($Config.TargetFolder)"
    $output += ""
    $output += "[サマリー]"
    $output += "同一ファイル: $($results.Identical.Count) 件"
    $output += "差分あり: $($results.Different.Count) 件"
    $output += "比較元のみ: $($results.OnlyInSource.Count) 件"
    $output += "比較先のみ: $($results.OnlyInTarget.Count) 件"
    $output += ""

    if ($results.Different.Count -gt 0) {
        $output += "[差分のあるファイル]"
        foreach ($item in $results.Different) {
            $output += "  $($item.Path)"
        }
        $output += ""
    }

    if ($results.OnlyInSource.Count -gt 0) {
        $output += "[比較元のみに存在]"
        foreach ($path in $results.OnlyInSource) {
            $output += "  $path"
        }
        $output += ""
    }

    if ($results.OnlyInTarget.Count -gt 0) {
        $output += "[比較先のみに存在]"
        foreach ($path in $results.OnlyInTarget) {
            $output += "  $path"
        }
        $output += ""
    }

    $output | Out-File -FilePath $outputFile -Encoding UTF8
    Write-Host "結果を出力しました: $outputFile" -ForegroundColor Green
}
#endregion

#region WinMerge連携
if ($results.Different.Count -gt 0) {
    Write-Host ""
    Write-Host "差分のあるファイルをWinMergeで比較しますか？" -ForegroundColor Cyan
    Write-Host " 1. 最初の差分ファイルを比較"
    Write-Host " 2. フォルダ全体を比較"
    Write-Host ""
    Write-Host " 0. 終了"
    Write-Host ""

    $choice = Read-Host "選択 (0-2)"

    # WinMergeパス検出
    $winmergePath = ""
    $possiblePaths = @(
        "${env:ProgramFiles}\WinMerge\WinMergeU.exe",
        "${env:ProgramFiles(x86)}\WinMerge\WinMergeU.exe",
        "${env:LOCALAPPDATA}\Programs\WinMerge\WinMergeU.exe"
    )
    foreach ($path in $possiblePaths) {
        if (Test-Path $path) {
            $winmergePath = $path
            break
        }
    }

    if ($winmergePath -eq "") {
        Write-Host "[情報] WinMergeが見つかりません" -ForegroundColor Yellow
    } elseif ($choice -eq "1") {
        $firstDiff = $results.Different[0]
        & $winmergePath $firstDiff.SourceFile $firstDiff.TargetFile
    } elseif ($choice -eq "2") {
        & $winmergePath "/r" $Config.SourceFolder $Config.TargetFolder
    }
}
#endregion

Write-Host ""
Write-Host "処理が完了しました。" -ForegroundColor Green
exit 0
