# =============================================================================
# TFS to Git Sync Logic (PowerShell)
# TFSとGitディレクトリを同期するPowerShellスクリプト
# =============================================================================

param(
    [Parameter(Mandatory=$true)]
    [string]$TfsDir,

    [Parameter(Mandatory=$true)]
    [string]$GitDir
)

# UTF-8出力設定
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host '差分チェック中...' -ForegroundColor Cyan
Write-Host ''

# TFSとGitのファイル一覧を取得
Write-Verbose "TFSディレクトリをスキャン中: $TfsDir"
$tfsFiles = Get-ChildItem -Path $TfsDir -Recurse -File -ErrorAction SilentlyContinue

Write-Verbose "Gitディレクトリをスキャン中: $GitDir"
$gitFiles = Get-ChildItem -Path $GitDir -Recurse -File -ErrorAction SilentlyContinue | Where-Object {
    $_.FullName -notlike '*\.git\*'
}

# ファイルを相対パスでハッシュテーブルに格納
$tfsFileDict = @{}
foreach ($file in $tfsFiles) {
    $relativePath = $file.FullName.Substring($TfsDir.Length).TrimStart('\')
    $tfsFileDict[$relativePath] = $file
}

$gitFileDict = @{}
foreach ($file in $gitFiles) {
    $relativePath = $file.FullName.Substring($GitDir.Length).TrimStart('\')
    $gitFileDict[$relativePath] = $file
}

# 統計カウンタ
$copiedCount = 0
$deletedCount = 0
$identicalCount = 0

Write-Host '=== ファイル差分チェック ===' -ForegroundColor Yellow
Write-Host ''

# TFSファイルをチェック（更新 & 新規追加）
foreach ($relativePath in $tfsFileDict.Keys) {
    $tfsFile = $tfsFileDict[$relativePath]
    $gitFilePath = Join-Path $GitDir $relativePath

    if (Test-Path $gitFilePath) {
        # ファイルが両方に存在 → MD5ハッシュで比較
        try {
            $tfsHash = (Get-FileHash -Path $tfsFile.FullName -Algorithm MD5).Hash
            $gitHash = (Get-FileHash -Path $gitFilePath -Algorithm MD5).Hash

            if ($tfsHash -ne $gitHash) {
                # ハッシュが異なる → 更新
                Write-Host '[更新] ' -ForegroundColor Yellow -NoNewline
                Write-Host $relativePath

                $targetDir = Split-Path -Path $gitFilePath -Parent
                if (-not (Test-Path $targetDir)) {
                    New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
                }

                Copy-Item -Path $tfsFile.FullName -Destination $gitFilePath -Force
                $copiedCount++
            } else {
                # ハッシュが同じ → 変更なし
                $identicalCount++
            }
        } catch {
            Write-Warning "ファイルハッシュ取得エラー: $relativePath - $_"
        }
    } else {
        # Gitに存在しない → 新規追加
        Write-Host '[新規] ' -ForegroundColor Green -NoNewline
        Write-Host $relativePath

        $targetDir = Split-Path -Path $gitFilePath -Parent
        if (-not (Test-Path $targetDir)) {
            New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
        }

        Copy-Item -Path $tfsFile.FullName -Destination $gitFilePath -Force
        $copiedCount++
    }
}

Write-Host ''
Write-Host '=== Gitのみに存在するファイル (削除対象) ===' -ForegroundColor Yellow
Write-Host ''

# Gitのみのファイルをチェック（削除）
foreach ($relativePath in $gitFileDict.Keys) {
    if (-not $tfsFileDict.ContainsKey($relativePath)) {
        $gitFile = $gitFileDict[$relativePath]
        Write-Host '[削除] ' -ForegroundColor Red -NoNewline
        Write-Host $relativePath

        try {
            Remove-Item -Path $gitFile.FullName -Force
            $deletedCount++
        } catch {
            Write-Warning "ファイル削除エラー: $relativePath - $_"
        }
    }
}

Write-Host ''
Write-Host '========================================' -ForegroundColor Cyan
Write-Host '同期完了' -ForegroundColor Cyan
Write-Host '========================================' -ForegroundColor Cyan
Write-Host ''
Write-Host "更新/新規ファイル: $copiedCount" -ForegroundColor Green
Write-Host "削除ファイル: $deletedCount" -ForegroundColor Red
Write-Host "変更なし: $identicalCount" -ForegroundColor Gray
Write-Host ''

# 終了コード: 0=成功
exit 0
