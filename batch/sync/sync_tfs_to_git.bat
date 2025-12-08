<# :
@echo off
chcp 65001 >nul
title TFS to Git 同期スクリプト
setlocal

rem UNCパス対応（PushD/PopDで自動マッピング）
pushd "%~dp0"

powershell -NoProfile -ExecutionPolicy Bypass -Command "$scriptDir=('%~dp0' -replace '\\$',''); try { iex ((gc '%~f0' -Encoding UTF8) -join \"`n\") } finally { Set-Location C:\ }"
set EXITCODE=%ERRORLEVEL%

popd

pause
exit /b %EXITCODE%
: #>

# =============================================================================
# TFS to Git Sync Script (PowerShell)
# TFSのファイルをGitリポジトリに同期します
# =============================================================================

# タイトル表示
Write-Host ""
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host "  TFS to Git 同期スクリプト" -ForegroundColor Cyan
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host ""

# UTF-8出力設定
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Gitコマンドの存在確認
$gitCommand = Get-Command git -ErrorAction SilentlyContinue
if (-not $gitCommand) {
    Write-Host ""
    Write-Host "========================================================================" -ForegroundColor Red
    Write-Host "  [エラー] Gitがインストールされていません" -ForegroundColor Red
    Write-Host "========================================================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "このスクリプトを実行するには、Gitがインストールされている必要があります。" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Gitのインストール方法:" -ForegroundColor Cyan
    Write-Host "  1. https://git-scm.com/download/win にアクセス" -ForegroundColor White
    Write-Host "  2. 「Download for Windows」をクリック" -ForegroundColor White
    Write-Host "  3. インストーラーをダウンロードして実行" -ForegroundColor White
    Write-Host "  4. インストール後、コマンドプロンプトを再起動" -ForegroundColor White
    Write-Host ""
    exit 1
}

# Git日本語表示設定
git config --global core.quotepath false 2>&1 | Out-Null

#region 設定セクション
# TFSとGitのディレクトリパスを設定（固定値）
$TFS_DIR = "C:\Users\$env:USERNAME\source"
$GIT_REPO_DIR = "C:\Users\$env:USERNAME\source\Git\project"
#endregion

#region パスの存在確認
if (-not (Test-Path $TFS_DIR)) {
    Write-Host "[エラー] TFSディレクトリが見つかりません: $TFS_DIR" -ForegroundColor Red
    exit 1
}

if (-not (Test-Path $GIT_REPO_DIR)) {
    Write-Host "[エラー] Gitディレクトリが見つかりません: $GIT_REPO_DIR" -ForegroundColor Red
    exit 1
}

# Gitリポジトリ確認（親ディレクトリを遡って.gitフォルダを探す）
Push-Location $GIT_REPO_DIR
$isGitRepo = $false
try {
    # git rev-parseコマンドでGitリポジトリかどうかを確認
    git rev-parse --git-dir 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) {
        $isGitRepo = $true
    }
} catch {
    $isGitRepo = $false
}
Pop-Location

if (-not $isGitRepo) {
    Write-Host "[エラー] 指定されたディレクトリはGit管理下にありません: $GIT_REPO_DIR" -ForegroundColor Red
    exit 1
}
#endregion

Write-Host ""
Write-Host "TFSディレクトリ: $TFS_DIR" -ForegroundColor White
Write-Host "Gitディレクトリ: $GIT_REPO_DIR" -ForegroundColor White
Write-Host ""

# Gitディレクトリに移動
Set-Location $GIT_REPO_DIR

#region ブランチ操作
Write-Host "------------------------------------------------------------------------" -ForegroundColor Yellow
Write-Host " 現在のGitブランチ:" -ForegroundColor Yellow
Write-Host "------------------------------------------------------------------------" -ForegroundColor Yellow
git branch
Write-Host ""

# ブランチ操作メニュー
:branchLoop while ($true) {
    Write-Host "ブランチ操作を選択してください:" -ForegroundColor Cyan
    Write-Host " 1. このまま続行"
    Write-Host " 2. ブランチを切り替える"
    Write-Host ""
    Write-Host " 0. 終了"
    Write-Host ""
    $branchChoice = Read-Host "選択 (0-2)"

    switch ($branchChoice) {
        "0" {
            # 終了
            exit 0
        }
        "1" {
            # 同期処理へ
            break branchLoop
        }
        "2" {
            Write-Host ""
            Write-Host "------------------------------------------------------------------------" -ForegroundColor Yellow
            Write-Host " 利用可能なブランチ:" -ForegroundColor Yellow
            Write-Host "------------------------------------------------------------------------" -ForegroundColor Yellow

            # ローカルブランチ一覧を取得
            $branches = git branch --format="%(refname:short)" | ForEach-Object { $_.Trim() }

            if ($branches.Count -eq 0) {
                Write-Host "[エラー] ブランチが見つかりません" -ForegroundColor Red
                Write-Host ""
                continue
            }

            # ブランチを番号付きで表示
            for ($i = 0; $i -lt $branches.Count; $i++) {
                $displayNum = $i + 1
                $branch = $branches[$i]
                Write-Host " $displayNum. $branch"
            }
            Write-Host ""

            # ブランチ選択
            $maxNum = $branches.Count
            $selection = Read-Host "ブランチ番号を入力してください (1-$maxNum)"

            # 入力検証
            if ($selection -match '^\d+$' -and [int]$selection -ge 1 -and [int]$selection -le $maxNum) {
                $selectedBranch = $branches[[int]$selection - 1]

                git checkout $selectedBranch
                if ($LASTEXITCODE -ne 0) {
                    Write-Host "[エラー] ブランチの切り替えに失敗しました" -ForegroundColor Red
                    Write-Host ""
                    continue
                }
                Write-Host "ブランチを切り替えました: $selectedBranch" -ForegroundColor Green
                Write-Host ""
            } else {
                Write-Host "[エラー] 無効な番号です" -ForegroundColor Red
                Write-Host ""
            }
        }
        default {
            Write-Host "無効な選択です。" -ForegroundColor Red
        }
    }
}
#endregion

#region 同期処理
Write-Host ""
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host " 差分チェックを開始します" -ForegroundColor Cyan
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "ファイルをスキャン中..." -ForegroundColor Cyan
Write-Host ""

# TFSとGitのファイル一覧を取得
Write-Verbose "TFSディレクトリをスキャン中: $TFS_DIR"
$tfsFiles = Get-ChildItem -Path $TFS_DIR -Recurse -File -ErrorAction SilentlyContinue

Write-Verbose "Gitディレクトリをスキャン中: $GIT_REPO_DIR"
$gitFiles = Get-ChildItem -Path $GIT_REPO_DIR -Recurse -File -ErrorAction SilentlyContinue | Where-Object {
    $_.FullName -notlike '*\.git\*'
}

# ファイルを相対パスでハッシュテーブルに格納
$tfsFileDict = @{}
foreach ($file in $tfsFiles) {
    $relativePath = $file.FullName.Substring($TFS_DIR.Length).TrimStart('\')
    $tfsFileDict[$relativePath] = $file
}

$gitFileDict = @{}
foreach ($file in $gitFiles) {
    $relativePath = $file.FullName.Substring($GIT_REPO_DIR.Length).TrimStart('\')
    $gitFileDict[$relativePath] = $file
}

# 差分を格納する配列
$newFiles = @()      # TFSにあってGitにない（新規追加）
$updateFiles = @()   # 両方にあるが内容が異なる（更新）
$deleteFiles = @()   # GitにあってTFSにない（削除対象）
$identicalCount = 0

# TFSファイルをチェック（更新 & 新規追加）
foreach ($relativePath in $tfsFileDict.Keys) {
    $tfsFile = $tfsFileDict[$relativePath]
    $gitFilePath = Join-Path $GIT_REPO_DIR $relativePath

    if (Test-Path $gitFilePath) {
        # ファイルが両方に存在 → MD5ハッシュで比較
        try {
            $tfsHash = (Get-FileHash -Path $tfsFile.FullName -Algorithm MD5).Hash
            $gitHash = (Get-FileHash -Path $gitFilePath -Algorithm MD5).Hash

            if ($tfsHash -ne $gitHash) {
                # ハッシュが異なる → 更新対象
                $updateFiles += [PSCustomObject]@{
                    RelativePath = $relativePath
                    TfsFile = $tfsFile
                    GitFilePath = $gitFilePath
                }
            } else {
                # ハッシュが同じ → 変更なし
                $identicalCount++
            }
        } catch {
            Write-Warning "ファイルハッシュ取得エラー: $relativePath - $_"
        }
    } else {
        # Gitに存在しない → 新規追加対象
        $newFiles += [PSCustomObject]@{
            RelativePath = $relativePath
            TfsFile = $tfsFile
            GitFilePath = $gitFilePath
        }
    }
}

# Gitのみのファイルをチェック（削除対象）
foreach ($relativePath in $gitFileDict.Keys) {
    if (-not $tfsFileDict.ContainsKey($relativePath)) {
        $gitFile = $gitFileDict[$relativePath]
        $deleteFiles += [PSCustomObject]@{
            RelativePath = $relativePath
            GitFile = $gitFile
        }
    }
}

# 差分サマリー表示
Write-Host "========================================================================" -ForegroundColor Yellow
Write-Host " 差分サマリー" -ForegroundColor Yellow
Write-Host "========================================================================" -ForegroundColor Yellow
Write-Host ""
Write-Host "新規ファイル (TFS → Git): " -NoNewline -ForegroundColor Green
Write-Host "$($newFiles.Count) 件"
Write-Host "更新ファイル (TFS → Git): " -NoNewline -ForegroundColor Yellow
Write-Host "$($updateFiles.Count) 件"
Write-Host "削除対象 (Gitのみ):       " -NoNewline -ForegroundColor Red
Write-Host "$($deleteFiles.Count) 件"
Write-Host "変更なし:                 " -NoNewline -ForegroundColor Gray
Write-Host "$identicalCount 件"
Write-Host ""

# 差分がない場合は終了
if ($newFiles.Count -eq 0 -and $updateFiles.Count -eq 0 -and $deleteFiles.Count -eq 0) {
    Write-Host "差分はありません。ファイルは同期されています。" -ForegroundColor Green
    Write-Host ""
    exit 0
}

# 差分詳細を表示
Write-Host "------------------------------------------------------------------------" -ForegroundColor Yellow
Write-Host " 差分詳細" -ForegroundColor Yellow
Write-Host "------------------------------------------------------------------------" -ForegroundColor Yellow
Write-Host ""

if ($newFiles.Count -gt 0) {
    Write-Host "=== 新規ファイル (TFSからGitへコピー) ===" -ForegroundColor Green
    foreach ($file in $newFiles) {
        Write-Host "  [新規] $($file.RelativePath)" -ForegroundColor Green
    }
    Write-Host ""
}

if ($updateFiles.Count -gt 0) {
    Write-Host "=== 更新ファイル (TFSの内容でGitを上書き) ===" -ForegroundColor Yellow
    foreach ($file in $updateFiles) {
        Write-Host "  [更新] $($file.RelativePath)" -ForegroundColor Yellow
    }
    Write-Host ""
}

if ($deleteFiles.Count -gt 0) {
    Write-Host "=== 削除対象ファイル (Gitから削除) ===" -ForegroundColor Red
    foreach ($file in $deleteFiles) {
        Write-Host "  [削除] $($file.RelativePath)" -ForegroundColor Red
    }
    Write-Host ""
}

# マージ確認
Write-Host "------------------------------------------------------------------------" -ForegroundColor Cyan
Write-Host " 上記の変更をGitに反映しますか？" -ForegroundColor Cyan
Write-Host " ※ TFSの内容でGitを同期します（TFS → Git）" -ForegroundColor White
Write-Host "------------------------------------------------------------------------" -ForegroundColor Cyan
Write-Host ""
Write-Host " 1. はい、同期を実行する"
Write-Host " 2. いいえ、キャンセルする"
Write-Host ""
$syncChoice = Read-Host "選択 (1-2)"

if ($syncChoice -ne "1") {
    Write-Host ""
    Write-Host "同期をキャンセルしました。" -ForegroundColor Yellow
    exit 0
}

# 同期実行
Write-Host ""
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host " 同期を実行中..." -ForegroundColor Cyan
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host ""

# 統計カウンタ
$copiedCount = 0
$deletedCount = 0

# 新規ファイルをコピー
foreach ($file in $newFiles) {
    try {
        $targetDir = Split-Path -Path $file.GitFilePath -Parent
        if (-not (Test-Path $targetDir)) {
            New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
        }
        Copy-Item -Path $file.TfsFile.FullName -Destination $file.GitFilePath -Force
        Write-Host "[コピー完了] $($file.RelativePath)" -ForegroundColor Green
        $copiedCount++
    } catch {
        Write-Warning "コピーエラー: $($file.RelativePath) - $_"
    }
}

# 更新ファイルをコピー
foreach ($file in $updateFiles) {
    try {
        $targetDir = Split-Path -Path $file.GitFilePath -Parent
        if (-not (Test-Path $targetDir)) {
            New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
        }
        Copy-Item -Path $file.TfsFile.FullName -Destination $file.GitFilePath -Force
        Write-Host "[更新完了] $($file.RelativePath)" -ForegroundColor Yellow
        $copiedCount++
    } catch {
        Write-Warning "更新エラー: $($file.RelativePath) - $_"
    }
}

# 削除ファイルを削除
foreach ($file in $deleteFiles) {
    try {
        Remove-Item -Path $file.GitFile.FullName -Force
        Write-Host "[削除完了] $($file.RelativePath)" -ForegroundColor Red
        $deletedCount++
    } catch {
        Write-Warning "削除エラー: $($file.RelativePath) - $_"
    }
}

Write-Host ""
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host " 同期完了" -ForegroundColor Cyan
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "コピー/更新ファイル: $copiedCount" -ForegroundColor Green
Write-Host "削除ファイル: $deletedCount" -ForegroundColor Red
Write-Host ""
#endregion

#region Gitステータス確認とコミット
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host " Gitステータスを確認してください" -ForegroundColor Cyan
Write-Host "========================================================================" -ForegroundColor Cyan
git status

Write-Host ""
Write-Host "------------------------------------------------------------------------" -ForegroundColor Yellow
Write-Host "次の操作を選択してください:" -ForegroundColor Cyan
Write-Host " 1. 変更をコミットする"
Write-Host ""
Write-Host " 0. 何もせず終了"
Write-Host "------------------------------------------------------------------------" -ForegroundColor Yellow
$commitChoice = Read-Host "選択 (0-1)"

if ($commitChoice -eq "0") {
    Write-Host ""
    Write-Host "処理を終了します。" -ForegroundColor Yellow
    exit 0
}

if ($commitChoice -eq "1") {
    Write-Host ""
    $commitMsg = Read-Host "コミットメッセージを入力してください"

    git add -A
    git commit -m $commitMsg

    if ($LASTEXITCODE -ne 0) {
        Write-Host "[警告] コミットに失敗しました、または変更がありませんでした" -ForegroundColor Yellow
    } else {
        Write-Host "コミットが完了しました" -ForegroundColor Green
    }
}
#endregion

Write-Host ""
Write-Host "処理が完了しました。" -ForegroundColor Cyan
exit 0
