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
    Write-Host " 3. 終了"
    Write-Host ""
    $branchChoice = Read-Host "選択 (1-3)"

    switch ($branchChoice) {
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
        "3" {
            exit 0
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
Write-Host " 同期処理を開始します" -ForegroundColor Cyan
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "差分チェック中..." -ForegroundColor Cyan
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

# 統計カウンタ
$copiedCount = 0
$deletedCount = 0
$identicalCount = 0

Write-Host "=== ファイル差分チェック ===" -ForegroundColor Yellow
Write-Host ""

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
                # ハッシュが異なる → 更新
                Write-Host "[更新] " -ForegroundColor Yellow -NoNewline
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
        Write-Host "[新規] " -ForegroundColor Green -NoNewline
        Write-Host $relativePath

        $targetDir = Split-Path -Path $gitFilePath -Parent
        if (-not (Test-Path $targetDir)) {
            New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
        }

        Copy-Item -Path $tfsFile.FullName -Destination $gitFilePath -Force
        $copiedCount++
    }
}

Write-Host ""
Write-Host "=== Gitのみに存在するファイル (削除対象) ===" -ForegroundColor Yellow
Write-Host ""

# Gitのみのファイルをチェック（削除）
foreach ($relativePath in $gitFileDict.Keys) {
    if (-not $tfsFileDict.ContainsKey($relativePath)) {
        $gitFile = $gitFileDict[$relativePath]
        Write-Host "[削除] " -ForegroundColor Red -NoNewline
        Write-Host $relativePath

        try {
            Remove-Item -Path $gitFile.FullName -Force
            $deletedCount++
        } catch {
            Write-Warning "ファイル削除エラー: $relativePath - $_"
        }
    }
}

Write-Host ""
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host " 同期完了" -ForegroundColor Cyan
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "更新/新規ファイル: $copiedCount" -ForegroundColor Green
Write-Host "削除ファイル: $deletedCount" -ForegroundColor Red
Write-Host "変更なし: $identicalCount" -ForegroundColor Gray
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
Write-Host " 2. 何もせず終了"
Write-Host "------------------------------------------------------------------------" -ForegroundColor Yellow
$commitChoice = Read-Host "選択 (1-2)"

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
