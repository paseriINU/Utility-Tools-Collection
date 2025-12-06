<# :
@echo off
chcp 65001 >nul
title Git ブランチ削除ツール
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
# Git Branch Delete Tool (PowerShell)
# リモートブランチとローカルブランチを数字で選択して削除
# =============================================================================

# タイトル表示
Write-Host ""
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host "  Git ブランチ削除ツール" -ForegroundColor Cyan
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host ""

# UTF-8出力設定
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# 環境変数PATHをシステム・ユーザーレベルから再読み込み（gitコマンドが見つからない問題対策）
$machinePath = [System.Environment]::GetEnvironmentVariable("Path", "Machine")
$userPath = [System.Environment]::GetEnvironmentVariable("Path", "User")
if ($machinePath) { $env:Path = $machinePath }
if ($userPath) { $env:Path += ";" + $userPath }

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

#region 設定セクション
# Gitプロジェクトのパス
$GIT_PROJECT_PATH = "C:\Users\$env:USERNAME\source\Git\project"

# 保護ブランチリスト（削除不可）
$ProtectedBranches = @("main", "master", "develop")
#endregion

#region 初期化
# パス存在確認
if (-not (Test-Path $GIT_PROJECT_PATH)) {
    Write-Host "[エラー] Gitプロジェクトフォルダが見つかりません: $GIT_PROJECT_PATH" -ForegroundColor Red
    exit 1
}

Write-Host "Gitプロジェクトパス: $GIT_PROJECT_PATH" -ForegroundColor White
Set-Location $GIT_PROJECT_PATH
Write-Host ""

# Gitリポジトリ確認
git rev-parse --git-dir 2>&1 | Out-Null
if ($LASTEXITCODE -ne 0) {
    Write-Host "[エラー] このフォルダはGit管理下にありません" -ForegroundColor Red
    exit 1
}
#endregion

#region メインメニュー
while ($true) {
    Clear-Host
    Write-Host "========================================================================" -ForegroundColor Cyan
    Write-Host "  Git ブランチ削除ツール - メインメニュー" -ForegroundColor Cyan
    Write-Host "========================================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host " 1. リモートブランチを削除"
    Write-Host " 2. ローカルブランチを削除"
    Write-Host " 3. リモート＆ローカル両方を削除"
    Write-Host " 4. 終了"
    Write-Host ""
    $choice = Read-Host "選択してください (1-4)"

    switch ($choice) {
        "1" { Delete-RemoteBranch }
        "2" { Delete-LocalBranch }
        "3" { Delete-BothBranches }
        "4" { exit 0 }
        default {
            Write-Host "無効な選択です" -ForegroundColor Red
            Start-Sleep -Seconds 1
        }
    }
}
#endregion

#region 関数: リモートブランチ削除
function Delete-RemoteBranch {
    Clear-Host
    Write-Host "========================================================================" -ForegroundColor Cyan
    Write-Host "  リモートブランチ削除" -ForegroundColor Cyan
    Write-Host "========================================================================" -ForegroundColor Cyan
    Write-Host ""

    # リモート名取得
    $remoteName = git remote | Select-Object -First 1
    if (-not $remoteName) {
        Write-Host "[エラー] リモートリポジトリが設定されていません" -ForegroundColor Red
        Read-Host "Enterキーで戻る"
        return
    }

    Write-Host "リモート名: $remoteName" -ForegroundColor White
    Write-Host ""
    Write-Host "リモートブランチ一覧を取得中..." -ForegroundColor Yellow
    Write-Host ""

    # リモートブランチ一覧取得（HEAD除外）
    $remoteBranches = git branch -r | Where-Object { $_ -notlike "*HEAD*" } | ForEach-Object { $_.Trim() }

    if ($remoteBranches.Count -eq 0) {
        Write-Host "リモートブランチが見つかりません" -ForegroundColor Yellow
        Read-Host "Enterキーで戻る"
        return
    }

    # ブランチを番号付きで表示
    for ($i = 0; $i -lt $remoteBranches.Count; $i++) {
        $displayNum = $i + 1
        Write-Host " $displayNum. $($remoteBranches[$i])"
    }
    Write-Host ""
    Write-Host " 0. キャンセル"
    Write-Host ""

    $maxNum = $remoteBranches.Count
    $selection = Read-Host "削除するブランチ番号を入力 (1-$maxNum, 0=キャンセル)"

    if ($selection -eq "0") { return }

    if ($selection -match '^\d+$' -and [int]$selection -ge 1 -and [int]$selection -le $maxNum) {
        $selectedBranch = $remoteBranches[[int]$selection - 1]
        $branchName = $selectedBranch -replace "$remoteName/", ""

        Write-Host ""
        Write-Host "========================================================================" -ForegroundColor Yellow
        Write-Host " 選択されたブランチ: $selectedBranch" -ForegroundColor Yellow
        Write-Host "========================================================================" -ForegroundColor Yellow
        Write-Host ""

        # 保護ブランチチェック
        if ($ProtectedBranches -contains $branchName) {
            Write-Host "[警告] $branchName は保護されています。このツールでは削除できません。" -ForegroundColor Red
            Read-Host "Enterキーで戻る"
            return
        }

        $confirm = Read-Host "このリモートブランチを削除しますか? (y/n)"
        if ($confirm -eq "y") {
            Write-Host ""
            Write-Host "リモートブランチを削除中..." -ForegroundColor Yellow
            git push $remoteName --delete $branchName

            if ($LASTEXITCODE -eq 0) {
                Write-Host ""
                Write-Host "リモートブランチを削除しました: $selectedBranch" -ForegroundColor Green
            } else {
                Write-Host ""
                Write-Host "[エラー] リモートブランチの削除に失敗しました" -ForegroundColor Red
            }
            Read-Host "Enterキーで戻る"
        }
    } else {
        Write-Host "無効な番号です" -ForegroundColor Red
        Read-Host "Enterキーで戻る"
    }
}
#endregion

#region 関数: ローカルブランチ削除
function Delete-LocalBranch {
    Clear-Host
    Write-Host "========================================================================" -ForegroundColor Cyan
    Write-Host "  ローカルブランチ削除" -ForegroundColor Cyan
    Write-Host "========================================================================" -ForegroundColor Cyan
    Write-Host ""

    # 現在のブランチ取得
    $currentBranch = git branch --show-current

    Write-Host "現在のブランチ: $currentBranch" -ForegroundColor White
    Write-Host ""
    Write-Host "ローカルブランチ一覧:" -ForegroundColor Yellow
    Write-Host ""

    # ローカルブランチ一覧取得（現在のブランチ除外）
    $localBranches = git branch --format="%(refname:short)" | Where-Object { $_ -ne $currentBranch }

    if ($localBranches.Count -eq 0) {
        Write-Host "削除可能なローカルブランチがありません" -ForegroundColor Yellow
        Write-Host "（現在のブランチ以外のブランチがありません）" -ForegroundColor Yellow
        Read-Host "Enterキーで戻る"
        return
    }

    # ブランチを番号付きで表示
    for ($i = 0; $i -lt $localBranches.Count; $i++) {
        $displayNum = $i + 1
        Write-Host " $displayNum. $($localBranches[$i])"
    }
    Write-Host ""
    Write-Host " 0. キャンセル"
    Write-Host ""

    $maxNum = $localBranches.Count
    $selection = Read-Host "削除するブランチ番号を入力 (1-$maxNum, 0=キャンセル)"

    if ($selection -eq "0") { return }

    if ($selection -match '^\d+$' -and [int]$selection -ge 1 -and [int]$selection -le $maxNum) {
        $selectedBranch = $localBranches[[int]$selection - 1]

        Write-Host ""
        Write-Host "========================================================================" -ForegroundColor Yellow
        Write-Host " 選択されたブランチ: $selectedBranch" -ForegroundColor Yellow
        Write-Host "========================================================================" -ForegroundColor Yellow
        Write-Host ""

        # 保護ブランチチェック
        if ($ProtectedBranches -contains $selectedBranch) {
            Write-Host "[警告] $selectedBranch は保護されています。このツールでは削除できません。" -ForegroundColor Red
            Read-Host "Enterキーで戻る"
            return
        }

        Write-Host "このブランチの削除方法を選択してください:" -ForegroundColor Cyan
        Write-Host " 1. 通常の削除 (マージ済みブランチのみ)"
        Write-Host " 2. 強制削除 (マージされていなくても削除)"
        Write-Host " 0. キャンセル"
        Write-Host ""
        $deleteMode = Read-Host "選択 (1-2, 0=キャンセル)"

        if ($deleteMode -eq "0") { return }

        if ($deleteMode -eq "1") {
            $confirm = Read-Host "ローカルブランチを削除しますか? (y/n)"
            if ($confirm -eq "y") {
                Write-Host ""
                Write-Host "ローカルブランチを削除中..." -ForegroundColor Yellow
                git branch -d $selectedBranch

                if ($LASTEXITCODE -eq 0) {
                    Write-Host ""
                    Write-Host "ローカルブランチを削除しました: $selectedBranch" -ForegroundColor Green
                } else {
                    Write-Host ""
                    Write-Host "[エラー] ローカルブランチの削除に失敗しました" -ForegroundColor Red
                    Write-Host "このブランチはマージされていない可能性があります" -ForegroundColor Yellow
                    Write-Host "強制削除する場合は、メニューから「強制削除」を選択してください" -ForegroundColor Yellow
                }
                Read-Host "Enterキーで戻る"
            }
        } elseif ($deleteMode -eq "2") {
            Write-Host ""
            Write-Host "[警告] 強制削除を選択しています" -ForegroundColor Red
            Write-Host "マージされていない変更は失われます" -ForegroundColor Red
            Write-Host ""
            $confirm = Read-Host "本当に強制削除しますか? (y/n)"
            if ($confirm -eq "y") {
                Write-Host ""
                Write-Host "ローカルブランチを強制削除中..." -ForegroundColor Yellow
                git branch -D $selectedBranch

                if ($LASTEXITCODE -eq 0) {
                    Write-Host ""
                    Write-Host "ローカルブランチを強制削除しました: $selectedBranch" -ForegroundColor Green
                } else {
                    Write-Host ""
                    Write-Host "[エラー] ローカルブランチの削除に失敗しました" -ForegroundColor Red
                }
                Read-Host "Enterキーで戻る"
            }
        } else {
            Write-Host "無効な選択です" -ForegroundColor Red
            Read-Host "Enterキーで戻る"
        }
    } else {
        Write-Host "無効な番号です" -ForegroundColor Red
        Read-Host "Enterキーで戻る"
    }
}
#endregion

#region 関数: リモート＆ローカル両方削除
function Delete-BothBranches {
    Clear-Host
    Write-Host "========================================================================" -ForegroundColor Cyan
    Write-Host "  リモート＆ローカルブランチ両方削除" -ForegroundColor Cyan
    Write-Host "========================================================================" -ForegroundColor Cyan
    Write-Host ""

    # リモート名取得
    $remoteName = git remote | Select-Object -First 1
    if (-not $remoteName) {
        Write-Host "[エラー] リモートリポジトリが設定されていません" -ForegroundColor Red
        Read-Host "Enterキーで戻る"
        return
    }

    # 現在のブランチ取得
    $currentBranch = git branch --show-current

    Write-Host "リモート名: $remoteName" -ForegroundColor White
    Write-Host "現在のブランチ: $currentBranch" -ForegroundColor White
    Write-Host ""
    Write-Host "共通するブランチを検索中..." -ForegroundColor Yellow
    Write-Host ""

    # ローカルブランチ一覧
    $localBranches = git branch --format="%(refname:short)" | Where-Object {
        $_ -ne $currentBranch -and $ProtectedBranches -notcontains $_
    }

    # リモートブランチ一覧
    $remoteBranches = git branch -r | Where-Object { $_ -notlike "*HEAD*" } | ForEach-Object {
        ($_ -replace "$remoteName/", "").Trim()
    }

    # 共通ブランチを検索
    $commonBranches = $localBranches | Where-Object { $remoteBranches -contains $_ }

    if ($commonBranches.Count -eq 0) {
        Write-Host "削除可能な共通ブランチがありません" -ForegroundColor Yellow
        Read-Host "Enterキーで戻る"
        return
    }

    # ブランチを番号付きで表示
    for ($i = 0; $i -lt $commonBranches.Count; $i++) {
        $displayNum = $i + 1
        Write-Host " $displayNum. $($commonBranches[$i]) (ローカル＆リモート)"
    }
    Write-Host ""
    Write-Host " 0. キャンセル"
    Write-Host ""

    $maxNum = $commonBranches.Count
    $selection = Read-Host "削除するブランチ番号を入力 (1-$maxNum, 0=キャンセル)"

    if ($selection -eq "0") { return }

    if ($selection -match '^\d+$' -and [int]$selection -ge 1 -and [int]$selection -le $maxNum) {
        $selectedBranch = $commonBranches[[int]$selection - 1]

        Write-Host ""
        Write-Host "========================================================================" -ForegroundColor Yellow
        Write-Host " 選択されたブランチ: $selectedBranch" -ForegroundColor Yellow
        Write-Host "========================================================================" -ForegroundColor Yellow
        Write-Host " リモート: $remoteName/$selectedBranch"
        Write-Host " ローカル: $selectedBranch"
        Write-Host ""

        Write-Host "このブランチの削除方法を選択してください:" -ForegroundColor Cyan
        Write-Host " 1. 通常の削除 (ローカルはマージ済みのみ)"
        Write-Host " 2. 強制削除 (ローカルはマージされていなくても削除)"
        Write-Host " 0. キャンセル"
        Write-Host ""
        $deleteMode = Read-Host "選択 (1-2, 0=キャンセル)"

        if ($deleteMode -eq "0") { return }

        $forceLocal = ($deleteMode -eq "2")

        $confirm = Read-Host "リモート＆ローカルブランチを削除しますか? (y/n)"
        if ($confirm -eq "y") {
            Write-Host ""
            Write-Host "リモートブランチを削除中..." -ForegroundColor Yellow
            git push $remoteName --delete $selectedBranch

            if ($LASTEXITCODE -eq 0) {
                Write-Host "ローカルブランチを削除中..." -ForegroundColor Yellow
                if ($forceLocal) {
                    git branch -D $selectedBranch
                } else {
                    git branch -d $selectedBranch
                }

                if ($LASTEXITCODE -eq 0) {
                    Write-Host ""
                    Write-Host "リモート＆ローカルブランチを削除しました: $selectedBranch" -ForegroundColor Green
                } else {
                    Write-Host ""
                    Write-Host "[エラー] ローカルブランチの削除に失敗しました" -ForegroundColor Red
                    Write-Host "リモートブランチは削除されましたが、ローカルブランチはマージされていない可能性があります" -ForegroundColor Yellow
                }
            } else {
                Write-Host ""
                Write-Host "[エラー] リモートブランチの削除に失敗しました" -ForegroundColor Red
            }
            Read-Host "Enterキーで戻る"
        }
    } else {
        Write-Host "無効な番号です" -ForegroundColor Red
        Read-Host "Enterキーで戻る"
    }
}
#endregion

exit 0
