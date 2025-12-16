<# :
@echo off
chcp 65001 >nul
title Git ブランチ上書きツール
setlocal

pushd "%~dp0"

powershell -NoProfile -ExecutionPolicy Bypass -Command "$scriptDir=('%~dp0' -replace '\\$',''); try { iex ((gc '%~f0') -join \"`n\") } finally { Set-Location C:\ }"
set EXITCODE=%ERRORLEVEL%

popd

pause
exit /b %EXITCODE%
: #>

#==============================================================================
# Git ブランチ上書きツール
#==============================================================================
# 概要:
#   指定したブランチの内容を、別のブランチの内容で完全に上書きします。
#   git reset --hard 方式を使用し、履歴も含めて完全に置き換えます。
#
# 使用例:
#   develop ブランチを main ブランチの内容で完全に上書きする
#
# 注意:
#   - 上書き後は元に戻せません（バックアップブランチ作成を推奨）
#   - リモートへのプッシュには --force-with-lease が必要です
#==============================================================================

# タイトル表示
Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Git ブランチ上書きツール" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  指定したブランチを別のブランチの内容で完全に上書きします。" -ForegroundColor Gray
Write-Host "  (git reset --hard 方式)" -ForegroundColor Gray
Write-Host ""

#==============================================================================
# Gitリポジトリの確認
#==============================================================================
function Test-GitRepository {
    $gitDir = git rev-parse --git-dir 2>$null
    return $LASTEXITCODE -eq 0
}

#==============================================================================
# ブランチ一覧を取得
#==============================================================================
function Get-GitBranches {
    param (
        [switch]$IncludeRemote
    )

    $branches = @()

    # ローカルブランチ
    $localBranches = git branch --format="%(refname:short)" 2>$null
    if ($localBranches) {
        $branches += $localBranches | ForEach-Object { $_.Trim() }
    }

    # リモートブランチ（オプション）
    if ($IncludeRemote) {
        $remoteBranches = git branch -r --format="%(refname:short)" 2>$null
        if ($remoteBranches) {
            $branches += $remoteBranches | ForEach-Object { $_.Trim() }
        }
    }

    return $branches | Sort-Object -Unique
}

#==============================================================================
# 現在のブランチを取得
#==============================================================================
function Get-CurrentBranch {
    return (git branch --show-current 2>$null)
}

#==============================================================================
# ブランチ選択メニュー
#==============================================================================
function Select-Branch {
    param (
        [string]$Title,
        [string[]]$Branches,
        [string]$ExcludeBranch = "",
        [string[]]$ExcludeBranches = @()
    )

    # 除外ブランチを除く（単一）
    if ($ExcludeBranch) {
        $Branches = $Branches | Where-Object { $_ -ne $ExcludeBranch }
    }

    # 除外ブランチを除く（複数）
    if ($ExcludeBranches.Count -gt 0) {
        $Branches = $Branches | Where-Object { $_ -notin $ExcludeBranches }
    }

    if ($Branches.Count -eq 0) {
        Write-Host "[エラー] 選択可能なブランチがありません。" -ForegroundColor Red
        return $null
    }

    Write-Host ""
    Write-Host "----------------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host "  $Title" -ForegroundColor Yellow
    Write-Host "----------------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host ""

    for ($i = 0; $i -lt $Branches.Count; $i++) {
        $num = $i + 1
        $branch = $Branches[$i]
        Write-Host "   $num. $branch" -ForegroundColor White
    }

    Write-Host ""
    Write-Host "   0. キャンセル" -ForegroundColor DarkGray
    Write-Host ""

    while ($true) {
        $input = Read-Host "選択 (0-$($Branches.Count))"

        if ($input -eq "0") {
            return $null
        }

        $index = 0
        if ([int]::TryParse($input, [ref]$index)) {
            if ($index -ge 1 -and $index -le $Branches.Count) {
                return $Branches[$index - 1]
            }
        }

        Write-Host "[エラー] 無効な選択です。0-$($Branches.Count) の数字を入力してください。" -ForegroundColor Red
    }
}

#==============================================================================
# メイン処理
#==============================================================================
function Main {
    # Gitリポジトリの確認
    if (-not (Test-GitRepository)) {
        Write-Host "[エラー] 現在のディレクトリはGitリポジトリではありません。" -ForegroundColor Red
        Write-Host ""
        Write-Host "Gitリポジトリのディレクトリで実行してください。" -ForegroundColor Yellow
        return 1
    }

    # リポジトリ情報を表示
    $repoRoot = git rev-parse --show-toplevel 2>$null
    $currentBranch = Get-CurrentBranch

    Write-Host "リポジトリ: $repoRoot" -ForegroundColor Gray
    Write-Host "現在のブランチ: $currentBranch" -ForegroundColor Gray
    Write-Host ""

    # 作業ツリーの状態確認
    $status = git status --porcelain 2>$null
    if ($status) {
        Write-Host "[警告] コミットされていない変更があります。" -ForegroundColor Yellow
        Write-Host ""
        git status --short
        Write-Host ""
        $confirm = Read-Host "続行しますか？ (y/N)"
        if ($confirm -ne "y" -and $confirm -ne "Y") {
            Write-Host "キャンセルしました。" -ForegroundColor Yellow
            return 0
        }
    }

    # ブランチ一覧を取得
    $branches = Get-GitBranches

    if ($branches.Count -lt 2) {
        Write-Host "[エラー] ブランチが2つ以上必要です。" -ForegroundColor Red
        return 1
    }

    #--------------------------------------------------------------------------
    # ソースブランチ選択（コピー元）
    #--------------------------------------------------------------------------
    $sourceBranch = Select-Branch -Title "ソースブランチを選択（コピー元）" -Branches $branches

    if (-not $sourceBranch) {
        Write-Host "キャンセルしました。" -ForegroundColor Yellow
        return 0
    }

    Write-Host ""
    Write-Host "[選択] ソースブランチ: $sourceBranch" -ForegroundColor Green

    #--------------------------------------------------------------------------
    # ターゲットブランチ選択（上書き対象）
    # ※ main/master は保護のため選択不可
    #--------------------------------------------------------------------------
    $protectedBranches = @("main", "master")
    $targetBranch = Select-Branch -Title "ターゲットブランチを選択（上書き対象）※main/masterは保護" -Branches $branches -ExcludeBranch $sourceBranch -ExcludeBranches $protectedBranches

    if (-not $targetBranch) {
        Write-Host "キャンセルしました。" -ForegroundColor Yellow
        return 0
    }

    Write-Host ""
    Write-Host "[選択] ターゲットブランチ: $targetBranch" -ForegroundColor Green

    #--------------------------------------------------------------------------
    # バックアップブランチの作成確認
    #--------------------------------------------------------------------------
    Write-Host ""
    Write-Host "----------------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host "  バックアップブランチの作成" -ForegroundColor Yellow
    Write-Host "----------------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  上書き前に '$targetBranch' のバックアップを作成しますか？" -ForegroundColor White
    Write-Host ""
    Write-Host "   1. バックアップを作成する（推奨）" -ForegroundColor White
    Write-Host "   2. バックアップを作成しない" -ForegroundColor White
    Write-Host ""
    Write-Host "   0. キャンセル" -ForegroundColor DarkGray
    Write-Host ""

    $backupChoice = Read-Host "選択 (0-2)"

    $createBackup = $false
    switch ($backupChoice) {
        "0" {
            Write-Host "キャンセルしました。" -ForegroundColor Yellow
            return 0
        }
        "1" { $createBackup = $true }
        "2" { $createBackup = $false }
        default {
            Write-Host "[エラー] 無効な選択です。" -ForegroundColor Red
            return 1
        }
    }

    #--------------------------------------------------------------------------
    # リモートプッシュの確認
    #--------------------------------------------------------------------------
    Write-Host ""
    Write-Host "----------------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host "  リモートへのプッシュ" -ForegroundColor Yellow
    Write-Host "----------------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  変更後、リモートリポジトリにプッシュしますか？" -ForegroundColor White
    Write-Host "  （--force-with-lease オプションが使用されます）" -ForegroundColor Gray
    Write-Host ""
    Write-Host "   1. リモートにプッシュする" -ForegroundColor White
    Write-Host "   2. ローカルのみ変更（プッシュしない）" -ForegroundColor White
    Write-Host ""
    Write-Host "   0. キャンセル" -ForegroundColor DarkGray
    Write-Host ""

    $pushChoice = Read-Host "選択 (0-2)"

    $pushToRemote = $false
    switch ($pushChoice) {
        "0" {
            Write-Host "キャンセルしました。" -ForegroundColor Yellow
            return 0
        }
        "1" { $pushToRemote = $true }
        "2" { $pushToRemote = $false }
        default {
            Write-Host "[エラー] 無効な選択です。" -ForegroundColor Red
            return 1
        }
    }

    #--------------------------------------------------------------------------
    # 最終確認
    #--------------------------------------------------------------------------
    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host "  実行内容の確認" -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  ソースブランチ（コピー元）: $sourceBranch" -ForegroundColor White
    Write-Host "  ターゲットブランチ（上書き対象）: $targetBranch" -ForegroundColor White
    Write-Host ""
    Write-Host "  実行される操作:" -ForegroundColor Yellow

    if ($createBackup) {
        $backupBranchName = "backup/${targetBranch}_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        Write-Host "    1. バックアップ作成: $backupBranchName" -ForegroundColor Gray
    }

    Write-Host "    $(if($createBackup){'2'}else{'1'}). git checkout $targetBranch" -ForegroundColor Gray
    Write-Host "    $(if($createBackup){'3'}else{'2'}). git reset --hard $sourceBranch" -ForegroundColor Gray

    if ($pushToRemote) {
        Write-Host "    $(if($createBackup){'4'}else{'3'}). git push --force-with-lease origin $targetBranch" -ForegroundColor Gray
    }

    Write-Host ""
    Write-Host "  [注意] この操作は取り消せません！" -ForegroundColor Red
    Write-Host ""

    $finalConfirm = Read-Host "実行しますか？ (yes/No)"

    if ($finalConfirm -ne "yes") {
        Write-Host ""
        Write-Host "キャンセルしました。" -ForegroundColor Yellow
        return 0
    }

    #--------------------------------------------------------------------------
    # 実行
    #--------------------------------------------------------------------------
    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Green
    Write-Host "  実行中..." -ForegroundColor Green
    Write-Host "================================================================" -ForegroundColor Green
    Write-Host ""

    # バックアップブランチの作成
    if ($createBackup) {
        Write-Host "[1/$(if($pushToRemote){4}else{3})] バックアップブランチを作成中..." -ForegroundColor Cyan

        git branch $backupBranchName $targetBranch 2>&1

        if ($LASTEXITCODE -ne 0) {
            Write-Host "[エラー] バックアップブランチの作成に失敗しました。" -ForegroundColor Red
            return 1
        }

        Write-Host "  作成完了: $backupBranchName" -ForegroundColor Green
        Write-Host ""
    }

    # ターゲットブランチにチェックアウト
    $step = if ($createBackup) { 2 } else { 1 }
    $totalSteps = if ($pushToRemote) { if ($createBackup) { 4 } else { 3 } } else { if ($createBackup) { 3 } else { 2 } }

    Write-Host "[$step/$totalSteps] ターゲットブランチにチェックアウト中..." -ForegroundColor Cyan

    git checkout $targetBranch 2>&1

    if ($LASTEXITCODE -ne 0) {
        Write-Host "[エラー] チェックアウトに失敗しました。" -ForegroundColor Red
        return 1
    }

    Write-Host "  完了: $targetBranch にチェックアウトしました" -ForegroundColor Green
    Write-Host ""

    # git reset --hard
    $step++
    Write-Host "[$step/$totalSteps] ブランチを上書き中 (git reset --hard)..." -ForegroundColor Cyan

    git reset --hard $sourceBranch 2>&1

    if ($LASTEXITCODE -ne 0) {
        Write-Host "[エラー] reset に失敗しました。" -ForegroundColor Red
        return 1
    }

    Write-Host "  完了: $targetBranch を $sourceBranch の内容で上書きしました" -ForegroundColor Green
    Write-Host ""

    # リモートにプッシュ
    if ($pushToRemote) {
        $step++
        Write-Host "[$step/$totalSteps] リモートにプッシュ中 (--force-with-lease)..." -ForegroundColor Cyan

        git push --force-with-lease origin $targetBranch 2>&1

        if ($LASTEXITCODE -ne 0) {
            Write-Host "[エラー] プッシュに失敗しました。" -ForegroundColor Red
            Write-Host "  ローカルの変更は完了しています。" -ForegroundColor Yellow
            Write-Host "  手動でプッシュしてください: git push --force-with-lease origin $targetBranch" -ForegroundColor Yellow
            return 1
        }

        Write-Host "  完了: リモートにプッシュしました" -ForegroundColor Green
        Write-Host ""
    }

    #--------------------------------------------------------------------------
    # 完了
    #--------------------------------------------------------------------------
    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Green
    Write-Host "  処理完了" -ForegroundColor Green
    Write-Host "================================================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "  $targetBranch ブランチを $sourceBranch の内容で上書きしました。" -ForegroundColor White
    Write-Host ""

    if ($createBackup) {
        Write-Host "  バックアップブランチ: $backupBranchName" -ForegroundColor Gray
        Write-Host "  元に戻す場合: git checkout $targetBranch && git reset --hard $backupBranchName" -ForegroundColor Gray
        Write-Host ""
    }

    # 最終状態を表示
    Write-Host "現在の状態:" -ForegroundColor Cyan
    git log --oneline -3
    Write-Host ""

    return 0
}

# 実行
$exitCode = Main
exit $exitCode
