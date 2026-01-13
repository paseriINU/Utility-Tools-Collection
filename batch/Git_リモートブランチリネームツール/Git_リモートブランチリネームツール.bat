<# :
@echo off
chcp 65001 >nul
title Git リモートブランチリネームツール
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
# Git Remote Branch Rename Tool (PowerShell)
# リモートブランチをリネーム（新規作成＋旧ブランチ削除）
# =============================================================================

# タイトル表示
Write-Host ""
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host "  Git リモートブランチリネームツール" -ForegroundColor Cyan
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

#region 設定セクション
# Gitプロジェクトのパス
$GIT_PROJECT_PATH = "C:\Users\$env:USERNAME\source\Git\project"

# 保護ブランチリスト（リネーム不可）
$ProtectedBranches = @("main", "master", "develop", "release")
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

# リモート情報を最新化
Write-Host "リモート情報を取得中..." -ForegroundColor Yellow
git fetch --prune 2>&1 | Out-Null
Write-Host ""
#endregion

#region 関数: 影響警告表示
function Show-ImpactWarning {
    param([string]$OldBranchName, [string]$NewBranchName)

    Write-Host ""
    Write-Host "========================================================================" -ForegroundColor Yellow
    Write-Host "  [警告] 他の開発者への影響" -ForegroundColor Yellow
    Write-Host "========================================================================" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "このブランチをリネームすると、以下の影響があります:" -ForegroundColor White
    Write-Host ""
    Write-Host "  1. このブランチを追跡している他の開発者は、" -ForegroundColor White
    Write-Host "     追跡ブランチの再設定が必要になります。" -ForegroundColor White
    Write-Host ""
    Write-Host "  2. 他の開発者は以下のコマンドを実行する必要があります:" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "     git fetch --prune" -ForegroundColor Green
    Write-Host "     git branch -m $OldBranchName $NewBranchName" -ForegroundColor Green
    Write-Host "     git branch -u origin/$NewBranchName $NewBranchName" -ForegroundColor Green
    Write-Host ""
    Write-Host "  3. プルリクエストやCI/CDパイプラインでこのブランチを" -ForegroundColor White
    Write-Host "     参照している場合は、設定の更新が必要です。" -ForegroundColor White
    Write-Host ""
}
#endregion

#region 関数: ブランチ名バリデーション
function Test-BranchName {
    param([string]$BranchName)

    # 空チェック
    if ([string]::IsNullOrWhiteSpace($BranchName)) {
        return @{ Valid = $false; Message = "ブランチ名が空です" }
    }

    # 禁止文字チェック
    if ($BranchName -match '[\s~^:?*\[\]\\]') {
        return @{ Valid = $false; Message = "ブランチ名に使用できない文字が含まれています（スペース、~、^、:、?、*、[、]、\）" }
    }

    # 先頭・末尾のドット/スラッシュチェック
    if ($BranchName -match '^[./]|[./]$') {
        return @{ Valid = $false; Message = "ブランチ名の先頭・末尾に . や / は使用できません" }
    }

    # 連続ドットチェック
    if ($BranchName -match '\.\.') {
        return @{ Valid = $false; Message = "ブランチ名に連続したドット(..)は使用できません" }
    }

    # @{ チェック
    if ($BranchName -match '@\{') {
        return @{ Valid = $false; Message = "ブランチ名に @{ は使用できません" }
    }

    return @{ Valid = $true; Message = "" }
}
#endregion

#region 関数: リモートブランチリネーム
function Rename-RemoteBranch {
    Write-Host ""
    Write-Host "========================================================================" -ForegroundColor Cyan
    Write-Host "  リモートブランチリネーム" -ForegroundColor Cyan
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
    Write-Host "リモートブランチ一覧を取得中..." -ForegroundColor Yellow
    Write-Host ""

    # リモートブランチ一覧取得（HEAD除外、保護ブランチ除外）
    $remoteBranches = @(git branch -r | Where-Object { $_ -notlike "*HEAD*" } | ForEach-Object {
        $branchFullName = $_.Trim()
        $branchShortName = $branchFullName -replace "^[^/]+/", ""
        # 保護ブランチは除外
        if ($ProtectedBranches -notcontains $branchShortName) {
            [PSCustomObject]@{
                FullName = $branchFullName
                ShortName = $branchShortName
            }
        }
    } | Where-Object { $_ })

    if ($remoteBranches.Count -eq 0) {
        Write-Host "リネーム可能なリモートブランチが見つかりません" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "保護ブランチ ($($ProtectedBranches -join ', ')) はリネームできません" -ForegroundColor Gray
        Read-Host "Enterキーで戻る"
        return
    }

    # ブランチを番号付きで表示
    Write-Host "リネーム可能なブランチ:" -ForegroundColor Cyan
    Write-Host ""
    for ($i = 0; $i -lt $remoteBranches.Count; $i++) {
        $displayNum = $i + 1
        $branchName = $remoteBranches[$i].ShortName
        if ($branchName -eq $currentBranch) {
            Write-Host " $displayNum. $branchName " -NoNewline
            Write-Host "[現在のブランチ]" -ForegroundColor Yellow
        } else {
            Write-Host " $displayNum. $branchName"
        }
    }
    Write-Host ""
    Write-Host " 0. キャンセル"
    Write-Host ""

    $maxNum = $remoteBranches.Count
    $selection = Read-Host "リネームするブランチ番号を入力 (0-$maxNum)"

    if ($selection -eq "0") { return }

    if ($selection -match '^\d+$' -and [int]$selection -ge 1 -and [int]$selection -le $maxNum) {
        $selectedBranch = $remoteBranches[[int]$selection - 1]
        $oldBranchName = $selectedBranch.ShortName

        Write-Host ""
        Write-Host "========================================================================" -ForegroundColor Yellow
        Write-Host " 選択されたブランチ: $oldBranchName" -ForegroundColor Yellow
        Write-Host "========================================================================" -ForegroundColor Yellow
        Write-Host ""

        # 保護ブランチチェック（念のため）
        if ($ProtectedBranches -contains $oldBranchName) {
            Write-Host "[エラー] $oldBranchName は保護されています。リネームできません。" -ForegroundColor Red
            Read-Host "Enterキーで戻る"
            return
        }

        # 新しいブランチ名を入力
        Write-Host "新しいブランチ名を入力してください" -ForegroundColor Cyan
        Write-Host "(空白でキャンセル)" -ForegroundColor Gray
        Write-Host ""
        $newBranchName = Read-Host "新しいブランチ名"

        if ([string]::IsNullOrWhiteSpace($newBranchName)) {
            Write-Host "キャンセルしました" -ForegroundColor Yellow
            Read-Host "Enterキーで戻る"
            return
        }

        # ブランチ名バリデーション
        $validation = Test-BranchName -BranchName $newBranchName
        if (-not $validation.Valid) {
            Write-Host ""
            Write-Host "[エラー] $($validation.Message)" -ForegroundColor Red
            Read-Host "Enterキーで戻る"
            return
        }

        # 同じ名前チェック
        if ($oldBranchName -eq $newBranchName) {
            Write-Host ""
            Write-Host "[エラー] 同じブランチ名です" -ForegroundColor Red
            Read-Host "Enterキーで戻る"
            return
        }

        # 保護ブランチ名へのリネームチェック
        if ($ProtectedBranches -contains $newBranchName) {
            Write-Host ""
            Write-Host "[エラー] $newBranchName は保護ブランチ名のため使用できません" -ForegroundColor Red
            Read-Host "Enterキーで戻る"
            return
        }

        # 既存ブランチ名チェック
        $existingBranches = @(git branch -r | ForEach-Object { ($_ -replace "^[^/]+/", "").Trim() })
        if ($existingBranches -contains $newBranchName) {
            Write-Host ""
            Write-Host "[エラー] ブランチ '$newBranchName' は既に存在します" -ForegroundColor Red
            Read-Host "Enterキーで戻る"
            return
        }

        # ローカルブランチ存在チェック
        $localBranchExists = git branch --format="%(refname:short)" | Where-Object { $_ -eq $oldBranchName }
        $isCurrentBranch = ($oldBranchName -eq $currentBranch)

        # 影響警告を表示
        Show-ImpactWarning -OldBranchName $oldBranchName -NewBranchName $newBranchName

        # リネーム内容の確認
        Write-Host "========================================================================" -ForegroundColor Cyan
        Write-Host "  リネーム内容の確認" -ForegroundColor Cyan
        Write-Host "========================================================================" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "  旧ブランチ名: $oldBranchName" -ForegroundColor White
        Write-Host "  新ブランチ名: $newBranchName" -ForegroundColor Green
        Write-Host ""
        Write-Host "  実行される操作:" -ForegroundColor Cyan
        if ($localBranchExists) {
            Write-Host "    [1] ローカルブランチをリネーム: $oldBranchName -> $newBranchName" -ForegroundColor White
        }
        Write-Host "    [$(if($localBranchExists){'2'}else{'1'})] 新しいブランチをリモートにプッシュ: $newBranchName" -ForegroundColor White
        Write-Host "    [$(if($localBranchExists){'3'}else{'2'})] 旧リモートブランチを削除: $remoteName/$oldBranchName" -ForegroundColor White
        if ($localBranchExists) {
            Write-Host "    [4] 追跡ブランチを設定: $newBranchName -> $remoteName/$newBranchName" -ForegroundColor White
        }
        Write-Host ""

        $confirm = Read-Host "リネームを実行しますか? (y/n)"
        if ($confirm -ne "y") {
            Write-Host "キャンセルしました" -ForegroundColor Yellow
            Read-Host "Enterキーで戻る"
            return
        }

        Write-Host ""
        Write-Host "リネームを実行中..." -ForegroundColor Yellow
        Write-Host ""

        $success = $true

        # ローカルブランチが存在する場合はリネーム
        if ($localBranchExists) {
            Write-Host "[1/4] ローカルブランチをリネーム中..." -ForegroundColor Cyan

            if ($isCurrentBranch) {
                # 現在のブランチの場合は git branch -m を使用
                git branch -m $oldBranchName $newBranchName 2>&1
            } else {
                # 現在のブランチでない場合
                git branch -m $oldBranchName $newBranchName 2>&1
            }

            if ($LASTEXITCODE -eq 0) {
                Write-Host "      [OK] ローカルブランチをリネームしました" -ForegroundColor Green
            } else {
                Write-Host "      [NG] ローカルブランチのリネームに失敗しました" -ForegroundColor Red
                $success = $false
            }
        }

        if ($success) {
            # 新しいブランチをリモートにプッシュ
            $stepNum = if ($localBranchExists) { "2/4" } else { "1/2" }
            Write-Host "[$stepNum] 新しいブランチをリモートにプッシュ中..." -ForegroundColor Cyan

            if ($localBranchExists) {
                git push -u $remoteName $newBranchName 2>&1
            } else {
                # ローカルブランチがない場合はリモートブランチから新規作成
                git push $remoteName "$remoteName/$oldBranchName`:refs/heads/$newBranchName" 2>&1
            }

            if ($LASTEXITCODE -eq 0) {
                Write-Host "      [OK] リモートにプッシュしました" -ForegroundColor Green
            } else {
                Write-Host "      [NG] リモートへのプッシュに失敗しました" -ForegroundColor Red
                $success = $false
            }
        }

        if ($success) {
            # 旧リモートブランチを削除
            $stepNum = if ($localBranchExists) { "3/4" } else { "2/2" }
            Write-Host "[$stepNum] 旧リモートブランチを削除中..." -ForegroundColor Cyan
            git push $remoteName --delete $oldBranchName 2>&1

            if ($LASTEXITCODE -eq 0) {
                Write-Host "      [OK] 旧リモートブランチを削除しました" -ForegroundColor Green
            } else {
                Write-Host "      [NG] 旧リモートブランチの削除に失敗しました" -ForegroundColor Red
                Write-Host "      新しいブランチは作成されていますが、旧ブランチが残っています" -ForegroundColor Yellow
                $success = $false
            }
        }

        if ($success -and $localBranchExists) {
            # 追跡ブランチを設定
            Write-Host "[4/4] 追跡ブランチを設定中..." -ForegroundColor Cyan
            git branch -u "$remoteName/$newBranchName" $newBranchName 2>&1

            if ($LASTEXITCODE -eq 0) {
                Write-Host "      [OK] 追跡ブランチを設定しました" -ForegroundColor Green
            } else {
                Write-Host "      [NG] 追跡ブランチの設定に失敗しました（手動で設定してください）" -ForegroundColor Yellow
            }
        }

        Write-Host ""
        if ($success) {
            Write-Host "========================================================================" -ForegroundColor Green
            Write-Host "  リネーム完了" -ForegroundColor Green
            Write-Host "========================================================================" -ForegroundColor Green
            Write-Host ""
            Write-Host "  $oldBranchName -> $newBranchName" -ForegroundColor Green
            Write-Host ""
            Write-Host "他の開発者への通知を忘れずに行ってください。" -ForegroundColor Yellow
        } else {
            Write-Host "========================================================================" -ForegroundColor Red
            Write-Host "  リネーム処理中にエラーが発生しました" -ForegroundColor Red
            Write-Host "========================================================================" -ForegroundColor Red
        }

        Read-Host "Enterキーで戻る"
    } else {
        Write-Host "無効な番号です" -ForegroundColor Red
        Read-Host "Enterキーで戻る"
    }
}
#endregion

#region 関数: 直接入力でリネーム
function Rename-RemoteBranchDirect {
    Write-Host ""
    Write-Host "========================================================================" -ForegroundColor Cyan
    Write-Host "  ブランチ名を直接入力してリネーム" -ForegroundColor Cyan
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

    # 旧ブランチ名を入力
    Write-Host "リネームするブランチ名を入力してください" -ForegroundColor Cyan
    Write-Host "(空白でキャンセル)" -ForegroundColor Gray
    Write-Host ""
    $oldBranchName = Read-Host "旧ブランチ名"

    if ([string]::IsNullOrWhiteSpace($oldBranchName)) {
        Write-Host "キャンセルしました" -ForegroundColor Yellow
        Read-Host "Enterキーで戻る"
        return
    }

    # 保護ブランチチェック
    if ($ProtectedBranches -contains $oldBranchName) {
        Write-Host ""
        Write-Host "[エラー] $oldBranchName は保護されています。リネームできません。" -ForegroundColor Red
        Read-Host "Enterキーで戻る"
        return
    }

    # リモートブランチ存在チェック
    $remoteBranchExists = git branch -r | Where-Object { ($_ -replace "^[^/]+/", "").Trim() -eq $oldBranchName }
    if (-not $remoteBranchExists) {
        Write-Host ""
        Write-Host "[エラー] リモートブランチ '$oldBranchName' が見つかりません" -ForegroundColor Red
        Read-Host "Enterキーで戻る"
        return
    }

    Write-Host ""
    Write-Host "新しいブランチ名を入力してください" -ForegroundColor Cyan
    Write-Host "(空白でキャンセル)" -ForegroundColor Gray
    Write-Host ""
    $newBranchName = Read-Host "新しいブランチ名"

    if ([string]::IsNullOrWhiteSpace($newBranchName)) {
        Write-Host "キャンセルしました" -ForegroundColor Yellow
        Read-Host "Enterキーで戻る"
        return
    }

    # ブランチ名バリデーション
    $validation = Test-BranchName -BranchName $newBranchName
    if (-not $validation.Valid) {
        Write-Host ""
        Write-Host "[エラー] $($validation.Message)" -ForegroundColor Red
        Read-Host "Enterキーで戻る"
        return
    }

    # 同じ名前チェック
    if ($oldBranchName -eq $newBranchName) {
        Write-Host ""
        Write-Host "[エラー] 同じブランチ名です" -ForegroundColor Red
        Read-Host "Enterキーで戻る"
        return
    }

    # 保護ブランチ名へのリネームチェック
    if ($ProtectedBranches -contains $newBranchName) {
        Write-Host ""
        Write-Host "[エラー] $newBranchName は保護ブランチ名のため使用できません" -ForegroundColor Red
        Read-Host "Enterキーで戻る"
        return
    }

    # 既存ブランチ名チェック
    $existingBranches = @(git branch -r | ForEach-Object { ($_ -replace "^[^/]+/", "").Trim() })
    if ($existingBranches -contains $newBranchName) {
        Write-Host ""
        Write-Host "[エラー] ブランチ '$newBranchName' は既に存在します" -ForegroundColor Red
        Read-Host "Enterキーで戻る"
        return
    }

    # ローカルブランチ存在チェック
    $localBranchExists = git branch --format="%(refname:short)" | Where-Object { $_ -eq $oldBranchName }
    $isCurrentBranch = ($oldBranchName -eq $currentBranch)

    # 影響警告を表示
    Show-ImpactWarning -OldBranchName $oldBranchName -NewBranchName $newBranchName

    # リネーム内容の確認
    Write-Host "========================================================================" -ForegroundColor Cyan
    Write-Host "  リネーム内容の確認" -ForegroundColor Cyan
    Write-Host "========================================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  旧ブランチ名: $oldBranchName" -ForegroundColor White
    Write-Host "  新ブランチ名: $newBranchName" -ForegroundColor Green
    Write-Host ""
    Write-Host "  実行される操作:" -ForegroundColor Cyan
    if ($localBranchExists) {
        Write-Host "    [1] ローカルブランチをリネーム: $oldBranchName -> $newBranchName" -ForegroundColor White
    }
    Write-Host "    [$(if($localBranchExists){'2'}else{'1'})] 新しいブランチをリモートにプッシュ: $newBranchName" -ForegroundColor White
    Write-Host "    [$(if($localBranchExists){'3'}else{'2'})] 旧リモートブランチを削除: $remoteName/$oldBranchName" -ForegroundColor White
    if ($localBranchExists) {
        Write-Host "    [4] 追跡ブランチを設定: $newBranchName -> $remoteName/$newBranchName" -ForegroundColor White
    }
    Write-Host ""

    $confirm = Read-Host "リネームを実行しますか? (y/n)"
    if ($confirm -ne "y") {
        Write-Host "キャンセルしました" -ForegroundColor Yellow
        Read-Host "Enterキーで戻る"
        return
    }

    Write-Host ""
    Write-Host "リネームを実行中..." -ForegroundColor Yellow
    Write-Host ""

    $success = $true

    # ローカルブランチが存在する場合はリネーム
    if ($localBranchExists) {
        Write-Host "[1/4] ローカルブランチをリネーム中..." -ForegroundColor Cyan

        if ($isCurrentBranch) {
            git branch -m $oldBranchName $newBranchName 2>&1
        } else {
            git branch -m $oldBranchName $newBranchName 2>&1
        }

        if ($LASTEXITCODE -eq 0) {
            Write-Host "      [OK] ローカルブランチをリネームしました" -ForegroundColor Green
        } else {
            Write-Host "      [NG] ローカルブランチのリネームに失敗しました" -ForegroundColor Red
            $success = $false
        }
    }

    if ($success) {
        $stepNum = if ($localBranchExists) { "2/4" } else { "1/2" }
        Write-Host "[$stepNum] 新しいブランチをリモートにプッシュ中..." -ForegroundColor Cyan

        if ($localBranchExists) {
            git push -u $remoteName $newBranchName 2>&1
        } else {
            git push $remoteName "$remoteName/$oldBranchName`:refs/heads/$newBranchName" 2>&1
        }

        if ($LASTEXITCODE -eq 0) {
            Write-Host "      [OK] リモートにプッシュしました" -ForegroundColor Green
        } else {
            Write-Host "      [NG] リモートへのプッシュに失敗しました" -ForegroundColor Red
            $success = $false
        }
    }

    if ($success) {
        $stepNum = if ($localBranchExists) { "3/4" } else { "2/2" }
        Write-Host "[$stepNum] 旧リモートブランチを削除中..." -ForegroundColor Cyan
        git push $remoteName --delete $oldBranchName 2>&1

        if ($LASTEXITCODE -eq 0) {
            Write-Host "      [OK] 旧リモートブランチを削除しました" -ForegroundColor Green
        } else {
            Write-Host "      [NG] 旧リモートブランチの削除に失敗しました" -ForegroundColor Red
            Write-Host "      新しいブランチは作成されていますが、旧ブランチが残っています" -ForegroundColor Yellow
            $success = $false
        }
    }

    if ($success -and $localBranchExists) {
        Write-Host "[4/4] 追跡ブランチを設定中..." -ForegroundColor Cyan
        git branch -u "$remoteName/$newBranchName" $newBranchName 2>&1

        if ($LASTEXITCODE -eq 0) {
            Write-Host "      [OK] 追跡ブランチを設定しました" -ForegroundColor Green
        } else {
            Write-Host "      [NG] 追跡ブランチの設定に失敗しました（手動で設定してください）" -ForegroundColor Yellow
        }
    }

    Write-Host ""
    if ($success) {
        Write-Host "========================================================================" -ForegroundColor Green
        Write-Host "  リネーム完了" -ForegroundColor Green
        Write-Host "========================================================================" -ForegroundColor Green
        Write-Host ""
        Write-Host "  $oldBranchName -> $newBranchName" -ForegroundColor Green
        Write-Host ""
        Write-Host "他の開発者への通知を忘れずに行ってください。" -ForegroundColor Yellow
    } else {
        Write-Host "========================================================================" -ForegroundColor Red
        Write-Host "  リネーム処理中にエラーが発生しました" -ForegroundColor Red
        Write-Host "========================================================================" -ForegroundColor Red
    }

    Read-Host "Enterキーで戻る"
}
#endregion

#region 関数: 保護ブランチ一覧表示
function Show-ProtectedBranches {
    Write-Host ""
    Write-Host "========================================================================" -ForegroundColor Cyan
    Write-Host "  保護ブランチ一覧" -ForegroundColor Cyan
    Write-Host "========================================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "以下のブランチはリネームできません:" -ForegroundColor Yellow
    Write-Host ""
    foreach ($branch in $ProtectedBranches) {
        Write-Host "  - $branch" -ForegroundColor White
    }
    Write-Host ""
    Write-Host "保護ブランチを変更する場合は、スクリプト内の" -ForegroundColor Gray
    Write-Host "`$ProtectedBranches 変数を編集してください。" -ForegroundColor Gray
    Write-Host ""
    Read-Host "Enterキーで戻る"
}
#endregion

#region メインメニュー
while ($true) {
    Write-Host ""
    Write-Host "========================================================================" -ForegroundColor Cyan
    Write-Host "  Git リモートブランチリネームツール - メインメニュー" -ForegroundColor Cyan
    Write-Host "========================================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host " 1. 一覧からブランチを選択してリネーム"
    Write-Host " 2. ブランチ名を直接入力してリネーム"
    Write-Host " 3. 保護ブランチ一覧を表示"
    Write-Host ""
    Write-Host " 0. 終了"
    Write-Host ""
    $choice = Read-Host "選択 (0-3)"

    switch ($choice) {
        "0" { exit 0 }
        "1" { Rename-RemoteBranch }
        "2" { Rename-RemoteBranchDirect }
        "3" { Show-ProtectedBranches }
        default {
            Write-Host "無効な選択です" -ForegroundColor Red
            Start-Sleep -Seconds 1
        }
    }
}
#endregion

exit 0
