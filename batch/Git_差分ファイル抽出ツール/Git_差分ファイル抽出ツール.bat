<# :
@echo off
chcp 65001 >nul
title Git 差分ファイル抽出ツール
setlocal

rem UNCパス対応（PushD/PopDで自動マッピング）
pushd "%~dp0"

powershell -NoProfile -ExecutionPolicy Bypass -Command "$scriptDir=('%~dp0' -replace '\\$',''); try { iex ((gc '%~f0' -Encoding UTF8) -join \"`n\") } finally { Set-Location C:\ }"
set EXITCODE=%ERRORLEVEL%

popd

rem WinMerge起動時（終了コード3）はpauseをスキップ
if %EXITCODE% neq 3 pause
exit /b %EXITCODE%
: #>

# =============================================================================
# Git Diff Extract Tool (PowerShell)
# Gitブランチ間/コミット間の差分ファイルを抽出してフォルダ構造を保ったままコピー
# =============================================================================

# UTF-8出力設定
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# タイトル表示
Write-Host ""
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host "  Git 差分ファイル抽出ツール" -ForegroundColor Cyan
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host ""

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
# Gitプロジェクトのパス
$GIT_PROJECT_PATH = "C:\Users\$env:USERNAME\source\Git\project"

# 削除されたファイルも含めるか（$true=含める, $false=含めない）
$INCLUDE_DELETED = $false

# WinMergeのパス（空文字列の場合は自動検出）
$WINMERGE_PATH = ""

# コミット履歴の表示件数
$COMMIT_HISTORY_COUNT = 20

# ネットワーク出力先のベースパス（空文字列の場合はデスクトップに出力）
# 例: "\\server\share\projects" または "Z:\projects"
$NETWORK_OUTPUT_BASE = ""
#endregion

#region Gitリポジトリ確認
# パス存在確認
if (-not (Test-Path $GIT_PROJECT_PATH)) {
    Write-Host "[エラー] Gitプロジェクトフォルダが見つかりません: $GIT_PROJECT_PATH" -ForegroundColor Red
    Write-Host ""
    Write-Host "フォルダが存在するか確認してください。" -ForegroundColor Yellow
    exit 1
}

Write-Host "Gitプロジェクトパス: $GIT_PROJECT_PATH" -ForegroundColor White
Set-Location $GIT_PROJECT_PATH

# Gitリポジトリ確認
git rev-parse --git-dir 2>&1 | Out-Null
if ($LASTEXITCODE -ne 0) {
    Write-Host "[エラー] このフォルダはGit管理下にありません: $GIT_PROJECT_PATH" -ForegroundColor Red
    exit 1
}

# リポジトリのルートディレクトリを取得
$REPO_ROOT = git rev-parse --show-toplevel
$REPO_ROOT = $REPO_ROOT -replace '/', '\'

Write-Host "リポジトリルート: $REPO_ROOT" -ForegroundColor White

# $GIT_PROJECT_PATH から $REPO_ROOT への相対パス（サブディレクトリパス）を計算
$subDirPath = ""
$repoRootNormalized = $REPO_ROOT.TrimEnd("\")
$projectPathNormalized = $GIT_PROJECT_PATH.TrimEnd("\")
if ($projectPathNormalized.StartsWith($repoRootNormalized + "\")) {
    $subDirPath = $projectPathNormalized.Substring($repoRootNormalized.Length + 1)
    Write-Host "対象サブディレクトリ: $subDirPath" -ForegroundColor White
}
Write-Host ""
#endregion

#region モード選択
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host "  比較モードを選択してください" -ForegroundColor Cyan
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host " 1. ブランチ間の比較（異なるブランチ同士を比較）"
Write-Host " 2. コミット間の比較（同ブランチ内の特定コミット間を比較）"
Write-Host ""
Write-Host " 0. 終了"
Write-Host ""

$modeSelection = Read-Host "選択 (0-2)"

if ($modeSelection -eq "0") {
    Write-Host "[キャンセル] 処理を中止しました" -ForegroundColor Yellow
    exit 0
}

if ($modeSelection -ne "1" -and $modeSelection -ne "2") {
    Write-Host "[エラー] 無効な選択です" -ForegroundColor Red
    exit 1
}

$compareMode = $modeSelection
#endregion

#region 共通関数: ブランチ選択/切り替え
function Select-Branch {
    param(
        [string]$CurrentBranch,
        [string]$Purpose = "操作"
    )

    $allBranches = @(git branch --format="%(refname:short)" | ForEach-Object { $_.Trim() })

    if ($allBranches.Count -eq 0) {
        Write-Host "[エラー] ブランチが見つかりません" -ForegroundColor Red
        return $null
    }

    Write-Host ""
    Write-Host "------------------------------------------------------------------------" -ForegroundColor Yellow
    Write-Host " 現在のブランチ: $CurrentBranch" -ForegroundColor Yellow
    Write-Host "------------------------------------------------------------------------" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "ブランチ操作を選択してください:" -ForegroundColor Cyan
    Write-Host " 1. このまま続行（$CurrentBranch）"
    Write-Host " 2. ブランチを切り替える"
    Write-Host ""
    Write-Host " 0. キャンセル"
    Write-Host ""

    $branchChoice = Read-Host "選択 (0-2)"

    switch ($branchChoice) {
        "0" {
            return $null
        }
        "1" {
            return $CurrentBranch
        }
        "2" {
            Write-Host ""
            Write-Host "利用可能なブランチ:" -ForegroundColor Yellow
            Write-Host ""

            for ($i = 0; $i -lt $allBranches.Count; $i++) {
                $displayNum = $i + 1
                $branch = $allBranches[$i]
                if ($branch -eq $CurrentBranch) {
                    Write-Host " $displayNum. $branch [現在]" -ForegroundColor Gray
                } else {
                    Write-Host " $displayNum. $branch"
                }
            }
            Write-Host ""
            Write-Host " 0. キャンセル"
            Write-Host ""

            $maxNum = $allBranches.Count
            $selection = Read-Host "ブランチ番号を入力 (0-$maxNum)"

            if ($selection -eq "0") {
                return $null
            }

            if ($selection -match '^\d+$' -and [int]$selection -ge 1 -and [int]$selection -le $maxNum) {
                $selectedBranch = $allBranches[[int]$selection - 1]

                if ($selectedBranch -ne $CurrentBranch) {
                    git checkout $selectedBranch 2>&1 | Out-Null
                    if ($LASTEXITCODE -ne 0) {
                        Write-Host "[エラー] ブランチの切り替えに失敗しました" -ForegroundColor Red
                        return $null
                    }
                    Write-Host "ブランチを切り替えました: $selectedBranch" -ForegroundColor Green
                }
                return $selectedBranch
            } else {
                Write-Host "[エラー] 無効な番号です" -ForegroundColor Red
                return $null
            }
        }
        default {
            Write-Host "[エラー] 無効な選択です" -ForegroundColor Red
            return $null
        }
    }
}
#endregion

#region モード別処理
if ($compareMode -eq "1") {
    #region ブランチ間比較モード
    Write-Host ""
    Write-Host "========================================================================" -ForegroundColor Cyan
    Write-Host "  ブランチ間比較モード" -ForegroundColor Cyan
    Write-Host "========================================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "[注意] ブランチ選択について" -ForegroundColor Yellow
    Write-Host "  比較元 = 修正前のブランチ（古いバージョン）" -ForegroundColor Gray
    Write-Host "  比較先 = 修正後のブランチ（新しいバージョン）" -ForegroundColor Gray
    Write-Host ""

    # すべてのブランチ一覧を取得
    $allBranches = @(git branch --format="%(refname:short)" | ForEach-Object { $_.Trim() })

    if ($allBranches.Count -eq 0) {
        Write-Host "[エラー] ブランチが見つかりません" -ForegroundColor Red
        exit 1
    }

    # 比較元ブランチの選択
    Write-Host "========================================================================" -ForegroundColor Cyan
    Write-Host "  比較元ブランチ（修正前 / 古いバージョン）を選択してください" -ForegroundColor Cyan
    Write-Host "========================================================================" -ForegroundColor Cyan
    Write-Host ""

    for ($i = 0; $i -lt $allBranches.Count; $i++) {
        $displayNum = $i + 1
        $branchName = $allBranches[$i]
        Write-Host " $displayNum. $branchName"
    }
    Write-Host ""
    Write-Host " 0. キャンセル"
    Write-Host ""

    $maxNum = $allBranches.Count
    $baseSelection = Read-Host "番号を選択してください (0-$maxNum)"

    if ($baseSelection -eq "0") {
        Write-Host "[キャンセル] 処理を中止しました" -ForegroundColor Yellow
        exit 0
    }

    if (-not $baseSelection -or $baseSelection -notmatch '^\d+$' -or [int]$baseSelection -lt 1 -or [int]$baseSelection -gt $maxNum) {
        Write-Host "[エラー] 無効な選択です" -ForegroundColor Red
        exit 1
    }

    $BASE_REF = $allBranches[[int]$baseSelection - 1]
    Write-Host ""
    Write-Host "[選択] 比較元ブランチ: $BASE_REF" -ForegroundColor Green
    Write-Host ""

    # 比較先ブランチの選択
    Write-Host "========================================================================" -ForegroundColor Cyan
    Write-Host "  比較先ブランチ（修正後 / 新しいバージョン）を選択してください" -ForegroundColor Cyan
    Write-Host "========================================================================" -ForegroundColor Cyan
    Write-Host ""

    for ($i = 0; $i -lt $allBranches.Count; $i++) {
        $displayNum = $i + 1
        $branchName = $allBranches[$i]

        if ($branchName -eq $BASE_REF) {
            Write-Host " $displayNum. $branchName [比較元]" -ForegroundColor Gray
        } else {
            Write-Host " $displayNum. $branchName"
        }
    }
    Write-Host ""
    Write-Host " 0. キャンセル"
    Write-Host ""

    $targetSelection = Read-Host "番号を選択してください (0-$maxNum)"

    if ($targetSelection -eq "0") {
        Write-Host "[キャンセル] 処理を中止しました" -ForegroundColor Yellow
        exit 0
    }

    if (-not $targetSelection -or $targetSelection -notmatch '^\d+$' -or [int]$targetSelection -lt 1 -or [int]$targetSelection -gt $maxNum) {
        Write-Host "[エラー] 無効な選択です" -ForegroundColor Red
        exit 1
    }

    $TARGET_REF = $allBranches[[int]$targetSelection - 1]

    if ($BASE_REF -eq $TARGET_REF) {
        Write-Host "[警告] 比較元と比較先が同じブランチです" -ForegroundColor Yellow
        $continue = Read-Host "続行しますか? (y/n)"
        if ($continue -ne "y") {
            Write-Host "処理を中止しました" -ForegroundColor Yellow
            exit 0
        }
    }

    Write-Host ""
    Write-Host "[選択] 比較先ブランチ: $TARGET_REF" -ForegroundColor Green

    $BASE_LABEL = $BASE_REF
    $TARGET_LABEL = $TARGET_REF
    $targetBranchName = $TARGET_REF  # ネットワーク出力用にブランチ名を保持
    #endregion

} else {
    #region コミット間比較モード
    Write-Host ""
    Write-Host "========================================================================" -ForegroundColor Cyan
    Write-Host "  コミット間比較モード" -ForegroundColor Cyan
    Write-Host "========================================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "[注意] コミット選択について" -ForegroundColor Yellow
    Write-Host "  比較元 = 修正前のコミット（古い時点）" -ForegroundColor Gray
    Write-Host "  比較先 = 修正後のコミット（新しい時点）" -ForegroundColor Gray
    Write-Host ""

    # 現在のブランチを取得
    $currentBranch = git branch --show-current

    # ブランチ選択/切り替え
    $selectedBranch = Select-Branch -CurrentBranch $currentBranch -Purpose "コミット比較"

    if ($null -eq $selectedBranch) {
        Write-Host "[キャンセル] 処理を中止しました" -ForegroundColor Yellow
        exit 0
    }

    $workingBranch = $selectedBranch

    # コミット履歴を取得
    Write-Host ""
    Write-Host "========================================================================" -ForegroundColor Cyan
    Write-Host "  ブランチ '$workingBranch' のコミット履歴" -ForegroundColor Cyan
    Write-Host "========================================================================" -ForegroundColor Cyan
    Write-Host ""

    $commitLog = @(git log --oneline -n $COMMIT_HISTORY_COUNT --format="%h|%ai|%s" | ForEach-Object {
        $parts = $_ -split '\|', 3
        [PSCustomObject]@{
            Hash = $parts[0]
            Date = ($parts[1] -split ' ')[0]
            Message = if ($parts[2].Length -gt 50) { $parts[2].Substring(0, 47) + "..." } else { $parts[2] }
        }
    })

    if ($commitLog.Count -eq 0) {
        Write-Host "[エラー] コミット履歴が見つかりません" -ForegroundColor Red
        exit 1
    }

    # ブランチの起点（分岐元）を検出 - すべてのブランチから最も近い分岐点を探す
    $branchBase = $null
    $branchBaseLabel = ""
    $baseBranchName = ""

    $allBranchesForBase = @(git branch --format="%(refname:short)" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne $workingBranch })

    if ($allBranchesForBase.Count -gt 0) {
        $bestMergeBase = $null
        $bestDistance = [int]::MaxValue
        $bestBranchName = ""

        foreach ($candidateBranch in $allBranchesForBase) {
            $mergeBase = git merge-base $workingBranch $candidateBranch 2>$null
            if ($LASTEXITCODE -eq 0 -and $mergeBase) {
                # 現在のブランチのHEADから分岐点までのコミット数を計算
                $distance = git rev-list --count "${mergeBase}..HEAD" 2>$null
                if ($LASTEXITCODE -eq 0 -and $distance -lt $bestDistance -and $distance -gt 0) {
                    $bestDistance = $distance
                    $bestMergeBase = $mergeBase
                    $bestBranchName = $candidateBranch
                }
            }
        }

        if ($bestMergeBase) {
            $branchBase = $bestMergeBase.Substring(0, 7)
            $baseBranchName = $bestBranchName
            # 分岐点のコミット情報を取得
            $baseInfo = git log -1 --format="%ai|%s" $bestMergeBase 2>$null
            if ($baseInfo) {
                $baseParts = $baseInfo -split '\|', 2
                $baseDate = ($baseParts[0] -split ' ')[0]
                $baseMsg = if ($baseParts[1].Length -gt 40) { $baseParts[1].Substring(0, 37) + "..." } else { $baseParts[1] }
                $branchBaseLabel = "[$branchBase] $baseDate $baseMsg"
            }
        }
    }

    # 比較元コミットの選択
    Write-Host "比較元コミット（修正前 / 古い時点）を選択してください:" -ForegroundColor Yellow
    Write-Host ""

    for ($i = 0; $i -lt $commitLog.Count; $i++) {
        $displayNum = $i + 1
        $commit = $commitLog[$i]
        Write-Host " $displayNum. [$($commit.Hash)] $($commit.Date) $($commit.Message)"
    }

    # ブランチ起点の選択肢を追加
    $hasBranchBase = $false
    if ($branchBase -and $branchBaseLabel) {
        $branchBaseNum = $commitLog.Count + 1
        Write-Host ""
        Write-Host " $branchBaseNum. $branchBaseLabel [${baseBranchName}からの分岐点]" -ForegroundColor Magenta
        $hasBranchBase = $true
    }

    Write-Host ""
    Write-Host " 0. キャンセル"
    Write-Host ""

    $maxNum = if ($hasBranchBase) { $commitLog.Count + 1 } else { $commitLog.Count }
    $baseSelection = Read-Host "番号を選択してください (0-$maxNum)"

    if ($baseSelection -eq "0") {
        Write-Host "[キャンセル] 処理を中止しました" -ForegroundColor Yellow
        exit 0
    }

    if (-not $baseSelection -or $baseSelection -notmatch '^\d+$' -or [int]$baseSelection -lt 1 -or [int]$baseSelection -gt $maxNum) {
        Write-Host "[エラー] 無効な選択です" -ForegroundColor Red
        exit 1
    }

    # ブランチ起点が選択された場合
    if ($hasBranchBase -and [int]$baseSelection -eq $branchBaseNum) {
        $BASE_REF = $branchBase
        $BASE_LABEL = "${branchBaseLabel} [${baseBranchName}からの分岐点]"
        $baseCommitIndex = $commitLog.Count  # 分岐点は最も古いとみなす
        Write-Host ""
        Write-Host "[選択] 比較元: ${baseBranchName}からの分岐点 [$branchBase]" -ForegroundColor Green
    } else {
        $baseCommit = $commitLog[[int]$baseSelection - 1]
        $BASE_REF = $baseCommit.Hash
        $BASE_LABEL = "[$($baseCommit.Hash)] $($baseCommit.Message)"
        $baseCommitIndex = [int]$baseSelection - 1  # 比較元のインデックスを保存
        Write-Host ""
        Write-Host "[選択] 比較元コミット: [$($baseCommit.Hash)] $($baseCommit.Message)" -ForegroundColor Green
    }
    Write-Host ""

    # 比較先コミットの選択
    Write-Host "========================================================================" -ForegroundColor Cyan
    Write-Host "  比較先コミット（修正後 / 新しい時点）を選択してください" -ForegroundColor Cyan
    Write-Host "========================================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host " ※ HEADを選択すると、現在の最新状態と比較します" -ForegroundColor Gray
    Write-Host ""

    # 比較元より新しいコミットのみを抽出（インデックスが小さいもの）
    $newerCommits = @()
    for ($i = 0; $i -lt $baseCommitIndex; $i++) {
        $newerCommits += $commitLog[$i]
    }

    # HEADオプションを追加
    Write-Host " 1. HEAD（現在の最新状態）" -ForegroundColor Cyan

    for ($i = 0; $i -lt $newerCommits.Count; $i++) {
        $displayNum = $i + 2
        $commit = $newerCommits[$i]
        Write-Host " $displayNum. [$($commit.Hash)] $($commit.Date) $($commit.Message)"
    }
    Write-Host ""
    Write-Host " 0. キャンセル"
    Write-Host ""

    $maxNum = $newerCommits.Count + 1
    $targetSelection = Read-Host "番号を選択してください (0-$maxNum)"

    if ($targetSelection -eq "0") {
        Write-Host "[キャンセル] 処理を中止しました" -ForegroundColor Yellow
        exit 0
    }

    if (-not $targetSelection -or $targetSelection -notmatch '^\d+$' -or [int]$targetSelection -lt 1 -or [int]$targetSelection -gt $maxNum) {
        Write-Host "[エラー] 無効な選択です" -ForegroundColor Red
        exit 1
    }

    if ($targetSelection -eq "1") {
        $TARGET_REF = "HEAD"
        $TARGET_LABEL = "HEAD（最新）"
        Write-Host ""
        Write-Host "[選択] 比較先: HEAD（現在の最新状態）" -ForegroundColor Green
    } else {
        $targetCommit = $newerCommits[[int]$targetSelection - 2]
        $TARGET_REF = $targetCommit.Hash
        $TARGET_LABEL = "[$($targetCommit.Hash)] $($targetCommit.Message)"

        Write-Host ""
        Write-Host "[選択] 比較先コミット: [$($targetCommit.Hash)] $($targetCommit.Message)" -ForegroundColor Green
    }
    $targetBranchName = $workingBranch  # ネットワーク出力用にブランチ名を保持
    #endregion
}
#endregion

#region 出力先フォルダ設定
Write-Host ""

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$OUTPUT_DIR = ""
$OUTPUT_DIR_BEFORE = ""
$OUTPUT_DIR_AFTER = ""

# ネットワーク出力モードの判定
if ($NETWORK_OUTPUT_BASE -ne "" -and (Test-Path $NETWORK_OUTPUT_BASE)) {
    Write-Host "========================================================================" -ForegroundColor Cyan
    Write-Host "  出力先フォルダの選択（ネットワーク出力モード）" -ForegroundColor Cyan
    Write-Host "========================================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "対象ブランチ名: $targetBranchName" -ForegroundColor White
    Write-Host "ネットワークパス: $NETWORK_OUTPUT_BASE" -ForegroundColor White
    Write-Host ""

    # ネットワークフォルダ配下のサブフォルダを取得
    $subFolders = @(Get-ChildItem -Path $NETWORK_OUTPUT_BASE -Directory -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name)

    # ブランチ名と完全一致するフォルダを検索
    $matchedFolder = $null
    foreach ($folder in $subFolders) {
        if ($folder -eq $targetBranchName) {
            $matchedFolder = $folder
            break
        }
    }

    if ($matchedFolder) {
        # 一致するフォルダが見つかった場合
        Write-Host "[自動検出] ブランチ名と一致するフォルダを発見: $matchedFolder" -ForegroundColor Green
        $OUTPUT_DIR = Join-Path $NETWORK_OUTPUT_BASE "$matchedFolder\30_M"
    } else {
        # 一致しない場合はフォルダ選択ダイアログを表示
        Write-Host "[情報] ブランチ名と一致するフォルダが見つかりません" -ForegroundColor Yellow
        Write-Host "フォルダ選択ダイアログを開きます..." -ForegroundColor Cyan
        Write-Host ""

        Add-Type -AssemblyName System.Windows.Forms

        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.Description = "出力先フォルダを選択してください（選択したフォルダ内の30_Mに出力されます）"
        $folderBrowser.RootFolder = [System.Environment+SpecialFolder]::Desktop
        $folderBrowser.SelectedPath = $NETWORK_OUTPUT_BASE
        $folderBrowser.ShowNewFolderButton = $true

        $dialogResult = $folderBrowser.ShowDialog()

        if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
            $selectedPath = $folderBrowser.SelectedPath
            Write-Host "[選択] フォルダ: $selectedPath" -ForegroundColor Green
            $OUTPUT_DIR = Join-Path $selectedPath "30_M"
        } else {
            Write-Host "[キャンセル] フォルダ選択がキャンセルされました" -ForegroundColor Yellow
            exit 0
        }
    }

    # 30_Mフォルダの存在確認と作成
    if (-not (Test-Path $OUTPUT_DIR)) {
        Write-Host "[作成] 30_Mフォルダを作成します: $OUTPUT_DIR" -ForegroundColor Yellow
        New-Item -ItemType Directory -Path $OUTPUT_DIR -Force | Out-Null
    }

    $OUTPUT_DIR_BEFORE = "$OUTPUT_DIR\01_修正前"
    $OUTPUT_DIR_AFTER = "$OUTPUT_DIR\02_修正後"

} elseif ($NETWORK_OUTPUT_BASE -ne "" -and -not (Test-Path $NETWORK_OUTPUT_BASE)) {
    # ネットワークパスが設定されているがアクセスできない場合
    Write-Host "[警告] ネットワークパスにアクセスできません: $NETWORK_OUTPUT_BASE" -ForegroundColor Yellow
    Write-Host "デスクトップに出力します" -ForegroundColor Yellow
    Write-Host ""

    $OUTPUT_DIR = "$env:USERPROFILE\Desktop\git_diff_$timestamp"
    $OUTPUT_DIR_BEFORE = "$OUTPUT_DIR\01_修正前"
    $OUTPUT_DIR_AFTER = "$OUTPUT_DIR\02_修正後"

} else {
    # 従来通りデスクトップに出力
    $OUTPUT_DIR = "$env:USERPROFILE\Desktop\git_diff_$timestamp"
    $OUTPUT_DIR_BEFORE = "$OUTPUT_DIR\01_修正前"
    $OUTPUT_DIR_AFTER = "$OUTPUT_DIR\02_修正後"
}

Write-Host "------------------------------------------------------------------------" -ForegroundColor White
Write-Host "比較元          : $BASE_LABEL" -ForegroundColor White
Write-Host "比較先          : $TARGET_LABEL" -ForegroundColor White
Write-Host "出力先フォルダ  : $OUTPUT_DIR" -ForegroundColor White
Write-Host "  01_修正前     : 比較元のファイル" -ForegroundColor Gray
Write-Host "  02_修正後     : 比較先のファイル" -ForegroundColor Gray
Write-Host "------------------------------------------------------------------------" -ForegroundColor White
Write-Host ""
#endregion

#region 差分ファイル取得（フォルダ操作の前に実行）
Write-Host "差分ファイルを検出中..." -ForegroundColor Cyan
Write-Host ""

# 差分ファイルリストを取得
if ($INCLUDE_DELETED) {
    $diffFiles = git diff --name-only "$BASE_REF..$TARGET_REF"
} else {
    $diffFiles = git diff --name-only --diff-filter=ACMR "$BASE_REF..$TARGET_REF"
}

if (-not $diffFiles -or $diffFiles.Count -eq 0) {
    Write-Host "[情報] 差分ファイルが見つかりませんでした" -ForegroundColor Yellow
    Write-Host "比較対象は同じ内容です" -ForegroundColor Yellow
    exit 0
}

# サブディレクトリ配下のファイルのみをフィルタリング
$filteredFiles = @()
foreach ($file in $diffFiles) {
    if ($subDirPath -ne "") {
        $subDirPathLinux = $subDirPath.Replace("\", "/")
        if ($file.StartsWith($subDirPathLinux + "/")) {
            $relativePath = $file.Substring($subDirPathLinux.Length + 1)
            $filteredFiles += [PSCustomObject]@{
                OriginalPath = $file
                RelativePath = $relativePath
            }
        }
    } else {
        $filteredFiles += [PSCustomObject]@{
            OriginalPath = $file
            RelativePath = $file
        }
    }
}

if ($filteredFiles.Count -eq 0) {
    Write-Host "[情報] 対象サブディレクトリ配下に差分ファイルが見つかりませんでした" -ForegroundColor Yellow
    exit 0
}

$FILE_COUNT = $filteredFiles.Count
Write-Host "検出された差分ファイル数: $FILE_COUNT 個" -ForegroundColor Green
Write-Host ""
#endregion

#region 出力先フォルダ確認（差分があることを確認した後に実行）
# 修正前・修正後フォルダが存在するかチェック
$beforeExists = Test-Path $OUTPUT_DIR_BEFORE
$afterExists = Test-Path $OUTPUT_DIR_AFTER

if ($beforeExists -or $afterExists) {
    Write-Host "[警告] 以下のフォルダが既に存在します" -ForegroundColor Yellow
    if ($beforeExists) {
        Write-Host "  - $OUTPUT_DIR_BEFORE" -ForegroundColor Yellow
    }
    if ($afterExists) {
        Write-Host "  - $OUTPUT_DIR_AFTER" -ForegroundColor Yellow
    }
    $overwrite = Read-Host "クリアして書き込みますか? (y/n)"

    if ($overwrite -ne "y") {
        Write-Host "処理を中止しました" -ForegroundColor Yellow
        exit 0
    }

    Write-Host "既存のフォルダをクリア中..." -ForegroundColor Yellow
    # 修正前・修正後フォルダのみを削除（30_M配下の他のフォルダは保持）
    if ($beforeExists) {
        Remove-Item -Path $OUTPUT_DIR_BEFORE -Recurse -Force
    }
    if ($afterExists) {
        Remove-Item -Path $OUTPUT_DIR_AFTER -Recurse -Force
    }
}

# 出力先フォルダを作成
if (-not (Test-Path $OUTPUT_DIR)) {
    New-Item -ItemType Directory -Path $OUTPUT_DIR -Force | Out-Null
}
New-Item -ItemType Directory -Path $OUTPUT_DIR_BEFORE -Force | Out-Null
New-Item -ItemType Directory -Path $OUTPUT_DIR_AFTER -Force | Out-Null
#endregion

#region ファイルコピー（高速一括抽出方式）
$COPY_COUNT_BEFORE = 0
$COPY_COUNT_AFTER = 0
$ERROR_COUNT = 0
$NEW_FILES = @()
$DELETED_FILES = @()

# 01_修正前（比較元）のファイルを抽出
Write-Host "[01_修正前] 比較元からファイルを抽出中..." -ForegroundColor Yellow

foreach ($fileObj in $filteredFiles) {
    $originalPath = $fileObj.OriginalPath
    $relativePath = $fileObj.RelativePath
    $relativePathWin = $relativePath -replace '/', '\'

    # 出力先パス
    $destFileBefore = Join-Path $OUTPUT_DIR_BEFORE $relativePathWin
    $destDirBefore = Split-Path -Path $destFileBefore -Parent

    # git show でファイル内容を取得
    $gitPath = $originalPath -replace '\\', '/'
    $contentBefore = git show "${BASE_REF}:${gitPath}" 2>&1

    if ($LASTEXITCODE -eq 0) {
        # ディレクトリ作成
        if (-not (Test-Path $destDirBefore)) {
            New-Item -ItemType Directory -Path $destDirBefore -Force | Out-Null
        }
        # ファイル書き込み
        [System.IO.File]::WriteAllText($destFileBefore, ($contentBefore -join "`n"), [System.Text.Encoding]::UTF8)
        $COPY_COUNT_BEFORE++
        Write-Host "  [OK] $relativePath" -ForegroundColor Gray
    } else {
        # 新規ファイル（比較元には存在しない）
        $NEW_FILES += $relativePath
        Write-Host "  [新規] $relativePath" -ForegroundColor DarkYellow
    }
}

Write-Host ""
Write-Host "  抽出完了: $COPY_COUNT_BEFORE 個" -ForegroundColor Green
if ($NEW_FILES.Count -gt 0) {
    Write-Host "  （新規ファイル: $($NEW_FILES.Count) 個）" -ForegroundColor DarkYellow
}
Write-Host ""

# 02_修正後（比較先）のファイルを抽出
Write-Host "[02_修正後] 比較先からファイルを抽出中..." -ForegroundColor Yellow

foreach ($fileObj in $filteredFiles) {
    $originalPath = $fileObj.OriginalPath
    $relativePath = $fileObj.RelativePath
    $relativePathWin = $relativePath -replace '/', '\'

    # 出力先パス
    $destFileAfter = Join-Path $OUTPUT_DIR_AFTER $relativePathWin
    $destDirAfter = Split-Path -Path $destFileAfter -Parent

    # git show でファイル内容を取得
    $gitPath = $originalPath -replace '\\', '/'
    $contentAfter = git show "${TARGET_REF}:${gitPath}" 2>&1

    if ($LASTEXITCODE -eq 0) {
        # ディレクトリ作成
        if (-not (Test-Path $destDirAfter)) {
            New-Item -ItemType Directory -Path $destDirAfter -Force | Out-Null
        }
        # ファイル書き込み
        [System.IO.File]::WriteAllText($destFileAfter, ($contentAfter -join "`n"), [System.Text.Encoding]::UTF8)
        $COPY_COUNT_AFTER++
        Write-Host "  [OK] $relativePath" -ForegroundColor Gray
    } else {
        # 削除ファイル（比較先には存在しない）
        $DELETED_FILES += $relativePath
        Write-Host "  [削除] $relativePath" -ForegroundColor DarkRed
    }
}

Write-Host ""
Write-Host "  抽出完了: $COPY_COUNT_AFTER 個" -ForegroundColor Green
if ($DELETED_FILES.Count -gt 0) {
    Write-Host "  （削除ファイル: $($DELETED_FILES.Count) 個）" -ForegroundColor DarkRed
}
Write-Host ""
#endregion

#region 結果表示
Write-Host ""
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host " 処理完了" -ForegroundColor Cyan
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "01_修正前 コピー数: $COPY_COUNT_BEFORE 個" -ForegroundColor Green
Write-Host "02_修正後 コピー数: $COPY_COUNT_AFTER 個" -ForegroundColor Green

if ($ERROR_COUNT -gt 0) {
    Write-Host "エラー            : $ERROR_COUNT 個" -ForegroundColor Red
}

Write-Host ""
Write-Host "出力先: $OUTPUT_DIR" -ForegroundColor White
Write-Host "  01_修正前: 比較元のファイル" -ForegroundColor Gray
Write-Host "  02_修正後: 比較先のファイル" -ForegroundColor Gray
Write-Host ""
#endregion

#region WinMerge比較
Write-Host ""

# WinMergeのパスを検出
$winmergePath = $WINMERGE_PATH
if ($winmergePath -eq "") {
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
}

if ($winmergePath -ne "" -and (Test-Path $winmergePath)) {
    $openWinMerge = Read-Host "WinMergeで比較しますか? (y/n)"
    if ($openWinMerge -eq "y") {
        Write-Host ""
        Write-Host "WinMergeを起動中..." -ForegroundColor Cyan
        Write-Host "（このウィンドウは自動的に閉じます）" -ForegroundColor Gray
        Start-Process -FilePath $winmergePath -ArgumentList "/r", "/e", "-cfg", "Settings/DirViewExpandSubdirs=1", $OUTPUT_DIR_BEFORE, $OUTPUT_DIR_AFTER
        exit 3  # WinMerge起動時は特別な終了コードを返す（pauseスキップ用）
    }
} else {
    Write-Host "[情報] WinMergeが見つかりません。手動で比較してください。" -ForegroundColor Yellow
    Write-Host "  WinMergeをインストールするか、設定セクションの WINMERGE_PATH を設定してください。" -ForegroundColor Gray
}
#endregion

exit 0
