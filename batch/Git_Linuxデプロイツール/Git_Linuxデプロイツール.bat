<# :
@echo off
chcp 65001 >nul
title Git Linuxデプロイツール
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
# Git Linuxデプロイツール
#==============================================================================
#
# 機能:
#   1. 複数環境から転送先を選択
#   2. 変更ファイルのみ OR すべてのファイルを選択
#   3. 拡張子フィルタ (.c .pc .h など)
#   4. 削除されたファイルは自動除外
#   5. 全部転送 or 個別選択
#   6. Linux側でディレクトリ自動作成・パーミッション設定
#   7. SCP対応
#
# 必要な環境:
#   - Git がインストールされていること
#   - SCP (Windows OpenSSH Client) が利用可能であること
#   - SSH公開鍵認証が設定されていること（推奨）
#
#==============================================================================

#region 設定 - ここを編集してください
#==============================================================================

# SSH接続情報
$SSH_USER = "youruser"
$SSH_HOST = "linux-server"
$SSH_PORT = 22

# SSH秘密鍵ファイル (公開鍵認証を使用する場合)
# パスワード認証の場合は空文字列 ""
$SSH_KEY = "$env:USERPROFILE\.ssh\id_rsa"

# Git リポジトリのルートディレクトリ (空文字列の場合は現在のディレクトリ)
$GIT_ROOT = ""

# 共通グループ（Linux側の所有者設定用）
$COMMON_GROUP = "common_group"

# 転送対象の拡張子（空文字列の場合は全ファイル、複数指定可能）
# 例: @(".c", ".pc", ".h")
# 全ファイル: @()
$TARGET_EXTENSIONS = @(".c", ".pc", ".h")

# 環境設定（複数環境対応）
# 環境名, 転送先パス, オーナー をハッシュテーブルで定義
$ENVIRONMENTS = @(
    @{
        Name = "tst1t"
        Path = "/path/to/tst1t/"
        Owner = "tzy_tst13"
    },
    @{
        Name = "tst2t"
        Path = "/path/to/tst2t/"
        Owner = "tzy_tst23"
    },
    @{
        Name = "tst3t"
        Path = "/path/to/tst3t/"
        Owner = "tzy_tst33"
    }
)

# Linux側のパーミッション設定
$LINUX_CHMOD_DIR = "777"   # ディレクトリのパーミッション
$LINUX_CHMOD_FILE = "777"  # ファイルのパーミッション

#==============================================================================
#endregion

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

# 色付き出力
function Write-Color {
    param(
        [string]$Text,
        [string]$Color = "White"
    )
    Write-Host $Text -ForegroundColor $Color
}

function Write-Header {
    param([string]$Text)
    Write-Host ""
    Write-Color "================================================================" "Cyan"
    Write-Color "  $Text" "Cyan"
    Write-Color "================================================================" "Cyan"
    Write-Host ""
}

# Git リポジトリチェック
if ($GIT_ROOT -eq "") {
    $GIT_ROOT = Get-Location
}

Set-Location $GIT_ROOT

# 最初にタイトルを表示
Write-Host ""
Write-Color "================================================================" "Cyan"
Write-Color "  Git Linuxデプロイツール" "Cyan"
Write-Color "================================================================" "Cyan"
Write-Host ""

# Gitリポジトリかどうか確認（.gitフォルダを親ディレクトリから探索）
$gitDir = git rev-parse --git-dir 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Color "[エラー] Gitリポジトリではありません: $GIT_ROOT" "Red"
    Write-Color "       git rev-parse --git-dir の実行に失敗しました" "Red"
    exit 1
}

# Gitリポジトリのルートディレクトリを取得
$gitRootDir = git rev-parse --show-toplevel 2>&1
if ($LASTEXITCODE -eq 0) {
    # Git形式（スラッシュ）をWindows形式（バックスラッシュ）に統一
    $gitRootDir = $gitRootDir.Replace("/", "\")
    Write-Color "[情報] Gitリポジトリルート: $gitRootDir" "Green"
}

# 作業ディレクトリも同じ形式で表示（Windows形式に統一）
Write-Color "[情報] 作業ディレクトリ: $GIT_ROOT" "Green"

# $GIT_ROOT から $gitRootDir への相対パス（サブディレクトリパス）を計算
# 例: $gitRootDir = C:\repo, $GIT_ROOT = C:\repo\src → $subDirPath = src
$subDirPath = ""
$gitRootNormalized = $gitRootDir.TrimEnd("\")
$workDirNormalized = "$GIT_ROOT".TrimEnd("\")
if ($workDirNormalized.StartsWith($gitRootNormalized + "\")) {
    $subDirPath = $workDirNormalized.Substring($gitRootNormalized.Length + 1)
    Write-Color "[情報] 転送対象サブディレクトリ: $subDirPath" "Green"
}
Write-Host ""

#region ブランチ確認・切り替え
Write-Color "================================================================" "Cyan"
Write-Color "ブランチの確認" "Cyan"
Write-Color "================================================================" "Cyan"
Write-Host ""

$currentBranch = git branch --show-current 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Color "[エラー] ブランチ情報の取得に失敗しました" "Red"
    exit 1
}

Write-Host "  現在のブランチ: " -NoNewline
Write-Color $currentBranch "Yellow"
Write-Host ""
Write-Host " 1. このブランチで続行"
Write-Host " 2. ブランチを切り替える"
Write-Host ""
Write-Host " 0. キャンセル"
Write-Host ""

do {
    $branchChoice = Read-Host "番号を入力 (0-2)"
    if ($branchChoice -eq "0") {
        Write-Color "[キャンセル] 処理を中止しました" "Yellow"
        exit 0
    }
} while ($branchChoice -notin @("1", "2"))

if ($branchChoice -eq "2") {
    # ブランチ一覧を取得
    Write-Host ""
    Write-Color "================================================================" "Cyan"
    Write-Color "ブランチを選択してください" "Cyan"
    Write-Color "================================================================" "Cyan"
    Write-Host ""

    $branches = @(git branch --list 2>&1 | ForEach-Object {
        $_ -replace '^\*?\s*', ''
    })

    $currentBranchTrimmed = $currentBranch -replace '\s', ''

    $branchIndex = 1
    foreach ($branch in $branches) {
        $branchTrimmed = $branch -replace '\s', ''
        if ($branchTrimmed -eq $currentBranchTrimmed) {
            Write-Host ("{0,3}. {1} (現在)" -f $branchIndex, $branch) -ForegroundColor Yellow
        } else {
            Write-Host ("{0,3}. {1}" -f $branchIndex, $branch)
        }
        $branchIndex++
    }
    Write-Host ""
    Write-Host " 0. キャンセル"
    Write-Host ""

    do {
        $selectedBranchNum = Read-Host "番号を入力 (0-$($branches.Count))"
        if ($selectedBranchNum -eq "0") {
            Write-Color "[キャンセル] 処理を中止しました" "Yellow"
            exit 0
        }
        $selectedBranchIndex = [int]$selectedBranchNum - 1
    } while ($selectedBranchIndex -lt 0 -or $selectedBranchIndex -ge $branches.Count)

    $targetBranch = $branches[$selectedBranchIndex]
    $targetBranchTrimmed = $targetBranch -replace '\s', ''

    if ($targetBranchTrimmed -ne $currentBranchTrimmed) {
        Write-Host ""
        Write-Color "[実行] ブランチを切り替え中: $targetBranch" "Yellow"

        $checkoutResult = git checkout $targetBranch 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Color "[エラー] ブランチの切り替えに失敗しました" "Red"
            Write-Host $checkoutResult
            exit 1
        }

        Write-Color "[OK] ブランチを切り替えました: $targetBranch" "Green"
        $currentBranch = $targetBranch
    } else {
        Write-Color "[情報] 同じブランチが選択されました" "Yellow"
    }
}

Write-Host ""
Write-Color "[選択] ブランチ: $currentBranch" "Green"
Write-Host ""
#endregion

#region 環境選択
Write-Color "================================================================" "Cyan"
Write-Color "転送先環境を選択してください" "Cyan"
Write-Color "================================================================" "Cyan"
Write-Host ""

for ($i = 0; $i -lt $ENVIRONMENTS.Count; $i++) {
    Write-Host " $($i + 1). $($ENVIRONMENTS[$i].Name)"
}
Write-Host ""
Write-Host " 0. キャンセル"
Write-Host ""

do {
    $envChoice = Read-Host "番号を入力 (0-$($ENVIRONMENTS.Count))"
    if ($envChoice -eq "0") {
        Write-Color "[キャンセル] 処理を中止しました" "Yellow"
        exit 0
    }
    $envIndex = [int]$envChoice - 1
} while ($envIndex -lt 0 -or $envIndex -ge $ENVIRONMENTS.Count)

$selectedEnv = $ENVIRONMENTS[$envIndex]
$REMOTE_DIR = $selectedEnv.Path
$OWNER = $selectedEnv.Owner

Write-Host ""
Write-Color "[選択] 環境: $($selectedEnv.Name)" "Green"
Write-Color "[情報] 転送先: ${SSH_USER}@${SSH_HOST}:${REMOTE_DIR}" "Green"
Write-Color "[情報] オーナー: ${OWNER}:${COMMON_GROUP}" "Green"
Write-Host ""
#endregion

#region 転送モード選択
Write-Color "================================================================" "Cyan"
Write-Color "転送するファイルを選択" "Cyan"
Write-Color "================================================================" "Cyan"
Write-Host ""
Write-Host " 1. 変更されたファイルのみ (git status)"
Write-Host " 2. すべてのファイル"
Write-Host ""
Write-Host " 0. キャンセル"
Write-Host ""

do {
    $modeChoice = Read-Host "番号を入力 (0-2)"
    if ($modeChoice -eq "0") {
        Write-Color "[キャンセル] 処理を中止しました" "Yellow"
        exit 0
    }
} while ($modeChoice -notin @("1", "2"))

$transferMode = if ($modeChoice -eq "1") { "changed" } else { "all" }
$modeName = if ($modeChoice -eq "1") { "変更ファイルのみ" } else { "すべてのファイル" }

Write-Host ""
Write-Color "[選択] モード: $modeName" "Green"

if ($TARGET_EXTENSIONS.Count -gt 0) {
    Write-Color "[情報] 対象拡張子: $($TARGET_EXTENSIONS -join ' ')" "Green"
} else {
    Write-Color "[情報] 対象拡張子: すべて" "Green"
}

Write-Host ""
#endregion

#region ファイルリスト取得
Write-Host ""
Write-Color "[実行] ファイルリストを取得中..." "Yellow"
Write-Host ""

$fileList = @()

if ($transferMode -eq "changed") {
    # 変更されたファイルのみ
    Write-Color "[情報] Git status から変更ファイルを取得中..." "Yellow"

    $gitStatusOutput = git status --porcelain 2>&1

    if ($LASTEXITCODE -ne 0) {
        Write-Color "[エラー] Git status の取得に失敗しました" "Red"
        Write-Host $gitStatusOutput
        exit 1
    }

    $gitStatusOutput -split "`n" | ForEach-Object {
        $line = $_
        if ($line.Length -lt 4) { return }

        # ステータスコードを取得（最初の2文字）
        $status = $line.Substring(0, 2)
        # ファイルパス（2文字目以降、先頭の空白のみ除去）
        $filePath = $line.Substring(2).TrimStart()

        # 引用符で囲まれている場合は除去
        if ($filePath.StartsWith('"') -and $filePath.EndsWith('"') -and $filePath.Length -ge 2) {
            $filePath = $filePath.Substring(1, $filePath.Length - 2)
        }

        if ($filePath -eq "") { return }

        # 削除されたファイルを除外（ステータスコードに'D'が含まれる場合）
        # D  = ステージング済み削除
        #  D = ステージングされていない削除
        # DD = 両方
        if ($status -match 'D') {
            return
        }

        # ../ で始まるパスを除外
        if ($filePath -match '^\.\./') {
            return
        }

        # サブディレクトリパスを除去（配下フォルダからの相対パスに変換）
        # 例: $subDirPath = "src", $filePath = "src/main.c" → $filePath = "main.c"
        if ($subDirPath -ne "") {
            $subDirPathLinux = $subDirPath.Replace("\", "/")
            if ($filePath.StartsWith($subDirPathLinux + "/")) {
                $filePath = $filePath.Substring($subDirPathLinux.Length + 1)
            } else {
                # 配下フォルダ外のファイルは除外
                return
            }
        }

        # 拡張子フィルタ
        if ($TARGET_EXTENSIONS.Count -eq 0) {
            # フィルタなし
            $fileList += [PSCustomObject]@{
                Status = $status
                Path = $filePath
            }
        } else {
            $fileExt = [System.IO.Path]::GetExtension($filePath)
            if ($TARGET_EXTENSIONS -contains $fileExt) {
                $fileList += [PSCustomObject]@{
                    Status = $status
                    Path = $filePath
                }
            }
        }
    }
} else {
    # すべてのファイル
    Write-Color "[情報] リポジトリ内のすべての対象ファイルを取得中..." "Yellow"

    Get-ChildItem -Recurse -File | ForEach-Object {
        $fullPath = $_.FullName
        $relativePath = $fullPath.Replace("$GIT_ROOT\", "").Replace("\", "/")

        # .git ディレクトリは除外
        if ($relativePath -notmatch '^\.git/') {
            # 拡張子フィルタ
            if ($TARGET_EXTENSIONS.Count -eq 0) {
                # フィルタなし
                $fileList += [PSCustomObject]@{
                    Status = "A "
                    Path = $relativePath.Replace("/", "\")
                }
            } else {
                $fileExt = $_.Extension
                if ($TARGET_EXTENSIONS -contains $fileExt) {
                    $fileList += [PSCustomObject]@{
                        Status = "A "
                        Path = $relativePath.Replace("/", "\")
                    }
                }
            }
        }
    }
}

if ($fileList.Count -eq 0) {
    Write-Host ""
    Write-Color "[情報] 転送対象のファイルがありません" "Yellow"
    exit 0
}

Write-Host ""
Write-Color "[成功] $($fileList.Count) 個のファイルが見つかりました" "Green"
Write-Host ""
#endregion

#region ファイルリスト表示
Write-Color "================================================================" "Cyan"
Write-Color "転送予定のファイル一覧" "Cyan"
Write-Color "================================================================" "Cyan"
Write-Host ""

$index = 1
foreach ($file in $fileList) {
    $statusLabel = switch -Regex ($file.Status) {
        'M'   { "[変更]" }
        'A'   { "[追加]" }
        '\?\?' { "[未追跡]" }
        'R'   { "[リネーム]" }
        default { "[$($file.Status)]" }
    }

    Write-Host ("{0,3}. {1,-12} {2}" -f $index, $statusLabel, $file.Path)
    $index++
}

Write-Host ""
#endregion

#region 転送確認
Write-Color "これらのファイルを転送しますか？" "Yellow"
Write-Host ""
Write-Host " 1. すべて転送"
Write-Host " 2. 個別に選択"
Write-Host ""
Write-Host " 0. キャンセル"
Write-Host ""

do {
    $choice = Read-Host "番号を入力 (0-2)"
} while ($choice -notin @("0", "1", "2"))

if ($choice -eq "0") {
    Write-Color "[キャンセル] 転送を中止しました" "Yellow"
    exit 0
}

# 転送するファイルリスト
$filesToTransfer = @()

if ($choice -eq "1") {
    # すべて転送
    $filesToTransfer = $fileList
    Write-Color "[選択] すべてのファイルを転送します" "Green"

    # 「すべてのファイル」モードの場合、転送方法を選択
    if ($transferMode -eq "all") {
        # 転送ファイルの親ディレクトリを収集（重複排除）
        $parentDirs = @{}
        foreach ($file in $filesToTransfer) {
            $linuxPath = $file.Path.Replace("\", "/")
            $parentDir = Split-Path $linuxPath -Parent
            if ($parentDir) {
                $parentDir = $parentDir.Replace("\", "/")
            } else {
                $parentDir = ""  # ルート直下のファイル
            }
            $remoteParentDir = "${REMOTE_DIR}${parentDir}".TrimEnd("/")
            if (-not $parentDirs.ContainsKey($remoteParentDir)) {
                $parentDirs[$remoteParentDir] = $true
            }
        }

        Write-Host ""
        Write-Color "================================================================" "Cyan"
        Write-Color "転送方法を選択してください" "Cyan"
        Write-Color "================================================================" "Cyan"
        Write-Host ""
        Write-Host " 1. 上書き転送（既存ファイルを残す）"
        Write-Host " 2. クリーンアップしてから転送（対象フォルダの中身を削除）"
        Write-Host ""
        Write-Host " 0. キャンセル"
        Write-Host ""

        do {
            $transferMethod = Read-Host "番号を入力 (0-2)"
            if ($transferMethod -eq "0") {
                Write-Color "[キャンセル] 処理を中止しました" "Yellow"
                exit 0
            }
        } while ($transferMethod -notin @("1", "2"))

        if ($transferMethod -eq "2") {
            Write-Host ""
            Write-Color "================================================================" "Yellow"
            Write-Color "  警告: 転送先フォルダのクリーンアップ" "Yellow"
            Write-Color "================================================================" "Yellow"
            Write-Host ""
            Write-Color "以下のフォルダ内のファイルが削除されます:" "Yellow"
            Write-Host ""
            Write-Color "削除対象フォルダ:" "Red"
            foreach ($dir in $parentDirs.Keys | Sort-Object) {
                Write-Host "  - ${dir}/*"
            }
            Write-Host ""
            Write-Color "※ ファイルのみ再帰的に削除（ディレクトリ構造は残ります）" "Gray"
            Write-Host ""

            do {
                $cleanupConfirm = Read-Host "本当にクリーンアップしますか？ (y/n)"
                $cleanupConfirm = $cleanupConfirm.ToLower()
            } while ($cleanupConfirm -notin @("y", "n"))

            if ($cleanupConfirm -eq "y") {
                $doCleanup = $true
                $cleanupDirs = $parentDirs.Keys
                Write-Color "[選択] 転送前にクリーンアップを実行します" "Green"
            } else {
                $doCleanup = $false
                Write-Color "[選択] クリーンアップをキャンセルしました（上書きモード）" "Yellow"
            }
        } else {
            $doCleanup = $false
            Write-Color "[選択] 上書きモードで転送します" "Green"
        }
    }
} else {
    # 個別選択
    Write-Host ""
    Write-Color "================================================================" "Cyan"
    Write-Color "個別ファイル選択" "Cyan"
    Write-Color "================================================================" "Cyan"
    Write-Host ""

    foreach ($file in $fileList) {
        do {
            $answer = Read-Host "転送: $($file.Path) (y/n)"
            $answer = $answer.ToLower()
        } while ($answer -notin @("y", "n"))

        if ($answer -eq "y") {
            $filesToTransfer += $file
        }
    }

    if ($filesToTransfer.Count -eq 0) {
        Write-Color "[情報] 転送するファイルが選択されませんでした" "Yellow"
        exit 0
    }

    Write-Host ""
    Write-Color "[選択] $($filesToTransfer.Count) 個のファイルを転送します" "Green"
}

Write-Host ""
#endregion

# SCP/SSHコマンドを設定
$scpCommand = "scp.exe"
$sshCommand = "ssh.exe"

#region クリーンアップ処理
if ($doCleanup -eq $true) {
    Write-Header "転送先フォルダのクリーンアップ"

    foreach ($dir in $cleanupDirs | Sort-Object) {
        Write-Color "[削除] ${dir}/* ..." "Yellow"

        $sshArgs = @()

        if ($SSH_KEY -ne "" -and (Test-Path $SSH_KEY)) {
            $sshArgs += "-i"
            $sshArgs += $SSH_KEY
        }

        if ($SSH_PORT -ne 22) {
            $sshArgs += "-p"
            $sshArgs += $SSH_PORT
        }

        $sshArgs += "${SSH_USER}@${SSH_HOST}"
        # フォルダ内のファイルを再帰的に削除（ディレクトリ構造は残す）
        $sshArgs += "find '${dir}' -type f -delete 2>/dev/null; echo 'done'"

        try {
            $result = & $sshCommand $sshArgs 2>&1
            Write-Color "  [OK] クリーンアップ完了" "Green"
        } catch {
            Write-Color "  [警告] クリーンアップ中にエラー: $($_.Exception.Message)" "Yellow"
        }
    }

    Write-Host ""
    Write-Color "[完了] クリーンアップが完了しました" "Green"
    Write-Host ""
}
#endregion

#region ファイル転送
Write-Header "ファイル転送開始"

$successCount = 0
$failCount = 0
$failedFiles = @()

foreach ($file in $filesToTransfer) {
    # ローカルパスは配下フォルダ（$GIT_ROOT）からの相対パスで計算
    $localPath = Join-Path $GIT_ROOT $file.Path

    # Windowsのパス区切り(\)をLinux形式(/)に変換
    $linuxPath = $file.Path.Replace("\", "/")

    # リモートパスを計算（$file.Pathは既に配下フォルダからの相対パス）
    $remotePath = "${REMOTE_DIR}${linuxPath}"

    Write-Color "[転送] $($file.Path)" "Cyan"
    Write-Host "  ローカル: $localPath"
    Write-Host "  リモート: ${SSH_USER}@${SSH_HOST}:${remotePath}"

    # Linux側で親ディレクトリを作成
    $parentDir = Split-Path $linuxPath -Parent
    if ($parentDir) {
        $parentDir = $parentDir.Replace("\", "/")
        $remoteParentDir = "${REMOTE_DIR}${parentDir}"

        $sshArgs = @()

        if ($SSH_KEY -ne "" -and (Test-Path $SSH_KEY)) {
            $sshArgs += "-i"
            $sshArgs += $SSH_KEY
        }

        if ($SSH_PORT -ne 22) {
            $sshArgs += "-p"
            $sshArgs += $SSH_PORT
        }

        $sshArgs += "${SSH_USER}@${SSH_HOST}"
        $sshArgs += "mkdir -p '$remoteParentDir' && chmod $LINUX_CHMOD_DIR '$remoteParentDir' && chown ${OWNER}:${COMMON_GROUP} '$remoteParentDir'"

        & $sshCommand $sshArgs 2>&1 | Out-Null
    }

    # SCPコマンド構築
    $scpArgs = @()

    if ($SSH_KEY -ne "" -and (Test-Path $SSH_KEY)) {
        $scpArgs += "-i"
        $scpArgs += $SSH_KEY
    }

    if ($SSH_PORT -ne 22) {
        $scpArgs += "-P"
        $scpArgs += $SSH_PORT
    }

    # ソースと宛先
    $scpArgs += $localPath
    $scpArgs += "${SSH_USER}@${SSH_HOST}:${remotePath}"

    # 実行
    try {
        & $scpCommand $scpArgs 2>&1 | Out-Null

        if ($LASTEXITCODE -eq 0) {
            # パーミッションと所有者を設定
            $sshArgs = @()

            if ($SSH_KEY -ne "" -and (Test-Path $SSH_KEY)) {
                $sshArgs += "-i"
                $sshArgs += $SSH_KEY
            }

            if ($SSH_PORT -ne 22) {
                $sshArgs += "-p"
                $sshArgs += $SSH_PORT
            }

            $sshArgs += "${SSH_USER}@${SSH_HOST}"
            $sshArgs += "chmod $LINUX_CHMOD_FILE '$remotePath' && chown ${OWNER}:${COMMON_GROUP} '$remotePath'"

            & $sshCommand $sshArgs 2>&1 | Out-Null

            Write-Color "  [OK] 成功" "Green"
            $successCount++
        } else {
            Write-Color "  [NG] 失敗 (終了コード: $LASTEXITCODE)" "Red"
            $failCount++
            $failedFiles += $file.Path
        }
    } catch {
        Write-Color "  [NG] 失敗: $($_.Exception.Message)" "Red"
        $failCount++
        $failedFiles += $file.Path
    }
}

Write-Host ""
#endregion

#region 結果サマリー
Write-Header "転送結果"

Write-Color "成功: $successCount ファイル" "Green"
if ($failCount -gt 0) {
    Write-Color "失敗: $failCount ファイル" "Red"
    Write-Host ""
    Write-Color "失敗したファイル:" "Yellow"
    foreach ($failedFile in $failedFiles) {
        Write-Host "  - $failedFile"
    }
}

Write-Host ""

if ($failCount -eq 0) {
    Write-Color "すべてのファイル転送が完了しました！" "Green"
    exit 0
} else {
    Write-Color "一部のファイル転送に失敗しました" "Yellow"
    exit 1
}
#endregion
