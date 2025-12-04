<# :
@echo off
setlocal
powershell -NoProfile -ExecutionPolicy Bypass -Command "iex ((gc '%~f0') -join \"`n\")"
exit /b %ERRORLEVEL%
: #>

#==============================================================================
# Git Deploy to Linux - 統合版
#==============================================================================
#
# 機能:
#   1. 複数環境から転送先を選択
#   2. 変更ファイルのみ OR すべてのファイルを選択
#   3. 拡張子フィルタ (.c .pc .h など)
#   4. 削除されたファイルは自動除外
#   5. 全部転送 or 個別選択
#   6. Linux側でディレクトリ自動作成・パーミッション設定
#   7. SCP/PSCP 自動検出
#
# 必要な環境:
#   - Git がインストールされていること
#   - SCP (Windows OpenSSH Client) または PSCP (PuTTY) が利用可能であること
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

if (-not (Test-Path ".git")) {
    Write-Color "[エラー] Gitリポジトリではありません: $GIT_ROOT" "Red"
    exit 1
}

Write-Header "Git Deploy to Linux"

Write-Color "[情報] Gitリポジトリ: $GIT_ROOT" "Green"
Write-Host ""

#region 環境選択
Write-Color "================================================================" "Cyan"
Write-Color "転送先環境を選択してください" "Cyan"
Write-Color "================================================================" "Cyan"
Write-Host ""

for ($i = 0; $i -lt $ENVIRONMENTS.Count; $i++) {
    Write-Host "$($i + 1). $($ENVIRONMENTS[$i].Name)"
}

Write-Host ""

do {
    $envChoice = Read-Host "番号を入力 (1-$($ENVIRONMENTS.Count))"
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
Write-Host "1. 変更されたファイルのみ (git status)"
Write-Host "2. すべてのファイル"
Write-Host ""

do {
    $modeChoice = Read-Host "番号を入力 (1-2)"
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
Write-Host "続行するには何かキーを押してください..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
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
        $line = $_.Trim()
        if ($line -eq "") { return }

        # ステータスコードを取得（最初の2文字）
        $status = $line.Substring(0, 2)
        $filePath = $line.Substring(3).Trim()

        # 削除されたファイル (D で始まる) を除外
        if ($status -notmatch '^.?D') {
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
        '^M'  { "[変更]" }
        '^A'  { "[追加]" }
        '^\?\?' { "[未追跡]" }
        '^R'  { "[リネーム]" }
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
Write-Host "  [A] すべて転送"
Write-Host "  [I] 個別に選択"
Write-Host "  [C] キャンセル"
Write-Host ""

do {
    $choice = Read-Host "選択してください (A/I/C)"
    $choice = $choice.ToUpper()
} while ($choice -notin @("A", "I", "C"))

if ($choice -eq "C") {
    Write-Color "[キャンセル] 転送を中止しました" "Yellow"
    exit 0
}

# 転送するファイルリスト
$filesToTransfer = @()

if ($choice -eq "A") {
    # すべて転送
    $filesToTransfer = $fileList
    Write-Color "[選択] すべてのファイルを転送します" "Green"
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
Write-Host "続行するには何かキーを押してください..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
#endregion

#region SCP/SSH検出
Write-Host ""
Write-Color "[チェック] SCP/SSHコマンドを検出中..." "Yellow"

$scpCommand = $null
$sshCommand = $null
$scpType = $null

# Windows OpenSSH Client の scp.exe を優先
$scpExe = Get-Command scp.exe -ErrorAction SilentlyContinue
$sshExe = Get-Command ssh.exe -ErrorAction SilentlyContinue

if ($scpExe -and $sshExe) {
    $scpCommand = "scp.exe"
    $sshCommand = "ssh.exe"
    $scpType = "OpenSSH"
    Write-Color "[検出] Windows OpenSSH Client (scp.exe, ssh.exe)" "Green"
} else {
    # PSCP/PLINK (PuTTY) をフォールバック
    $pscpExe = Get-Command pscp.exe -ErrorAction SilentlyContinue
    $plinkExe = Get-Command plink.exe -ErrorAction SilentlyContinue

    if ($pscpExe -and $plinkExe) {
        $scpCommand = "pscp.exe"
        $sshCommand = "plink.exe"
        $scpType = "PuTTY"
        Write-Color "[検出] PuTTY PSCP/PLINK (pscp.exe, plink.exe)" "Green"
    } else {
        Write-Color "[エラー] SCP/SSHコマンドが見つかりません" "Red"
        Write-Host ""
        Write-Host "以下のいずれかをインストールしてください："
        Write-Host "  1. Windows OpenSSH Client (推奨)"
        Write-Host "     設定 > アプリ > オプション機能 > OpenSSH クライアント"
        Write-Host ""
        Write-Host "  2. PuTTY PSCP/PLINK"
        Write-Host "     https://www.putty.org/ からダウンロード"
        exit 1
    }
}

Write-Host ""
#endregion

#region ファイル転送
Write-Header "ファイル転送開始"

$successCount = 0
$failCount = 0
$failedFiles = @()

foreach ($file in $filesToTransfer) {
    $localPath = Join-Path $GIT_ROOT $file.Path

    # Windowsのパス区切り(\)をLinux形式(/)に変換
    $linuxPath = $file.Path.Replace("\", "/")
    $remotePath = "${REMOTE_DIR}${linuxPath}"

    Write-Color "[転送] $($file.Path)" "Cyan"

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
            if ($scpType -eq "OpenSSH") {
                $sshArgs += "-p"
            } else {
                $sshArgs += "-P"
            }
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
        if ($scpType -eq "OpenSSH") {
            $scpArgs += "-P"
        } else {
            $scpArgs += "-P"
        }
        $scpArgs += $SSH_PORT
    }

    # バッチモード（対話なし）
    if ($scpType -eq "PuTTY") {
        $scpArgs += "-batch"
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
                if ($scpType -eq "OpenSSH") {
                    $sshArgs += "-p"
                } else {
                    $sshArgs += "-P"
                }
                $sshArgs += $SSH_PORT
            }

            $sshArgs += "${SSH_USER}@${SSH_HOST}"
            $sshArgs += "chmod $LINUX_CHMOD_FILE '$remotePath' && chown ${OWNER}:${COMMON_GROUP} '$remotePath'"

            & $sshCommand $sshArgs 2>&1 | Out-Null

            Write-Color "  ✓ 成功" "Green"
            $successCount++
        } else {
            Write-Color "  ✗ 失敗 (終了コード: $LASTEXITCODE)" "Red"
            $failCount++
            $failedFiles += $file.Path
        }
    } catch {
        Write-Color "  ✗ 失敗: $($_.Exception.Message)" "Red"
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
