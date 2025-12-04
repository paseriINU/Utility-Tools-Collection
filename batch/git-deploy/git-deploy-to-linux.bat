<# :
@echo off
setlocal
powershell -NoProfile -ExecutionPolicy Bypass -Command "iex ((gc '%~f0') -join \"`n\")"
exit /b %ERRORLEVEL%
: #>

#==============================================================================
# Git Deploy to Linux - Gitの変更ファイルをLinuxサーバーに転送
#==============================================================================
#
# 機能:
#   1. Git status から変更/追加されたファイルを取得
#   2. 削除されたファイルは除外
#   3. ユーザーに全部転送 or 個別選択を確認
#   4. SCP/PSCP でLinuxサーバーに転送
#
# 必要な環境:
#   - Git がインストールされていること
#   - SCP (Windows OpenSSH Client) または PSCP (PuTTY) が利用可能であること
#   - SSH公開鍵認証またはパスワード認証が設定されていること
#
#==============================================================================

#region 設定 - ここを編集してください
#==============================================================================

# 転送先サーバー情報
$SSH_USER = "youruser"              # SSHユーザー名
$SSH_HOST = "192.168.1.100"         # SSHホスト名またはIPアドレス
$SSH_PORT = 22                      # SSHポート番号

# 転送先ディレクトリ (Linuxサーバー上のパス)
$REMOTE_DIR = "/home/youruser/project"

# SSH秘密鍵ファイル (公開鍵認証を使用する場合)
# パスワード認証の場合は空文字列 ""
$SSH_KEY = "$env:USERPROFILE\.ssh\id_rsa"

# Git リポジトリのルートディレクトリ (空文字列の場合は現在のディレクトリ)
$GIT_ROOT = ""

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
    Write-Color "========================================" "Cyan"
    Write-Color "  $Text" "Cyan"
    Write-Color "========================================" "Cyan"
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
Write-Color "[情報] 転送先: ${SSH_USER}@${SSH_HOST}:${REMOTE_DIR}" "Green"
Write-Host ""

#region Git Status取得とフィルタリング
Write-Color "[実行] Git status を取得中..." "Yellow"

$gitStatusOutput = git status --porcelain 2>&1

if ($LASTEXITCODE -ne 0) {
    Write-Color "[エラー] Git status の取得に失敗しました" "Red"
    Write-Host $gitStatusOutput
    exit 1
}

# ファイルリストを配列に変換（削除されたファイルを除外）
$fileList = @()
$gitStatusOutput -split "`n" | ForEach-Object {
    $line = $_.Trim()
    if ($line -eq "") { return }

    # ステータスコードを取得（最初の2文字）
    $status = $line.Substring(0, 2)
    $filePath = $line.Substring(3).Trim()

    # 削除されたファイル (D で始まる) を除外
    if ($status -notmatch '^.?D') {
        $fileList += [PSCustomObject]@{
            Status = $status
            Path = $filePath
        }
    }
}

if ($fileList.Count -eq 0) {
    Write-Color "[情報] 転送するファイルがありません" "Yellow"
    Write-Color "       変更されたファイルが見つからないか、すべて削除されたファイルです" "Yellow"
    exit 0
}

Write-Color "[成功] $($fileList.Count) 個のファイルが見つかりました（削除ファイルを除く）" "Green"
Write-Host ""

#endregion

#region ファイルリスト表示
Write-Color "========================================" "Cyan"
Write-Color "転送予定のファイル一覧" "Cyan"
Write-Color "========================================" "Cyan"

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

#region ユーザー確認（全部 or 個別）
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
    Write-Color "========================================" "Cyan"
    Write-Color "個別ファイル選択" "Cyan"
    Write-Color "========================================" "Cyan"
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

#region SCP/PSCP検出
Write-Color "[チェック] SCPコマンドを検出中..." "Yellow"

$scpCommand = $null
$scpType = $null

# Windows OpenSSH Client の scp.exe を優先
$scpExe = Get-Command scp.exe -ErrorAction SilentlyContinue
if ($scpExe) {
    $scpCommand = "scp.exe"
    $scpType = "OpenSSH"
    Write-Color "[検出] Windows OpenSSH Client (scp.exe)" "Green"
} else {
    # PSCP (PuTTY) をフォールバック
    $pscpExe = Get-Command pscp.exe -ErrorAction SilentlyContinue
    if ($pscpExe) {
        $scpCommand = "pscp.exe"
        $scpType = "PuTTY"
        Write-Color "[検出] PuTTY PSCP (pscp.exe)" "Green"
    } else {
        Write-Color "[エラー] SCPコマンドが見つかりません" "Red"
        Write-Host ""
        Write-Host "以下のいずれかをインストールしてください："
        Write-Host "  1. Windows OpenSSH Client (推奨)"
        Write-Host "     設定 > アプリ > オプション機能 > OpenSSH クライアント"
        Write-Host ""
        Write-Host "  2. PuTTY PSCP"
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
    $remotePath = "${SSH_USER}@${SSH_HOST}:${REMOTE_DIR}/$($file.Path)"

    Write-Color "[転送] $($file.Path)" "Cyan"

    # SCPコマンド構築
    $scpArgs = @()

    if ($SSH_KEY -ne "" -and (Test-Path $SSH_KEY)) {
        # 公開鍵認証
        if ($scpType -eq "OpenSSH") {
            $scpArgs += "-i"
            $scpArgs += $SSH_KEY
        } else {
            # PSCP
            $scpArgs += "-i"
            $scpArgs += $SSH_KEY
        }
    }

    if ($SSH_PORT -ne 22) {
        if ($scpType -eq "OpenSSH") {
            $scpArgs += "-P"
            $scpArgs += $SSH_PORT
        } else {
            # PSCP
            $scpArgs += "-P"
            $scpArgs += $SSH_PORT
        }
    }

    # バッチモード（対話なし）
    if ($scpType -eq "PuTTY") {
        $scpArgs += "-batch"
    }

    # ソースと宛先
    $scpArgs += $localPath
    $scpArgs += $remotePath

    # 実行
    try {
        & $scpCommand $scpArgs 2>&1 | Out-Null

        if ($LASTEXITCODE -eq 0) {
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

Write-Color "成功: $successCount 個" "Green"
if ($failCount -gt 0) {
    Write-Color "失敗: $failCount 個" "Red"
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
