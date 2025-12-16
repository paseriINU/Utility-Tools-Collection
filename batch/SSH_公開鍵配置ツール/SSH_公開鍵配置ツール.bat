<# :
@echo off
chcp 65001 >nul
title SSH公開鍵配置ツール
setlocal

powershell -NoProfile -ExecutionPolicy Bypass -Command "iex ((gc '%~f0' -Encoding UTF8) -join \"`n\")"
set EXITCODE=%ERRORLEVEL%

pause
exit /b %EXITCODE%
: #>

#==============================================================================
# SSH公開鍵配置ツール
#==============================================================================
#
# 機能:
#   1. SSH鍵ペア（公開鍵・秘密鍵）の生成
#   2. Linuxサーバーへの公開鍵の自動配置（authorized_keysに追記）
#   3. 接続テストの実行
#
# 必要な環境:
#   - Windows OpenSSH Client（Windows 10 1809以降は標準搭載）
#   - Linuxサーバーへのパスワード認証が有効であること
#
#==============================================================================

#region 設定 - ここを編集してください
#==============================================================================

# Linuxサーバー接続情報
$SSH_HOST = "linux-server"      # ホスト名またはIPアドレス
$SSH_USER = "youruser"          # SSHユーザー名
$SSH_PORT = 22                  # SSHポート番号

# 鍵ファイル設定
$KEY_NAME = "id_ed25519"        # 鍵ファイル名（id_ed25519, id_rsa など）
$KEY_TYPE = "ed25519"           # 鍵の種類: "ed25519"（推奨） または "rsa"
$KEY_BITS = 4096                # RSAの場合のビット数（ed25519では無視）
$USE_PASSPHRASE = $false        # パスフレーズを使用するか（$true / $false）

#==============================================================================
#endregion

# UTF-8出力設定
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# PCのユーザー名とコンピューター名を自動取得してコメントを生成
$KEY_COMMENT = "$env:USERNAME@$env:COMPUTERNAME"

#region 関数定義
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
#endregion

# タイトル表示
Write-Host ""
Write-Color "================================================================" "Cyan"
Write-Color "  SSH公開鍵配置ツール" "Cyan"
Write-Color "================================================================" "Cyan"
Write-Host ""

#region OpenSSH確認
$sshKeygen = Get-Command ssh-keygen.exe -ErrorAction SilentlyContinue
$sshCommand = Get-Command ssh.exe -ErrorAction SilentlyContinue

if (-not $sshKeygen -or -not $sshCommand) {
    Write-Color "[エラー] Windows OpenSSH Clientがインストールされていません" "Red"
    Write-Host ""
    Write-Color "インストール方法:" "Yellow"
    Write-Host "  1. 設定 > アプリ > オプション機能"
    Write-Host "  2. 機能の追加 > OpenSSH クライアント"
    Write-Host ""
    exit 1
}

Write-Color "[OK] OpenSSH Client が利用可能です" "Green"
Write-Host ""
#endregion

#region 設定表示
Write-Color "================================================================" "Cyan"
Write-Color "設定内容" "Cyan"
Write-Color "================================================================" "Cyan"
Write-Host ""
Write-Host "  接続先: ${SSH_USER}@${SSH_HOST}:${SSH_PORT}"
Write-Host "  鍵の種類: $KEY_TYPE $(if ($KEY_TYPE -eq 'rsa') { "($KEY_BITS bit)" })"
Write-Host "  鍵ファイル名: $KEY_NAME"
Write-Host "  鍵のコメント: $KEY_COMMENT"
Write-Host "  パスフレーズ: $(if ($USE_PASSPHRASE) { 'あり' } else { 'なし' })"
Write-Host "  配置方法: authorized_keys に追記（重複チェックあり）"
Write-Host ""

do {
    $confirm = Read-Host "この設定で実行しますか？ (y/n)"
    $confirm = $confirm.ToLower()
} while ($confirm -notin @("y", "n"))

if ($confirm -eq "n") {
    Write-Color "[キャンセル] 処理を中止しました" "Yellow"
    Write-Host "  設定を変更する場合は、バッチファイルの「設定」セクションを編集してください"
    exit 0
}
#endregion

#region 鍵ファイルパス設定
$sshDir = "$env:USERPROFILE\.ssh"
$keyPath = "$sshDir\$KEY_NAME"

# .sshディレクトリがなければ作成
if (-not (Test-Path $sshDir)) {
    New-Item -ItemType Directory -Path $sshDir -Force | Out-Null
    Write-Color "[作成] $sshDir ディレクトリを作成しました" "Green"
}
#endregion

#region 鍵生成
Write-Header "SSH鍵ペアの生成"

# 既存の鍵があるか確認
if (Test-Path $keyPath) {
    Write-Color "[警告] 既に鍵ファイルが存在します: $keyPath" "Yellow"
    Write-Host ""
    do {
        $overwrite = Read-Host "上書きしますか？ (y/n)"
        $overwrite = $overwrite.ToLower()
    } while ($overwrite -notin @("y", "n"))

    if ($overwrite -eq "n") {
        Write-Color "[スキップ] 既存の鍵を使用します" "Yellow"
    } else {
        # 既存の鍵を削除
        Remove-Item -Path $keyPath -Force -ErrorAction SilentlyContinue
        Remove-Item -Path "$keyPath.pub" -Force -ErrorAction SilentlyContinue
        Write-Color "[削除] 既存の鍵を削除しました" "Yellow"
    }
}

# 鍵が存在しない場合のみ生成
if (-not (Test-Path $keyPath)) {
    Write-Color "[実行] 鍵ペアを生成中..." "Yellow"
    Write-Host ""

    if ($KEY_TYPE -eq "rsa") {
        if ($USE_PASSPHRASE) {
            & ssh-keygen.exe -t rsa -b $KEY_BITS -C $KEY_COMMENT -f $keyPath
        } else {
            & ssh-keygen.exe -t rsa -b $KEY_BITS -C $KEY_COMMENT -f $keyPath -N '""'
        }
    } else {
        if ($USE_PASSPHRASE) {
            & ssh-keygen.exe -t ed25519 -C $KEY_COMMENT -f $keyPath
        } else {
            & ssh-keygen.exe -t ed25519 -C $KEY_COMMENT -f $keyPath -N '""'
        }
    }

    if ($LASTEXITCODE -eq 0 -and (Test-Path $keyPath)) {
        Write-Host ""
        Write-Color "[成功] 鍵ペアを生成しました" "Green"
        Write-Host "  秘密鍵: $keyPath"
        Write-Host "  公開鍵: $keyPath.pub"
        Write-Host "  コメント: $KEY_COMMENT"
    } else {
        Write-Color "[エラー] 鍵の生成に失敗しました" "Red"
        exit 1
    }
}
#endregion

#region 公開鍵の配置
Write-Header "公開鍵のLinuxサーバーへの配置"

Write-Color "パスワード認証でLinuxサーバーに接続し、公開鍵を配置します。" "Yellow"
Write-Color "パスワードの入力を求められたら、Linuxサーバーのパスワードを入力してください。" "Yellow"
Write-Host ""
Write-Color "※ 既存の authorized_keys に追記します（重複する場合はスキップ）" "Gray"
Write-Host ""

# 公開鍵の内容を取得
$pubKeyContent = Get-Content "$keyPath.pub" -Raw
$pubKeyContent = $pubKeyContent.Trim()

Write-Color "[実行] 公開鍵を配置中..." "Yellow"
Write-Host ""

# SSHコマンドで公開鍵を配置（追記モード）
$sshScript = @"
mkdir -p ~/.ssh && chmod 700 ~/.ssh && touch ~/.ssh/authorized_keys && chmod 600 ~/.ssh/authorized_keys && if ! grep -qF '$pubKeyContent' ~/.ssh/authorized_keys 2>/dev/null; then echo '$pubKeyContent' >> ~/.ssh/authorized_keys && echo '[OK] 公開鍵を追記しました'; else echo '[情報] 公開鍵は既に登録済みです'; fi
"@

$sshArgs = @()
if ($SSH_PORT -ne 22) {
    $sshArgs += "-p"
    $sshArgs += $SSH_PORT
}
$sshArgs += "${SSH_USER}@${SSH_HOST}"
$sshArgs += $sshScript

& ssh.exe $sshArgs

if ($LASTEXITCODE -eq 0) {
    Write-Host ""
    Write-Color "[成功] 公開鍵の配置が完了しました" "Green"
} else {
    Write-Host ""
    Write-Color "[エラー] 公開鍵の配置に失敗しました" "Red"
    Write-Host "  - パスワードが正しいか確認してください"
    Write-Host "  - Linuxサーバーでパスワード認証が有効か確認してください"
    exit 1
}
#endregion

#region 接続テスト
Write-Header "SSH接続テスト（公開鍵認証）"

Write-Color "[実行] 公開鍵認証で接続テスト中..." "Yellow"
Write-Host ""

$testArgs = @()
$testArgs += "-o"
$testArgs += "BatchMode=yes"
$testArgs += "-o"
$testArgs += "ConnectTimeout=10"

if ($SSH_PORT -ne 22) {
    $testArgs += "-p"
    $testArgs += $SSH_PORT
}

$testArgs += "-i"
$testArgs += $keyPath
$testArgs += "${SSH_USER}@${SSH_HOST}"
$testArgs += "echo '[OK] SSH公開鍵認証で接続成功！'"

& ssh.exe $testArgs

if ($LASTEXITCODE -eq 0) {
    Write-Host ""
    Write-Color "================================================================" "Green"
    Write-Color "  SSH公開鍵認証の設定が完了しました！" "Green"
    Write-Color "================================================================" "Green"
    Write-Host ""
    Write-Color "以下のコマンドでパスワードなしで接続できます:" "Cyan"
    Write-Host ""
    Write-Host "  ssh -i `"$keyPath`" ${SSH_USER}@${SSH_HOST}"
    Write-Host ""
    Write-Color "他のツールで使用する場合の設定値:" "Cyan"
    Write-Host ""
    Write-Host "  `$SSH_KEY = `"$keyPath`""
    Write-Host "  `$SSH_USER = `"$SSH_USER`""
    Write-Host "  `$SSH_HOST = `"$SSH_HOST`""
    Write-Host "  `$SSH_PORT = $SSH_PORT"
    Write-Host ""
} else {
    Write-Host ""
    Write-Color "[警告] 公開鍵認証での接続に失敗しました" "Yellow"
    Write-Host ""
    Write-Host "考えられる原因:"
    Write-Host "  - Linuxサーバー側の設定問題"
    Write-Host "    - /etc/ssh/sshd_config で PubkeyAuthentication yes が必要"
    Write-Host "  - パーミッションの問題"
    Write-Host "    - ~/.ssh は 700"
    Write-Host "    - ~/.ssh/authorized_keys は 600"
    Write-Host ""
    exit 1
}
#endregion

Write-Host ""
Write-Color "[完了] 処理が終了しました" "Green"
