<# :
@echo off
chcp 65001 >nul
title Git 開発環境初期設定ツール
setlocal

powershell -NoProfile -ExecutionPolicy Bypass -Command "iex ((gc '%~f0' -Encoding UTF8) -join \"`n\")"
set EXITCODE=%ERRORLEVEL%

pause
exit /b %EXITCODE%
: #>

#==============================================================================
# Git 開発環境初期設定ツール
#==============================================================================
#
# 機能:
#   1. Git globalの user.name / user.email 自動設定（未設定の場合）
#   2. SSH鍵ペア（公開鍵・秘密鍵）の生成
#   3. Linuxサーバーへの公開鍵の自動配置（authorized_keysに追記）
#   4. 接続テストの実行
#   5. Gitリポジトリのクローン（設定されている場合）
#   6. WinMergeフィルター設定（.gitignore等を除外）
#
# 必要な環境:
#   - Windows OpenSSH Client（Windows 10 1809以降は標準搭載）
#   - Linuxサーバーへのパスワード認証が有効であること
#   - Git for Windows（クローン機能を使用する場合）
#   - WinMerge（フィルター設定機能を使用する場合）
#
#==============================================================================

#region 設定 - ここを編集してください
#==============================================================================

# Linuxサーバー接続情報
$SSH_HOST = "linux-server"      # ホスト名またはIPアドレス
$SSH_USER = "youruser"          # SSHユーザー名
$SSH_PORT = 22                  # SSHポート番号
$SSH_PASSWORD = ""              # SSHパスワード（空欄の場合は実行時に入力を求めます）

# 鍵ファイル設定
$KEY_NAME = "id_ed25519"        # 鍵ファイル名（id_ed25519, id_rsa など）
$KEY_TYPE = "ed25519"           # 鍵の種類: "ed25519"（推奨） または "rsa"
$KEY_BITS = 4096                # RSAの場合のビット数（ed25519では無視）
$USE_PASSPHRASE = $false        # パスフレーズを使用するか（$true / $false）

# Gitリポジトリクローン設定（空欄の場合はクローンをスキップ）
$GIT_CLONE_URL = ""             # クローンするリポジトリのURL（SSH形式）
                                # 例: $GIT_CLONE_URL = "git@github.com:user/repo.git"
                                # 例: $GIT_CLONE_URL = "ssh://git@linux-server/path/to/repo.git"
$GIT_LOCAL_PATH = ""            # クローン先のローカルパス
                                # 例: $GIT_LOCAL_PATH = "C:\Projects\MyRepo"

# WinMergeフィルター設定
$SETUP_WINMERGE_FILTER = $true  # WinMergeフィルターを設定するか（$true / $false）
$WINMERGE_EXCLUDE_FILES = @(    # 比較時に除外するファイル名パターン
    ".gitignore",
    ".gitkeep",
    ".gitattributes"
)

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
Write-Color "  Git 開発環境初期設定ツール" "Cyan"
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

#region Git設定確認・自動設定
Write-Header "Git設定の確認"

# Gitがインストールされているか確認
$gitCommand = Get-Command git.exe -ErrorAction SilentlyContinue
if (-not $gitCommand) {
    Write-Color "[警告] Gitがインストールされていません" "Yellow"
    Write-Host "  Gitリポジトリのクローン機能は使用できません"
    Write-Host ""
    $GIT_AVAILABLE = $false
} else {
    $GIT_AVAILABLE = $true
    Write-Color "[OK] Git が利用可能です" "Green"
    Write-Host ""

    # user.name の確認・設定
    $currentUserName = git config --global user.name 2>$null
    if ($currentUserName) {
        Write-Host "  user.name: $currentUserName (設定済み)"
    } else {
        $newUserName = $env:USERNAME
        git config --global user.name $newUserName
        Write-Color "  user.name: $newUserName (自動設定しました)" "Green"
    }

    # user.email の確認・設定
    $currentUserEmail = git config --global user.email 2>$null
    if ($currentUserEmail) {
        Write-Host "  user.email: $currentUserEmail (設定済み)"
    } else {
        $newUserEmail = "$env:USERNAME@$env:COMPUTERNAME"
        git config --global user.email $newUserEmail
        Write-Color "  user.email: $newUserEmail (自動設定しました)" "Green"
    }
    Write-Host ""
}

Write-Host ""
Read-Host "続行するには Enter を押してください"
#endregion

#region 設定表示
Write-Color "================================================================" "Cyan"
Write-Color "設定内容" "Cyan"
Write-Color "================================================================" "Cyan"
Write-Host ""
Write-Host "  接続先: ${SSH_USER}@${SSH_HOST}:${SSH_PORT}"
Write-Host "  SSHパスワード: $(if ($SSH_PASSWORD) { '設定済み' } else { '実行時に入力' })"
Write-Host "  鍵の種類: $KEY_TYPE $(if ($KEY_TYPE -eq 'rsa') { "($KEY_BITS bit)" })"
Write-Host "  鍵ファイル名: $KEY_NAME"
Write-Host "  鍵のコメント: $KEY_COMMENT"
Write-Host "  パスフレーズ: $(if ($USE_PASSPHRASE) { 'あり' } else { 'なし' })"
Write-Host "  配置方法: authorized_keys に追記（重複チェックあり）"
Write-Host ""
if ($GIT_CLONE_URL -and $GIT_LOCAL_PATH) {
    Write-Host "  Gitクローン: 有効"
    Write-Host "    URL: $GIT_CLONE_URL"
    Write-Host "    ローカルパス: $GIT_LOCAL_PATH"
} else {
    Write-Host "  Gitクローン: スキップ（設定なし）"
}
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

Write-Host ""
Read-Host "続行するには Enter を押してください"
#endregion

#region 公開鍵の配置
Write-Header "公開鍵のLinuxサーバーへの配置"

if ($SSH_PASSWORD) {
    Write-Color "設定されたパスワードを使用してLinuxサーバーに接続します。" "Yellow"
} else {
    Write-Color "パスワード認証でLinuxサーバーに接続し、公開鍵を配置します。" "Yellow"
    Write-Color "パスワードの入力を求められたら、Linuxサーバーのパスワードを入力してください。" "Yellow"
}
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
$sshArgs += "-o"
$sshArgs += "StrictHostKeyChecking=accept-new"
if ($SSH_PORT -ne 22) {
    $sshArgs += "-p"
    $sshArgs += $SSH_PORT
}
$sshArgs += "${SSH_USER}@${SSH_HOST}"
$sshArgs += $sshScript

# パスワードが設定されている場合はSSH_ASKPASSを使用
if ($SSH_PASSWORD) {
    # 一時的なパスワード応答スクリプトを作成
    $askpassScript = "$env:TEMP\ssh_askpass_$([System.Guid]::NewGuid().ToString('N')).bat"
    "@echo off`necho $SSH_PASSWORD" | Set-Content -Path $askpassScript -Encoding ASCII

    # 環境変数を設定
    $env:SSH_ASKPASS = $askpassScript
    $env:SSH_ASKPASS_REQUIRE = "force"
    $env:DISPLAY = "dummy:0"

    try {
        & ssh.exe $sshArgs
    } finally {
        # 一時ファイルを削除
        Remove-Item -Path $askpassScript -Force -ErrorAction SilentlyContinue
        # 環境変数をクリア
        Remove-Item Env:\SSH_ASKPASS -ErrorAction SilentlyContinue
        Remove-Item Env:\SSH_ASKPASS_REQUIRE -ErrorAction SilentlyContinue
        Remove-Item Env:\DISPLAY -ErrorAction SilentlyContinue
    }
} else {
    & ssh.exe $sshArgs
}

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

Write-Host ""
Read-Host "続行するには Enter を押してください"
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

Write-Host ""
Read-Host "続行するには Enter を押してください"
#endregion

#region Gitリポジトリのクローン
if ($GIT_AVAILABLE -and $GIT_CLONE_URL -and $GIT_LOCAL_PATH) {
    Write-Header "Gitリポジトリのクローン"

    # 既にクローン済みか確認
    if (Test-Path "$GIT_LOCAL_PATH\.git") {
        Write-Color "[スキップ] リポジトリは既にクローン済みです: $GIT_LOCAL_PATH" "Yellow"
        Write-Host ""

        # リモート情報を表示
        Push-Location $GIT_LOCAL_PATH
        $remoteUrl = git remote get-url origin 2>$null
        Pop-Location

        if ($remoteUrl) {
            Write-Host "  リモートURL: $remoteUrl"
        }
    } else {
        # 親ディレクトリが存在するか確認
        $parentDir = Split-Path $GIT_LOCAL_PATH -Parent
        if ($parentDir -and -not (Test-Path $parentDir)) {
            New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
            Write-Color "[作成] ディレクトリを作成しました: $parentDir" "Green"
        }

        Write-Color "[実行] リポジトリをクローン中..." "Yellow"
        Write-Host "  URL: $GIT_CLONE_URL"
        Write-Host "  ローカルパス: $GIT_LOCAL_PATH"
        Write-Host ""

        # SSH鍵を指定してクローン
        $env:GIT_SSH_COMMAND = "ssh -i `"$keyPath`" -o StrictHostKeyChecking=accept-new"

        try {
            & git clone $GIT_CLONE_URL $GIT_LOCAL_PATH

            if ($LASTEXITCODE -eq 0) {
                Write-Host ""
                Write-Color "[成功] リポジトリをクローンしました" "Green"
                Write-Host "  パス: $GIT_LOCAL_PATH"
            } else {
                Write-Host ""
                Write-Color "[エラー] クローンに失敗しました" "Red"
                Write-Host "  - URLが正しいか確認してください"
                Write-Host "  - リポジトリへのアクセス権限があるか確認してください"
            }
        } finally {
            # 環境変数をクリア
            Remove-Item Env:\GIT_SSH_COMMAND -ErrorAction SilentlyContinue
        }
    }
    Write-Host ""
}
#endregion

#region WinMergeフィルター設定
if ($SETUP_WINMERGE_FILTER) {
    Write-Header "WinMergeフィルター設定"

    # WinMergeがインストールされているか確認
    $winmergePath = $null
    $winmergeLocations = @(
        "${env:ProgramFiles}\WinMerge\WinMergeU.exe",
        "${env:ProgramFiles(x86)}\WinMerge\WinMergeU.exe"
    )

    foreach ($path in $winmergeLocations) {
        if (Test-Path $path) {
            $winmergePath = $path
            break
        }
    }

    if (-not $winmergePath) {
        Write-Color "[スキップ] WinMergeがインストールされていません" "Yellow"
        Write-Host "  WinMergeをインストール後、再度実行してください"
        Write-Host ""
    } else {
        Write-Color "[OK] WinMergeが見つかりました: $winmergePath" "Green"
        Write-Host ""

        # フィルターファイルのパス
        $filterDir = "$env:APPDATA\WinMerge\Filters"
        $filterFile = "$filterDir\GitFiles.flt"

        # フィルターディレクトリがなければ作成
        if (-not (Test-Path $filterDir)) {
            New-Item -ItemType Directory -Path $filterDir -Force | Out-Null
            Write-Color "[作成] フィルターディレクトリを作成しました: $filterDir" "Green"
        }

        # 除外パターンを生成
        $excludePatterns = $WINMERGE_EXCLUDE_FILES | ForEach-Object {
            "f: \\$_`$"
        }

        # フィルターファイルの内容
        $filterContent = @"
## This is a directory/file filter for WinMerge
## This filter was auto-generated by Git開発環境初期設定ツール
name: Git管理ファイル除外フィルター
desc: .gitignore, .gitkeep, .gitattributes などのGit管理ファイルを除外します

## ファイルフィルター（比較から除外するファイル）
## f: は正規表現パターン。ファイル名の末尾にマッチさせるには `$` を使用
$($excludePatterns -join "`n")
"@

        # 既存のフィルターファイルがあるか確認
        if (Test-Path $filterFile) {
            Write-Color "[情報] 既存のフィルターファイルが見つかりました" "Yellow"
            Write-Host "  パス: $filterFile"
            Write-Host ""

            do {
                $overwriteFilter = Read-Host "上書きしますか？ (y/n)"
                $overwriteFilter = $overwriteFilter.ToLower()
            } while ($overwriteFilter -notin @("y", "n"))

            if ($overwriteFilter -eq "n") {
                Write-Color "[スキップ] 既存のフィルターを保持します" "Yellow"
            } else {
                $filterContent | Out-File -FilePath $filterFile -Encoding UTF8
                Write-Color "[更新] フィルターファイルを更新しました" "Green"
            }
        } else {
            $filterContent | Out-File -FilePath $filterFile -Encoding UTF8
            Write-Color "[作成] フィルターファイルを作成しました" "Green"
        }

        Write-Host ""
        Write-Host "  フィルターファイル: $filterFile"
        Write-Host ""
        Write-Color "除外対象ファイル:" "Cyan"
        foreach ($file in $WINMERGE_EXCLUDE_FILES) {
            Write-Host "  - $file"
        }
        Write-Host ""
        Write-Color "[完了] フィルターファイルを作成しました" "Green"
        Write-Host ""
        Write-Color "【フィルターの適用方法】" "Yellow"
        Write-Host "  WinMergeでフォルダ比較時に、以下の手順でフィルターを適用してください："
        Write-Host ""
        Write-Host "  1. WinMergeを起動し、フォルダ比較ダイアログを開く"
        Write-Host "  2. ダイアログ左下の「フォルダー: フィルター」欄の [...] ボタンをクリック"
        Write-Host "  3. 「Git管理ファイル除外フィルター」を選択して [OK]"
        Write-Host "  4. 比較を実行"
        Write-Host ""
        Write-Host "  ※ フィルターは毎回選択する必要があります（WinMergeの仕様）"
    }
    Write-Host ""
}
#endregion

Write-Host ""
Write-Color "[完了] 処理が終了しました" "Green"
