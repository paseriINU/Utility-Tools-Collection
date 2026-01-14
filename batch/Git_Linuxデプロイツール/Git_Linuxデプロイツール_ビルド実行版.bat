<# :
@echo off
chcp 65001 >nul
title Git Linuxデプロイツール（ビルド実行版）
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
# Git Linuxデプロイツール（ビルド実行版）
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
#   8. 【追加】ファイル転送後にビルドシェルを自動実行
#   9. 【追加】業務ID単位ビルド / フルコンパイル自動判定
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
# 環境名, 転送先パス, オーナー, ビルドシェルの環境選択番号 をハッシュテーブルで定義
$ENVIRONMENTS = @(
    @{
        Name = "tst1t"
        Path = "/path/to/tst1t/"
        Owner = "tzy_tst13"
        BuildEnvNumber = "1"      # ビルドシェルの環境選択番号
    },
    @{
        Name = "tst2t"
        Path = "/path/to/tst2t/"
        Owner = "tzy_tst23"
        BuildEnvNumber = "2"      # ビルドシェルの環境選択番号
    },
    @{
        Name = "tst3t"
        Path = "/path/to/tst3t/"
        Owner = "tzy_tst33"
        BuildEnvNumber = "3"      # ビルドシェルの環境選択番号
    }
)

# Linux側のパーミッション設定
$LINUX_CHMOD_DIR = "777"   # ディレクトリのパーミッション
$LINUX_CHMOD_FILE = "777"  # ファイルのパーミッション

# 一括転送の自動切り替えしきい値
# このファイル数以上で一括転送に自動切り替え（変更ファイルモード時）
# 目安: 個別転送は1ファイル約3秒、10ファイルで約30秒
$BULK_TRANSFER_THRESHOLD = 10

#------------------------------------------------------------------------------
# ビルドシェル設定
#------------------------------------------------------------------------------

# ビルドシェルスクリプトのパス
$BUILD_SCRIPT = "/opt/build/build.sh"

# ラッパースクリプト設定
# ローカルのラッパースクリプト（batファイルと同じフォルダに配置）
$BUILD_WRAPPER_LOCAL = "build_wrapper.sh"
# Linux側の一時配置先（終了後に自動削除）
$BUILD_WRAPPER_REMOTE = "/tmp/build_wrapper.sh"

# プロンプト文字列（ビルドシェルが表示するプロンプトに合わせて設定）
$BUILD_PROMPT_ENV = "環境を選択"       # 環境選択プロンプト
$BUILD_PROMPT_OPTION = "オプション"     # ビルドオプション選択プロンプト
$BUILD_PROMPT_GYOMU = "業務ID"          # 業務ID選択プロンプト
$BUILD_PROMPT_CONFIRM = "(y/n)"         # 確認プロンプト
$BUILD_CONFIRM_YES = "y"                # 確認プロンプトへの応答

# ビルドオプション選択番号（共通）
$BUILD_OPTION_NORMAL = "1"      # 業務ID単位ビルド時の追加選択番号
$BUILD_OPTION_FULL = "2"        # フルコンパイル時の追加選択番号
$BUILD_OPTION_EXIT = "99"       # ビルドシェル終了番号

# エラー検出文字列（ビルド出力にこの文字列が含まれていたらエラー）
$BUILD_ERROR_PATTERN = "エラー"

# 業務ID → ビルドシェルの選択番号 マッピング
# 転送ファイルの親フォルダ名（業務ID）に対応するビルドシェルの番号を定義
# 例: gyoumu/online/AAA1/AAA1001.c → 業務ID: AAA1
$GYOMU_BUILD_MAP = @{
    "AAA1" = "1"
    "BBB2" = "2"
    "CCC3" = "3"
    "DDD4" = "4"
    # 必要に応じて追加（業務IDは親フォルダ名をそのまま指定）
}

# ビルド対象フォルダのプレフィックス（これらのフォルダ配下のみ業務ID単位ビルド対象）
# 例: gyoumu/online/業務ID/xxx.c, gyoumu/remote/業務ID/xxx.c
$BUILD_TARGET_PREFIXES = @("gyoumu/online/", "gyoumu/remote/", "gyoumu/batch/")

# comフォルダ用のプレフィックス（特別ルール適用）
$BUILD_TARGET_PREFIX_COM = "gyoumu/com/"

# comフォルダ内で業務ID単位ビルド対象とする業務IDリスト
# これ以外の業務IDはフルコンパイル対象となる
$COM_ALLOWED_GYOMU_IDS = @("ZSG030", "ZSG060", "ZSG100", "ZSG920", "ZSG960", "ZSG970")

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
Write-Color "  Git Linuxデプロイツール（ビルド実行版）" "Cyan"
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

#region ブランチ確認・環境選択
Write-Color "================================================================" "Cyan"
Write-Color "転送先環境を選択してください" "Cyan"
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
Write-Host " 2. すべてのファイル" -NoNewline
Write-Host " ※転送先を事前にクリーンアップします" -ForegroundColor Yellow
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

#region 業務ID抽出とビルドモード判定
Write-Header "ビルドモード判定"

$gyomuIds = @{}
$hasNonGyomuFiles = $false

# 「すべてのファイル」モードの場合はフルコンパイル
if ($transferMode -eq "all") {
    $hasNonGyomuFiles = $true
    Write-Color "[判定] 「すべてのファイル」モード → フルコンパイル" "Yellow"
} else {
    # 転送ファイルから業務IDを抽出
    foreach ($file in $fileList) {
        $filePath = $file.Path.Replace("\", "/")
        $isGyomuFile = $false

        # 通常のプレフィックス（online, remote, batch）をチェック
        foreach ($prefix in $BUILD_TARGET_PREFIXES) {
            if ($filePath.StartsWith($prefix)) {
                # gyoumu/online/業務ID/xxx.c の形式
                $pathParts = $filePath.Substring($prefix.Length).Split("/")
                if ($pathParts.Count -ge 2) {
                    # 業務IDは親フォルダ名（.c/.hファイルの直上フォルダ）
                    $gyomuId = $pathParts[0]
                    if ($gyomuId.Length -gt 0) {
                        if (-not $gyomuIds.ContainsKey($gyomuId)) {
                            $gyomuIds[$gyomuId] = @()
                        }
                        $gyomuIds[$gyomuId] += $filePath
                        $isGyomuFile = $true
                    }
                }
                break
            }
        }

        # comフォルダの特別処理
        if (-not $isGyomuFile -and $filePath.StartsWith($BUILD_TARGET_PREFIX_COM)) {
            # gyoumu/com/業務ID/xxx.c の形式
            $pathParts = $filePath.Substring($BUILD_TARGET_PREFIX_COM.Length).Split("/")
            if ($pathParts.Count -ge 2) {
                $gyomuId = $pathParts[0]
                if ($gyomuId.Length -gt 0) {
                    # comフォルダ内で許可された業務IDかチェック
                    if ($COM_ALLOWED_GYOMU_IDS -contains $gyomuId) {
                        if (-not $gyomuIds.ContainsKey($gyomuId)) {
                            $gyomuIds[$gyomuId] = @()
                        }
                        $gyomuIds[$gyomuId] += $filePath
                        $isGyomuFile = $true
                    } else {
                        # 許可リストにない業務ID → フルコンパイル対象
                        Write-Color "[判定] comフォルダ内の非許可業務ID: $gyomuId → フルコンパイル対象" "Yellow"
                        $hasNonGyomuFiles = $true
                        $isGyomuFile = $true  # 後続の重複判定を防ぐ
                    }
                }
            }
        }

        if (-not $isGyomuFile) {
            $hasNonGyomuFiles = $true
            Write-Color "[判定] 業務ID対象外ファイル: $filePath" "Yellow"
        }
    }
}

# ビルドモード決定
$buildMode = ""
$buildInputs = @()

if ($hasNonGyomuFiles -or $gyomuIds.Count -eq 0) {
    # フルコンパイルモード
    $buildMode = "full"
    Write-Host ""
    Write-Color "[決定] ビルドモード: フルコンパイル" "Magenta"
    Write-Host ""
} else {
    # 業務ID単位ビルドモード
    $buildMode = "gyomu"
    Write-Host ""
    Write-Color "[決定] ビルドモード: 業務ID単位ビルド" "Magenta"
    Write-Host ""
    Write-Host "  対象業務ID:"
    foreach ($gyomuId in $gyomuIds.Keys | Sort-Object) {
        $buildNumber = $GYOMU_BUILD_MAP[$gyomuId]
        if ($buildNumber) {
            Write-Host "    - $gyomuId → ビルド番号: $buildNumber"
        } else {
            Write-Color "    - $gyomuId → [警告] マッピングなし（フルコンパイルに変更）" "Yellow"
            $buildMode = "full"
        }
    }
    Write-Host ""
}

# マッピングがない業務IDがあればフルコンパイルに変更
if ($buildMode -eq "gyomu") {
    foreach ($gyomuId in $gyomuIds.Keys) {
        if (-not $GYOMU_BUILD_MAP.ContainsKey($gyomuId)) {
            $buildMode = "full"
            Write-Color "[変更] マッピングがない業務IDがあるため、フルコンパイルに変更" "Yellow"
            break
        }
    }
}
#endregion

#region 転送設定
# 転送するファイルリスト
$filesToTransfer = $fileList

# 転送モードの自動判断
Write-Color "================================================================" "Cyan"
Write-Color "転送モードの決定" "Cyan"
Write-Color "================================================================" "Cyan"
Write-Host ""

if ($transferMode -eq "all") {
    # 「すべてのファイル」モード → 常に一括転送
    $useBulkTransfer = $true
    Write-Color "[自動選択] 一括転送モード（すべてのファイル転送のため）" "Green"
} else {
    # 「変更されたファイルのみ」モード → ファイル数で自動判断
    $fileCount = $filesToTransfer.Count
    $estimatedTimeIndividual = $fileCount * 3  # 個別転送: 1ファイル約3秒
    $estimatedTimeBulk = 8 + [math]::Ceiling($fileCount / 10)  # 一括転送: 固定8秒 + α

    if ($fileCount -ge $BULK_TRANSFER_THRESHOLD) {
        # しきい値以上 → 一括転送
        $useBulkTransfer = $true
        Write-Color "[自動選択] 一括転送モード" "Green"
        Write-Host ""
        Write-Host "  ファイル数: $fileCount 個（しきい値: $BULK_TRANSFER_THRESHOLD 個以上）"
        Write-Host "  推定時間: 約${estimatedTimeBulk}秒（個別転送の場合: 約${estimatedTimeIndividual}秒）"
    } else {
        # しきい値未満 → 個別転送（進捗表示あり）
        $useBulkTransfer = $false
        Write-Color "[自動選択] 個別転送モード（進捗表示あり）" "Green"
        Write-Host ""
        Write-Host "  ファイル数: $fileCount 個（しきい値: $BULK_TRANSFER_THRESHOLD 個未満）"
        Write-Host "  推定時間: 約${estimatedTimeIndividual}秒"
    }
}

# 「すべてのファイル」モードの場合、クリーンアップを自動実行
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

    # クリーンアップ固定
    $doCleanup = $true
    $cleanupDirs = $parentDirs.Keys

    Write-Host ""
    Write-Color "[自動設定] 転送前にクリーンアップを実行します" "Yellow"
    Write-Host ""
    Write-Host "  削除対象フォルダ:"
    foreach ($dir in $parentDirs.Keys | Sort-Object) {
        Write-Host "    - ${dir}/*"
    }
    Write-Host ""
    Write-Color "  ※ ファイルのみ再帰的に削除（ディレクトリ構造は残ります）" "Gray"
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

if ($useBulkTransfer -eq $true) {
    # 一括転送モード（tar圧縮）
    Write-Color "[実行] 一括転送モードで転送します..." "Yellow"
    Write-Host ""

    # 一時ディレクトリを作成
    $tempDir = Join-Path $env:TEMP "git_deploy_$(Get-Date -Format 'yyyyMMddHHmmss')"
    $tarFileName = "deploy_$(Get-Date -Format 'yyyyMMddHHmmss').tar"
    $tarFilePath = Join-Path $env:TEMP $tarFileName

    try {
        # 一時ディレクトリにファイルをコピー（ディレクトリ構造を維持）
        Write-Color "[準備] ファイルを収集中..." "Yellow"
        foreach ($file in $filesToTransfer) {
            $localPath = Join-Path $GIT_ROOT $file.Path
            $destPath = Join-Path $tempDir $file.Path
            $destDir = Split-Path $destPath -Parent

            if (-not (Test-Path $destDir)) {
                New-Item -ItemType Directory -Path $destDir -Force | Out-Null
            }
            Copy-Item -Path $localPath -Destination $destPath -Force
        }

        # tarアーカイブを作成
        Write-Color "[圧縮] tarアーカイブを作成中..." "Yellow"
        Push-Location $tempDir
        $tarResult = & tar -cvf $tarFilePath * 2>&1
        Pop-Location

        if ($LASTEXITCODE -ne 0) {
            Write-Color "[エラー] tarアーカイブの作成に失敗しました" "Red"
            throw "tar creation failed"
        }

        $tarSize = (Get-Item $tarFilePath).Length / 1KB
        Write-Color "[情報] アーカイブサイズ: $([math]::Round($tarSize, 2)) KB" "Cyan"

        # tarファイルをリモートに転送
        Write-Color "[転送] アーカイブを転送中..." "Yellow"
        $scpArgs = @()
        if ($SSH_KEY -ne "" -and (Test-Path $SSH_KEY)) {
            $scpArgs += "-i"
            $scpArgs += $SSH_KEY
        }
        if ($SSH_PORT -ne 22) {
            $scpArgs += "-P"
            $scpArgs += $SSH_PORT
        }
        $scpArgs += $tarFilePath
        $scpArgs += "${SSH_USER}@${SSH_HOST}:/tmp/${tarFileName}"

        & $scpCommand $scpArgs 2>&1 | Out-Null

        if ($LASTEXITCODE -ne 0) {
            Write-Color "[エラー] アーカイブの転送に失敗しました" "Red"
            throw "scp failed"
        }

        # リモートで展開
        Write-Color "[展開] リモートでアーカイブを展開中..." "Yellow"
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

        # 展開先ディレクトリを作成してtar展開
        $extractCmd = "cd '${REMOTE_DIR}' && tar -xvf /tmp/${tarFileName} && rm -f /tmp/${tarFileName}"
        $sshArgs += $extractCmd

        & $sshCommand $sshArgs 2>&1 | Out-Null

        if ($LASTEXITCODE -ne 0) {
            Write-Color "[エラー] アーカイブの展開に失敗しました" "Red"
            throw "tar extract failed"
        }

        # パーミッションと所有者を一括設定
        Write-Color "[設定] パーミッションと所有者を設定中..." "Yellow"
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
        $sshArgs += "find '${REMOTE_DIR}' -type f -exec chmod $LINUX_CHMOD_FILE {} \; -exec chown ${OWNER}:${COMMON_GROUP} {} \; && find '${REMOTE_DIR}' -type d -exec chmod $LINUX_CHMOD_DIR {} \; -exec chown ${OWNER}:${COMMON_GROUP} {} \;"

        & $sshCommand $sshArgs 2>&1 | Out-Null

        $successCount = $filesToTransfer.Count
        Write-Host ""
        Write-Color "[OK] $successCount 個のファイルを一括転送しました" "Green"

    } catch {
        Write-Color "[エラー] 一括転送に失敗しました: $($_.Exception.Message)" "Red"
        $failCount = $filesToTransfer.Count
    } finally {
        # 一時ファイルをクリーンアップ
        if (Test-Path $tempDir) {
            Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        }
        if (Test-Path $tarFilePath) {
            Remove-Item -Path $tarFilePath -Force -ErrorAction SilentlyContinue
        }
    }

} else {
    # 個別転送モード（従来方式）
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
}

Write-Host ""
#endregion

#region 転送結果サマリー
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

if ($failCount -gt 0) {
    Write-Color "一部のファイル転送に失敗しました" "Yellow"
    Write-Host ""
    $continueChoice = Read-Host "ビルドを続行しますか？ (y/n)"
    if ($continueChoice -ne "y") {
        Write-Color "[キャンセル] ビルドを中止しました" "Yellow"
        exit 1
    }
}
#endregion

#region ラッパースクリプト転送
Write-Header "ラッパースクリプト転送"

# ローカルのラッパースクリプトパス
$wrapperLocalPath = Join-Path $scriptDir $BUILD_WRAPPER_LOCAL

if (-not (Test-Path $wrapperLocalPath)) {
    Write-Color "[エラー] ラッパースクリプトが見つかりません: $wrapperLocalPath" "Red"
    Write-Color "        build_wrapper.sh をバッチファイルと同じフォルダに配置してください" "Yellow"
    exit 1
}

Write-Color "[転送] ラッパースクリプトを転送中..." "Yellow"
Write-Host "  ローカル: $wrapperLocalPath"
Write-Host "  リモート: ${SSH_USER}@${SSH_HOST}:${BUILD_WRAPPER_REMOTE}"

$scpArgs = @()
if ($SSH_KEY -ne "" -and (Test-Path $SSH_KEY)) {
    $scpArgs += "-i"
    $scpArgs += $SSH_KEY
}
if ($SSH_PORT -ne 22) {
    $scpArgs += "-P"
    $scpArgs += $SSH_PORT
}
$scpArgs += $wrapperLocalPath
$scpArgs += "${SSH_USER}@${SSH_HOST}:${BUILD_WRAPPER_REMOTE}"

try {
    & $scpCommand $scpArgs 2>&1 | Out-Null

    if ($LASTEXITCODE -eq 0) {
        # 実行権限を付与
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
        $sshArgs += "chmod +x '$BUILD_WRAPPER_REMOTE'"

        & $sshCommand $sshArgs 2>&1 | Out-Null

        Write-Color "[OK] ラッパースクリプトの転送完了" "Green"
    } else {
        Write-Color "[エラー] ラッパースクリプトの転送に失敗しました" "Red"
        exit 1
    }
} catch {
    Write-Color "[エラー] ラッパースクリプトの転送に失敗しました: $($_.Exception.Message)" "Red"
    exit 1
}

Write-Host ""
#endregion

#region ビルド実行
Write-Header "ビルド実行"

Write-Color "[情報] ビルドシェル: $BUILD_SCRIPT" "Cyan"
Write-Color "[情報] ラッパー: $BUILD_WRAPPER_REMOTE" "Cyan"
Write-Color "[情報] ビルドモード: $(if ($buildMode -eq 'full') { 'フルコンパイル' } else { '業務ID単位ビルド' })" "Cyan"
Write-Host ""

# SSH引数の共通部分を構築
$sshBaseArgs = @()
if ($SSH_KEY -ne "" -and (Test-Path $SSH_KEY)) {
    $sshBaseArgs += "-i"
    $sshBaseArgs += $SSH_KEY
}
if ($SSH_PORT -ne 22) {
    $sshBaseArgs += "-p"
    $sshBaseArgs += $SSH_PORT
}

if ($buildMode -eq "full") {
    # フルコンパイルモード
    Write-Color "[実行] フルコンパイルを実行します..." "Yellow"
    Write-Host ""

    # 入力値を構築: 環境選択番号 + フルコンパイル選択番号
    $buildEnvNumber = $selectedEnv.BuildEnvNumber

    Write-Host "  環境選択番号: $buildEnvNumber"
    Write-Host "  追加選択番号: $BUILD_OPTION_FULL (フルコンパイル)"
    Write-Host ""

    # ラッパースクリプト経由でビルド実行（プロンプト待機モード: -w）
    $wrapperArgs = "-w '$BUILD_SCRIPT' '${BUILD_PROMPT_ENV}:${buildEnvNumber}' '${BUILD_PROMPT_OPTION}:${BUILD_OPTION_FULL}'"

    $sshArgs = $sshBaseArgs + @("${SSH_USER}@${SSH_HOST}")
    $sshArgs += "$BUILD_WRAPPER_REMOTE $wrapperArgs"

    Write-Color "[コマンド] $BUILD_WRAPPER_REMOTE $wrapperArgs" "Gray"
    Write-Host ""

    try {
        & $sshCommand $sshArgs 2>&1 | ForEach-Object { Write-Host $_ }

        if ($LASTEXITCODE -eq 0) {
            Write-Host ""
            Write-Color "[OK] フルコンパイルが完了しました" "Green"
        } else {
            Write-Host ""
            Write-Color "[警告] ビルドが終了しました (終了コード: $LASTEXITCODE)" "Yellow"
        }
    } catch {
        Write-Color "[エラー] ビルド実行に失敗しました: $($_.Exception.Message)" "Red"
    }

} else {
    # 業務ID単位ビルドモード（マルチビルドモード: -m）
    $buildEnvNumber = $selectedEnv.BuildEnvNumber

    # 業務IDをビルド番号に変換してカンマ区切りで結合
    $gyomuBuildNumbers = @()
    foreach ($gyomuId in $gyomuIds.Keys | Sort-Object) {
        $buildNumber = $GYOMU_BUILD_MAP[$gyomuId]
        $gyomuBuildNumbers += $buildNumber
        Write-Host "  業務ID: $gyomuId → ビルド番号: $buildNumber"
    }
    $gyomuBuildNumbersStr = $gyomuBuildNumbers -join ","

    Write-Host ""
    Write-Host "  環境選択番号: $buildEnvNumber"
    Write-Host "  追加選択番号: $BUILD_OPTION_NORMAL (業務ID単位)"
    Write-Host "  業務ID選択番号: $gyomuBuildNumbersStr"
    Write-Host "  終了番号: $BUILD_OPTION_EXIT"
    Write-Host ""

    # ラッパースクリプト経由でビルド実行（マルチビルドモード: -m）
    # 引数: env_prompt:env option_prompt:opt gyomu_prompt:id1,id2 confirm_prompt:y option_prompt:exit [error_pattern]
    $wrapperArgs = "-m '$BUILD_SCRIPT' '${BUILD_PROMPT_ENV}:${buildEnvNumber}' '${BUILD_PROMPT_OPTION}:${BUILD_OPTION_NORMAL}' '${BUILD_PROMPT_GYOMU}:${gyomuBuildNumbersStr}' '${BUILD_PROMPT_CONFIRM}:${BUILD_CONFIRM_YES}' '${BUILD_PROMPT_OPTION}:${BUILD_OPTION_EXIT}' '${BUILD_ERROR_PATTERN}'"

    $sshArgs = $sshBaseArgs + @("${SSH_USER}@${SSH_HOST}")
    $sshArgs += "$BUILD_WRAPPER_REMOTE $wrapperArgs"

    Write-Color "[コマンド] $BUILD_WRAPPER_REMOTE $wrapperArgs" "Gray"
    Write-Host ""

    try {
        & $sshCommand $sshArgs 2>&1 | ForEach-Object { Write-Host $_ }

        if ($LASTEXITCODE -eq 0) {
            Write-Host ""
            Write-Color "[OK] すべての業務IDのビルドが完了しました" "Green"
        } else {
            Write-Host ""
            Write-Color "[警告] ビルドが終了しました (終了コード: $LASTEXITCODE)" "Yellow"
        }
    } catch {
        Write-Color "[エラー] ビルド実行に失敗しました: $($_.Exception.Message)" "Red"
    }
}
#endregion

#region ラッパースクリプト削除
Write-Header "クリーンアップ"

Write-Color "[削除] ラッパースクリプトを削除中..." "Yellow"

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
$sshArgs += "rm -f '$BUILD_WRAPPER_REMOTE'"

try {
    & $sshCommand $sshArgs 2>&1 | Out-Null
    Write-Color "[OK] ラッパースクリプトを削除しました" "Green"
} catch {
    Write-Color "[警告] ラッパースクリプトの削除に失敗しました: $($_.Exception.Message)" "Yellow"
}

Write-Host ""
#endregion

#region 最終結果
Write-Header "処理完了"

Write-Color "ファイル転送: $successCount 件成功" "Green"
if ($failCount -gt 0) {
    Write-Color "ファイル転送: $failCount 件失敗" "Red"
}
Write-Color "ビルド実行: 完了" "Green"

Write-Host ""
Write-Color "すべての処理が完了しました！" "Green"

exit 0
#endregion
