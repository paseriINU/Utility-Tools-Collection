<# :
@echo off
chcp 65001 >nul
title TFS to Git 同期ツール
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
# TFS to Git Sync Script (PowerShell)
# TFSのファイルをGitリポジトリに同期します
# =============================================================================

# タイトル表示
Write-Host ""
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host "  TFS to Git 同期ツール" -ForegroundColor Cyan
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

# Git日本語表示設定
git config --global core.quotepath false 2>&1 | Out-Null

#region 設定セクション
# TFSとGitのディレクトリパスを設定（固定値）
# 重要: TFS_DIRとGIT_REPO_DIRは同じフォルダ構造を指すようにしてください
# 例: TFSの "C:\TFS\project\src" とGitの "C:\Git\project\src" を比較する場合
$TFS_DIR = "C:\Users\$env:USERNAME\source"
$GIT_REPO_DIR = "C:\Users\$env:USERNAME\source\Git\project"

# TFS最新取得を実行するか
$UPDATE_TFS_BEFORE_COMPARE = $true

# tf.exeのパス（空の場合はPATHから検索）
# PATHが通っていない場合は、以下のようにフルパスを指定してください
# 例: Visual Studio 2022 Enterprise
#   $TF_EXE_PATH = "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer\tf.exe"
# 例: Visual Studio 2022 Professional
#   $TF_EXE_PATH = "C:\Program Files\Microsoft Visual Studio\2022\Professional\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer\tf.exe"
# 例: Visual Studio 2019
#   $TF_EXE_PATH = "C:\Program Files (x86)\Microsoft Visual Studio\2019\Enterprise\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer\tf.exe"
$TF_EXE_PATH = ""

# 比較対象から除外するファイル（Git固有ファイルなど）
$EXCLUDE_FILES = @(
    ".gitignore",
    ".gitattributes"
)

# 空フォルダに作成する.gitignoreの内容
$GITIGNORE_CONTENT = @"
# このファイルは空フォルダをGitで管理するために自動生成されました
# This file was auto-generated to keep this empty folder in Git

# このフォルダ内のすべてのファイルを無視（.gitignore自身を除く）
*
!.gitignore
"@
#endregion

#region パスの存在確認
if (-not (Test-Path $TFS_DIR)) {
    Write-Host "[エラー] TFSディレクトリが見つかりません: $TFS_DIR" -ForegroundColor Red
    exit 1
}

if (-not (Test-Path $GIT_REPO_DIR)) {
    Write-Host "[エラー] Gitディレクトリが見つかりません: $GIT_REPO_DIR" -ForegroundColor Red
    exit 1
}

# Gitリポジトリ確認（親ディレクトリを遡って.gitフォルダを探す）
Push-Location $GIT_REPO_DIR
$isGitRepo = $false
try {
    # git rev-parseコマンドでGitリポジトリかどうかを確認
    git rev-parse --git-dir 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) {
        $isGitRepo = $true
    }
} catch {
    $isGitRepo = $false
}
Pop-Location

if (-not $isGitRepo) {
    Write-Host "[エラー] 指定されたディレクトリはGit管理下にありません: $GIT_REPO_DIR" -ForegroundColor Red
    exit 1
}
#endregion

Write-Host ""
Write-Host "TFSディレクトリ: $TFS_DIR" -ForegroundColor White
Write-Host "Gitディレクトリ: $GIT_REPO_DIR" -ForegroundColor White
Write-Host ""

#region TFS最新取得
if ($UPDATE_TFS_BEFORE_COMPARE) {
    # tf.exeのパスを決定
    $tfExePath = $null

    # 設定で直接パスが指定されている場合
    if (-not [string]::IsNullOrWhiteSpace($TF_EXE_PATH)) {
        if (Test-Path $TF_EXE_PATH) {
            $tfExePath = $TF_EXE_PATH
            Write-Host "[情報] 設定で指定されたtf.exeを使用" -ForegroundColor Gray
        } else {
            Write-Host "[エラー] 指定されたtf.exeが見つかりません: $TF_EXE_PATH" -ForegroundColor Red
            exit 1
        }
    } else {
        # 1. PATHから検索
        $tfCommand = Get-Command tf -ErrorAction SilentlyContinue
        if ($tfCommand) {
            $tfExePath = $tfCommand.Source
            Write-Host "[情報] PATHからtf.exeを検出" -ForegroundColor Gray
        } else {
            # 2. Visual Studioインストールパスから自動検索
            Write-Host "[情報] tf.exeを自動検索中..." -ForegroundColor Gray

            # 検索対象のパス一覧（新しいバージョン順）
            $searchPaths = @(
                # Visual Studio 2022
                "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer\tf.exe",
                "C:\Program Files\Microsoft Visual Studio\2022\Professional\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer\tf.exe",
                "C:\Program Files\Microsoft Visual Studio\2022\Community\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer\tf.exe",
                # Visual Studio 2019
                "C:\Program Files (x86)\Microsoft Visual Studio\2019\Enterprise\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer\tf.exe",
                "C:\Program Files (x86)\Microsoft Visual Studio\2019\Professional\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer\tf.exe",
                "C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer\tf.exe",
                # Visual Studio 2017
                "C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer\tf.exe",
                "C:\Program Files (x86)\Microsoft Visual Studio\2017\Professional\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer\tf.exe",
                "C:\Program Files (x86)\Microsoft Visual Studio\2017\Community\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer\tf.exe"
            )

            foreach ($path in $searchPaths) {
                if (Test-Path $path) {
                    $tfExePath = $path
                    Write-Host "[情報] tf.exeを自動検出: $tfExePath" -ForegroundColor Green
                    break
                }
            }
        }
    }

    if ($tfExePath) {
        Write-Host "------------------------------------------------------------------------" -ForegroundColor Cyan
        Write-Host " TFSワークスペースを最新に更新しています..." -ForegroundColor Cyan
        Write-Host "------------------------------------------------------------------------" -ForegroundColor Cyan
        Write-Host ""

        Push-Location $TFS_DIR
        try {
            # TFS最新取得（再帰的に全ファイル）
            $tfResult = & $tfExePath get /recursive /noprompt 2>&1
            if ($LASTEXITCODE -eq 0) {
                Write-Host "[完了] TFSワークスペースを最新に更新しました" -ForegroundColor Green
            } else {
                Write-Host "[警告] TFS更新で問題が発生しました（続行します）" -ForegroundColor Yellow
                Write-Host $tfResult -ForegroundColor Gray
            }
        } catch {
            Write-Host "[警告] TFS更新でエラーが発生しました: $_" -ForegroundColor Yellow
        }
        Pop-Location
        Write-Host ""
    } else {
        Write-Host ""
        Write-Host "========================================================================" -ForegroundColor Red
        Write-Host " [エラー] tfコマンドが見つかりません" -ForegroundColor Red
        Write-Host "========================================================================" -ForegroundColor Red
        Write-Host ""
        Write-Host "TFS最新取得が有効ですが、tf.exeが見つかりません。" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "解決方法:" -ForegroundColor Cyan
        Write-Host "  1. Visual Studio開発者コマンドプロンプトから実行する" -ForegroundColor White
        Write-Host "  2. バッチファイル内の `$TF_EXE_PATH` にtf.exeのパスを直接指定する" -ForegroundColor White
        Write-Host ""
        Write-Host "tf.exeの場所（例）:" -ForegroundColor Cyan
        Write-Host "  VS2022 Enterprise:" -ForegroundColor Gray
        Write-Host "    C:\Program Files\Microsoft Visual Studio\2022\Enterprise\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer\tf.exe" -ForegroundColor Gray
        Write-Host "  VS2022 Professional:" -ForegroundColor Gray
        Write-Host "    C:\Program Files\Microsoft Visual Studio\2022\Professional\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer\tf.exe" -ForegroundColor Gray
        Write-Host "  VS2019:" -ForegroundColor Gray
        Write-Host "    C:\Program Files (x86)\Microsoft Visual Studio\2019\Enterprise\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer\tf.exe" -ForegroundColor Gray
        Write-Host ""
        Write-Host "TFS更新をスキップして続行しますか？" -ForegroundColor Yellow
        Write-Host " 1. はい、TFS更新なしで続行する（ローカルの状態で比較）"
        Write-Host ""
        Write-Host " 0. いいえ、終了する"
        Write-Host ""
        $skipChoice = Read-Host "選択 (0-1)"

        if ($skipChoice -ne "1") {
            Write-Host ""
            Write-Host "処理を終了します。" -ForegroundColor Yellow
            exit 1
        }
        Write-Host ""
        Write-Host "[情報] TFS更新をスキップして続行します。" -ForegroundColor Yellow
        Write-Host ""
    }
}
#endregion

# Gitディレクトリに移動
Set-Location $GIT_REPO_DIR

#region ブランチ操作
Write-Host "------------------------------------------------------------------------" -ForegroundColor Yellow
Write-Host " 現在のGitブランチ:" -ForegroundColor Yellow
Write-Host "------------------------------------------------------------------------" -ForegroundColor Yellow
git branch
Write-Host ""

# ブランチ操作メニュー
:branchLoop while ($true) {
    Write-Host "ブランチ操作を選択してください:" -ForegroundColor Cyan
    Write-Host " 1. このまま続行"
    Write-Host " 2. ブランチを切り替える"
    Write-Host ""
    Write-Host " 0. 終了"
    Write-Host ""
    $branchChoice = Read-Host "選択 (0-2)"

    switch ($branchChoice) {
        "0" {
            # 終了
            exit 0
        }
        "1" {
            # 同期処理へ
            break branchLoop
        }
        "2" {
            Write-Host ""
            Write-Host "------------------------------------------------------------------------" -ForegroundColor Yellow
            Write-Host " 利用可能なブランチ:" -ForegroundColor Yellow
            Write-Host "------------------------------------------------------------------------" -ForegroundColor Yellow

            # ローカルブランチ一覧を取得（@()で配列を強制）
            $branches = @(git branch --format="%(refname:short)" | ForEach-Object { $_.Trim() })

            if ($branches.Count -eq 0) {
                Write-Host "[エラー] ブランチが見つかりません" -ForegroundColor Red
                Write-Host ""
                continue
            }

            # ブランチを番号付きで表示
            for ($i = 0; $i -lt $branches.Count; $i++) {
                $displayNum = $i + 1
                $branch = $branches[$i]
                Write-Host " $displayNum. $branch"
            }
            Write-Host ""

            # ブランチ選択
            $maxNum = $branches.Count
            $selection = Read-Host "ブランチ番号を入力してください (1-$maxNum)"

            # 入力検証
            if ($selection -match '^\d+$' -and [int]$selection -ge 1 -and [int]$selection -le $maxNum) {
                $selectedBranch = $branches[[int]$selection - 1]

                git checkout $selectedBranch
                if ($LASTEXITCODE -ne 0) {
                    Write-Host "[エラー] ブランチの切り替えに失敗しました" -ForegroundColor Red
                    Write-Host ""
                    continue
                }
                Write-Host "ブランチを切り替えました: $selectedBranch" -ForegroundColor Green
                Write-Host ""
            } else {
                Write-Host "[エラー] 無効な番号です" -ForegroundColor Red
                Write-Host ""
            }
        }
        default {
            Write-Host "無効な選択です。" -ForegroundColor Red
        }
    }
}
#endregion

#region 同期処理
Write-Host ""
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host " 差分チェックを開始します" -ForegroundColor Cyan
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "ファイルをスキャン中..." -ForegroundColor Cyan
Write-Host ""

# TFSとGitのファイル・フォルダ一覧を取得
Write-Host "TFSディレクトリをスキャン中: $TFS_DIR" -ForegroundColor Gray
$tfsFiles = Get-ChildItem -Path $TFS_DIR -Recurse -File -ErrorAction SilentlyContinue
$tfsFolders = Get-ChildItem -Path $TFS_DIR -Recurse -Directory -ErrorAction SilentlyContinue

Write-Host "Gitディレクトリをスキャン中: $GIT_REPO_DIR" -ForegroundColor Gray
$gitFiles = Get-ChildItem -Path $GIT_REPO_DIR -Recurse -File -ErrorAction SilentlyContinue | Where-Object {
    $_.FullName -notlike '*\.git\*' -and $_.Name -notin $EXCLUDE_FILES
}
$gitFolders = Get-ChildItem -Path $GIT_REPO_DIR -Recurse -Directory -ErrorAction SilentlyContinue | Where-Object {
    # .gitフォルダ自体とその配下のみを除外（.githubなどは除外しない）
    $_.FullName -notlike '*\.git\*' -and $_.Name -ne '.git'
}

# パスの正規化関数（末尾のバックスラッシュを統一）
function Normalize-Path($path) {
    return $path.TrimEnd('\', '/')
}

$TFS_DIR_NORMALIZED = Normalize-Path $TFS_DIR
$GIT_REPO_DIR_NORMALIZED = Normalize-Path $GIT_REPO_DIR

# ファイルを相対パスでハッシュテーブルに格納
$tfsFileDict = @{}
foreach ($file in $tfsFiles) {
    $relativePath = $file.FullName.Substring($TFS_DIR_NORMALIZED.Length).TrimStart('\', '/')
    $tfsFileDict[$relativePath] = $file
}

$gitFileDict = @{}
foreach ($file in $gitFiles) {
    $relativePath = $file.FullName.Substring($GIT_REPO_DIR_NORMALIZED.Length).TrimStart('\', '/')
    $gitFileDict[$relativePath] = $file
}

# フォルダを相対パスでハッシュテーブルに格納
$tfsFolderDict = @{}
foreach ($folder in $tfsFolders) {
    $relativePath = $folder.FullName.Substring($TFS_DIR_NORMALIZED.Length).TrimStart('\', '/')
    $tfsFolderDict[$relativePath] = $folder
}

$gitFolderDict = @{}
foreach ($folder in $gitFolders) {
    $relativePath = $folder.FullName.Substring($GIT_REPO_DIR_NORMALIZED.Length).TrimStart('\', '/')
    $gitFolderDict[$relativePath] = $folder
}

Write-Host "TFS: $($tfsFileDict.Count) ファイル, $($tfsFolderDict.Count) フォルダ" -ForegroundColor Gray
Write-Host "Git: $($gitFileDict.Count) ファイル, $($gitFolderDict.Count) フォルダ" -ForegroundColor Gray
Write-Host ""

# 差分を格納する配列
$newFiles = @()       # TFSにあってGitにない（新規追加）
$updateFiles = @()    # 両方にあるが内容が異なる（更新）
$deleteFiles = @()    # GitにあってTFSにない（削除対象）
$newFolders = @()     # TFSにあってGitにない空フォルダ
$deleteFolders = @()  # GitにあってTFSにないフォルダ
$identicalCount = 0

# TFSファイルをチェック（更新 & 新規追加）
foreach ($relativePath in $tfsFileDict.Keys) {
    $tfsFile = $tfsFileDict[$relativePath]
    $gitFilePath = Join-Path $GIT_REPO_DIR $relativePath

    if (Test-Path $gitFilePath) {
        # ファイルが両方に存在 → MD5ハッシュで比較
        try {
            $tfsHash = (Get-FileHash -Path $tfsFile.FullName -Algorithm MD5).Hash
            $gitHash = (Get-FileHash -Path $gitFilePath -Algorithm MD5).Hash

            if ($tfsHash -ne $gitHash) {
                # ハッシュが異なる → 更新対象
                $updateFiles += [PSCustomObject]@{
                    RelativePath = $relativePath
                    TfsFile = $tfsFile
                    GitFilePath = $gitFilePath
                }
            } else {
                # ハッシュが同じ → 変更なし
                $identicalCount++
            }
        } catch {
            Write-Warning "ファイルハッシュ取得エラー: $relativePath - $_"
        }
    } else {
        # Gitに存在しない → 新規追加対象
        $newFiles += [PSCustomObject]@{
            RelativePath = $relativePath
            TfsFile = $tfsFile
            GitFilePath = $gitFilePath
        }
    }
}

# Gitのみのファイルをチェック（削除対象）
foreach ($relativePath in $gitFileDict.Keys) {
    if (-not $tfsFileDict.ContainsKey($relativePath)) {
        $gitFile = $gitFileDict[$relativePath]
        $deleteFiles += [PSCustomObject]@{
            RelativePath = $relativePath
            GitFile = $gitFile
        }
    }
}

# TFSのみのフォルダをチェック（新規フォルダ - 空フォルダ対応）
foreach ($relativePath in $tfsFolderDict.Keys) {
    $gitFolderPath = Join-Path $GIT_REPO_DIR $relativePath
    if (-not (Test-Path $gitFolderPath)) {
        # Gitに存在しないフォルダ
        # そのフォルダ配下にファイルがあるかチェック（パス区切り文字を正規化して比較）
        $normalizedPath = $relativePath.Replace('/', '\').ToLower()
        $hasFiles = $false
        foreach ($filePath in $tfsFileDict.Keys) {
            $normalizedFilePath = $filePath.Replace('/', '\').ToLower()
            if ($normalizedFilePath.StartsWith("$normalizedPath\")) {
                $hasFiles = $true
                break
            }
        }
        if (-not $hasFiles) {
            # 空フォルダとして追加
            $newFolders += [PSCustomObject]@{
                RelativePath = $relativePath
                TfsFolder = $tfsFolderDict[$relativePath]
                GitFolderPath = $gitFolderPath
            }
        }
    }
}

# Gitのみのフォルダをチェック（削除対象フォルダ）
# デバッグ: Gitフォルダ数を表示
Write-Host "[DEBUG] Gitフォルダ数: $($gitFolderDict.Count)" -ForegroundColor Magenta
Write-Host "[DEBUG] TFSフォルダ数: $($tfsFolderDict.Count)" -ForegroundColor Magenta

foreach ($relativePath in $gitFolderDict.Keys) {
    # TFSにフォルダが存在するかチェック（大文字小文字を区別しない）
    $normalizedGitPath = $relativePath.Replace('/', '\').ToLower()
    $folderExistsInTfs = $false
    foreach ($tfsPath in $tfsFolderDict.Keys) {
        $normalizedTfsPath = $tfsPath.Replace('/', '\').ToLower()
        if ($normalizedGitPath -eq $normalizedTfsPath) {
            $folderExistsInTfs = $true
            break
        }
    }

    Write-Host "[DEBUG] チェック中: $relativePath -> TFSにフォルダ存在: $folderExistsInTfs" -ForegroundColor Magenta

    if (-not $folderExistsInTfs) {
        # TFSにフォルダが存在しない場合、ファイルとしても存在しないか確認
        $existsInTfs = $false
        $searchPattern = "$normalizedGitPath\"
        foreach ($filePath in $tfsFileDict.Keys) {
            $normalizedFilePath = $filePath.Replace('/', '\').ToLower()
            if ($normalizedFilePath.StartsWith($searchPattern)) {
                $existsInTfs = $true
                Write-Host "[DEBUG]   TFSファイル存在: $filePath" -ForegroundColor Magenta
                break
            }
        }
        Write-Host "[DEBUG]   TFSにファイル存在: $existsInTfs" -ForegroundColor Magenta
        if (-not $existsInTfs) {
            $gitFolder = $gitFolderDict[$relativePath]
            $deleteFolders += [PSCustomObject]@{
                RelativePath = $relativePath
                GitFolder = $gitFolder
            }
            Write-Host "[DEBUG]   -> 削除対象に追加" -ForegroundColor Yellow
        }
    }
}

# 差分がない場合は終了
if ($newFiles.Count -eq 0 -and $updateFiles.Count -eq 0 -and $deleteFiles.Count -eq 0 -and $newFolders.Count -eq 0 -and $deleteFolders.Count -eq 0) {
    Write-Host "差分はありません。ファイルは同期されています。" -ForegroundColor Green
    Write-Host ""
    exit 0
}

# 差分詳細を表示
Write-Host "------------------------------------------------------------------------" -ForegroundColor Yellow
Write-Host " 差分詳細" -ForegroundColor Yellow
Write-Host "------------------------------------------------------------------------" -ForegroundColor Yellow
Write-Host ""

if ($newFiles.Count -gt 0) {
    Write-Host "=== 新規ファイル (TFSからGitへコピー) ===" -ForegroundColor Green
    foreach ($file in $newFiles) {
        Write-Host "  [新規] $($file.RelativePath)" -ForegroundColor Green
    }
    Write-Host ""
}

if ($updateFiles.Count -gt 0) {
    Write-Host "=== 更新ファイル (TFSの内容でGitを上書き) ===" -ForegroundColor Yellow
    foreach ($file in $updateFiles) {
        Write-Host "  [更新] $($file.RelativePath)" -ForegroundColor Yellow
    }
    Write-Host ""
}

if ($deleteFiles.Count -gt 0) {
    Write-Host "=== 削除対象ファイル (Gitから削除) ===" -ForegroundColor Red
    foreach ($file in $deleteFiles) {
        Write-Host "  [削除] $($file.RelativePath)" -ForegroundColor Red
    }
    Write-Host ""
}

if ($newFolders.Count -gt 0) {
    Write-Host "=== 新規フォルダ (空フォルダをGitに作成) ===" -ForegroundColor Magenta
    foreach ($folder in $newFolders) {
        Write-Host "  [フォルダ新規] $($folder.RelativePath)" -ForegroundColor Magenta
    }
    Write-Host ""
}

if ($deleteFolders.Count -gt 0) {
    Write-Host "=== 削除対象フォルダ (Gitから削除) ===" -ForegroundColor DarkRed
    foreach ($folder in $deleteFolders) {
        Write-Host "  [フォルダ削除] $($folder.RelativePath)" -ForegroundColor DarkRed
    }
    Write-Host ""
}

# 差分サマリー表示
Write-Host "========================================================================" -ForegroundColor Yellow
Write-Host " 差分サマリー" -ForegroundColor Yellow
Write-Host "========================================================================" -ForegroundColor Yellow
Write-Host ""
Write-Host "新規ファイル (TFS → Git):   " -NoNewline -ForegroundColor Green
Write-Host "$($newFiles.Count) 件"
Write-Host "更新ファイル (TFS → Git):   " -NoNewline -ForegroundColor Yellow
Write-Host "$($updateFiles.Count) 件"
Write-Host "削除対象ファイル (Gitのみ): " -NoNewline -ForegroundColor Red
Write-Host "$($deleteFiles.Count) 件"
Write-Host "新規フォルダ (空フォルダ):  " -NoNewline -ForegroundColor Magenta
Write-Host "$($newFolders.Count) 件"
Write-Host "削除対象フォルダ:           " -NoNewline -ForegroundColor DarkRed
Write-Host "$($deleteFolders.Count) 件"
Write-Host "変更なし:                   " -NoNewline -ForegroundColor Gray
Write-Host "$identicalCount 件"
Write-Host ""

# マージ確認
Write-Host "------------------------------------------------------------------------" -ForegroundColor Cyan
Write-Host " 上記の変更をGitに反映しますか？" -ForegroundColor Cyan
Write-Host " ※ TFSの内容でGitを同期します（TFS → Git）" -ForegroundColor White
Write-Host "------------------------------------------------------------------------" -ForegroundColor Cyan
Write-Host ""
Write-Host " 1. はい、同期を実行する"
Write-Host " 2. いいえ、キャンセルする"
Write-Host ""
$syncChoice = Read-Host "選択 (1-2)"

if ($syncChoice -ne "1") {
    Write-Host ""
    Write-Host "同期をキャンセルしました。" -ForegroundColor Yellow
    exit 0
}

# 同期実行
Write-Host ""
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host " 同期を実行中..." -ForegroundColor Cyan
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host ""

# 統計カウンタ
$copiedCount = 0
$deletedCount = 0
$folderCreatedCount = 0
$folderDeletedCount = 0

# 新規フォルダを作成（空フォルダ対応）
foreach ($folder in $newFolders) {
    try {
        if (-not (Test-Path $folder.GitFolderPath)) {
            New-Item -ItemType Directory -Path $folder.GitFolderPath -Force | Out-Null
            # .gitignoreファイルを作成（Gitで空フォルダを追跡するため）
            $gitignorePath = Join-Path $folder.GitFolderPath ".gitignore"
            $GITIGNORE_CONTENT | Out-File -FilePath $gitignorePath -Encoding UTF8 -NoNewline
            Write-Host "[フォルダ作成] $($folder.RelativePath)" -ForegroundColor Magenta
            $folderCreatedCount++
        }
    } catch {
        Write-Warning "フォルダ作成エラー: $($folder.RelativePath) - $_"
    }
}

# 新規ファイルをコピー
foreach ($file in $newFiles) {
    try {
        $targetDir = Split-Path -Path $file.GitFilePath -Parent
        if (-not (Test-Path $targetDir)) {
            New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
        }
        Copy-Item -Path $file.TfsFile.FullName -Destination $file.GitFilePath -Force
        Write-Host "[コピー完了] $($file.RelativePath)" -ForegroundColor Green
        $copiedCount++
    } catch {
        Write-Warning "コピーエラー: $($file.RelativePath) - $_"
    }
}

# 更新ファイルをコピー
foreach ($file in $updateFiles) {
    try {
        $targetDir = Split-Path -Path $file.GitFilePath -Parent
        if (-not (Test-Path $targetDir)) {
            New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
        }
        Copy-Item -Path $file.TfsFile.FullName -Destination $file.GitFilePath -Force
        Write-Host "[更新完了] $($file.RelativePath)" -ForegroundColor Yellow
        $copiedCount++
    } catch {
        Write-Warning "更新エラー: $($file.RelativePath) - $_"
    }
}

# 削除ファイルを削除
foreach ($file in $deleteFiles) {
    try {
        Remove-Item -Path $file.GitFile.FullName -Force
        Write-Host "[削除完了] $($file.RelativePath)" -ForegroundColor Red
        $deletedCount++
    } catch {
        Write-Warning "削除エラー: $($file.RelativePath) - $_"
    }
}

# 削除フォルダを削除（深い階層から削除するためソート）
$sortedDeleteFolders = $deleteFolders | Sort-Object { $_.RelativePath.Length } -Descending
foreach ($folder in $sortedDeleteFolders) {
    try {
        if (Test-Path $folder.GitFolder.FullName) {
            Remove-Item -Path $folder.GitFolder.FullName -Recurse -Force
            Write-Host "[フォルダ削除完了] $($folder.RelativePath)" -ForegroundColor DarkRed
            $folderDeletedCount++
        }
    } catch {
        Write-Warning "フォルダ削除エラー: $($folder.RelativePath) - $_"
    }
}

Write-Host ""
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host " 同期完了" -ForegroundColor Cyan
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "コピー/更新ファイル: $copiedCount" -ForegroundColor Green
Write-Host "削除ファイル:         $deletedCount" -ForegroundColor Red
Write-Host "作成フォルダ:         $folderCreatedCount" -ForegroundColor Magenta
Write-Host "削除フォルダ:         $folderDeletedCount" -ForegroundColor DarkRed
Write-Host ""
#endregion

#region Gitステータス確認とコミット
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host " Gitステータスを確認してください" -ForegroundColor Cyan
Write-Host "========================================================================" -ForegroundColor Cyan
git status

Write-Host ""
Write-Host "------------------------------------------------------------------------" -ForegroundColor Yellow
Write-Host "次の操作を選択してください:" -ForegroundColor Cyan
Write-Host " 1. 変更をコミットする"
Write-Host ""
Write-Host " 0. 何もせず終了"
Write-Host "------------------------------------------------------------------------" -ForegroundColor Yellow
$commitChoice = Read-Host "選択 (0-1)"

if ($commitChoice -eq "0") {
    Write-Host ""
    Write-Host "処理を終了します。" -ForegroundColor Yellow
    exit 0
}

if ($commitChoice -eq "1") {
    Write-Host ""
    $commitMsg = Read-Host "コミットメッセージを入力してください"

    git add -A
    git commit -m $commitMsg

    if ($LASTEXITCODE -ne 0) {
        Write-Host "[警告] コミットに失敗しました、または変更がありませんでした" -ForegroundColor Yellow
    } else {
        Write-Host "コミットが完了しました" -ForegroundColor Green
    }
}
#endregion

Write-Host ""
Write-Host "処理が完了しました。" -ForegroundColor Cyan
exit 0
