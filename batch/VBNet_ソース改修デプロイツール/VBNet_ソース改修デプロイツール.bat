<# :
@echo off
chcp 65001 >nul
title VB.NET ソース改修デプロイツール
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
# VB.NET ソース改修デプロイツール
# 改修済みソースをメインフォルダにデプロイし、コンパイル後にDLLを取得
# =============================================================================

# UTF-8出力設定
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

#region 設定セクション
# =============================================================================
# 設定（環境に合わせて編集してください）
# =============================================================================

# メインフォルダ（本番ソースのパス）
$MAIN_FOLDER = "C:\Projects\MainSource"

# 改修済みソースフォルダ（メインと同じ構成）
$MODIFIED_FOLDER = "C:\Projects\ModifiedSource"

# バックアップフォルダの保存先（空の場合はメインフォルダの親に作成）
$BACKUP_BASE = ""

# MSBuildのパス（空の場合は自動検出）
$MSBUILD_PATH = ""

# ビルド構成（Release / Debug）
$BUILD_CONFIG = "Release"
#endregion

# タイトル表示
Write-Host ""
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host "  VB.NET ソース改修デプロイツール" -ForegroundColor Cyan
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host ""

#region MSBuild検出
function Find-MSBuild {
    # 設定されている場合はそれを使用
    if ($MSBUILD_PATH -ne "" -and (Test-Path $MSBUILD_PATH)) {
        return $MSBUILD_PATH
    }

    # Visual Studio 2022
    $vs2022Paths = @(
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
    )

    # Visual Studio 2019
    $vs2019Paths = @(
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
    )

    # Visual Studio 2017
    $vs2017Paths = @(
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2017\Enterprise\MSBuild\15.0\Bin\MSBuild.exe",
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2017\Professional\MSBuild\15.0\Bin\MSBuild.exe",
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2017\Community\MSBuild\15.0\Bin\MSBuild.exe",
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2017\BuildTools\MSBuild\15.0\Bin\MSBuild.exe"
    )

    $allPaths = $vs2022Paths + $vs2019Paths + $vs2017Paths

    foreach ($path in $allPaths) {
        if (Test-Path $path) {
            return $path
        }
    }

    # vswhere を使用した検出
    $vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
    if (Test-Path $vswhere) {
        $vsPath = & $vswhere -latest -requires Microsoft.Component.MSBuild -find "MSBuild\**\Bin\MSBuild.exe" | Select-Object -First 1
        if ($vsPath -and (Test-Path $vsPath)) {
            return $vsPath
        }
    }

    return $null
}
#endregion

#region フォルダ検証
Write-Host "[検証] フォルダを確認中..." -ForegroundColor Yellow
Write-Host ""

# メインフォルダの確認
if (-not (Test-Path $MAIN_FOLDER)) {
    Write-Host "[エラー] メインフォルダが見つかりません: $MAIN_FOLDER" -ForegroundColor Red
    Write-Host ""
    Write-Host "設定セクションの MAIN_FOLDER を正しいパスに編集してください。" -ForegroundColor Yellow
    exit 1
}
Write-Host "  メインフォルダ    : $MAIN_FOLDER" -ForegroundColor White

# 改修済みフォルダの確認
if (-not (Test-Path $MODIFIED_FOLDER)) {
    Write-Host "[エラー] 改修済みソースフォルダが見つかりません: $MODIFIED_FOLDER" -ForegroundColor Red
    Write-Host ""
    Write-Host "設定セクションの MODIFIED_FOLDER を正しいパスに編集してください。" -ForegroundColor Yellow
    exit 1
}
Write-Host "  改修済みフォルダ  : $MODIFIED_FOLDER" -ForegroundColor White
Write-Host ""
#endregion

#region 改修済みファイルの取得
Write-Host "[検索] 改修済みフォルダのファイルを取得中..." -ForegroundColor Yellow
Write-Host ""

# 改修済みフォルダ内のすべてのファイルを取得（相対パス）
$modifiedFiles = @()
Get-ChildItem -Path $MODIFIED_FOLDER -Recurse -File | ForEach-Object {
    $relativePath = $_.FullName.Substring($MODIFIED_FOLDER.Length).TrimStart('\')
    $modifiedFiles += [PSCustomObject]@{
        RelativePath = $relativePath
        FullPath = $_.FullName
    }
}

if ($modifiedFiles.Count -eq 0) {
    Write-Host "[エラー] 改修済みフォルダにファイルが見つかりません" -ForegroundColor Red
    exit 1
}

Write-Host "  検出されたファイル数: $($modifiedFiles.Count) 個" -ForegroundColor Green
Write-Host ""
#endregion

#region ソリューションファイルの検出
Write-Host "[検索] ソリューションファイルを検索中..." -ForegroundColor Yellow
Write-Host ""

# 改修済みフォルダ内の.slnファイルを検索
$slnFiles = Get-ChildItem -Path $MODIFIED_FOLDER -Filter "*.sln" -Recurse
if ($slnFiles.Count -eq 0) {
    Write-Host "[エラー] 改修済みフォルダ内にソリューションファイル(.sln)が見つかりません" -ForegroundColor Red
    exit 1
}

if ($slnFiles.Count -eq 1) {
    $selectedSln = $slnFiles[0]
    Write-Host "  検出されたソリューション: $($selectedSln.Name)" -ForegroundColor Green
} else {
    Write-Host "  複数のソリューションファイルが見つかりました:" -ForegroundColor Yellow
    Write-Host ""
    for ($i = 0; $i -lt $slnFiles.Count; $i++) {
        $displayNum = $i + 1
        Write-Host "   $displayNum. $($slnFiles[$i].Name)"
    }
    Write-Host ""
    Write-Host "   0. キャンセル"
    Write-Host ""

    $maxNum = $slnFiles.Count
    $selection = Read-Host "ビルドするソリューションを選択してください (0-$maxNum)"

    if ($selection -eq "0") {
        Write-Host "[キャンセル] 処理を中止しました" -ForegroundColor Yellow
        exit 0
    }

    if ($selection -match '^\d+$' -and [int]$selection -ge 1 -and [int]$selection -le $maxNum) {
        $selectedSln = $slnFiles[[int]$selection - 1]
    } else {
        Write-Host "[エラー] 無効な選択です" -ForegroundColor Red
        exit 1
    }
}

# メインフォルダ内の対応するソリューションファイルのパスを計算
$slnRelativePath = $selectedSln.FullName.Substring($MODIFIED_FOLDER.Length).TrimStart('\')
$mainSlnPath = Join-Path $MAIN_FOLDER $slnRelativePath

Write-Host ""
Write-Host "  ソリューション名: $($selectedSln.BaseName)" -ForegroundColor White
Write-Host "  メインのSLNパス : $mainSlnPath" -ForegroundColor White
Write-Host ""
#endregion

#region MSBuild検出
Write-Host "[検索] MSBuildを検索中..." -ForegroundColor Yellow

$msbuild = Find-MSBuild
if (-not $msbuild) {
    Write-Host "[エラー] MSBuildが見つかりません" -ForegroundColor Red
    Write-Host ""
    Write-Host "Visual StudioまたはBuild Toolsをインストールしてください。" -ForegroundColor Yellow
    Write-Host "または、設定セクションの MSBUILD_PATH を直接指定してください。" -ForegroundColor Yellow
    exit 1
}

Write-Host "  MSBuild: $msbuild" -ForegroundColor Green
Write-Host ""
#endregion

#region 処理確認
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host "  処理内容の確認" -ForegroundColor Cyan
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "以下の処理を実行します:" -ForegroundColor Yellow
Write-Host ""
Write-Host "  1. メインフォルダから対象ファイルをバックアップ（YYYYMMDD_bk）"
Write-Host "  2. 改修済みファイルをメインフォルダに上書きコピー"
Write-Host "  3. ソリューションをReleaseビルド"
Write-Host "  4. 生成されたDLLをデスクトップにコピー"
Write-Host ""
Write-Host "------------------------------------------------------------------------" -ForegroundColor White
Write-Host "  対象ファイル数    : $($modifiedFiles.Count) 個" -ForegroundColor White
Write-Host "  ソリューション    : $($selectedSln.BaseName).sln" -ForegroundColor White
Write-Host "  取得するDLL       : $($selectedSln.BaseName).dll" -ForegroundColor White
Write-Host "------------------------------------------------------------------------" -ForegroundColor White
Write-Host ""

$confirm = Read-Host "続行しますか? (y/n)"
if ($confirm -ne "y") {
    Write-Host "[キャンセル] 処理を中止しました" -ForegroundColor Yellow
    exit 0
}
Write-Host ""
#endregion

#region バックアップ処理
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host "  Step 1: バックアップ" -ForegroundColor Cyan
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host ""

# バックアップフォルダ名の生成（YYYYMMDD_bk）
$dateStr = Get-Date -Format "yyyyMMdd"
$backupFolderName = "${dateStr}_bk"

# バックアップ先の決定
if ($BACKUP_BASE -ne "" -and (Test-Path $BACKUP_BASE)) {
    $backupFolder = Join-Path $BACKUP_BASE $backupFolderName
} else {
    $mainParent = Split-Path -Path $MAIN_FOLDER -Parent
    $backupFolder = Join-Path $mainParent $backupFolderName
}

Write-Host "  バックアップ先: $backupFolder" -ForegroundColor White
Write-Host ""

# バックアップフォルダが既に存在する場合
if (Test-Path $backupFolder) {
    Write-Host "[警告] バックアップフォルダが既に存在します: $backupFolder" -ForegroundColor Yellow
    $overwrite = Read-Host "上書きしますか? (y/n)"
    if ($overwrite -ne "y") {
        Write-Host "[キャンセル] 処理を中止しました" -ForegroundColor Yellow
        exit 0
    }
}

# バックアップ実行
$backupCount = 0
$backupSkipped = 0

foreach ($file in $modifiedFiles) {
    $mainFilePath = Join-Path $MAIN_FOLDER $file.RelativePath
    $backupFilePath = Join-Path $backupFolder $file.RelativePath

    if (Test-Path $mainFilePath) {
        # バックアップ先ディレクトリの作成
        $backupDir = Split-Path -Path $backupFilePath -Parent
        if (-not (Test-Path $backupDir)) {
            New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
        }

        # ファイルコピー
        Copy-Item -Path $mainFilePath -Destination $backupFilePath -Force
        $backupCount++
        Write-Host "  [OK] $($file.RelativePath)" -ForegroundColor Gray
    } else {
        # メインフォルダに存在しないファイル（新規ファイル）
        $backupSkipped++
        Write-Host "  [新規] $($file.RelativePath)" -ForegroundColor DarkYellow
    }
}

Write-Host ""
Write-Host "  バックアップ完了: $backupCount 個" -ForegroundColor Green
if ($backupSkipped -gt 0) {
    Write-Host "  新規ファイル（スキップ）: $backupSkipped 個" -ForegroundColor DarkYellow
}
Write-Host ""
#endregion

#region ファイルコピー処理
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host "  Step 2: 改修ファイルのコピー" -ForegroundColor Cyan
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host ""

$copyCount = 0
$copyErrors = @()

foreach ($file in $modifiedFiles) {
    $sourceFilePath = $file.FullPath
    $destFilePath = Join-Path $MAIN_FOLDER $file.RelativePath

    try {
        # コピー先ディレクトリの作成
        $destDir = Split-Path -Path $destFilePath -Parent
        if (-not (Test-Path $destDir)) {
            New-Item -ItemType Directory -Path $destDir -Force | Out-Null
        }

        # ファイルコピー
        Copy-Item -Path $sourceFilePath -Destination $destFilePath -Force
        $copyCount++
        Write-Host "  [OK] $($file.RelativePath)" -ForegroundColor Gray
    } catch {
        $copyErrors += $file.RelativePath
        Write-Host "  [NG] $($file.RelativePath): $_" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "  コピー完了: $copyCount 個" -ForegroundColor Green
if ($copyErrors.Count -gt 0) {
    Write-Host "  エラー: $($copyErrors.Count) 個" -ForegroundColor Red
    Write-Host ""
    Write-Host "[エラー] コピーに失敗したファイルがあります。処理を中止しますか?" -ForegroundColor Yellow
    $continueOnError = Read-Host "(y=続行 / n=中止)"
    if ($continueOnError -ne "y") {
        Write-Host "[中止] 処理を中止しました" -ForegroundColor Yellow
        exit 1
    }
}
Write-Host ""
#endregion

#region ビルド処理
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host "  Step 3: ソリューションのビルド" -ForegroundColor Cyan
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host ""

# ソリューションファイルの存在確認
if (-not (Test-Path $mainSlnPath)) {
    Write-Host "[エラー] メインフォルダにソリューションファイルが見つかりません: $mainSlnPath" -ForegroundColor Red
    exit 1
}

Write-Host "  ソリューション: $mainSlnPath" -ForegroundColor White
Write-Host "  構成: $BUILD_CONFIG" -ForegroundColor White
Write-Host ""
Write-Host "  ビルド中..." -ForegroundColor Yellow

# MSBuild実行
$buildArgs = @(
    "`"$mainSlnPath`"",
    "/t:Rebuild",
    "/p:Configuration=$BUILD_CONFIG",
    "/p:Platform=`"Any CPU`"",
    "/v:minimal",
    "/nologo"
)

$buildProcess = Start-Process -FilePath $msbuild -ArgumentList $buildArgs -NoNewWindow -Wait -PassThru

if ($buildProcess.ExitCode -ne 0) {
    Write-Host ""
    Write-Host "[エラー] ビルドに失敗しました（終了コード: $($buildProcess.ExitCode)）" -ForegroundColor Red
    Write-Host ""
    Write-Host "詳細なビルドログを確認するには、Visual Studioでソリューションを開いてビルドしてください。" -ForegroundColor Yellow
    exit 1
}

Write-Host ""
Write-Host "  ビルド成功!" -ForegroundColor Green
Write-Host ""
#endregion

#region DLL取得
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host "  Step 4: DLLの取得" -ForegroundColor Cyan
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host ""

# DLL名（ソリューション名と同名）
$dllName = "$($selectedSln.BaseName).dll"

# ソリューションのあるディレクトリを基準にDLLを検索
$slnDir = Split-Path -Path $mainSlnPath -Parent

# bin\Release フォルダからDLLを検索
$dllSearchPaths = @(
    (Join-Path $slnDir "bin\$BUILD_CONFIG\$dllName"),
    (Join-Path $slnDir "$($selectedSln.BaseName)\bin\$BUILD_CONFIG\$dllName"),
    (Join-Path $slnDir "bin\$BUILD_CONFIG\net*\$dllName")  # .NET Core/.NET 5+
)

$foundDll = $null

# 直接パスでの検索
foreach ($searchPath in $dllSearchPaths) {
    if ($searchPath -like "*`**") {
        # ワイルドカードを含むパスの場合
        $matchedFiles = Get-ChildItem -Path $searchPath -ErrorAction SilentlyContinue
        if ($matchedFiles) {
            $foundDll = $matchedFiles | Select-Object -First 1
            break
        }
    } else {
        if (Test-Path $searchPath) {
            $foundDll = Get-Item $searchPath
            break
        }
    }
}

# 見つからない場合は再帰検索
if (-not $foundDll) {
    Write-Host "  DLLを検索中..." -ForegroundColor Yellow
    $foundDlls = Get-ChildItem -Path $slnDir -Filter $dllName -Recurse -ErrorAction SilentlyContinue |
                 Where-Object { $_.FullName -like "*\bin\$BUILD_CONFIG\*" }

    if ($foundDlls) {
        $foundDll = $foundDlls | Select-Object -First 1
    }
}

if (-not $foundDll) {
    Write-Host "[エラー] DLLが見つかりません: $dllName" -ForegroundColor Red
    Write-Host ""
    Write-Host "ビルドは成功しましたが、DLLが見つかりませんでした。" -ForegroundColor Yellow
    Write-Host "手動で bin\$BUILD_CONFIG フォルダを確認してください。" -ForegroundColor Yellow
    exit 1
}

Write-Host "  検出されたDLL: $($foundDll.FullName)" -ForegroundColor White

# デスクトップにコピー
$desktopPath = [Environment]::GetFolderPath("Desktop")
$destDllPath = Join-Path $desktopPath $dllName

# 既存ファイルがある場合の確認
if (Test-Path $destDllPath) {
    Write-Host ""
    Write-Host "[警告] デスクトップに同名のDLLが既に存在します" -ForegroundColor Yellow
    $overwriteDll = Read-Host "上書きしますか? (y/n)"
    if ($overwriteDll -ne "y") {
        # タイムスタンプ付きで別名保存
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $destDllPath = Join-Path $desktopPath "$($selectedSln.BaseName)_$timestamp.dll"
        Write-Host "  別名で保存します: $destDllPath" -ForegroundColor Yellow
    }
}

Copy-Item -Path $foundDll.FullName -Destination $destDllPath -Force

Write-Host ""
Write-Host "  DLLをデスクトップにコピーしました: $destDllPath" -ForegroundColor Green
Write-Host ""
#endregion

#region 完了
Write-Host ""
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host "  処理完了" -ForegroundColor Cyan
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  [完了] すべての処理が正常に完了しました" -ForegroundColor Green
Write-Host ""
Write-Host "------------------------------------------------------------------------" -ForegroundColor White
Write-Host "  バックアップ先  : $backupFolder" -ForegroundColor White
Write-Host "  コピー済みファイル: $copyCount 個" -ForegroundColor White
Write-Host "  DLL出力先       : $destDllPath" -ForegroundColor White
Write-Host "------------------------------------------------------------------------" -ForegroundColor White
Write-Host ""
#endregion

exit 0
