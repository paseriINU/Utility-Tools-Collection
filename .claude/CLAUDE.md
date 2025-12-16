# CLAUDE.md
必ず日本語で回答してください。
This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 重要な作業ルール

### 広範囲な修正を行う場合の事前確認

**必須**: 以下のいずれかに該当する場合、**修正を実行する前に必ずユーザーに確認すること**：

1. **複数ファイルの修正**
   - 3つ以上のファイルを修正する場合
   - リポジトリ全体に影響する変更の場合

2. **実装方法の選択肢がある場合**
   - 複数のアプローチが考えられる場合
   - 異なる技術的トレードオフがある場合

3. **大規模なリファクタリング**
   - ファイル構造の変更
   - コーディングパターンの変更
   - 既存の動作を大きく変更する場合

**確認時に提示すべき情報**:
- 修正対象のファイルリスト
- 各アプローチの説明
- メリット・デメリットの比較
- 推奨する方法とその理由

**例**:
```
以下の7つのバッチファイルを修正します：
1. batch/Git_差分ファイル抽出ツール/Git_差分ファイル抽出ツール.bat
2. batch/TFS_to_Git_同期ツール/TFS_to_Git_同期ツール.bat
...

UNCパス対応には以下の方法があります：
1. 一時ドライブマッピング方式（確実だが複雑）
2. PushD/PopD方式（シンプル・推奨）
3. PowerShell Set-Location方式（環境依存性高い）

どの方式を採用しますか？
```

### プログラム作成時の品質チェック（必須）

**重要**: プログラムを作成・修正した場合、以下のチェックを**必ず実施してから完成**とすること：

1. **構文チェック**
   - 作成したコードの構文が正しいかを確認する
   - 構文エラーがある場合は修正する

2. **コンパイル/実行テスト（すべてのプログラムに対して必須）**
   - **C/C++**: `gcc` または `g++` でコンパイルし、エラーがないことを確認
     ```bash
     gcc -o output_name source.c
     # または警告も確認する場合
     gcc -Wall -o output_name source.c
     ```
   - **Python**: `python3 -m py_compile` で構文チェック
     ```bash
     python3 -m py_compile script.py
     ```
   - **Bash/Shell**: `bash -n` で構文チェック
     ```bash
     bash -n script.sh
     ```
   - **PowerShell**: 可能であれば構文解析を実行
   - **JavaScript**: `node --check` で構文チェック
     ```bash
     node --check script.js
     ```
   - **その他の言語**: 該当する構文チェックコマンドを使用

3. **エラー発生時の対応**
   - コンパイルエラーや構文エラーが発生した場合は、**必ず修正**してから完成とする
   - 修正後、再度チェックを実行して成功を確認する

4. **完了報告**
   - チェックが成功したことをユーザーに報告する
   - 例: 「コンパイル成功しました」「構文チェック完了しました」

**例**:
```
[作成完了後]
構文チェックを実行します...
$ gcc -Wall -o winrm_exec winrm_exec.c
コンパイル成功しました。エラー・警告なしです。
```

### ソースコード修正時のREADME同期（必須）

**重要**: ソースコードを作成・修正した場合、**関連するREADME.mdも必ず同時に更新すること**。

1. **更新が必要なケース**
   - 新しいツール・スクリプトを追加した場合
   - 既存ツールの機能を追加・変更した場合
   - コマンドライン引数やオプションを変更した場合
   - 依存関係や動作環境が変更された場合
   - ファイル名やフォルダ構成を変更した場合

2. **更新対象のREADME**
   - **ツール固有のREADME**: 該当ツールフォルダ内の`README.md`
   - **カテゴリREADME**: 言語カテゴリフォルダの`README.md`（例: `batch/README.md`, `vba/README.md`）
   - **ルートREADME**: 新規ツール追加時は`README.md`（リポジトリルート）も更新

3. **同期チェックリスト**
   - [ ] 機能説明がソースの実際の動作と一致しているか
   - [ ] 使用方法・コマンド例が正確か
   - [ ] 依存関係・動作環境の記載が正確か
   - [ ] ファイル一覧やフォルダ構成が現状と一致しているか

4. **完了報告**
   - ソース修正とREADME更新を同一コミットに含める
   - コミットメッセージに両方の変更を記載する

**例**:
```
[ソース修正後]
README.mdを更新します...
- 新機能の説明を追加
- コマンドライン引数の更新
- 実行例の追加

ソースコードとREADMEの同期完了しました。
```

### 確認プロンプトの表記統一（必須）

**重要**: 確認プロンプト（y/n形式）は、すべて**小文字で統一**すること。

- ✅ 正しい: `(y/n)`
- ❌ 誤り: `(Y/N)`, `(y/N)`, `(Y/n)`

**例**:
```powershell
$confirm = Read-Host "続行しますか？ (y/n)"
$answer = Read-Host "削除しますか? (y/n)"
```

**理由**:
- 大文字・小文字が混在するとユーザーが混乱する
- 小文字で統一することで一貫性のあるUIを提供

## Project Overview

このリポジトリは個人の開発効率化・業務自動化のための便利ツール集です。言語・用途別に整理されたスクリプトとマクロを管理しています。

## Repository Structure

```
.
├── batch/                           # Windowsバッチスクリプト
│   ├── TFS_to_Git_同期ツール/       # TFS→Git同期ツール
│   ├── リモートバッチ実行ツール/    # リモート実行ツール
│   ├── Git_差分ファイル抽出ツール/  # Git差分ファイル抽出
│   ├── Git_ブランチ削除ツール/      # Gitブランチ管理
│   ├── Git_Linuxデプロイツール/     # Git→Linux転送
│   ├── JP1_リモートジョブ起動ツール/ # JP1ジョブネット起動
│   └── サーバ構成情報収集ツール/    # サーバ構成情報収集
│
├── vba/                             # Excel VBAマクロ
│   ├── Word_しおり整理ツール/       # Wordしおり整理
│   ├── Git_Log_可視化ツール/        # Git Log可視化
│   └── Excel_Word_ファイル比較ツール/ # ファイル比較
│
├── linux/                           # Linuxスクリプト
│   ├── winrm-client/                # WinRMクライアント
│   └── opentp1-deploy/              # OpenTP1デプロイ自動化
│
└── javascript/                      # JavaScriptツール（準備中）
```

## Development Guidelines

### 新しいツールを追加する場合

1. **適切なフォルダを選択**
   - バッチスクリプト（.bat / .ps1） → `batch/[用途]/`
   - VBAマクロ → `vba/[用途]/`
   - JavaScriptツール → `javascript/[用途]/`

2. **ファイル配置**
   - スクリプトファイルを配置
   - 必要に応じてREADME.mdを作成
   - ライセンス情報が必要な場合はLICENSEファイルを追加

3. **ドキュメント更新**
   - ルートのREADME.mdの「現在利用可能なツール」セクションを更新
   - 各ツールのREADME.mdに使い方を記載

### コーディング規約

#### Batch Scripts (.bat / .ps1)
- **バッチファイル (.bat)**:
  - **エンコーディング**: 基本的にShift_JISで保存
  - 日本語のファイル名・パスに対応すること
  - コマンドプロンプトでの実行を想定
  - 先頭にスクリプトの目的をコメントで記載
  - **文字コード対応**:
    - **Shift_JIS互換文字のみ使用**: レ点チェックマーク（✓）は `[OK]`、バツ印（✗）は `[NG]` など、Shift_JISで表現できる文字のみを使用すること
    - **UTF-8入力対応**: `git status` など外部コマンドからUTF-8で返ってくる文字がある場合は、スクリプト冒頭に `chcp 65001 >nul` を追加すること
  - **タイトル表示**:
    - すべてのプログラムにおいて、実行開始時にタイトルまたはヘッダーをコマンドプロンプトに表示すること
    - 例: `title ツール名` または PowerShell の `Write-Host` でヘッダーを表示
  - **メニュー操作（番号選択式）**:
    - 数字で選択肢を提示する場合、**「0」はキャンセル/終了用に統一**すること
    - 機能選択は「1」から開始すること
    - 表示形式例:
      ```
      選択してください:
       1. 機能A
       2. 機能B

       0. キャンセル
      ```
    - 入力プロンプト例: `選択 (0-2)` のように範囲を明示すること

- **PowerShellスクリプト (.ps1) → ハイブリッド.bat形式を推奨**:
  - **重要**: PowerShellを使用する場合は、`.ps1`ファイルではなく、ポリグロットパターンを使用した`.bat`形式で作成すること
  - ポリグロットパターン（標準版）:
    ```batch
    <# :
    @echo off
    chcp 65001 >nul
    title ツール名
    setlocal
    powershell -NoProfile -ExecutionPolicy Bypass -Command "iex ((gc '%~f0') -join \"`n\")"
    set EXITCODE=%ERRORLEVEL%
    pause
    exit /b %EXITCODE%
    : #>

    # PowerShellコードはここから
    # 最初にタイトルを表示
    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host "  ツール名" -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""
    ```
  - ポリグロットパターン（管理者権限必須版）:
    ```batch
    <# :
    @echo off
    chcp 65001 >nul
    title ツール名
    setlocal

    rem 管理者権限チェック
    net session >nul 2>&1
    if %errorLevel% neq 0 (
        echo 管理者権限が必要です。管理者として再起動します...
        powershell -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
        exit /b
    )

    powershell -NoProfile -ExecutionPolicy Bypass -Command "$scriptDir=('%~dp0' -replace '\\$',''); iex ((gc '%~f0') -join \"`n\")"
    set EXITCODE=%ERRORLEVEL%
    pause
    exit /b %EXITCODE%
    : #>

    # PowerShellコードはここから
    # 最初にタイトルを表示
    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host "  ツール名" -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""

    # $scriptDir 変数でバッチファイルのディレクトリパスが利用可能
    ```
  - ポリグロットパターン（UNCパス対応版 - PushD/PopD方式）:
    **重要**: UNCパス（`\\server\share\folder`形式のネットワークパス）に配置されたバッチファイルを実行する場合、PowerShellが直接UNCパスから実行できないため、PushD/PopDで一時的にドライブマッピングを行います。
    ```batch
    <# :
    @echo off
    chcp 65001 >nul
    title ツール名
    setlocal

    rem UNCパス対応（PushD/PopDで自動マッピング）
    pushd "%~dp0"

    powershell -NoProfile -ExecutionPolicy Bypass -Command "$scriptDir=('%~dp0' -replace '\\$',''); try { iex ((gc '%~f0') -join \"`n\") } finally { Set-Location C:\ }"
    set EXITCODE=%ERRORLEVEL%

    popd

    pause
    exit /b %EXITCODE%
    : #>

    # PowerShellコードはここから
    # 最初にタイトルを表示
    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host "  ツール名" -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""

    # $scriptDir 変数でバッチファイルのディレクトリパスが利用可能
    ```
    **UNCパス対応版（PushD/PopD方式）の動作**:
    - `pushd "%~dp0"` でUNCパスを自動検出し、一時ドライブをマッピング
    - PowerShell内で `try-finally` を使用してエラー時も確実にクリーンアップ
    - `finally { Set-Location C:\ }` でカレントディレクトリを変更してからpopd実行
    - `popd` で一時ドライブマッピングを自動解除
    - **×ボタンで閉じた場合もWindowsが自動的にドライブを解除**（pushd/popdの仕組み）
    - ローカルパスの場合は通常通り実行（オーバーヘッドなし）

    **従来の明示的なドライブマッピング方式との比較**:
    - ✅ コードが大幅にシンプル（約50%削減）
    - ✅ ドライブレター検索ループが不要
    - ✅ Windowsの標準機能でクリーンアップを保証
    - ✅ エラー処理がシンプル
  - PowerShell 5.1以降で動作すること
  - コメントベースのヘルプを記載
  - **スタンドアローン版を基本とする**: ソース内の設定セクションを編集するだけで使用可能にすること
  - 管理者権限が必要な場合は自動昇格パターンを使用すること

- **既存の.ps1ファイルの変換**:
  - 既存の`.ps1`ファイルを見つけた場合は、自動的にハイブリッド`.bat`形式に変換すること
  - `.ps1`ファイルは保守性のため削除または非推奨とすること

- **WinRM関連ツールの標準仕様**:

  WinRM（Windows Remote Management）を使用するツールは、以下の実装を**必須**とします：

  **必須実装項目**:
  1. **管理者権限の自動昇格**
     - `net session` チェックで管理者権限を確認
     - 権限がない場合は自動的にUACプロンプトで再起動

  2. **WinRMサービスの自動起動と復元**
     - スクリプト開始時にWinRMサービスの状態を確認
     - 停止している場合のみ起動し、`$winrmServiceWasStarted` フラグを設定
     - finally ブロックでフラグを確認し、起動した場合のみ停止して元の状態に復元

  3. **TrustedHostsの自動設定と復元**
     - 現在の TrustedHosts 設定を `$originalTrustedHosts` に保存
     - 接続先が含まれていない場合のみ一時的に追加
     - finally ブロックで必ず元の設定に復元

  4. **確認プロンプトの完全無効化**
     - すべての `Set-Item WSMan:\localhost\Client\TrustedHosts` コマンドに `-Confirm:$false` を追加
     - 言語環境に依存しない動作を保証

  5. **エラー時の自動復元**
     - try-catch-finally パターンを使用
     - finally ブロックで設定の復元を保証（エラー時も実行される）
     - WinRM設定エラー時は `exit 1` で終了

  6. **環境変数を使わない実装**
     - `$scriptDir` 変数をバッチ起動時に PowerShell に直接渡す
     - 環境変数の汚染を防ぐ

  **実装パターン例**:
  ```powershell
  #region WinRM設定の保存と自動設定
  $originalTrustedHosts = $null
  $winrmConfigChanged = $false
  $winrmServiceWasStarted = $false

  try {
      # 現在のTrustedHostsを取得
      $originalTrustedHosts = (Get-Item WSMan:\localhost\Client\TrustedHosts -ErrorAction SilentlyContinue).Value

      # WinRMサービスの起動確認
      $winrmService = Get-Service -Name WinRM -ErrorAction SilentlyContinue
      if ($winrmService.Status -ne 'Running') {
          Start-Service -Name WinRM -ErrorAction Stop
          $winrmServiceWasStarted = $true
      }

      # TrustedHostsに接続先を追加（必要な場合のみ）
      if ($needsConfig) {
          Set-Item WSMan:\localhost\Client\TrustedHosts -Value $newValue -Force -Confirm:$false
          $winrmConfigChanged = $true
      }
  } catch {
      Write-Host "[エラー] WinRM設定の自動構成に失敗しました" -ForegroundColor Red
      exit 1
  }
  #endregion

  # メイン処理
  try {
      # リモート実行処理
  } catch {
      # エラー処理
      $exitCode = 1
  } finally {
      #region WinRM設定の復元
      if ($winrmConfigChanged) {
          Set-Item WSMan:\localhost\Client\TrustedHosts -Value $originalTrustedHosts -Force -Confirm:$false
      }

      if ($winrmServiceWasStarted) {
          Stop-Service -Name WinRM -Force -ErrorAction Stop
      }
      #endregion
  }
  ```

  **基本方針**:
  - **スタンドアローン版のみ**: config ファイルを使わず、ソース内に設定を記述
  - **ダブルクリックで実行可能**: .bat 形式で保存し、設定編集のみで使用可能
  - **自動設定・自動復元**: ユーザーが手動でWinRM設定を行う必要がない

#### VBA Macros (.bas, .xlsm)
- Excel 2010以降で動作すること
- マクロ有効ブック(.xlsm)として保存
- 変数は明示的に宣言（Option Explicit使用）
- 日本語のコメントで処理内容を説明

- **メインシートの作成（必須）**:
  - すべてのVBAツールは初期化時に「メインシート」を作成すること
  - メインシートには以下を含める：
    - ツールタイトル（色付きヘッダー）
    - 設定入力欄（必要に応じてドロップダウン）
    - 実行ボタン
    - 使い方・説明文
    - 必要な環境・動作条件
  - 初期化用マクロ名: `Initialize[ツール名]`
  - フォーマット用サブ: `Format[シート名]Sheet`
  - 例:
    ```vba
    Public Sub InitializeSQLGenerator()
        CreateSheet SHEET_MAIN
        FormatMainSheet  ' メインシートのフォーマット
        ' ...
    End Sub
    ```

- **ドロップダウン選択時の自動処理**:
  - ドロップダウンで値を選択した際に関連する項目を自動更新する場合は、`Worksheet_Change`イベントを使用すること
  - 初期化時にシートモジュールへイベントコードを自動追加する実装を推奨
  - VBAプロジェクトへのアクセスが許可されていない環境用に、手動更新ボタンも併設すること
  - 例:
    ```vba
    ' 標準モジュール側
    Public Sub OnSelectionChanged(ByVal changedRange As Range)
        ' 変更されたセルに応じて処理
    End Sub

    ' シートモジュール側（自動追加）
    Private Sub Worksheet_Change(ByVal Target As Range)
        ModuleName.OnSelectionChanged Target
    End Sub
    ```

#### JavaScript (.js)
- ブラウザコンソールでの実行を想定
- ブックマークレット形式も考慮
- 対象サイトの構造変更に注意
- 先頭にコメントで対象サイト・目的を記載

### ドキュメント規約

各ツールのREADME.mdには以下を含めること：

1. **概要** - ツールの目的と機能
2. **必要な環境** - 動作環境・依存関係
3. **インストール/セットアップ** - 準備手順
4. **使い方** - 具体的な実行方法
5. **実行例** - サンプル出力
6. **注意事項** - 使用上の制限・注意点
7. **トラブルシューティング** - よくある問題と解決方法
8. **ライセンス** - 必要に応じて

## Git Workflow

### コミットメッセージ

- 日本語または英語で記載
- 変更内容を明確に記述
- Claude Codeで作成した場合は自動的に署名が追加されます

### ブランチ戦略

- **mainブランチへの直接プッシュは禁止**
- **必ずフィーチャーブランチを作成してプルリクエスト（PR）を作成すること**
- ブランチ命名規則: `claude/[機能名]-[session-id]`
  - 例: `claude/remote-exec-consolidation-01BGfeHT5izXrCtTTXPVzWdZ`
- 作業フロー:
  1. フィーチャーブランチを作成: `git checkout -b claude/feature-name-sessionid`
  2. 変更をコミット
  3. リモートにプッシュ: `git push -u origin claude/feature-name-sessionid`
  4. GitHubでプルリクエストを作成
  5. レビュー後、mainにマージ

## Special Considerations

- **個人用リポジトリ**: mainブランチは保護されており、すべての変更はPRを通してマージ
- **多言語対応**: ファイル名やコメントに日本語を使用可能
- **エンコーディング注意**: 特にバッチファイルはShift_JISで保存（UTF-8入力がある場合は chcp 65001 使用）
- **プライバシー**: 機密情報（パスワード、APIキーなど）をコミットしないこと

## Tools Currently Available

### Batch Scripts (.bat)
- **TFS_to_Git_同期ツール** (`batch/TFS_to_Git_同期ツール/`): TFS→Git同期
- **リモートバッチ実行ツール** (`batch/リモートバッチ実行ツール/`): リモートサーバでバッチ実行
- **Git_差分ファイル抽出ツール** (`batch/Git_差分ファイル抽出ツール/`): Gitブランチ間の差分ファイル抽出
- **Git_ブランチ削除ツール** (`batch/Git_ブランチ削除ツール/`): Gitブランチを対話的に削除
- **Git_Linuxデプロイツール** (`batch/Git_Linuxデプロイツール/`): Git変更ファイルをLinuxへ転送
- **JP1_リモートジョブ起動ツール** (`batch/JP1_リモートジョブ起動ツール/`): JP1ジョブネット起動
- **サーバ構成情報収集ツール** (`batch/サーバ構成情報収集ツール/`): サーバ構成情報収集

**注**: すべてのツールは`.bat`ファイル単体で動作し、PowerShellが必要な場合もポリグロットパターンで埋め込まれています。

### VBA Macros
- **Word_しおり整理ツール** (`vba/Word_しおり整理ツール/`): Word文書のしおり整理とPDF出力
- **Git_Log_可視化ツール** (`vba/Git_Log_可視化ツール/`): Gitコミット履歴の可視化
- **Excel_Word_ファイル比較ツール** (`vba/Excel_Word_ファイル比較ツール/`): Excel/Wordファイル比較

### Linux Tools
- **WinRM-Client** (`linux/winrm-client/`): LinuxからWindowsへWinRM接続
- **OpenTP1-Deploy** (`linux/opentp1-deploy/`): OpenTP1環境でのCプログラムデプロイ自動化

## Development Approach

このリポジトリで作業する際は：

1. **既存パターンに従う**: 同じ言語の既存ツールのスタイルを参考に
2. **シンプルに保つ**: 複雑さを避け、目的に集中
3. **ドキュメントを充実**: 後で見返したときに理解できるように
4. **再利用性を考慮**: 他のプロジェクトでも使えるように汎用的に

## IT制限環境対応方針

**重要**: このリポジトリのツールは、IT制限環境でも動作することを最優先とします。

### 基本原則

1. **標準ライブラリ・標準コマンドを最優先使用**
   - **Python**: 標準ライブラリのみで実装（`pip install`不要）
   - **Bash/Shell**: curl、grep、sed、awk等の標準コマンドのみ使用
   - **PowerShell**: 標準コマンドレットのみ使用
   - 追加パッケージのインストールを極力避ける

2. **IT制限環境を常に考慮**
   以下の環境でも動作することを保証：
   - インターネット接続が制限されている環境
   - pip/yum/apt等のパッケージ管理ツールが使用できない環境
   - 管理者権限がない環境（クライアント側）
   - ファイアウォールで通信が制限されている環境

3. **Python開発における標準ライブラリ優先**
   - ✅ 使用可能: `urllib`, `xml.etree.ElementTree`, `base64`, `uuid`, `socket`, `ssl`, `json`, `re`, `os`, `sys`等
   - ❌ 避ける: `requests`, `pywinrm`, `paramiko`, `xmltodict`, `beautifulsoup4`等の外部パッケージ
   - 例外: セキュリティ上または機能上やむを得ない場合のみ

4. **Bash/Shell開発における標準コマンド優先**
   - ✅ 使用可能: `curl`, `grep`, `sed`, `awk`, `base64`, `date`, `cut`, `tr`, `sort`, `uniq`等
   - ❌ 避ける: `jq`, `xmllint`, `python`, `perl`（インストールされていない可能性）
   - 例外: 複雑なJSON/XMLパースが必須の場合のみ

### 例外的に追加パッケージを提案する場合

以下の**すべての条件**を満たす場合のみ、追加パッケージの使用を提案できます：

1. **技術的正当性**
   - 標準ライブラリでの実装が著しく複雑になる（500行以上）
   - セキュリティ上の重大なリスクがある（暗号化、認証等）
   - パフォーマンスに重大な影響がある（10倍以上の速度差）

2. **提案方法**
   - **必ず標準ライブラリ版も併せて提供**すること
   - メリット・デメリットを明確に説明
   - インストール方法と制限事項を明記
   - ユーザーに選択を委ねる（強制しない）

3. **提案フォーマット例**:
   ```
   標準ライブラリ版とパッケージ版の2つの実装を提供します：

   【標準ライブラリ版】（推奨）
   - メリット: pip install不要、IT制限環境で動作
   - デメリット: コードが長い、機能が限定的

   【パッケージ版】（オプション）
   - 必要パッケージ: requests, beautifulsoup4
   - メリット: コードが簡潔、高機能
   - デメリット: pip installが必要、IT制限環境では使用不可
   ```

### 実装時のチェックリスト

新しいツールを作成する際は、以下を確認すること：

- [ ] 追加パッケージのインストールは不要か？
- [ ] Python 3.6以降（RHEL 7 / CentOS 7標準）で動作するか？
- [ ] インターネット接続なしで動作するか？
- [ ] 一般ユーザー権限で実行可能か？（必要な場合を除く）
- [ ] README.mdに依存関係を明記したか？
- [ ] IT制限環境での動作確認方法を記載したか？

### 対応例

**良い例**:
```python
# 標準ライブラリのみ使用
import urllib.request
import xml.etree.ElementTree as ET
import base64

# WinRMプロトコルを標準ライブラリで実装
```

**避けるべき例**:
```python
# 外部パッケージに依存
import requests  # ❌ pip install requests が必要
import pywinrm   # ❌ pip install pywinrm が必要
```

この方針により、どのような制限された環境でも確実に動作するツールを提供します。

IMPORTANT: Claudeはこのコンテキストが現在のタスクに関連する場合のみ応答してください。関連性がない場合は、このコンテキストに言及しないでください。
