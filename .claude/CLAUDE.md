# CLAUDE.md
必ず日本語で回答してください。
This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

このリポジトリは個人の開発効率化・業務自動化のための便利ツール集です。言語・用途別に整理されたスクリプトとマクロを管理しています。

## Repository Structure

```
.
├── batch/                    # Windowsバッチスクリプト
│   └── sync/                # 同期ツール
│       ├── sync_tfs_to_git.bat  # TFS→Git同期スクリプト
│       ├── README.md
│       └── LICENSE
│
├── vba/                     # Excel VBAマクロ
│   └── excel-automation/    # Excel自動化ツール
│       └── README.md
│
└── javascript/              # JavaScriptツール
    └── browser-automation/  # ブラウザ自動化スクリプト
        └── README.md
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
- **TFS-Git-sync** (`batch/sync/`): TFS（Team Foundation Server）とGitリポジトリを同期（PowerShellロジック埋め込み版）
- **Remote-Batch-Executor** (`batch/remote-exec/`): リモートサーバでバッチ実行（schtasks/WinRM/PowerShell Remoting）
  - パラメータ版とスタンドアローン版
  - 環境選択機能付きハイブリッド版
- **JP1-Job-Executor** (`batch/jp1-job-executor/`): JP1/AJS3 REST APIを使用したジョブネット起動
  - パラメータ版とスタンドアローン版
- **Git-Diff-Extract** (`batch/git-diff-extract/`): Gitブランチ間の差分ファイル抽出
- **Git-Branch-Manager** (`batch/git-branch-manager/`): Gitブランチを対話的に削除

**注**: すべてのツールは`.bat`ファイル単体で動作し、PowerShellが必要な場合もポリグロットパターンで埋め込まれています。

### Other Categories
VBA、JavaScriptのツールは今後追加予定

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
