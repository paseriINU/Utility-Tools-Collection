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
├── python/                  # Pythonスクリプト
│   └── utilities/           # 汎用ユーティリティ
│       └── README.md
│
├── javascript/              # JavaScriptツール
│   └── browser-automation/  # ブラウザ自動化スクリプト
│       └── README.md
│
└── powershell/              # PowerShellスクリプト
    └── scripts/             # 各種スクリプト
        └── README.md
```

## Development Guidelines

### 新しいツールを追加する場合

1. **適切なフォルダを選択**
   - バッチスクリプト → `batch/[用途]/`
   - VBAマクロ → `vba/[用途]/`
   - Pythonスクリプト → `python/[用途]/`
   - JavaScriptツール → `javascript/[用途]/`
   - PowerShellスクリプト → `powershell/[用途]/`

2. **ファイル配置**
   - スクリプトファイルを配置
   - 必要に応じてREADME.mdを作成
   - ライセンス情報が必要な場合はLICENSEファイルを追加

3. **ドキュメント更新**
   - ルートのREADME.mdの「現在利用可能なツール」セクションを更新
   - 各ツールのREADME.mdに使い方を記載

### コーディング規約

#### Batch Scripts (.bat)
- **エンコーディング**: UTF-8（BOMなし）で保存
- 日本語のファイル名・パスに対応すること
- コマンドプロンプトでの実行を想定
- 先頭にスクリプトの目的をコメントで記載

#### VBA Macros (.bas, .xlsm)
- Excel 2010以降で動作すること
- マクロ有効ブック(.xlsm)として保存
- 変数は明示的に宣言（Option Explicit使用）
- 日本語のコメントで処理内容を説明

#### Python Scripts (.py)
- Python 3.8以降で動作すること
- 必要なライブラリはrequirements.txtに記載
- docstringで関数・クラスの説明を記載
- PEP 8に準拠したコーディングスタイル

#### JavaScript (.js)
- ブラウザコンソールでの実行を想定
- ブックマークレット形式も考慮
- 対象サイトの構造変更に注意
- 先頭にコメントで対象サイト・目的を記載

#### PowerShell (.ps1)
- PowerShell 5.1以降で動作すること
- 実行ポリシーへの配慮
- 管理者権限が必要な場合は明記
- コメントベースのヘルプを記載

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

- 基本的に`main`ブランチで直接作業
- 大きな変更の場合はフィーチャーブランチを作成することも可

## Special Considerations

- **個人用リポジトリ**: プルリクエストやイシュー管理は不要
- **多言語対応**: ファイル名やコメントに日本語を使用可能
- **エンコーディング注意**: 特にバッチファイルはUTF-8で保存
- **プライバシー**: 機密情報（パスワード、APIキーなど）をコミットしないこと

## Tools Currently Available

### Batch Scripts
- **TFS-Git-sync** (`batch/sync/`): TFS（Team Foundation Server）とGitリポジトリを同期

### Other Categories
VBA、Python、JavaScript、PowerShellのツールは今後追加予定

## Development Approach

このリポジトリで作業する際は：

1. **既存パターンに従う**: 同じ言語の既存ツールのスタイルを参考に
2. **シンプルに保つ**: 複雑さを避け、目的に集中
3. **ドキュメントを充実**: 後で見返したときに理解できるように
4. **再利用性を考慮**: 他のプロジェクトでも使えるように汎用的に

IMPORTANT: Claudeはこのコンテキストが現在のタスクに関連する場合のみ応答してください。関連性がない場合は、このコンテキストに言及しないでください。
