# 便利ツール集 (Utility Tools Collection)

開発効率化・業務自動化のための便利ツールを言語・用途別に整理したリポジトリです。

## 📁 フォルダ構成

```
.
├── batch/                    # Windowsバッチスクリプト
│   └── sync/                # 同期ツール
│       └── TFS-Git-sync     # TFS→Git同期スクリプト
│
├── vba/                     # Excel VBAマクロ
│   └── excel-automation/    # Excel自動化ツール
│
├── python/                  # Pythonスクリプト
│   └── utilities/           # 汎用ユーティリティ
│
├── javascript/              # JavaScriptツール
│   └── browser-automation/  # ブラウザ自動化スクリプト
│
└── powershell/              # PowerShellスクリプト
    └── scripts/             # 各種スクリプト
```

## 🛠️ 現在利用可能なツール

### Batch Scripts
- **[TFS-Git-sync](batch/sync/)**: TFS（Team Foundation Server）とGitリポジトリを同期するバッチスクリプト
  - MD5ハッシュによる高速差分チェック
  - 自動的なファイル更新・追加・削除
  - 日本語ファイル名対応

### VBA Macros
*準備中 - 今後追加予定*

### Python Utilities
*準備中 - 今後追加予定*

### JavaScript Tools
*準備中 - 今後追加予定*

### PowerShell Scripts
*準備中 - 今後追加予定*

## 🚀 使い方

各ツールの詳細な使い方は、それぞれのフォルダ内のREADME.mdを参照してください。

1. 使いたいツールのフォルダに移動
2. そのフォルダ内のREADME.mdを確認
3. 必要に応じてスクリプトをカスタマイズ
4. 実行

## 📋 ツール追加ガイド

新しいツールを追加する際は、以下の手順に従ってください：

1. 適切な言語フォルダを選択（batch/vba/python/javascript/powershell）
2. 用途に応じたサブフォルダを作成または既存のものを使用
3. スクリプトファイルとREADME.mdを配置
4. このREADME.mdの「現在利用可能なツール」セクションを更新

## 📝 ライセンス

各ツールのライセンスは、それぞれのフォルダ内で指定されています。特に記載がない場合はMIT Licenseとします。

## 🤝 貢献

プルリクエストやイシューの報告を歓迎します。

## 📧 お問い合わせ

質問や提案がある場合は、Issuesセクションで報告してください。

---

**Note**: このリポジトリは個人の開発効率化・業務自動化を目的としています。
