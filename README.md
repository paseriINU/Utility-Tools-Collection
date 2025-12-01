# 便利ツール集 (Utility Tools Collection)

開発効率化・業務自動化のための便利ツールを言語・用途別に整理したリポジトリです。

## 📁 フォルダ構成

```
.
├── batch/                    # Windowsバッチスクリプト
│   ├── sync/                # 同期ツール
│   │   └── TFS-Git-sync     # TFS→Git同期スクリプト
│   ├── remote-exec/         # リモート実行ツール
│   │   ├── schtasks/        # タスクスケジューラ版
│   │   ├── winrm/           # WinRM版
│   │   └── powershell-remoting/  # PowerShell Remoting版
│   ├── git-diff-extract/    # Git差分ファイル抽出ツール
│   └── git-branch-manager/  # Gitブランチ管理ツール
│
├── vba/                     # Excel VBAマクロ
│   └── excel-automation/    # Excel自動化ツール
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

- **[Remote-Batch-Executor](batch/remote-exec/)**: リモートWindowsサーバでバッチファイルをCMDから実行
  - **タスクスケジューラ版**: schtasksコマンドでリモート実行（セットアップ簡単）
  - **WinRM版**: PowerShell Remotingで実行結果をリアルタイム取得
  - **PowerShell Remoting版**: 純粋なPowerShellスクリプトで高度な制御
  - Windows標準機能のみ使用（追加ツール不要）
  - セキュアな環境対応
  - 設定ファイル対応で複数サーバ管理が容易

- **[Git-Diff-Extract](batch/git-diff-extract/)**: Gitブランチ間の差分ファイルを抽出
  - フォルダ構造を保ったまま差分ファイルをコピー
  - main と develop などブランチ間の差分を簡単に抽出
  - デプロイ用差分ファイル作成に最適
  - ダブルクリックで実行可能

- **[Git-Branch-Manager](batch/git-branch-manager/)**: Gitブランチを数字で選択して削除
  - リモートブランチを対話的に削除
  - ローカルブランチを対話的に削除
  - リモート＆ローカル両方を一度に削除
  - main/master/develop は保護機能付き
  - 通常削除・強制削除を選択可能

### VBA Macros
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

1. 適切な言語フォルダを選択（batch/vba/javascript/powershell）
2. 用途に応じたサブフォルダを作成または既存のものを使用
3. スクリプトファイルとREADME.mdを配置
4. このREADME.mdの「現在利用可能なツール」セクションを更新

## 📝 ライセンス

このリポジトリのすべてのツール・スクリプトは **MIT License** で提供されています。

### 重要な免責事項

**このソフトウェアは「現状のまま」提供され、明示的または暗黙的な保証は一切ありません。**

作者は以下を含む、いかなる責任も負いません：
- ソフトウェアの使用により生じた損害
- データの損失や破損
- システムの障害や予期しない動作
- 業務上の損失

使用する場合は、**自己責任**でお願いします。

詳細は [LICENSE](LICENSE) ファイルをご確認ください。

---

**Note**: このリポジトリは個人の開発効率化・業務自動化を目的としています。
