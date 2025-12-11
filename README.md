# 便利ツール集 (Utility Tools Collection)

開発効率化・業務自動化のための便利ツールを言語・用途別に整理したリポジトリです。

## 📁 フォルダ構成

```
.
├── batch/                    # Windowsバッチスクリプト
│   ├── TFS_to_Git_同期ツール/   # TFS→Git同期ツール
│   ├── リモートバッチ実行ツール/   # リモート実行ツール（PowerShell Remoting）
│   ├── Git_差分ファイル抽出ツール/ # Git差分ファイル抽出ツール
│   ├── Git_ブランチ削除ツール/   # Gitブランチ管理ツール
│   ├── Git_Linuxデプロイツール/  # Git→Linux転送ツール
│   ├── jp1-job-executor/    # JP1ジョブネット起動ツール
│   │   ├── JP1_リモートジョブ起動ツール/
│   │   └── JP1_ジョブネット起動ツール/
│   └── サーバ構成情報収集ツール/  # サーバ構成情報収集ツール
│
├── linux/                   # Linuxスクリプト
│   └── winrm-client/        # WinRMクライアント（Python/Bash/C）
│
├── vba/                     # Excel VBAマクロ
│   ├── word-bookmark-organizer/  # Word文書しおり整理ツール
│   ├── git-log-visualizer/  # Git Log 可視化ツール
│   └── excel-file-comparator/  # Excel/Word ファイル比較ツール
│
├── git-hooks/               # Git Hooks（ブランチ保護等）
│
├── samples/                 # サンプルツール・テンプレート集
│
└── javascript/              # JavaScriptツール（準備中）
```

## 🛠️ 現在利用可能なツール

### Batch Scripts

**Note**: このカテゴリにはWindowsバッチファイル(.bat)およびPowerShellスクリプト(.ps1)が含まれます。
- **[TFS to Git 同期ツール](batch/TFS_to_Git_同期ツール/)**: TFS（Team Foundation Server）とGitリポジトリを同期するバッチスクリプト
  - MD5ハッシュによる高速差分チェック
  - 自動的なファイル更新・追加・削除
  - 日本語ファイル名対応

- **[リモートバッチ実行ツール](batch/リモートバッチ実行ツール/)**: リモートWindowsサーバでバッチファイルを実行（PowerShell Remoting）
  - ダブルクリックで実行可能な.batハイブリッドスクリプト
  - WinRM設定の自動構成と復元（TrustedHosts自動設定）
  - 環境選択機能（tst1t/tst2t）
  - 実行結果のリアルタイム表示と終了コード取得
  - ログファイル自動保存
  - ネットワークパス（UNCパス）対応

- **[Git 差分ファイル抽出ツール](batch/Git_差分ファイル抽出ツール/)**: Gitブランチ間の差分ファイルを抽出
  - フォルダ構造を保ったまま差分ファイルをコピー
  - main と develop などブランチ間の差分を簡単に抽出
  - デプロイ用差分ファイル作成に最適
  - ダブルクリックで実行可能

- **[Git ブランチ削除ツール](batch/Git_ブランチ削除ツール/)**: Gitブランチを数字で選択して削除
  - リモートブランチを対話的に削除
  - ローカルブランチを対話的に削除
  - リモート＆ローカル両方を一度に削除
  - main/master/develop は保護機能付き
  - 通常削除・強制削除を選択可能

- **[Git Linuxデプロイツール](batch/Git_Linuxデプロイツール/)**: Git変更ファイルをLinuxサーバーに転送
  - 複数環境対応・拡張子フィルタ・削除ファイル除外
  - すべて転送 or 個別選択の2つのモード
  - Linux側の自動設定（mkdir/chmod/chown）
  - SCP/SSH対応（Windows OpenSSH Client）
  - ネットワークパス（UNCパス）からの実行対応
  - ダブルクリックで実行可能

- **[JP1 ジョブ起動ツール](batch/jp1-job-executor/)**: JP1/AJS3ジョブネットを起動
  - [JP1 リモートジョブ起動ツール](batch/jp1-job-executor/JP1_リモートジョブ起動ツール/): PowerShell Remoting版
  - [JP1 ジョブネット起動ツール](batch/jp1-job-executor/JP1_ジョブネット起動ツール/): REST API版
  - ローカルPCにJP1インストール不要
  - ダブルクリックで実行可能なスタンドアローン版

- **[サーバ構成情報収集ツール](batch/サーバ構成情報収集ツール/)**: サーバ構成情報収集ツール
  - ネットワーク・セキュリティ設定をExcel出力
  - WinRM設定、ファイアウォール、開放ポート、レジストリ
  - WinRM実行前の事前調査に最適

### Linux Tools

- **[WinRM-Client](linux/winrm-client/)**: LinuxからWindowsへWinRM接続してコマンド実行
  - Python版・Bash版・C言語版の3種類を提供
  - 追加パッケージ不要（標準ライブラリのみ）
  - IT制限環境対応

### Git Hooks

- **[Git-Hooks](git-hooks/)**: リモートリポジトリ保護設定
  - master/mainブランチへの直接プッシュ防止
  - 機密情報・大きなファイルの検出
  - コミットメッセージ品質チェック

### VBA Macros

- **[Word-Bookmark-Organizer](vba/word-bookmark-organizer/)**: Word文書のしおり（ブックマーク）整理とPDF出力
  - ExcelからWordを操作し、スタイルに基づいてアウトラインレベルを自動設定
  - 「表題1」「表題2」「表題3」などの独自スタイルに対応
  - しおり付きPDFを自動出力（目次で適切な箇所に飛べる）
  - Input/Output方式でファイル管理が明確
  - スタイル名のカスタマイズ可能

- **[Git-Log-Visualizer](vba/git-log-visualizer/)**: Gitコミット履歴の可視化ツール
  - Excelから実行し、git logを表形式・統計・グラフで視覚化
  - 4つのシート：Dashboard、CommitHistory、Statistics、Charts
  - 作者別・日別の統計情報を自動集計
  - 全ブランチ対応で最近のN件を取得
  - コミットメッセージ・変更行数も表示
  - フィルター機能でデータを絞り込み

- **[Excel-File-Comparator](vba/excel-file-comparator/)**: Excel/Wordファイル比較ツール
  - 2つのExcelファイルを比較し、差異を一覧表示
  - 2つのWordファイルを比較（段落単位/詳細比較）
  - シート単位・セル単位での差異検出
  - 差異の種類を識別（値変更、追加、削除）
  - 結果をExcelシートに出力
  - 差異セルのハイライト表示（黄色:変更、緑:追加、赤:削除）

### JavaScript Tools
*準備中 - 今後追加予定*

### Samples（提案ツール）

今後作成を検討しているツールの一覧は [samples/README.md](samples/README.md) を参照してください。

## 🚀 使い方

各ツールの詳細な使い方は、それぞれのフォルダ内のREADME.mdを参照してください。

1. 使いたいツールのフォルダに移動
2. そのフォルダ内のREADME.mdを確認
3. 必要に応じてスクリプトをカスタマイズ
4. 実行

## 📋 ツール追加ガイド

新しいツールを追加する際は、以下の手順に従ってください：

1. 適切な言語フォルダを選択（batch/vba/javascript）
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
