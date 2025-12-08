# VBA Tools

Excel VBAマクロによる自動化ツール集です。

## 📂 サブフォルダ

### word-bookmark-organizer/
Word文書のしおり（ブックマーク）整理とPDF出力ツールです。

- ExcelからWordを操作し、スタイルに基づいてアウトラインレベルを自動設定
- 「表題1」「表題2」「表題3」などの独自スタイルに対応
- しおり付きPDFを自動出力
- Input/Output方式でファイル管理が明確

### git-log-visualizer/
Gitコミット履歴の可視化ツールです。

- Excelから実行し、git logを表形式・統計・グラフで視覚化
- 5つのシート：Dashboard、CommitHistory、Statistics、Charts、BranchGraph
- 作者別・日別の統計情報を自動集計
- 全ブランチ対応で最近のN件を取得
- ブランチ構造の図形可視化機能

## 使い方

1. 各サブフォルダ内のREADME.mdを参照
2. .xlsmファイルまたは.basファイルをダウンロード
3. Excelで開くか、VBAエディタでインポート
4. マクロを実行

## 注意事項

- マクロを有効にする必要があります
- Excel 2010以降を推奨
- ファイルは必ずマクロ有効ブック(.xlsm)として保存してください
