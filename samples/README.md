# サンプルツール集

業務効率化のためのサンプルツール提案です。
各言語の特性を活かした実用的なツールを提供します。

---

## 作成済みツール

| フォルダ | ツール名 | 説明 |
|----------|----------|------|
| `server-config-collector/` | サーバ構成情報収集 | ネットワーク・セキュリティ設定をExcel出力 |

---

## 提案ツール一覧

### PowerShell / Batch（Windows運用向け）

| ツール名 | 概要 | 用途 |
|----------|------|------|
| **log-collector** | 複数サーバからログを一括収集 | 障害調査、監査 |
| **service-monitor** | Windowsサービスの状態確認・再起動 | 運用監視 |
| **port-checker** | サーバのポート開放確認 | ネットワーク確認 |
| **disk-usage-report** | ディスク使用量レポート作成 | 容量管理 |
| **scheduled-task-export** | タスクスケジューラ設定のエクスポート | 設定バックアップ |
| **event-log-analyzer** | Windowsイベントログ解析 | 障害分析 |
| **file-hash-checker** | ファイル改ざん検知（ハッシュ比較） | セキュリティ |
| **backup-with-rotation** | 世代管理付きバックアップ | データ保護 |
| **config-diff** | 環境間の設定ファイル比較 | 構成管理 |

### Shell（Linux/Unix運用向け）

| ツール名 | 概要 | 用途 |
|----------|------|------|
| **log-rotate-custom** | カスタムログローテーション | ログ管理 |
| **process-monitor** | プロセス監視・自動再起動 | 運用監視 |
| **db-backup** | データベースバックアップ（MySQL/PostgreSQL） | データ保護 |
| **deploy-script** | アプリケーションデプロイスクリプト | リリース自動化 |
| **ssl-cert-checker** | SSL証明書有効期限チェック | セキュリティ |
| **cron-job-manager** | cronジョブ管理・可視化 | ジョブ管理 |
| **server-health-check** | サーバヘルスチェック | 運用監視 |

### VBA（Excel業務効率化）

| ツール名 | 概要 | 用途 |
|----------|------|------|
| **csv-importer** | 複数CSV一括インポート・整形 | データ集計 |
| **excel-to-pdf** | Excel帳票のPDF一括変換 | 帳票作成 |
| **data-validator** | 入力データ検証・エラーハイライト | 品質管理 |
| **pivot-table-generator** | ピボットテーブル自動生成 | データ分析 |
| **sheet-merger** | 複数ブックのシート統合 | データ統合 |
| **mail-sender** | Outlookメール一括送信 | 業務連絡 |
| **schedule-calendar** | ガントチャート自動生成 | プロジェクト管理 |
| **invoice-generator** | 請求書自動生成 | 経理業務 |

### JavaScript（ブラウザ自動化・Webツール）

| ツール名 | 概要 | 用途 |
|----------|------|------|
| **form-auto-filler** | Webフォーム自動入力（ブックマークレット） | 入力効率化 |
| **table-to-csv** | Webテーブル→CSV変換 | データ抽出 |
| **page-scraper** | ページ情報抽出（ブックマークレット） | 情報収集 |
| **json-formatter** | JSON整形・検証ツール | 開発支援 |
| **regex-tester** | 正規表現テスター | 開発支援 |
| **diff-viewer** | テキスト差分表示（HTML） | 比較確認 |
| **markdown-preview** | Markdownプレビュー | ドキュメント作成 |
| **color-picker** | カラーピッカー・変換ツール | デザイン支援 |

### Python（データ処理・自動化）

| ツール名 | 概要 | 用途 |
|----------|------|------|
| **excel-processor** | Excel一括処理（標準ライブラリ版） | データ処理 |
| **file-organizer** | ファイル自動整理（日付・拡張子別） | ファイル管理 |
| **text-converter** | 文字コード一括変換 | データ変換 |
| **duplicate-finder** | 重複ファイル検出 | ディスク整理 |
| **report-generator** | レポート自動生成（HTML/PDF） | 帳票作成 |
| **api-client** | REST API クライアント | システム連携 |
| **log-parser** | ログファイル解析・集計 | 障害分析 |

---

## 優先度の高いツール（推奨）

業務でよく使われる機能を優先的に実装することを推奨します：

### 即効性が高いもの
1. **server-config-collector** (PowerShell) - サーバ設定の可視化 ✅作成済み
2. **log-collector** (PowerShell) - 障害調査で頻繁に使用
3. **csv-importer** (VBA) - Excel業務で頻繁に使用
4. **form-auto-filler** (JavaScript) - 繰り返し入力の効率化

### 運用品質向上
1. **service-monitor** (PowerShell) - サービス監視の自動化
2. **ssl-cert-checker** (Shell) - 証明書切れ防止
3. **event-log-analyzer** (PowerShell) - 障害の早期発見

### 開発効率向上
1. **json-formatter** (JavaScript) - API開発で必須
2. **regex-tester** (JavaScript) - 正規表現の検証
3. **diff-viewer** (JavaScript) - コードレビュー支援

---

## リクエスト

作成してほしいツールがあれば、以下の情報と共にリクエストしてください：

1. **ツール名**
2. **目的・用途**
3. **入力・出力**
4. **実行環境**（Windows/Linux/ブラウザ など）
5. **制約事項**（IT制限環境かどうか など）
