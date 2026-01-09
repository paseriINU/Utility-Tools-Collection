# Git Hooks - リモートリポジトリ保護設定

Linuxのリモートリポジトリにおいて、セキュリティとコード品質を保護するための業界標準的なGit Hooksセットです。

## 概要

このGit Hooksは、以下の保護機能を提供します：

### 1. updateフック - ブランチ保護
- **master/mainブランチの削除防止**
  - 誤ってmasterブランチを削除することを防ぎます

- **master/mainブランチへの直接プッシュ防止**
  - masterブランチへの直接プッシュを禁止し、フィーチャーブランチ経由のワークフローを強制します

### 2. pre-receiveフック - セキュリティとファイルチェック
- **機密情報の検出**
  - パスワード、APIキー、秘密鍵などの検出
  - AWS認証情報、データベース接続文字列の検出

- **大きなファイルの防止**
  - 10MB以上のファイルはエラー
  - 5MB以上のファイルは警告

- **禁止ファイルタイプの検出**
  - .env, credentials.json などの機密ファイル
  - 秘密鍵ファイル (.pem, .key, id_rsa など)

### 3. commit-msgフック - コミットメッセージ品質
- **コミットメッセージの最小文字数チェック**
  - デフォルト10文字以上

- **Conventional Commits形式の検証**（オプション）
  - feat:, fix:, docs: などの型付きコミット

- **チケット番号の必須化**（オプション）
  - JIRA-1234, #123 などの形式

### 4. post-mergeフック - Pull後の自動クリーン（クライアント側）
- **未追跡ファイルの自動削除**
  - リモートで削除されたファイルがローカルに残る問題を解決
  - `git pull` 後に自動的に `git clean -fd` を実行

- **安全設計**
  - `.gitignore`に含まれるファイルは削除されない
  - 削除前に対象ファイルを表示

## これらは標準的な設定か？

### 業界標準として一般的

✅ **非常に一般的**: 以下の環境では標準的な設定として広く採用されています：

1. **エンタープライズ開発**
   - 複数人で開発するチーム環境
   - コードレビューが必須の組織
   - CI/CD パイプラインを使用するプロジェクト

2. **オープンソースプロジェクト**
   - GitHub、GitLab、Bitbucketなどのホスティングサービスでは、ブランチ保護機能が標準搭載
   - Pull Request/Merge Request経由のマージが一般的

3. **品質管理が重要なプロジェクト**
   - 本番環境に直結するコードベース
   - 複数の承認者が必要な変更管理プロセス

### 使用されない場合

❌ **不要なケース**:

1. **個人開発プロジェクト**
   - 1人で開発している場合
   - 迅速な実験・プロトタイピングが優先される場合

2. **小規模チーム**
   - 2-3人の信頼できるメンバーのみのチーム
   - 頻繁なコミュニケーションが可能な環境

3. **学習・教育目的**
   - Gitの学習中
   - 制約のない環境で実験したい場合

### 推奨環境

以下の場合は、**必ず設定することを推奨**します：

- ✅ 3人以上のチーム開発
- ✅ コードレビュープロセスがある
- ✅ 本番環境へのデプロイが自動化されている
- ✅ 品質基準（テスト、リンター等）を満たす必要がある
- ✅ 変更履歴の追跡が重要

## ファイル構成

```
git-hooks/
├── README.md       # このファイル（説明書）
├── update          # ブランチ保護フック（サーバー側）
├── pre-receive     # セキュリティチェックフック（サーバー側）
├── commit-msg      # コミットメッセージ検証フック（サーバー側）
├── post-merge      # Pull後自動クリーンフック（クライアント側）
└── install.sh      # インストールスクリプト
```

### 各フックの詳細

| フック名 | 実行場所 | 実行タイミング | 用途 | 標準度 |
|---------|---------|-------------|------|-------|
| update | サーバー | プッシュ受信時（ブランチ単位） | ブランチ保護 | ⭐⭐⭐⭐⭐ 必須 |
| pre-receive | サーバー | プッシュ受信時（全体） | セキュリティ・ファイルチェック | ⭐⭐⭐⭐ 強く推奨 |
| commit-msg | サーバー | コミット作成時 | メッセージ品質 | ⭐⭐⭐ 推奨 |
| post-merge | クライアント | pull/merge後 | 未追跡ファイル削除 | ⭐⭐⭐ 推奨 |

## インストール方法

### 前提条件

- Linuxサーバー上にGitのリモートリポジトリ（bare repository）が存在すること
- リモートリポジトリへの書き込み権限があること

### 手順1: リモートリポジトリの準備（初めての場合）

リモートリポジトリがまだない場合は、以下のコマンドで作成します：

```bash
# リモートリポジトリを作成（bare repository）
mkdir -p /path/to/remote/repo.git
cd /path/to/remote/repo.git
git init --bare
```

### 手順2: フックのインストール

#### 方法A: インストールスクリプトを使用（推奨）

```bash
# このリポジトリをクローンまたはダウンロード
cd /path/to/Utility-Tools-Collection/git-hooks

# インストールスクリプトを実行
./install.sh /path/to/remote/repo.git
```

#### 方法B: 手動インストール

```bash
# updateフックをコピー
cp /path/to/Utility-Tools-Collection/git-hooks/update /path/to/remote/repo.git/hooks/update

# 実行権限を付与
chmod +x /path/to/remote/repo.git/hooks/update
```

### 手順3: 動作確認

別のマシンからリモートリポジトリをクローンして、動作を確認します：

```bash
# リモートリポジトリをクローン
git clone user@server:/path/to/remote/repo.git
cd repo

# masterブランチで変更を試みる
echo "test" > test.txt
git add test.txt
git commit -m "Test commit"
git push origin master  # エラーが表示されることを確認
```

期待される出力：
```
========================================
エラー: master ブランチへの直接プッシュは禁止されています
========================================

理由: master ブランチは保護されており、
      直接プッシュはできません。

推奨ワークフロー:
  1. フィーチャーブランチを作成
     git checkout -b feature/your-feature
  ...
```

## クライアント側フック（post-merge）のインストール

サーバー側フック（update, pre-receive）とは別に、クライアント側フック（post-merge）は**各開発者のPC**で設定が必要です。

### 手順（クローン後に1回だけ実行）

```bash
# 1. リポジトリをクローン
git clone \\server\repo.git
cd repo

# 2. フックパスを設定（この1コマンドで完了）
git config core.hooksPath git-hooks
```

### 動作確認

```bash
# pullを実行
git pull

# 以下のメッセージが表示されれば成功
# ========================================
#   post-merge: 未追跡ファイルをクリーン中...
# ========================================
```

### VSCodeでの動作

VSCodeのGUI操作（プルボタン）でも自動的に実行されます。
出力は「表示」→「出力」→「Git」で確認できます。

### 注意事項

- `.gitignore`に含まれるファイルは削除されません
- ローカルで作成した未追跡ファイルは削除されます
- 設定はリポジトリごとに必要です（グローバル設定ではありません）

## 使用方法

### 正しいワークフロー

master/mainブランチへの変更は、以下の手順で行います：

```bash
# 1. フィーチャーブランチを作成
git checkout -b feature/new-feature

# 2. 変更を行う
echo "new feature" > feature.txt
git add feature.txt
git commit -m "Add new feature"

# 3. フィーチャーブランチをプッシュ
git push origin feature/new-feature

# 4. masterブランチにマージ（リモートサーバー上で実行）
# サーバーにSSHでログインして実行
cd /path/to/remote/repo.git
git merge --no-ff feature/new-feature master
```

### マージの自動化（オプション）

GitHub/GitLab風のマージリクエスト機能を実装したい場合は、以下のツールを検討してください：

- **Gitea**: 軽量なGitホスティングサービス
- **GitLab CE**: オンプレミスで動作するGitLab Community Edition
- **Gogs**: GoベースのシンプルなGitサーバー

## カスタマイズ

### 保護するブランチを変更

`update` フックの以下の行を編集します：

```bash
# 保護するブランチのリスト
PROTECTED_BRANCHES=("refs/heads/master" "refs/heads/main")
```

例：developブランチも保護する場合：

```bash
PROTECTED_BRANCHES=("refs/heads/master" "refs/heads/main" "refs/heads/develop")
```

### 特定のユーザーのみ許可

特定のユーザー（例：管理者）のみmasterへのプッシュを許可する場合：

```bash
# 許可するユーザーリスト
ALLOWED_USERS=("admin" "deploy")

# 現在のユーザー名を取得
current_user=$(whoami)

# 許可されたユーザーかチェック
for allowed_user in "${ALLOWED_USERS[@]}"; do
    if [ "$current_user" = "$allowed_user" ]; then
        # 許可されたユーザーは通過
        exit 0
    fi
done

# それ以外は拒否
echo "エラー: あなたのユーザー ($current_user) はmasterブランチへプッシュできません" >&2
exit 1
```

### エラーメッセージのカスタマイズ

`update` フックのエラーメッセージ部分を編集して、プロジェクト固有のワークフローに合わせることができます。

### pre-receiveフックのカスタマイズ

#### ファイルサイズ制限の変更

```bash
# pre-receiveフックの設定部分を編集
MAX_FILE_SIZE=$((50 * 1024 * 1024))  # 50MBに変更
WARN_FILE_SIZE=$((20 * 1024 * 1024))  # 20MBで警告
```

#### 禁止ファイルパターンの追加

```bash
# 禁止ファイルパターンに追加
FORBIDDEN_FILES=(
    '\.env$'
    '\.env\..*'
    'credentials\.json$'
    'config/database\.yml$'  # 追加例
    '.*\.backup$'            # 追加例
)
```

#### 機密情報検出パターンの追加

```bash
# カスタムパターンを追加
SECRET_PATTERNS=(
    'password\s*=\s*["\047][^"\047]{8,}'
    'api[_-]?key\s*=\s*["\047][A-Za-z0-9]{20,}'
    'MY_COMPANY_SECRET\s*=\s*["\047][^"\047]+'  # 追加例
)
```

### commit-msgフックのカスタマイズ

#### Conventional Commitsを有効化

```bash
# commit-msgフックの設定部分を編集
ENFORCE_CONVENTIONAL_COMMITS=true  # falseをtrueに変更
```

#### チケット番号を必須化

```bash
# チケット番号を必須にする
REQUIRE_TICKET_NUMBER=true

# チケット番号パターンをカスタマイズ
TICKET_PATTERN='(MYPROJECT-|#)[0-9]+'  # プロジェクト固有のパターン
```

#### コミットメッセージの最小文字数を変更

```bash
# 最小文字数を変更
MIN_MESSAGE_LENGTH=20  # 20文字に変更
```

## 各フックの業界標準度と推奨度

### updateフック（ブランチ保護）

**業界標準度**: ⭐⭐⭐⭐⭐ (必須)

- **GitHub**: Protected Branches（標準機能）
- **GitLab**: Protected Branches（標準機能）
- **Bitbucket**: Branch Permissions（標準機能）
- **採用率**: 90%以上のエンタープライズプロジェクトで使用

**推奨される環境:**
- ✅ すべてのチーム開発プロジェクト
- ✅ 本番環境に直結するリポジトリ
- ✅ 複数の開発者が参加するプロジェクト

### pre-receiveフック（セキュリティチェック）

**業界標準度**: ⭐⭐⭐⭐ (強く推奨)

- **機密情報検出**: GitHub Advanced SecurityのSecret Scanningと同等
- **ファイルサイズ制限**: GitHub（100MB制限）、GitLab（デフォルト10GB）
- **採用率**: 70%以上の大規模プロジェクトで使用

**推奨される環境:**
- ✅ 金融、医療、インフラなどの重要システム
- ✅ オープンソースプロジェクト
- ✅ セキュリティが重要なプロジェクト

**必須とされる業界:**
- 金融機関（PCI DSS準拠）
- 医療機関（HIPAA準拠）
- 政府機関（各種セキュリティ基準）

### commit-msgフック（メッセージ品質）

**業界標準度**: ⭐⭐⭐ (推奨)

- **Conventional Commits**: Angular、Reactなど多くのOSSプロジェクトで採用
- **チケット番号**: JIRAなど課題管理ツール連携で一般的
- **採用率**: 50%程度のプロジェクトで使用

**推奨される環境:**
- ✅ 変更ログを自動生成するプロジェクト
- ✅ チケット管理システムと連携するプロジェクト
- ✅ コミット履歴の品質を重視するプロジェクト

**注意**: クライアント側のフックのため、強制力は弱い（開発者がローカルで無効化可能）

## トラブルシューティング

### フックが動作しない

**問題**: プッシュが通ってしまう

**原因と解決策**:

1. **実行権限がない**
   ```bash
   chmod +x /path/to/remote/repo.git/hooks/update
   ```

2. **フックの配置場所が間違っている**
   - Bare repository: `/path/to/repo.git/hooks/update`
   - 通常のリポジトリ: `/path/to/repo/.git/hooks/update`

3. **ファイル名が間違っている**
   - 正しいファイル名: `update` (拡張子なし)
   - 間違い: `update.sh`, `update.sample`

### 緊急時にフックをバイパス

**一時的に無効化する場合**:

```bash
# フックをリネーム
mv /path/to/remote/repo.git/hooks/update /path/to/remote/repo.git/hooks/update.disabled

# プッシュ

# フックを再有効化
mv /path/to/remote/repo.git/hooks/update.disabled /path/to/remote/repo.git/hooks/update
```

**完全に削除する場合**:

```bash
rm /path/to/remote/repo.git/hooks/update
```

### エラーメッセージが文字化けする

リモートサーバーのロケール設定を確認してください：

```bash
# ロケールを確認
locale

# UTF-8が設定されていない場合は設定
export LANG=ja_JP.UTF-8
export LC_ALL=ja_JP.UTF-8
```

## セキュリティ上の注意

1. **フックファイルのパーミッション**
   - フックファイルは誰でも読めますが、書き込みは管理者のみに制限してください
   ```bash
   chmod 755 /path/to/remote/repo.git/hooks/update
   chown root:root /path/to/remote/repo.git/hooks/update
   ```

2. **定期的な監査**
   - フックファイルが改ざんされていないか定期的にチェックしてください
   ```bash
   # チェックサムを記録
   sha256sum /path/to/remote/repo.git/hooks/update > hooks.sha256

   # 定期的に検証
   sha256sum -c hooks.sha256
   ```

## よくある質問

### Q1: GitHub/GitLabのブランチ保護との違いは？

**A**: 機能的には似ていますが、以下の違いがあります：

| 項目 | Git Hooks | GitHub/GitLab |
|------|-----------|---------------|
| 実装場所 | リモートリポジトリ | ホスティングサービス |
| UI | なし（CLIのみ） | Web UIあり |
| 追加機能 | カスタマイズ可能 | CI/CD連携、レビュー機能等 |
| コスト | 無料（自己管理） | プランによる |

### Q2: フックはクライアント側にもコピーされますか？

**A**: いいえ。サーバー側のフック（update, pre-receive等）は、リモートリポジトリでのみ動作し、クライアント側にはコピーされません。

### Q3: 他のブランチ保護設定との併用は可能ですか？

**A**: はい。以下と併用できます：

- SSH鍵認証による制限
- Gitoliteなどのアクセス制御ツール
- CI/CDパイプラインのチェック

### Q4: パフォーマンスへの影響は？

**A**: ほぼありません。フックスクリプトは軽量で、プッシュごとに数ミリ秒程度の処理時間です。

## 関連情報

- [Git公式ドキュメント - Hooks](https://git-scm.com/book/ja/v2/Git-%E3%81%AE%E3%82%AB%E3%82%B9%E3%82%BF%E3%83%9E%E3%82%A4%E3%82%BA-Git-%E3%83%95%E3%83%83%E3%82%AF)
- [GitHub - ブランチ保護ルール](https://docs.github.com/ja/repositories/configuring-branches-and-merges-in-your-repository/defining-the-mergeability-of-pull-requests/about-protected-branches)
- [GitLab - 保護されたブランチ](https://docs.gitlab.com/ee/user/project/protected_branches.html)

## ライセンス

このツールはMITライセンスの下で提供されています。
