# Git Deploy to Linux

## 概要

Gitで変更されたファイルを自動的にLinuxサーバーに転送するWindowsバッチスクリプトです。

### 主な機能

- **Git変更ファイルの自動検出**: `git status` から変更/追加されたファイルを自動的に取得
- **削除ファイルの自動除外**: 削除されたファイルは転送対象から除外
- **柔軟な転送モード**: すべて転送 or 個別選択の2つのモードをサポート
- **SCP/PSCP対応**: Windows OpenSSH Client または PuTTY PSCP を自動検出
- **SSH公開鍵認証対応**: パスワード認証と公開鍵認証の両方に対応

## 必要な環境

### 必須
- **Git**: コマンドラインからgitコマンドが実行可能であること
- **PowerShell**: Windows 7以降（標準搭載）
- **SCP/PSCP**: 以下のいずれか
  - Windows OpenSSH Client (推奨)
  - PuTTY PSCP

### 推奨
- **SSH公開鍵認証**: パスワード入力を省略できます

## インストール/セットアップ

### 1. SCPコマンドのインストール

#### Windows OpenSSH Client（推奨）

Windows 10 (1809以降) / Windows 11の場合:

1. **設定** > **アプリ** > **オプション機能**
2. **機能の追加** をクリック
3. **OpenSSH クライアント** を検索してインストール

確認方法:
```cmd
scp -V
```

#### PuTTY PSCP（代替）

1. [PuTTY公式サイト](https://www.putty.org/) からダウンロード
2. `pscp.exe` をPATHが通っているフォルダに配置

確認方法:
```cmd
pscp -V
```

### 2. SSH公開鍵認証の設定（推奨）

#### 秘密鍵/公開鍵ペアの生成

```cmd
ssh-keygen -t rsa -b 4096 -f %USERPROFILE%\.ssh\id_rsa
```

#### 公開鍵をLinuxサーバーに配置

```cmd
type %USERPROFILE%\.ssh\id_rsa.pub | ssh user@hostname "cat >> ~/.ssh/authorized_keys"
```

または手動で:
1. `%USERPROFILE%\.ssh\id_rsa.pub` の内容をコピー
2. Linuxサーバーの `~/.ssh/authorized_keys` に追記

### 3. スクリプトの設定

`git-deploy-to-linux.bat` をテキストエディタで開き、設定セクションを編集:

```powershell
#region 設定 - ここを編集してください
#==============================================================================

# 転送先サーバー情報
$SSH_USER = "youruser"              # SSHユーザー名
$SSH_HOST = "192.168.1.100"         # SSHホスト名またはIPアドレス
$SSH_PORT = 22                      # SSHポート番号

# 転送先ディレクトリ (Linuxサーバー上のパス)
$REMOTE_DIR = "/home/youruser/project"

# SSH秘密鍵ファイル (公開鍵認証を使用する場合)
# パスワード認証の場合は空文字列 ""
$SSH_KEY = "$env:USERPROFILE\.ssh\id_rsa"

# Git リポジトリのルートディレクトリ (空文字列の場合は現在のディレクトリ)
$GIT_ROOT = ""

#==============================================================================
#endregion
```

## 使い方

### 基本的な使い方

1. Gitリポジトリのルートディレクトリに移動
2. バッチファイルをダブルクリックまたはコマンドプロンプトから実行

```cmd
git-deploy-to-linux.bat
```

### 実行の流れ

1. **Git変更ファイル検出**
   - `git status` から変更/追加されたファイルを取得
   - 削除されたファイルは自動的に除外

2. **ファイルリスト表示**
   ```
   ========================================
   転送予定のファイル一覧
   ========================================
     1. [変更]      src/main.py
     2. [追加]      src/utils.py
     3. [未追跡]    config/settings.json
   ```

3. **転送モード選択**
   ```
   これらのファイルを転送しますか？

     [A] すべて転送
     [I] 個別に選択
     [C] キャンセル

   選択してください (A/I/C):
   ```

4. **転送実行**
   - **すべて転送 (A)**: 全ファイルを一括転送
   - **個別選択 (I)**: 各ファイルごとに転送確認

### 個別選択モードの例

```
========================================
個別ファイル選択
========================================

転送: src/main.py (y/n): y
転送: src/utils.py (y/n): n
転送: config/settings.json (y/n): y

[選択] 2 個のファイルを転送します
```

## 実行例

### 例1: すべて転送

```
========================================
  Git Deploy to Linux
========================================

[情報] Gitリポジトリ: C:\Users\user\project
[情報] 転送先: deploy@192.168.1.100:/var/www/html

[実行] Git status を取得中...
[成功] 3 個のファイルが見つかりました（削除ファイルを除く）

========================================
転送予定のファイル一覧
========================================
  1. [変更]      index.html
  2. [変更]      style.css
  3. [追加]      script.js

これらのファイルを転送しますか？

  [A] すべて転送
  [I] 個別に選択
  [C] キャンセル

選択してください (A/I/C): A
[選択] すべてのファイルを転送します

[チェック] SCPコマンドを検出中...
[検出] Windows OpenSSH Client (scp.exe)

========================================
  ファイル転送開始
========================================

[転送] index.html
  ✓ 成功
[転送] style.css
  ✓ 成功
[転送] script.js
  ✓ 成功

========================================
  転送結果
========================================

成功: 3 個

すべてのファイル転送が完了しました！
```

### 例2: 個別選択

```
これらのファイルを転送しますか？

  [A] すべて転送
  [I] 個別に選択
  [C] キャンセル

選択してください (A/I/C): I

========================================
個別ファイル選択
========================================

転送: index.html (y/n): y
転送: style.css (y/n): y
転送: script.js (y/n): n

[選択] 2 個のファイルを転送します

[転送] index.html
  ✓ 成功
[転送] style.css
  ✓ 成功

========================================
  転送結果
========================================

成功: 2 個

すべてのファイル転送が完了しました！
```

## 注意事項

### Git Status の注意点

- **ステージング不要**: `git add` していないファイルも転送対象になります
- **削除ファイルは除外**: `git rm` や手動削除したファイルは自動的に除外されます
- **未追跡ファイルも対象**: 新規作成したファイルも転送対象に含まれます

### SSH接続の注意点

- **初回接続**: 初回接続時にホスト鍵の確認が表示される場合があります
- **パスワード認証**: 公開鍵認証が設定されていない場合、各ファイルごとにパスワード入力が必要です
- **タイムアウト**: ネットワーク接続が不安定な場合、転送に失敗することがあります

### ファイル転送の注意点

- **ディレクトリ構造**: 転送先にディレクトリが存在しない場合、転送に失敗します
- **上書き**: 同名ファイルは上書きされます（確認なし）
- **パーミッション**: 転送後のファイルパーミッションはSSHユーザーのumaskに依存します

### 制限事項

- **ディレクトリ作成**: 転送先ディレクトリは事前に作成されている必要があります
- **シンボリックリンク**: シンボリックリンクの扱いはSCPの実装に依存します
- **バイナリファイル**: 大きなバイナリファイルの転送は時間がかかる場合があります

## トラブルシューティング

### SCPコマンドが見つからない

**エラー**:
```
[エラー] SCPコマンドが見つかりません
```

**解決方法**:
1. Windows OpenSSH Clientをインストール（推奨）
2. または PuTTY PSCPをインストール

### SSH接続に失敗する

**エラー**:
```
✗ 失敗 (終了コード: 255)
```

**原因**:
- SSHホスト名/IPアドレスが間違っている
- SSHポート番号が間違っている
- ファイアウォールでブロックされている
- SSH公開鍵認証が正しく設定されていない

**解決方法**:
1. SSH接続を手動でテスト:
   ```cmd
   ssh user@hostname
   ```

2. 設定を確認:
   - `$SSH_USER`
   - `$SSH_HOST`
   - `$SSH_PORT`
   - `$SSH_KEY` (公開鍵認証の場合)

### Permission denied エラー

**エラー**:
```
Permission denied (publickey,password)
```

**解決方法**:

#### 公開鍵認証の場合
1. 公開鍵がLinuxサーバーに登録されているか確認:
   ```bash
   cat ~/.ssh/authorized_keys
   ```

2. パーミッション確認:
   ```bash
   chmod 700 ~/.ssh
   chmod 600 ~/.ssh/authorized_keys
   ```

#### パスワード認証の場合
1. スクリプトの設定を変更:
   ```powershell
   $SSH_KEY = ""  # 空文字列に設定
   ```

### 転送先ディレクトリが存在しない

**エラー**:
```
✗ 失敗: No such file or directory
```

**解決方法**:
1. Linuxサーバーにログインしてディレクトリを作成:
   ```bash
   mkdir -p /path/to/destination
   ```

2. または `$REMOTE_DIR` の設定を確認

### Git変更ファイルが検出されない

**メッセージ**:
```
[情報] 転送するファイルがありません
```

**原因**:
- 変更されたファイルがない
- すべて削除されたファイル
- Gitリポジトリではないディレクトリで実行している

**解決方法**:
1. Git statusを確認:
   ```cmd
   git status
   ```

2. Gitリポジトリのルートで実行しているか確認

### ファイル転送が途中で失敗する

**解決方法**:

1. **ネットワーク接続を確認**: ping でサーバーに到達できるか確認
2. **ディスク容量を確認**: 転送先のディスク容量が十分か確認
3. **ファイルサイズを確認**: 大きなファイルはタイムアウトする可能性があります
4. **個別選択モード**: 失敗したファイルのみ再転送

## 応用例

### 別のGitリポジトリを指定して実行

スクリプト内の `$GIT_ROOT` を変更:

```powershell
$GIT_ROOT = "C:\Users\user\another-project"
```

### 複数の転送先に対応

スクリプトをコピーして、それぞれ異なる設定にする:

```
git-deploy-to-prod.bat   # 本番環境用
git-deploy-to-staging.bat # ステージング環境用
git-deploy-to-dev.bat    # 開発環境用
```

### タスクスケジューラで自動実行

すべて転送モードを選択したい場合、スクリプトを修正:

```powershell
# ユーザー確認をスキップして常に全部転送
$choice = "A"
```

## ライセンス

MIT License

Copyright (c) 2025

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
