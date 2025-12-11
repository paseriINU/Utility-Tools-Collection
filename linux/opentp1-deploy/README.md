# OpenTP1 デプロイ自動化ツール

## 概要

OpenTP1環境でのソース配置を自動化するシェルスクリプトです。

以下の一連の流れを自動で実行します：

1. **OpenTP1 停止** (`dcstop -f`)
2. **ソース配置**（バックアップ付き）
3. **OpenTP1 起動** (`dcstart`)

---

## 必要な環境

- Linux（Red Hat / CentOS / RHEL 等）
- OpenTP1がインストールされていること
- OpenTP1の停止・起動権限を持つユーザー

---

## ファイル構成

```
opentp1-deploy/
├── opentp1_deploy.sh   # デプロイ自動化スクリプト
└── README.md           # このファイル
```

---

## 使い方

### 1. 設定を編集

`opentp1_deploy.sh` をテキストエディタで開き、設定セクションを環境に合わせて編集します：

```bash
#==============================================================================
# 設定セクション（環境に合わせて編集してください）
#==============================================================================

# OpenTP1のインストールパス
OPENTP1_HOME="/opt/OpenTP1"

# OpenTP1コマンドのパス（通常はOPENTP1_HOME/bin）
OPENTP1_BIN="${OPENTP1_HOME}/bin"

# コピー元ファイル（フルパスで指定）
SOURCE_FILE="/home/user/src/myprogram"

# 配置先ディレクトリ
DEPLOY_DIR="/opt/OpenTP1/aplib"

# バックアップを作成するか（true/false）
CREATE_BACKUP=true

# 停止待機時間（秒）
STOP_WAIT_TIME=10

# 起動待機時間（秒）
START_WAIT_TIME=10
```

### 2. 実行権限を付与

```bash
chmod +x opentp1_deploy.sh
```

### 3. 実行

```bash
./opentp1_deploy.sh
```

---

## 実行例

```
================================================================
  OpenTP1 デプロイ自動化ツール
================================================================

実行日時    : 2025-12-11 10:30:45
コピー元    : /home/user/src/myprogram
配置先      : /opt/OpenTP1/aplib/myprogram
ログファイル: opentp1_deploy_20251211_103045.log

以下の処理を実行します:
  1. OpenTP1 停止 (dcstop -f)
  2. ソース配置
  3. OpenTP1 起動 (dcstart)

実行しますか? (y/n): y

======================================
  事前チェック
======================================

[2025-12-11 10:30:47] [INFO] OpenTP1 bin: /opt/OpenTP1/bin [OK]
[2025-12-11 10:30:47] [INFO] コピー元: /home/user/src/myprogram [OK]
[2025-12-11 10:30:47] [INFO] 配置先Dir: /opt/OpenTP1/aplib [OK]
[2025-12-11 10:30:47] [INFO] 事前チェック完了

======================================
  OpenTP1 停止
======================================

[2025-12-11 10:30:47] [INFO] OpenTP1の状態を確認中...
[2025-12-11 10:30:47] [INFO] OpenTP1は稼働中です。停止します...
[2025-12-11 10:30:47] [INFO] dcstop -f を実行中...
[2025-12-11 10:30:48] [INFO] 10秒待機中...
[2025-12-11 10:30:58] [INFO] 停止を確認中...
[2025-12-11 10:30:58] [INFO] OpenTP1の停止を確認しました

======================================
  ソース配置
======================================

[2025-12-11 10:30:58] [INFO] 既存ファイルをバックアップ: myprogram.bak.20251211_103058
[2025-12-11 10:30:58] [INFO] ファイルをコピー: /home/user/src/myprogram → /opt/OpenTP1/aplib/
[2025-12-11 10:30:58] [INFO] 配置完了:
-rwxr-xr-x 1 user user 45678 Dec 11 10:30 /opt/OpenTP1/aplib/myprogram

======================================
  OpenTP1 起動
======================================

[2025-12-11 10:30:59] [INFO] dcstart を実行中...
[2025-12-11 10:31:00] [INFO] 10秒待機中...
[2025-12-11 10:31:10] [INFO] 起動を確認中...
[2025-12-11 10:31:10] [INFO] OpenTP1の起動を確認しました

================================================================
  デプロイ完了
================================================================

実行結果    : 成功
終了日時    : 2025-12-11 10:31:10
ログファイル: opentp1_deploy_20251211_103045.log
```

---

## 設定項目

| 設定項目 | 説明 | 例 |
|---------|------|---|
| `OPENTP1_HOME` | OpenTP1のインストールパス | `/opt/OpenTP1` |
| `OPENTP1_BIN` | OpenTP1コマンドのパス | `${OPENTP1_HOME}/bin` |
| `SOURCE_FILE` | コピー元ファイル（フルパス） | `/home/user/src/myprogram` |
| `DEPLOY_DIR` | 配置先ディレクトリ | `/opt/OpenTP1/aplib` |
| `CREATE_BACKUP` | バックアップ作成の有無 | `true` |
| `STOP_WAIT_TIME` | 停止後の待機時間（秒） | `10` |
| `START_WAIT_TIME` | 起動後の待機時間（秒） | `10` |

---

## エラー時の動作

- **デプロイ失敗時**: OpenTP1を自動的に再起動してから終了
- **OpenTP1起動失敗時**: エラーメッセージを表示して終了

---

## ログファイル

実行ごとにログファイルが作成されます：

```
opentp1_deploy_YYYYMMDD_HHMMSS.log
```

エラーが発生した場合は、このログファイルで詳細を確認できます。

---

## トラブルシューティング

### dcstop/dcstartコマンドが見つからない

**原因**: `OPENTP1_BIN` のパスが間違っている

**対処法**:
```bash
# OpenTP1コマンドの場所を確認
which dcstop
which dcstart

# 設定を修正
OPENTP1_BIN="/actual/path/to/opentp1/bin"
```

---

### 権限エラー

**原因**: OpenTP1の操作権限がない

**対処法**:
```bash
# root または OpenTP1 管理ユーザーで実行
sudo ./opentp1_deploy.sh

# または適切なユーザーに切り替え
su - opentp1_user
./opentp1_deploy.sh
```

---

### 配置先への書き込み権限がない

**原因**: `DEPLOY_DIR` への書き込み権限がない

**対処法**:
```bash
# 権限を確認
ls -la /opt/OpenTP1/aplib/

# 必要に応じて権限を変更（管理者が実施）
chmod 775 /opt/OpenTP1/aplib/
```

---

## 注意事項

1. **本番環境での実行前に必ずテスト環境で動作確認してください**
2. OpenTP1の停止・起動には適切な権限が必要です
3. バックアップファイルは自動削除されません。定期的に整理してください

---

## 複数ファイル対応

複数のファイルをデプロイする場合：

```bash
# ファイルごとにスクリプトを作成
cp opentp1_deploy.sh deploy_program1.sh
cp opentp1_deploy.sh deploy_program2.sh

# それぞれの設定を編集
vim deploy_program1.sh  # SOURCE_FILE="/path/to/program1"
vim deploy_program2.sh  # SOURCE_FILE="/path/to/program2"
```

---

## ライセンス

このツールはMITライセンスの下で公開されています。
