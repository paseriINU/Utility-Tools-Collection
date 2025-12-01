# JP1ジョブネット起動ツール - ajsentryコマンド版

JP1の標準コマンド`ajsentry`を使用してジョブネットを起動するツールです。

---

## 📋 概要

このツールは、ローカルPCにインストールされたJP1/AJS3の`ajsentry`コマンドを使用して、
リモートのJP1サーバ上のジョブネットを起動します。

---

## ✅ 必要な環境

### ローカルPC（実行する側）

- Windows OS
- **JP1/AJS3 - View** または **JP1/AJS3 - Manager** がインストール済み
- `ajsentry`コマンドが使用可能（環境変数PATHに含まれている）

### JP1サーバ（JP1が稼働している側）

- JP1/AJS3 - Manager が稼働中
- ジョブネットが登録済み
- ネットワーク経由でアクセス可能

---

## 🚀 使い方

### 方法1: 直接編集版（jp1_start_job.bat）

#### 1. バッチファイルを編集

`jp1_start_job.bat`をテキストエディタで開き、以下の設定項目を編集します：

```batch
rem JP1/AJS3のホスト名またはIPアドレス
set JP1_HOST=192.168.1.100

rem JP1ユーザー名
set JP1_USER=jp1admin

rem JP1パスワード（空の場合は実行時に入力を求めます）
set JP1_PASSWORD=

rem 起動するジョブネットのフルパス
set JOBNET_PATH=/main_unit/jobgroup1/daily_batch

rem ajsentryコマンドのパス
set AJSENTRY_CMD=ajsentry
```

#### 2. 実行

`jp1_start_job.bat`をダブルクリックまたはコマンドプロンプトから実行：

```cmd
jp1_start_job.bat
```

---

### 方法2: 設定ファイル版（jp1_start_job_config.bat）

#### 1. 設定ファイルを作成

`config.ini.sample`を`config.ini`にコピーして編集：

```cmd
copy config.ini.sample config.ini
notepad config.ini
```

**config.ini の内容**:

```ini
[JP1]
JP1_HOST=192.168.1.100
JP1_USER=jp1admin
JP1_PASSWORD=
JOBNET_PATH=/main_unit/jobgroup1/daily_batch
AJSENTRY_CMD=ajsentry
```

#### 2. 実行

`jp1_start_job_config.bat`をダブルクリックまたはコマンドプロンプトから実行：

```cmd
jp1_start_job_config.bat
```

---

## 📖 実行例

### 成功時の出力

```
========================================
JP1ジョブネット起動ツール
（ajsentryコマンド版）
========================================

JP1ホスト      : 192.168.1.100
JP1ユーザー    : jp1admin
ジョブネットパス: /main_unit/jobgroup1/daily_batch

ジョブネットを起動しますか？
実行する場合はYを押してください [Y,N]?Y

========================================
ジョブネット起動中...
========================================

KAVS1820-I ajsentryコマンドが正常終了しました。

========================================
ジョブネットの起動に成功しました
========================================

ジョブネット: /main_unit/jobgroup1/daily_batch
ホスト      : 192.168.1.100

続行するには何かキーを押してください . . .
```

### パスワード入力が必要な場合

```
[注意] パスワードが設定されていません。
JP1パスワードを入力してください: ********
```

---

## ⚙️ 設定項目の詳細

| 設定項目 | 説明 | 例 |
|---------|------|---|
| JP1_HOST | JP1サーバのホスト名またはIPアドレス | `192.168.1.100` |
| JP1_USER | JP1ユーザー名 | `jp1admin` |
| JP1_PASSWORD | JP1パスワード（空の場合は実行時入力） | ` `（空推奨） |
| JOBNET_PATH | ジョブネットのフルパス | `/main_unit/jobgroup1/daily_batch` |
| AJSENTRY_CMD | ajsentryコマンドのパス | `ajsentry`または`C:\...\ajsentry.exe` |

---

## 🔧 ajsentryコマンドの確認方法

### コマンドが使用可能か確認

コマンドプロンプトで以下を実行：

```cmd
where ajsentry
```

**出力例（成功）**:
```
C:\Program Files\HITACHI\JP1AJS3\bin\ajsentry.exe
```

**出力例（失敗）**:
```
情報: 与えられたパターンのファイルが見つかりませんでした。
```

### PATHに追加する場合

JP1/AJS3のインストールディレクトリをPATHに追加：

```
C:\Program Files\HITACHI\JP1AJS3\bin
```

**手順**:
1. システムのプロパティ → 環境変数
2. システム環境変数の「Path」を編集
3. 上記パスを追加
4. コマンドプロンプトを再起動

---

## ⚠️ 注意事項

### セキュリティ

- **パスワード**: バッチファイルやconfig.iniにパスワードを記載しないことを推奨
- **ファイル保護**: config.iniは`.gitignore`に追加してコミットしない

### JP1環境

- **実行権限**: JP1ユーザーにジョブネット実行権限が必要
- **ジョブネット状態**: 既に実行中の場合の動作に注意
- **本番環境**: 本番実行前に必ずテスト環境で確認

---

## 🐛 トラブルシューティング

### エラー: 「ajsentryコマンドが見つかりません」

**原因**:
- JP1/AJS3がインストールされていない
- 環境変数PATHにajsentryのパスが含まれていない

**対処法**:
1. JP1/AJS3 - Viewまたは- Managerをインストール
2. 環境変数PATHに`C:\Program Files\HITACHI\JP1AJS3\bin`を追加
3. またはバッチファイルで`AJSENTRY_CMD`にフルパスを指定：
   ```batch
   set AJSENTRY_CMD=C:\Program Files\HITACHI\JP1AJS3\bin\ajsentry.exe
   ```

---

### エラー: 「認証に失敗しました」

**原因**:
- JP1ユーザー名またはパスワードが間違っている
- JP1サーバに接続できない

**対処法**:
1. JP1ユーザー名、パスワードを確認
2. JP1サーバのホスト名/IPアドレスを確認
3. ネットワーク接続を確認（pingコマンドで確認）
   ```cmd
   ping 192.168.1.100
   ```

---

### エラー: 「ジョブネットが見つかりません」

**原因**:
- ジョブネットパスが間違っている
- ジョブネットが登録されていない

**対処法**:
1. JP1/AJS3 - Viewでジョブネットパスを確認
2. パスの先頭に`/`があることを確認
3. 大文字小文字を正確に記載

**正しいパスの例**:
```
/main_unit/jobgroup1/daily_batch
```

---

### エラーコード一覧

ajsentryコマンドの主なエラーコード：

| エラーコード | 意味 | 対処法 |
|------------|------|-------|
| KAVS1821-E | ジョブネットが見つからない | ジョブネットパスを確認 |
| KAVS1822-E | 認証エラー | ユーザー名、パスワードを確認 |
| KAVS1823-E | 接続エラー | ホスト名、ネットワークを確認 |
| KAVS1824-E | 実行権限なし | JP1ユーザーの権限を確認 |

詳細はJP1/AJS3のマニュアルを参照してください。

---

## 💡 応用例

### 複数のジョブネットを順次起動

複数のバッチファイルを作成し、それぞれ異なるジョブネットを設定：

```
jp1_start_morning.bat   → /main_unit/jobs/morning_batch
jp1_start_noon.bat      → /main_unit/jobs/noon_batch
jp1_start_evening.bat   → /main_unit/jobs/evening_batch
```

### タスクスケジューラで定期実行

Windows タスクスケジューラに登録して自動実行：

1. タスクスケジューラを開く
2. 「基本タスクの作成」
3. トリガー：毎日 6:00
4. 操作：プログラムの開始
5. プログラム：`jp1_start_job_config.bat`のフルパス

---

## 📚 参考資料

- [JP1/Automatic Job Management System 3 コマンドリファレンス](https://www.hitachi.co.jp/Prod/comp/soft1/manual/pc/d3K2211/AJSK01/EU210049.HTM)
- [ajsentryコマンド](https://www.hitachi.co.jp/Prod/comp/soft1/manual/pc/d3K2211/AJSK01/EU210278.HTM)

---

**更新日**: 2025-12-02
