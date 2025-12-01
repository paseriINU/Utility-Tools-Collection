# Git差分ファイル抽出ツール

## 概要

Gitの2つのブランチ（例: `main` と `develop`）間で差分があるファイルのみを抽出し、**フォルダ構造を保ったまま**別フォルダにコピーするツールです。

### 特徴

- ✅ **フォルダ構造を保持** - 元のディレクトリ構造のまま抽出
- ✅ **差分ファイルのみ抽出** - 変更・追加されたファイルだけをコピー
- ✅ **Windows標準機能のみ** - 追加ツール不要
- ✅ **簡単操作** - ダブルクリックで実行

### 使用例

```
main ブランチと develop ブランチを比較
↓
差分ファイルのみを抽出
↓
フォルダ構造を保ったまま diff_output/ にコピー
```

---

## 必要な環境

- Windows 10 / Windows 11
- Git for Windows がインストールされていること

---

## 使い方

### 方法1: 基本版（最も簡単）

1. **Gitリポジトリのルートフォルダ**にこのツールをコピー
   ```
   your-project/
   ├── .git/
   ├── src/
   ├── extract_diff.bat  ← ここに配置
   └── ...
   ```

2. **`extract_diff.bat`** を編集して設定を変更（必要に応じて）
   ```batch
   rem 比較元ブランチ（基準）
   set BASE_BRANCH=main

   rem 比較先ブランチ（差分を取得したいブランチ）
   set TARGET_BRANCH=develop

   rem 出力先フォルダ
   set OUTPUT_DIR=diff_output
   ```

3. **`extract_diff.bat`** をダブルクリックで実行

4. 完了！`diff_output/` フォルダに差分ファイルがコピーされます

---

### 方法2: 設定ファイル版

1. **`config.ini.sample`** を **`config.ini`** にコピー
   ```cmd
   copy config.ini.sample config.ini
   ```

2. **`config.ini`** を編集
   ```ini
   [Branches]
   BASE_BRANCH=main
   TARGET_BRANCH=develop

   [Output]
   OUTPUT_DIR=diff_output
   INCLUDE_DELETED=0
   ```

3. **`extract_diff_config.bat`** をダブルクリックで実行

---

## 実行例

### ケース1: main と develop の差分を抽出

```
リポジトリ構造:
your-project/
├── src/
│   ├── main.js          (変更あり)
│   ├── utils.js         (変更なし)
│   └── components/
│       └── Header.js    (新規追加)
├── docs/
│   └── README.md        (変更あり)
└── package.json         (変更なし)

実行後:
diff_output/
├── src/
│   ├── main.js          ← コピーされた
│   └── components/
│       └── Header.js    ← コピーされた
└── docs/
    └── README.md        ← コピーされた
```

**フォルダ構造が完全に保たれます！**

---

### 実行時の出力例

```
========================================
Git差分ファイル抽出ツール
========================================

リポジトリルート: C:\Projects\MyApp
比較元ブランチ  : main
比較先ブランチ  : develop
出力先フォルダ  : diff_output

差分ファイルを検出中...

検出された差分ファイル数: 15 個

ファイルをコピー中...

[コピー] src\main.js
[コピー] src\components\Header.js
[コピー] src\components\Footer.js
[コピー] docs\README.md
[コピー] config\app.config.js
...

========================================
処理完了
========================================

コピーしたファイル数: 15 個
出力先: diff_output

出力先フォルダを開きますか? (Y/N)
```

---

## パラメータ説明

### 基本設定

| パラメータ | 説明 | デフォルト値 | 例 |
|----------|------|------------|-----|
| `BASE_BRANCH` | 比較元ブランチ（基準） | `main` | `main`, `master`, `release` |
| `TARGET_BRANCH` | 比較先ブランチ（差分取得） | `develop` | `develop`, `feature/xxx` |
| `OUTPUT_DIR` | 出力先フォルダ | `diff_output` | `diff_output`, `C:\Deploy\Diff` |
| `INCLUDE_DELETED` | 削除されたファイルも含めるか | `0` | `0`（含めない）, `1`（含める） |

---

## よくある使用例

### 例1: リリース前の差分確認

```batch
set BASE_BRANCH=main
set TARGET_BRANCH=develop
set OUTPUT_DIR=release_diff
```

**用途**: 次回リリースで変更されるファイルを確認

---

### 例2: フィーチャーブランチの差分

```batch
set BASE_BRANCH=develop
set TARGET_BRANCH=feature/new-login
set OUTPUT_DIR=feature_diff
```

**用途**: 特定の機能追加で変更されたファイルを確認

---

### 例3: タグ間の差分

```batch
set BASE_BRANCH=v1.0.0
set TARGET_BRANCH=v2.0.0
set OUTPUT_DIR=v1_to_v2_diff
```

**用途**: バージョン間の変更ファイルを確認

---

### 例4: デプロイ用差分ファイル作成

```batch
set BASE_BRANCH=production
set TARGET_BRANCH=staging
set OUTPUT_DIR=C:\Deploy\差分ファイル_%DATE%
```

**用途**: 本番環境へのデプロイ用に差分ファイルを準備

---

## 応用例

### 差分ファイルをZIP圧縮

1. 差分ファイルを抽出
   ```cmd
   extract_diff.bat
   ```

2. PowerShellでZIP圧縮
   ```powershell
   Compress-Archive -Path "diff_output\*" -DestinationPath "diff_files.zip"
   ```

または、バッチファイルで自動化：

```batch
rem extract_and_zip.bat
call extract_diff.bat

powershell Compress-Archive -Path "diff_output\*" -DestinationPath "diff_files_%DATE:~0,4%%DATE:~5,2%%DATE:~8,2%.zip"

echo ZIP圧縮完了
```

---

### 差分ファイル一覧を出力

```batch
rem 差分ファイルリストをテキストファイルに保存
git diff --name-only main...develop > diff_list.txt
```

---

### 特定の拡張子のみ抽出

バッチファイルを改造：

```batch
rem .jsファイルのみ抽出
git diff --name-only --diff-filter=ACMR main...develop | findstr /i "\.js$" > "%TEMP_FILE%"

rem .cssと.scssのみ抽出
git diff --name-only --diff-filter=ACMR main...develop | findstr /i "\.css$ \.scss$" > "%TEMP_FILE%"
```

---

## トラブルシューティング

### エラー: "このフォルダはGitリポジトリではありません"

**原因**: `.git` フォルダがない場所で実行している

**対処法**:
- Gitリポジトリのルートフォルダで実行してください
- または、リポジトリをcloneしてください

---

### エラー: "ブランチが見つかりません"

**原因**: 指定したブランチが存在しない

**対処法**:

1. ブランチ一覧を確認
   ```cmd
   git branch -a
   ```

2. 正しいブランチ名を設定
   ```batch
   set BASE_BRANCH=main    ← 正しいブランチ名に変更
   set TARGET_BRANCH=develop
   ```

---

### 差分ファイルが見つからない

**原因**: 2つのブランチが同じ内容

**対処法**:
- ブランチを間違えていないか確認
- `git status` で現在のブランチを確認

---

### 日本語ファイル名が文字化けする

**原因**: Gitの文字コード設定

**対処法**:

```cmd
git config --global core.quotepath false
```

---

## Git差分オプションについて

### `--diff-filter` オプション

| オプション | 説明 |
|----------|------|
| `A` | Added（追加されたファイル） |
| `C` | Copied（コピーされたファイル） |
| `M` | Modified（変更されたファイル） |
| `R` | Renamed（名前変更されたファイル） |
| `D` | Deleted（削除されたファイル） |

**デフォルト設定**: `ACMR`（削除を除くすべて）

**削除も含める場合**: `ACMRD` または `--diff-filter` を指定しない

---

## 仕組み

```
1. git diff --name-only BASE_BRANCH...TARGET_BRANCH
   ↓ 差分ファイルのリストを取得

2. 各ファイルについて：
   ├─ ファイルが存在するか確認
   ├─ 出力先のディレクトリ構造を作成
   └─ ファイルをコピー

3. 完了
   └─ 出力先フォルダを開く（オプション）
```

---

## 注意事項

### 1. 作業ディレクトリの変更は含まれない

このツールはブランチ間の差分のみを抽出します。
現在の作業ディレクトリで変更したがコミットしていないファイルは含まれません。

### 2. バイナリファイルも抽出される

画像・PDFなどのバイナリファイルも差分があればコピーされます。

### 3. 出力先フォルダは上書きされる

既存の `diff_output/` フォルダがある場合、確認後に削除されます。
重要なファイルは事前にバックアップしてください。

### 4. Gitサブモジュールは未対応

サブモジュール内の差分は抽出されません。

---

## カスタマイズ例

### 日付付きフォルダに出力

```batch
rem 現在の日時をフォルダ名に含める
set OUTPUT_DIR=diff_%DATE:~0,4%%DATE:~5,2%%DATE:~8,2%_%TIME:~0,2%%TIME:~3,2%
```

結果: `diff_20251201_1530/`

---

### 複数ブランチの差分を一括抽出

```batch
rem extract_multiple.bat
call extract_diff.bat
ren diff_output diff_main_to_develop

set BASE_BRANCH=develop
set TARGET_BRANCH=feature/new-ui
call extract_diff.bat
ren diff_output diff_develop_to_feature
```

---

### 除外ファイルパターンを追加

特定のファイルを除外したい場合：

```batch
rem node_modules や .lock ファイルを除外
git diff --name-only main...develop | findstr /v "node_modules package-lock.json" > "%TEMP_FILE%"
```

---

## ライセンス

このツールはMITライセンスの下で公開されています。

---

## 関連コマンド

### Git差分確認コマンド

```bash
# ファイル一覧のみ表示
git diff --name-only main...develop

# ファイルごとの変更内容も表示
git diff main...develop

# 統計情報を表示
git diff --stat main...develop

# 追加・削除された行数を表示
git diff --numstat main...develop
```

---

**作成日**: 2025-12-01
**バージョン**: 1.0
