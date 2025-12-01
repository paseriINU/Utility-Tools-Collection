# Gitツール 共通仕様

## 📍 固定Gitプロジェクトパス

すべてのGit関連バッチファイルは、以下の固定パスを使用します：

```
C:\Users\%username%\source\Git\project\
```

### 利点

- ✅ **どこからでも実行可能** - バッチファイルをどこに置いても動作
- ✅ **デスクトップから実行** - デスクトップにショートカットを置いて実行可能
- ✅ **統一された動作** - すべてのGitツールが同じリポジトリを操作

---

## 🚀 使い方

### 前提条件

Gitリポジトリが以下のパスに存在すること：

```
C:\Users\<あなたのユーザー名>\source\Git\project\
```

例：
```
C:\Users\yamada\source\Git\project\
C:\Users\tanaka\source\Git\project\
```

### 実行方法

#### 方法1: どこからでもダブルクリック

```
デスクトップ\
├── extract_diff.bat          ← ダブルクリック
└── delete_branches.bat       ← ダブルクリック
```

**動作**:
1. バッチファイルを実行
2. 自動的に `C:\Users\%username%\source\Git\project\` に移動
3. Git操作を実行

---

#### 方法2: ショートカットをデスクトップに配置

```
1. バッチファイルを右クリック
2. 「送る」→「デスクトップ（ショートカットを作成）」
3. デスクトップのショートカットから実行
```

---

#### 方法3: 任意のフォルダに配置

```
C:\Tools\GitScripts\
├── extract_diff.bat
├── delete_branches.bat
└── ...

→ どこに置いても C:\Users\%username%\source\Git\project\ で実行される
```

---

## 📋 対象ツール

以下のツールが固定パスを使用します：

### Git差分ファイル抽出ツール

- `batch/git-diff-extract/extract_diff.bat`
- `batch/git-diff-extract/extract_diff_config.bat`

**動作確認例**:
```
========================================
Git差分ファイル抽出ツール
========================================

Gitプロジェクトパス: C:\Users\yamada\source\Git\project

リポジトリルート: C:\Users\yamada\source\Git\project
比較元ブランチ  : main
比較先ブランチ  : develop
...
```

---

### Gitブランチ管理ツール

- `batch/git-branch-manager/delete_branches.bat`
- `batch/git-branch-manager/delete_remote_branch.bat`
- `batch/git-branch-manager/delete_local_branch.bat`

**動作確認例**:
```
========================================
Gitブランチ削除ツール
========================================

Gitプロジェクトパス: C:\Users\yamada\source\Git\project

[1] リモートブランチを削除
[2] ローカルブランチを削除
...
```

---

## 🔧 パスのカスタマイズ

デフォルトパス `C:\Users\%username%\source\Git\project\` を変更したい場合：

### 全ツール共通で変更する場合

各バッチファイルの以下の行を編集：

```batch
rem 変更前
set GIT_PROJECT_PATH=C:\Users\%username%\source\Git\project

rem 変更後（例）
set GIT_PROJECT_PATH=D:\MyProjects\MainProject
```

**対象ファイル**:
- `git-diff-extract/extract_diff.bat`
- `git-diff-extract/extract_diff_config.bat`
- `git-branch-manager/delete_branches.bat`
- `git-branch-manager/delete_remote_branch.bat`
- `git-branch-manager/delete_local_branch.bat`

---

## ⚠️ トラブルシューティング

### エラー: "Gitプロジェクトフォルダが見つかりません"

**原因**: 指定されたパスにフォルダが存在しない

**対処法1**: フォルダを作成

```cmd
mkdir C:\Users\%username%\source\Git\project
```

**対処法2**: 既存のGitリポジトリを移動

```cmd
# 既存のリポジトリを移動
move C:\Projects\MyProject C:\Users\%username%\source\Git\project
```

**対処法3**: シンボリックリンクを作成

既存のリポジトリを残したまま、固定パスからリンク：

```cmd
# 管理者権限のコマンドプロンプトで実行
mklink /D C:\Users\%username%\source\Git\project D:\RealProject
```

---

### エラー: "このフォルダはGitリポジトリではありません"

**原因**: 指定されたパスに `.git` フォルダが存在しない

**対処法**: Gitリポジトリを初期化またはクローン

```bash
# 新規リポジトリを初期化
cd C:\Users\%username%\source\Git\project
git init

# または既存リポジトリをクローン
cd C:\Users\%username%\source\Git
git clone https://github.com/user/repo.git project
```

---

## 💡 活用例

### 例1: デスクトップから差分抽出

```
1. extract_diff.bat をデスクトップにコピー
2. ダブルクリック
3. 自動的に C:\Users\%username%\source\Git\project で差分抽出
4. diff_output フォルダがデスクトップに作成される
```

---

### 例2: タスクバーにピン留め

```
1. バッチファイルのショートカットを作成
2. タスクバーにピン留め
3. タスクバーから1クリックで実行
```

---

### 例3: 複数のGitプロジェクトがある場合

プロジェクトごとにバッチファイルをコピーしてパスを変更：

```
extract_diff_ProjectA.bat → set GIT_PROJECT_PATH=C:\Projects\ProjectA
extract_diff_ProjectB.bat → set GIT_PROJECT_PATH=C:\Projects\ProjectB
extract_diff_Main.bat     → set GIT_PROJECT_PATH=C:\Users\%username%\source\Git\project
```

---

## 📊 実行フロー

```
[バッチファイル実行]
    ↓
[固定パスに移動]
C:\Users\%username%\source\Git\project
    ↓
[パス存在確認]
    ↓ YES
[Gitリポジトリ確認]
    ↓ YES
[Git操作実行]
    ↓
[完了]
```

---

## 🎯 推奨構成

### 推奨フォルダ構成

```
C:\Users\<username>\
└── source\
    └── Git\
        └── project\              ← メインのGitリポジトリ
            ├── .git\
            ├── src\
            ├── docs\
            └── ...
```

### バッチファイルの配置

```
# パターン1: デスクトップに配置
C:\Users\<username>\Desktop\
├── extract_diff.bat
├── delete_branches.bat
└── ...

# パターン2: 専用フォルダに配置
C:\Tools\GitScripts\
├── extract_diff.bat
├── delete_branches.bat
└── ...

# パターン3: スタートメニューに登録
C:\Users\<username>\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\
└── GitTools\
    ├── Git差分抽出.lnk
    └── Gitブランチ削除.lnk
```

---

## まとめ

### 重要ポイント

1. **固定パス**: すべてのGitツールは `C:\Users\%username%\source\Git\project\` を使用
2. **どこからでも実行可能**: バッチファイルをどこに置いても動作
3. **自動移動**: 実行時に自動的にGitリポジトリに移動

### メリット

- ✅ バッチファイルの配置場所を気にしなくて良い
- ✅ デスクトップやツールフォルダから実行可能
- ✅ ショートカットやタスクバーピン留めが簡単
- ✅ 統一された動作で管理しやすい

---

**更新日**: 2025-12-01
