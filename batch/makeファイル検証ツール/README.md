# makeファイル検証ツール

フォルダ内のmakeファイル（.mk）をスキャンし、各makeファイル内の.oファイル参照がmakeファイル名と一致しているかを検証するツールです。

## 機能

- フォルダ内のすべての.mkファイルをスキャン（**サブフォルダも含む**）
- 各makeファイル内の.oファイル参照を抽出（パス付きでもファイル名部分を認識）
- makeファイル名と.oファイル名の一致を検証
- 例外ファイルの指定（makeファイル名と異なっていてもOK）
- 部分一致モードのサポート

## 必要な環境

- Windows 10以降
- PowerShell 5.1以降（Windows標準搭載）

## 使い方

### 基本的な使い方

1. スクリプトをダブルクリックで実行
2. 同じフォルダ内の.mkファイルが検証される
3. 結果がコンソールに表示される

### 設定のカスタマイズ

スクリプト内の「設定」セクションを編集してください:

```powershell
#region 設定
# 対象フォルダ（空欄の場合はスクリプトと同じフォルダ）
$TARGET_FOLDER = ""

# 例外ファイル（makeファイル名と異なっていてもOKな.oファイル）
# 拡張子なしで指定（例: "common" は common.o にマッチ）
$EXCEPTION_FILES = @(
    "common",
    "util",
    "shared"
)

# 部分一致を許可するか（library.mk に library_sub.o があってもOK）
$ALLOW_PARTIAL_MATCH = $true
#endregion
```

## 設定項目

| 設定項目 | 説明 | 初期値 |
|----------|------|--------|
| `$TARGET_FOLDER` | 検証対象フォルダのパス（空欄=スクリプトと同じフォルダ） | `""` |
| `$EXCEPTION_FILES` | 例外ファイルリスト（.o拡張子なしで指定） | `@("common", "util", "shared")` |
| `$ALLOW_PARTIAL_MATCH` | 部分一致を許可するか | `$true` |

## 一致判定ロジック

### 完全一致モード（`$ALLOW_PARTIAL_MATCH = $false`）
- `myprogram.mk` と `myprogram.o` → OK
- `myprogram.mk` と `myprogram_sub.o` → NG

### 部分一致モード（`$ALLOW_PARTIAL_MATCH = $true`）
- `myprogram.mk` と `myprogram.o` → OK
- `myprogram.mk` と `myprogram_sub.o` → OK（makeファイル名が含まれている）
- `myprogram.mk` と `otherfile.o` → NG

## 出力例

```
================================================================
  makeファイル .o 検証ツール
================================================================

対象フォルダ: C:\projects\src
例外ファイル: common, util, shared
部分一致許可: はい

----------------------------------------
[検証中] myprogram.mk
  [OK] myprogram.o
  [NG] otherfile.o
  [除外] common.o

----------------------------------------
[検証中] subdir\library.mk
  [OK] library.o
  [OK] library_sub.o

----------------------------------------
[検証中] subdir\module\helper.mk
  [OK] helper.o

========================================
検証結果サマリー
========================================

検証makeファイル数: 3
検証.oファイル数:   6
  OK:    4
  NG:    1
  除外:  1

不一致一覧:
  myprogram.mk:
    - otherfile.o
```

## 対応する.oファイル形式

以下のような様々な形式の.oファイル参照を認識します:

| 形式 | 例 | 抽出されるファイル名 |
|------|-----|---------------------|
| 単純 | `myprogram.o` | myprogram |
| パス付き | `$(OBJ_DIR)/myprogram.o` | myprogram |
| 相対パス | `../obj/myprogram.o` | myprogram |
| 絶対パス | `/home/user/obj/myprogram.o` | myprogram |

## 終了コード

| コード | 説明 |
|--------|------|
| 0 | すべてのファイルが一致 |
| 1 | 不一致があった |

## ライセンス

このツールは個人利用を目的としています。
