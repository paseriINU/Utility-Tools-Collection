# C言語 メモリ安全性・デッドロジック チェックツール

C言語ソースファイルのメモリ安全性とデッドロジック（到達不能コード、未使用コード等）を検出するツールです。

## 機能

### 検出できる問題

| カテゴリ | 検出項目 | 検出方法 |
|----------|----------|----------|
| **メモリ安全性** | バッファオーバーフロー | AddressSanitizer, cppcheck |
| | 解放済みメモリアクセス（use-after-free） | AddressSanitizer, Valgrind |
| | メモリリーク | Valgrind, AddressSanitizer |
| | 未初期化変数の使用 | GCC警告, cppcheck |
| | NULLポインタ参照 | cppcheck, AddressSanitizer |
| | スタックバッファオーバーフロー | AddressSanitizer |
| **デッドロジック** | 到達不能コード | GCC警告, cppcheck |
| | 未使用変数 | GCC警告, cppcheck |
| | 未使用関数 | GCC警告, cppcheck |
| | 未使用パラメータ | GCC警告 |

### チェックフロー

```
1. GCC警告チェック
   └─ 厳格な警告オプションでコンパイル（構文チェック）

2. cppcheck静的解析（オプション）
   └─ メモリリーク、デッドコード等の詳細解析

3. AddressSanitizerチェック
   └─ 実行時のメモリエラー検出

4. Valgrindチェック（オプション）
   └─ 詳細なメモリリーク検出
```

## 必要な環境

### 必須

- **GCC** (gcc)
  - AddressSanitizer対応版（GCC 4.8以降）

### オプション（インストール推奨）

- **cppcheck** - 静的解析の強化
  ```bash
  # Ubuntu/Debian
  sudo apt install cppcheck

  # RHEL/CentOS
  sudo yum install cppcheck
  ```

- **valgrind** - 動的メモリチェック
  ```bash
  # Ubuntu/Debian
  sudo apt install valgrind

  # RHEL/CentOS
  sudo yum install valgrind
  ```

## インストール

```bash
# リポジトリをクローン（または直接ダウンロード）
git clone <repository-url>

# 実行権限を付与
chmod +x linux/c-memory-safety-check/c-memory-safety-check.sh

# パスを通す（オプション）
sudo ln -s $(pwd)/linux/c-memory-safety-check/c-memory-safety-check.sh /usr/local/bin/c-memory-check
```

## 使い方

### 基本的な使用方法

```bash
./c-memory-safety-check.sh <source.c>
```

### インクルードパスを指定

別ディレクトリのヘッダーファイルを参照する場合：

```bash
# 単一のインクルードパス
./c-memory-safety-check.sh main.c -I ./include

# 複数のインクルードパス
./c-memory-safety-check.sh main.c -I ./include -I ./lib -I ../common
```

### テスト実行時に引数を渡す

```bash
# -- 以降がテスト実行時の引数になります
./c-memory-safety-check.sh myprogram.c -- arg1 arg2
```

### インクルードパスとテスト引数の両方を指定

```bash
./c-memory-safety-check.sh main.c -I ./include -- input.txt output.txt
```

### ヘルプ表示

```bash
./c-memory-safety-check.sh --help
```

## 実行例

### サンプルコード（問題あり）

```c
// test_memory.c
#include <stdio.h>
#include <stdlib.h>

void unused_function() {  // デッドコード
    printf("This is never called\n");
}

int main() {
    int *ptr = malloc(sizeof(int) * 10);

    // バッファオーバーフロー
    ptr[10] = 100;

    // メモリリーク（freeしていない）
    return 0;
}
```

### 実行結果

```
============================================
 C言語 メモリ安全性チェックツール v1.1.0
============================================

チェック対象: test_memory.c
インクルードパス: ./include ./lib  (指定した場合)

========================================
 環境チェック
========================================
[OK] GCC: gcc (Ubuntu 11.4.0-1ubuntu1~22.04) 11.4.0
[OK] cppcheck: Cppcheck 2.7
[OK] valgrind: valgrind-3.18.1

========================================
 GCC警告チェック
========================================
[WARNING] GCC: 警告 1件

========================================
 cppcheck静的解析
========================================
[WARNING] cppcheck: 2件の問題を検出

========================================
 AddressSanitizer チェック
========================================
[INFO] AddressSanitizerでビルド中...
[OK] ビルド成功
[INFO] テスト実行中...
[ERROR] AddressSanitizer: 1件のメモリエラーを検出

========================================
 Valgrind メモリチェック
========================================
[INFO] デバッグビルド中...
[OK] ビルド成功
[INFO] Valgrind実行中...
[WARNING] Valgrind: メモリリークを検出

========================================
 チェック結果サマリー
========================================

チェック対象: test_memory.c
レポート: ./c-check-results/test_memory_report.txt

エラー: 1件
警告: 3件
メモリ問題: 2件
デッドコード: 1件

[結果] 修正が必要な問題が検出されました

[INFO] 詳細は ./c-check-results/test_memory_report.txt を参照してください
```

## 出力ファイル

チェック結果は以下に保存されます：

```
./c-check-results/
└── <ソース名>_report.txt    # 詳細レポート
```

## チェック項目の詳細

### 1. GCC警告オプション

以下の警告オプションを使用：

| オプション | 検出内容 |
|------------|----------|
| `-Wall` | 一般的な警告 |
| `-Wextra` | 追加の警告 |
| `-Wuninitialized` | 未初期化変数 |
| `-Wunused` | 未使用の変数・関数 |
| `-Wunreachable-code` | 到達不能コード |
| `-Wformat=2` | フォーマット文字列の問題 |
| `-Warray-bounds=2` | 配列境界チェック |
| `-Wnull-dereference` | NULLポインタ参照 |
| `-Wshadow` | 変数のシャドウイング |

### 2. cppcheck

| チェック項目 | 説明 |
|--------------|------|
| `memleak` | メモリリーク |
| `nullPointer` | NULLポインタ参照 |
| `uninitvar` | 未初期化変数 |
| `bufferAccessOutOfBounds` | バッファオーバーフロー |
| `unreachableCode` | 到達不能コード |
| `unusedFunction` | 未使用関数 |

### 3. AddressSanitizer

| 検出項目 | 説明 |
|----------|------|
| heap-buffer-overflow | ヒープバッファオーバーフロー |
| stack-buffer-overflow | スタックバッファオーバーフロー |
| use-after-free | 解放済みメモリアクセス |
| double-free | 二重解放 |
| memory-leak | メモリリーク |

### 4. Valgrind

| 検出項目 | 説明 |
|----------|------|
| Invalid read/write | 不正なメモリアクセス |
| Conditional jump on uninitialised value | 未初期化値の条件判定使用 |
| definitely lost | 確実なメモリリーク |
| indirectly lost | 間接的なメモリリーク |
| possibly lost | 可能性のあるメモリリーク |

## 終了コード

| コード | 意味 |
|--------|------|
| 0 | エラーなし（警告のみ、または問題なし） |
| 1 | エラーまたはメモリ問題を検出 |

## 注意事項

- **動的チェック（AddressSanitizer, Valgrind）では実際にプログラムが実行されます**
  - テスト用の入力データを準備してください
  - 危険な操作（ファイル削除等）を含むプログラムには注意してください

- **複数ファイルプロジェクト**
  - 現在は単一ファイルのチェックに対応
  - 複数ファイルの場合はメインファイルを指定（リンクエラーが発生する可能性あり）

- **cppcheck/valgrindがない環境**
  - GCCとAddressSanitizerのみでも基本的なチェックは可能
  - より詳細な解析にはcppcheck/valgrindのインストールを推奨

## トラブルシューティング

### AddressSanitizerでビルドエラーが発生する

```
# GCCのバージョンを確認（4.8以降が必要）
gcc --version

# libasan がインストールされているか確認
ldconfig -p | grep asan
```

### Valgrindで「Permission denied」エラー

```
# 実行権限を確認
chmod +x ./c-check-results/*_test
```

### 日本語が文字化けする

```
# ロケールを設定
export LANG=ja_JP.UTF-8
```

## ライセンス

MIT License
