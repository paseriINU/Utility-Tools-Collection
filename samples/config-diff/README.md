# 設定ファイル比較ツール (config-diff)

環境間の設定ファイルを比較し、差分を検出するツールです。

## 機能

- 2つのフォルダ内の設定ファイルを再帰的に比較
- 対応形式: `.ini`, `.json`, `.xml`, `.properties`, `.conf`, `.cfg`, `.yaml`, `.yml`
- 差分結果をテキストファイルに出力
- WinMerge連携で詳細比較

## 使い方

1. `config-diff.bat` の設定セクションを編集:
   ```powershell
   $Config = @{
       SourceFolder = "C:\Config\Production"   # 比較元（本番など）
       TargetFolder = "C:\Config\Development"  # 比較先（開発など）
   }
   ```

2. ダブルクリックで実行

## 出力例

```
========================================================================
 比較結果
========================================================================

同一ファイル      : 15 件
差分あり          : 3 件
比較元のみ        : 1 件
比較先のみ        : 2 件

--- 差分のあるファイル ---
  [差分] app\settings.json
  [差分] db\connection.properties
  [差分] web\config.xml
```

## 設定オプション

| 設定項目 | 説明 | デフォルト |
|----------|------|------------|
| `SourceFolder` | 比較元フォルダ | - |
| `TargetFolder` | 比較先フォルダ | - |
| `Extensions` | 対象拡張子 | 設定ファイル系 |
| `IgnorePatterns` | 無視するファイル名パターン | `.bak`, `.backup`, `~` |
| `ExportResult` | 結果ファイル出力 | `$true` |
| `OutputFolder` | 出力先フォルダ | デスクトップ |

## 動作環境

- Windows 10/11
- PowerShell 5.1以降
- WinMerge（オプション、詳細比較用）
