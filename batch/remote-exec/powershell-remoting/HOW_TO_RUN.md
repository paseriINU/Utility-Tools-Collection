# PowerShellスクリプトをダブルクリックで実行する方法

## 概要

デフォルトでは、PowerShellスクリプト（.ps1）をダブルクリックすると**メモ帳で開いてしまいます**。
ここでは、ダブルクリックで実行する複数の方法を紹介します。

---

## 方法1: バッチファイル経由（最も簡単）✅ 推奨

すでに用意されている**バッチファイル**を使用する方法です。

### 使い方

1. `remote_exec_ps.bat` または `remote_exec_ps_config.bat` を編集
2. **ダブルクリックで実行**

```cmd
remote_exec_ps.bat  ← これをダブルクリック
```

**メリット:**
- ✅ 追加設定不要
- ✅ すぐに使える
- ✅ 実行ポリシーを気にしなくて良い

**デメリット:**
- ❌ バッチファイル経由なので少し間接的

---

## 方法2: ショートカットを作成（簡単）✅ 推奨

PowerShellスクリプトを実行するショートカットを作成する方法です。

### 手順

#### ステップ1: ショートカットを作成

1. デスクトップまたは任意のフォルダで**右クリック**
2. 「新規作成」→「ショートカット」を選択
3. 以下のコマンドを入力：

```
powershell.exe -ExecutionPolicy Bypass -File "C:\Tools\remote-exec\powershell-remoting\Invoke-RemoteBatch.ps1" -ComputerName "192.168.1.100" -UserName "Administrator" -BatchPath "C:\Scripts\backup.bat"
```

4. ショートカット名を入力（例: 「リモートバックアップ実行」）
5. 完成

#### ステップ2: ショートカットをダブルクリック

作成したショートカットをダブルクリックすると実行されます。

### パラメータ付きショートカットの例

#### 例1: 基本的な実行
```
powershell.exe -ExecutionPolicy Bypass -File "C:\Tools\Invoke-RemoteBatch.ps1" -ComputerName "192.168.1.100" -UserName "Administrator" -BatchPath "C:\Scripts\test.bat"
```

#### 例2: ログファイル保存
```
powershell.exe -ExecutionPolicy Bypass -File "C:\Tools\Invoke-RemoteBatch.ps1" -ComputerName "192.168.1.100" -UserName "Administrator" -BatchPath "C:\Scripts\backup.bat" -OutputLog "C:\Logs\backup.log"
```

#### 例3: HTTPS使用
```
powershell.exe -ExecutionPolicy Bypass -File "C:\Tools\Invoke-RemoteBatch.ps1" -ComputerName "server.example.com" -UserName "Administrator" -BatchPath "C:\Scripts\secure.bat" -UseSSL
```

#### 例4: ウィンドウを非表示にする
```
powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File "C:\Tools\Invoke-RemoteBatch.ps1" -ComputerName "192.168.1.100" -UserName "Administrator" -BatchPath "C:\Scripts\silent.bat"
```

**メリット:**
- ✅ パラメータを事前設定できる
- ✅ 複数のショートカットを作れる（用途別）
- ✅ デスクトップから簡単に実行

**デメリット:**
- ❌ フルパスを指定する必要がある

---

## 方法3: .ps1ファイルの関連付けを変更（上級者向け）⚠️

PowerShellスクリプト（.ps1）のデフォルト動作を「実行」に変更する方法です。

### ⚠️ 注意

この方法は**セキュリティリスク**があります：
- すべての.ps1ファイルがダブルクリックで実行されるようになる
- 悪意のあるスクリプトを誤って実行してしまう可能性がある

**推奨しません**が、参考として記載します。

### 手順（レジストリ編集）

1. **レジストリエディタを開く**
   - `Win + R` → `regedit` と入力

2. **以下のキーに移動**
   ```
   HKEY_CLASSES_ROOT\Microsoft.PowerShellScript.1\Shell
   ```

3. **デフォルト値を変更**
   - 現在: `Edit` または `Open`
   - 変更後: `0` （実行を意味する）

4. **Windowsを再起動**

### 元に戻す方法

デフォルト値を `Edit` または `Open` に戻す

---

## 方法4: VBScriptラッパーを使用（クリック実行）✅ 推奨

VBScriptを使ってPowerShellを実行する方法です。ウィンドウを表示せずに実行できます。

### 手順

#### ステップ1: VBScriptファイルを作成

`run_invoke_remotebatch.vbs` という名前で以下の内容を保存：

```vbscript
' PowerShellスクリプトをバックグラウンドで実行
Dim objShell, command

Set objShell = CreateObject("WScript.Shell")

' 設定項目
computerName = "192.168.1.100"
userName = "Administrator"
batchPath = "C:\Scripts\backup.bat"
outputLog = "C:\Logs\backup.log"

' PowerShellスクリプトのパス（カレントディレクトリ）
scriptPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName) & "\Invoke-RemoteBatch.ps1"

' コマンドを構築
command = "powershell.exe -ExecutionPolicy Bypass -File """ & scriptPath & """ -ComputerName " & computerName & " -UserName " & userName & " -BatchPath """ & batchPath & """ -OutputLog """ & outputLog & """"

' 実行（ウィンドウ非表示: 0, 表示: 1）
objShell.Run command, 1, True

Set objShell = Nothing

MsgBox "リモートバッチ実行が完了しました。", vbInformation, "実行完了"
```

#### ステップ2: VBScriptをダブルクリック

`run_invoke_remotebatch.vbs` をダブルクリックすると実行されます。

**メリット:**
- ✅ ダブルクリックで実行可能
- ✅ ウィンドウの表示/非表示を制御できる
- ✅ 実行完了を通知できる

**デメリット:**
- ❌ VBScriptファイルを別途作成する必要がある

---

## 方法5: 専用のランチャーバッチを作成（簡易版）✅ 推奨

すでに用意されている方法ですが、より詳しく説明します。

### すでに用意されているファイル

- `remote_exec_ps.bat` - 基本版
- `remote_exec_ps_config.bat` - 設定ファイル版

### カスタムランチャーの作成例

特定のタスク専用のバッチを作成する例：

#### 例: バックアップ実行専用バッチ

`backup_server01.bat` を作成：

```batch
@echo off
echo ========================================
echo サーバ01のバックアップを実行します
echo ========================================
echo.

powershell.exe -ExecutionPolicy Bypass -File "%~dp0Invoke-RemoteBatch.ps1" -ComputerName "192.168.1.100" -UserName "Administrator" -BatchPath "C:\Scripts\backup.bat" -OutputLog "%~dp0backup_server01.log"

if errorlevel 1 (
    echo.
    echo [エラー] バックアップに失敗しました。
    pause
    exit /b 1
)

echo.
echo バックアップが完了しました。
pause
```

このバッチをダブルクリックするだけで実行できます。

---

## 比較表

| 方法 | 簡単さ | 安全性 | カスタマイズ性 | 推奨度 |
|-----|--------|--------|--------------|--------|
| **バッチファイル経由** | ⭐⭐⭐⭐⭐ | ⭐⭐⭐⭐⭐ | ⭐⭐⭐⭐ | ✅ 最推奨 |
| **ショートカット** | ⭐⭐⭐⭐ | ⭐⭐⭐⭐⭐ | ⭐⭐⭐⭐⭐ | ✅ 推奨 |
| **関連付け変更** | ⭐⭐ | ⚠️ ⭐⭐ | ⭐⭐⭐ | ❌ 非推奨 |
| **VBScriptラッパー** | ⭐⭐⭐ | ⭐⭐⭐⭐ | ⭐⭐⭐⭐ | ✅ 推奨 |

---

## 実行ポリシーについて

### 実行ポリシーとは？

PowerShellスクリプトの実行を制限するセキュリティ機能です。

### 確認方法

```powershell
Get-ExecutionPolicy
```

### 主なポリシー

| ポリシー | 説明 |
|---------|------|
| `Restricted` | スクリプト実行不可（デフォルト） |
| `AllSigned` | 署名されたスクリプトのみ実行可 |
| `RemoteSigned` | ローカルは実行可、リモートは署名必要 |
| `Unrestricted` | すべて実行可（非推奨） |

### 変更方法（管理者権限必要）

```powershell
# RemoteSignedに変更（推奨）
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### 一時的にバイパスする方法

実行ポリシーを変更せずに実行：

```powershell
powershell.exe -ExecutionPolicy Bypass -File ".\Invoke-RemoteBatch.ps1" -ComputerName "192.168.1.100" ...
```

**すでに用意されているバッチファイルはこの方法を使用しています**。

---

## おすすめの使い方

### ケース1: たまに使う場合

**方法**: バッチファイル経由

```cmd
remote_exec_ps.bat  ← 編集して実行
```

---

### ケース2: 頻繁に使う場合

**方法**: ショートカットをデスクトップに作成

1. ショートカット作成（パラメータ設定済み）
2. アイコンを変更（見やすく）
3. デスクトップに配置

---

### ケース3: 複数のタスクがある場合

**方法**: タスクごとにバッチファイルを作成

```
backup_server01.bat  ← サーバ01のバックアップ
backup_server02.bat  ← サーバ02のバックアップ
check_server01.bat   ← サーバ01のヘルスチェック
```

フォルダに整理：

```
C:\RemoteTools\
├── backup_server01.bat
├── backup_server02.bat
├── check_server01.bat
└── Invoke-RemoteBatch.ps1
```

---

## VBScriptラッパーの完全版

より高機能なVBScriptラッパーの例：

`RemoteExec_GUI.vbs` を作成：

```vbscript
' PowerShellスクリプトをGUIで実行
Option Explicit

Dim objShell, fso, scriptPath
Dim computerName, userName, batchPath, outputLog
Dim command, result

Set objShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' 設定項目
computerName = "192.168.1.100"
userName = "Administrator"
batchPath = "C:\Scripts\backup.bat"
outputLog = fso.GetParentFolderName(WScript.ScriptFullName) & "\remote_output.log"

' PowerShellスクリプトのパス
scriptPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\Invoke-RemoteBatch.ps1"

' 確認ダイアログ
result = MsgBox("以下の設定でリモートバッチを実行します：" & vbCrLf & vbCrLf & _
                "リモートサーバ: " & computerName & vbCrLf & _
                "実行ユーザー: " & userName & vbCrLf & _
                "実行ファイル: " & batchPath & vbCrLf & vbCrLf & _
                "実行しますか？", vbQuestion + vbYesNo, "リモートバッチ実行")

If result = vbNo Then
    WScript.Quit
End If

' コマンドを構築
command = "powershell.exe -ExecutionPolicy Bypass -File """ & scriptPath & """ " & _
          "-ComputerName """ & computerName & """ " & _
          "-UserName """ & userName & """ " & _
          "-BatchPath """ & batchPath & """ " & _
          "-OutputLog """ & outputLog & """"

' 実行（ウィンドウ表示: 1）
result = objShell.Run(command, 1, True)

' 結果を表示
If result = 0 Then
    MsgBox "リモートバッチ実行が正常に完了しました。" & vbCrLf & vbCrLf & _
           "ログファイル: " & outputLog, vbInformation, "実行完了"
Else
    MsgBox "リモートバッチ実行に失敗しました。" & vbCrLf & vbCrLf & _
           "終了コード: " & result & vbCrLf & _
           "ログファイルを確認してください: " & outputLog, vbCritical, "実行失敗"
End If

Set fso = Nothing
Set objShell = Nothing
```

**特徴:**
- ✅ 実行前に確認ダイアログを表示
- ✅ 実行結果を通知
- ✅ 終了コードを判定

---

## まとめ

### 最も簡単な方法

**すでに用意されているバッチファイルを使う**

```cmd
remote_exec_ps.bat  ← これをダブルクリック
```

### 最も柔軟な方法

**ショートカットを作成**
- 用途別に複数作成できる
- パラメータを事前設定できる

### 最も洗練された方法

**VBScriptラッパーを使用**
- GUIで確認・通知
- ウィンドウ表示制御
- エラーハンドリング

---

どの方法を選んでも、PowerShellスクリプトをダブルクリック（または簡単に）実行できます！
