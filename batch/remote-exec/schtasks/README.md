# リモートバッチ実行ツール (Remote Batch Executor)

## 概要

リモートのWindowsサーバ上でバッチファイルをCMDから実行するためのツールです。
Windowsタスクスケジューラ（`schtasks`コマンド）を利用して、リモートでバッチファイルを起動します。

### 特徴
- ✅ **Windows標準機能のみ使用** - 追加ツール不要
- ✅ **セキュアな環境対応** - 外部ツールのインストールが制限されている環境でも使用可能
- ✅ **簡単な操作** - バッチファイルをダブルクリックするだけ
- ✅ **設定ファイル対応** - 複数サーバの管理が容易

## 必要な環境

### ローカル（実行元）
- Windows 10 / Windows 11 / Windows Server 2016以降
- コマンドプロンプト（cmd.exe）

### リモートサーバ（実行先）
- Windows Server 2008 R2以降 / Windows 10以降
- リモートレジストリサービスが有効
- ファイアウォールで以下のポートが開放されている：
  - TCP 135 (RPC)
  - TCP 445 (SMB)
  - 動的ポート範囲（通常は49152-65535）

### 必要な権限
- リモートサーバの**管理者権限**を持つアカウント
- タスクスケジューラへのアクセス権限

## インストール/セットアップ

### 1. ファイルのダウンロード
このフォルダ内のファイルを任意の場所にコピーします。

### 2. リモートサーバ側の準備

リモートサーバで以下のサービスが起動していることを確認してください：

```cmd
sc query "Schedule"
sc query "RemoteRegistry"
```

停止している場合は起動します：

```cmd
sc start "Schedule"
sc start "RemoteRegistry"
```

### 3. ファイアウォール設定の確認

リモートサーバのファイアウォールで以下の受信規則が有効になっているか確認：

- **リモート スケジュール タスク管理**
- **ファイルとプリンターの共有 (SMB受信)**

PowerShellで確認する場合：

```powershell
Get-NetFirewallRule | Where-Object {$_.DisplayName -like "*スケジュール*" -or $_.DisplayName -like "*SMB*"}
```

## 使い方

### 方法1: 基本版（毎回パスワード入力）

1. **`remote_exec.bat`** をテキストエディタで開く
2. 以下の設定項目を編集：
   ```batch
   set REMOTE_SERVER=192.168.1.100          ← サーバ名またはIPアドレス
   set REMOTE_USER=Administrator            ← ユーザー名
   set REMOTE_BATCH_PATH=C:\Scripts\target_script.bat  ← 実行するバッチファイルのフルパス
   ```
3. **`remote_exec.bat`** をダブルクリックまたはCMDから実行
4. パスワードを入力
5. 実行結果を確認

### 方法2: 設定ファイル版（複数サーバ管理向け）

1. **`config.ini.sample`** を **`config.ini`** にコピー
   ```cmd
   copy config.ini.sample config.ini
   ```

2. **`config.ini`** をテキストエディタで開いて編集：
   ```ini
   [Server]
   REMOTE_SERVER=192.168.1.100
   REMOTE_USER=Administrator
   REMOTE_BATCH_PATH=C:\Scripts\target_script.bat

   [Options]
   AUTO_DELETE=1
   TASK_NAME=RemoteExec
   ```

3. **`remote_exec_config.bat`** を実行
4. パスワードを入力（設定ファイルに記載していない場合）
5. 実行結果を確認

## 実行例

### 成功時の出力例

```
========================================
リモートバッチ実行ツール
========================================

リモートサーバ: 192.168.1.100
実行ユーザー  : Administrator
実行ファイル  : C:\Scripts\backup.bat
タスク名      : RemoteExec_12345

リモートサーバのパスワードを入力してください：
********

タスクを作成中...
タスク作成成功

タスクを実行中...
タスク実行開始

実行状態を確認中（5秒待機）...

タスク名: RemoteExec_12345
次回の実行時刻: なし
状態: 実行中

========================================
注意: タスクはバックグラウンドで実行されます。
      実行結果を確認するには、リモートサーバの
      ログファイルやタスクスケジューラを確認してください。
========================================

タスクを削除中...
タスク削除完了

処理が完了しました。
```

## パラメータ説明

### 基本設定

| パラメータ | 説明 | 例 |
|----------|------|-----|
| `REMOTE_SERVER` | リモートサーバのコンピュータ名またはIPアドレス | `192.168.1.100` |
| `REMOTE_USER` | 管理者ユーザー名<br>（ドメイン環境の場合は `DOMAIN\User` 形式） | `Administrator`<br>`COMPANY\admin` |
| `REMOTE_BATCH_PATH` | リモートサーバで実行するバッチファイルの**絶対パス** | `C:\Scripts\backup.bat` |
| `TASK_NAME` | タスクスケジューラに作成するタスク名<br>（重複しない名前を推奨） | `RemoteExec_%RANDOM%` |
| `AUTO_DELETE` | 実行後にタスクを自動削除するか<br>`1`=削除する / `0`=削除しない | `1` |

## 注意事項

### セキュリティ

⚠️ **パスワードの取り扱い**
- 設定ファイルにパスワードを記載しないことを強く推奨
- 必要な場合は、ファイルのアクセス権限を適切に設定してください
- `.gitignore` に `config.ini` を追加してGitにコミットしないようにしてください

⚠️ **実行権限**
- リモートサーバの管理者権限が必要です
- タスクは `SYSTEM` アカウントで実行されます

### 実行結果の確認

このツールはバッチファイルを**起動するだけ**で、実行結果は取得できません。
実行結果を確認するには以下の方法を使用してください：

1. **リモートサーバのログファイルを確認**
   - バッチファイル内でログ出力を実装してください

2. **タスクスケジューラの履歴を確認**
   ```cmd
   schtasks /Query /S 192.168.1.100 /U Administrator /P password /TN TaskName /V /FO LIST
   ```

3. **リモートデスクトップで直接確認**

### トラブルシューティング

#### エラー: "タスクの作成に失敗しました"

**原因と対処法：**

1. **認証エラー**
   - ユーザー名とパスワードが正しいか確認
   - ドメイン環境の場合は `DOMAIN\User` 形式で指定

2. **ネットワークエラー**
   - `ping <サーバ名>` でサーバに到達可能か確認
   - ファイアウォール設定を確認

3. **サービス停止**
   - リモートサーバで「Task Scheduler」サービスが起動しているか確認
   ```cmd
   sc \\192.168.1.100 query "Schedule"
   ```

4. **権限不足**
   - 管理者権限を持つアカウントを使用しているか確認

#### エラー: "アクセスが拒否されました"

**対処法：**
- リモートサーバのUACを一時的に無効化（テスト目的のみ）
- または、リモートレジストリで以下を設定：
  ```
  HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System
  LocalAccountTokenFilterPolicy = 1 (DWORD)
  ```

#### タスクは作成されたが実行されない

**対処法：**
1. リモートサーバでタスクスケジューラを開く
2. タスクの履歴を確認
3. バッチファイルのパスが正しいか確認
4. バッチファイルが実行可能か確認

## 応用例

### 1. 定期的なバックアップスクリプトの実行

```batch
rem config.ini の設定例
REMOTE_SERVER=backup-server
REMOTE_USER=backup_admin
REMOTE_BATCH_PATH=C:\Backup\daily_backup.bat
```

### 2. 複数サーバの一括再起動準備

複数のサーバに対して順次実行する場合：

```batch
@echo off
call remote_exec_config.bat server1_config.ini
timeout /t 10
call remote_exec_config.bat server2_config.ini
timeout /t 10
call remote_exec_config.bat server3_config.ini
```

### 3. ログ収集の自動化

リモートサーバでログを集約するバッチを実行：

```batch
REMOTE_BATCH_PATH=C:\Scripts\collect_logs.bat
```

## ライセンス

このツールはMITライセンスの下で公開されています。
個人・商用問わず自由に使用・改変できます。

## 参考情報

### schtasksコマンドの詳細

```cmd
rem タスク作成
schtasks /Create /?

rem タスク実行
schtasks /Run /?

rem タスク削除
schtasks /Delete /?

rem タスク照会
schtasks /Query /?
```

### 関連リンク

- [Microsoftドキュメント: Schtasks](https://docs.microsoft.com/ja-jp/windows-server/administration/windows-commands/schtasks)
- [タスクスケジューラのトラブルシューティング](https://docs.microsoft.com/ja-jp/troubleshoot/windows-server/system-management-components/troubleshoot-task-scheduler)

---

**作成日:** 2025-12-01
**バージョン:** 1.0
