# リモートバッチ実行ツール (Remote Batch Executor)

## 概要

リモートのWindowsサーバ上でバッチファイルをCMDから実行するためのツール集です。
接続方法ごとに3つのバージョンを用意しています。

## 📁 フォルダ構成

```
remote-exec/
├── schtasks/                      # タスクスケジューラ版
│   ├── remote_exec.bat            # 基本版（パラメータ直接編集）
│   ├── remote_exec_config.bat     # 設定ファイル版
│   ├── config.ini.sample          # 設定ファイルサンプル
│   └── README.md                  # 詳細ドキュメント
│
├── winrm/                         # WinRM版（バッチ）
│   ├── remote_exec_winrm.bat      # 基本版（パラメータ直接編集）
│   ├── remote_exec_winrm_config.bat  # 設定ファイル版
│   ├── config_winrm.ini.sample    # 設定ファイルサンプル
│   └── README.md                  # 詳細ドキュメント
│
└── powershell-remoting/           # PowerShell Remoting版
    ├── Invoke-RemoteBatch.ps1     # PowerShellスクリプト（本体）
    ├── remote_exec_ps.bat         # バッチラッパー（基本版）
    ├── remote_exec_ps_config.bat  # バッチラッパー（設定ファイル版）
    ├── config_ps.ini.sample       # 設定ファイルサンプル
    └── README.md                  # 詳細ドキュメント
```

## 🔍 どれを使うべきか？

### タスクスケジューラ版（schtasks/）

**推奨ケース：**
- ✅ **セットアップを簡単にしたい**
- ✅ **実行するだけで結果は不要**（ログはリモートサーバで確認）
- ✅ **WinRMの設定が難しい環境**

**特徴：**
- Windows標準のタスクスケジューラを使用
- セットアップが簡単（サービスが標準で有効）
- 実行結果はリアルタイムで取得できない

**使用ポート：** TCP 135, 445

📖 [詳細はこちら](schtasks/README.md)

---

### WinRM版（winrm/）

**推奨ケース：**
- ✅ **実行結果をリアルタイムで確認したい**
- ✅ **エラーを即座に把握したい**
- ✅ **バッチファイルで完結させたい**

**特徴：**
- バッチファイル内でPowerShell Remotingを使用
- 実行結果をリアルタイムで取得可能
- セットアップがやや複雑（WinRM有効化が必要）

**使用ポート：** TCP 5985 (HTTP) / 5986 (HTTPS)

📖 [詳細はこちら](winrm/README.md)

---

### PowerShell Remoting版（powershell-remoting/）

**推奨ケース：**
- ✅ **PowerShellの全機能を活用したい**
- ✅ **終了コードを取得したい**
- ✅ **引数を柔軟に渡したい**
- ✅ **複数サーバへの並列実行など高度な処理をしたい**

**特徴：**
- 純粋なPowerShellスクリプト（.ps1）として実装
- バッチファイル経由でも実行可能
- 終了コード取得、詳細エラーハンドリング
- Get-Credentialなどネイティブ機能を活用

**使用ポート：** TCP 5985 (HTTP) / 5986 (HTTPS)

📖 [詳細はこちら](powershell-remoting/README.md)

---

## 📊 比較表

| 項目 | タスクスケジューラ版 | WinRM版（バッチ） | PowerShell Remoting版 |
|-----|------------------|----------------|---------------------|
| **実行結果の取得** | ❌ 取得不可 | ✅ リアルタイム取得 | ✅ リアルタイム取得 |
| **標準出力の表示** | ❌ 表示されない | ✅ 画面に表示 | ✅ 画面に表示 |
| **終了コード取得** | ❌ 取得不可 | ❌ 取得不可 | ✅ 取得可能 |
| **引数の柔軟性** | ⚠️ 基本的 | ⚠️ 基本的 | ✅ 完全対応 |
| **実装言語** | バッチ | バッチ | PowerShell |
| **カスタマイズ性** | ⭐⭐ | ⭐⭐⭐ | ⭐⭐⭐⭐⭐ |
| **セットアップ難易度** | ⭐⭐ 簡単 | ⭐⭐⭐⭐ やや複雑 | ⭐⭐⭐⭐ やや複雑 |
| **使用ポート** | 135, 445 | 5985, 5986 | 5985, 5986 |
| **追加ツール** | 不要 | 不要 | 不要 |
| **Windows標準機能** | ✅ | ✅ | ✅ |

## 🚀 クイックスタート

### タスクスケジューラ版を使う場合

1. `schtasks/` フォルダに移動
2. `remote_exec.bat` を編集して設定を記入
3. 実行

```cmd
cd schtasks
notepad remote_exec.bat  # 設定を編集
remote_exec.bat          # 実行
```

### WinRM版を使う場合

1. **リモートサーバでWinRMを有効化**（初回のみ）
   ```powershell
   winrm quickconfig
   ```

2. `winrm/` フォルダに移動
3. `remote_exec_winrm.bat` を編集して設定を記入
4. 実行

```cmd
cd winrm
notepad remote_exec_winrm.bat  # 設定を編集
remote_exec_winrm.bat          # 実行
```

### PowerShell Remoting版を使う場合

1. **リモートサーバでWinRMを有効化**（初回のみ）
   ```powershell
   winrm quickconfig
   ```

2. `powershell-remoting/` フォルダに移動
3. PowerShellから直接実行、またはバッチファイル経由で実行

```powershell
# PowerShellから直接実行
cd powershell-remoting
.\Invoke-RemoteBatch.ps1 -ComputerName "192.168.1.100" -UserName "Administrator" -BatchPath "C:\Scripts\backup.bat"
```

または

```cmd
# バッチファイル経由で実行
cd powershell-remoting
notepad remote_exec_ps.bat  # 設定を編集
remote_exec_ps.bat          # 実行
```

## 💡 使い分けガイド

### ユースケース別推奨

| やりたいこと | 推奨バージョン |
|------------|--------------|
| 定期バックアップスクリプトを起動 | タスクスケジューラ版 |
| リモートでコマンド実行して結果を見たい | WinRM版 / PowerShell Remoting版 |
| ログ収集スクリプトを実行（ログはリモートに保存） | タスクスケジューラ版 |
| リモートサーバの状態確認スクリプトを実行 | PowerShell Remoting版 |
| バッチ実行の成功/失敗を終了コードで判定したい | PowerShell Remoting版 |
| 複数サーバへ並列実行したい | PowerShell Remoting版 |
| PowerShellに精通していてフル活用したい | PowerShell Remoting版 |
| バッチファイルで完結させたい | WinRM版 |
| セットアップ時間を最小限にしたい | タスクスケジューラ版 |

## ⚙️ 必要な環境

### ローカルPC（実行元）
- Windows 10 / Windows 11 / Windows Server 2016以降
- コマンドプロンプト（cmd.exe）
- PowerShell 5.1以降（WinRM版のみ）

### リモートサーバ（実行先）

**タスクスケジューラ版：**
- Task Schedulerサービスが起動していること
- ファイアウォールでポート135, 445が開放

**WinRM版 / PowerShell Remoting版：**
- WinRMサービスが有効化されていること
- PowerShell Remotingが有効
- ファイアウォールでポート5985（または5986）が開放

### 共通
- リモートサーバの**管理者権限**を持つアカウント

## 🔐 セキュリティ注意事項

1. **パスワードの管理**
   - 設定ファイルにパスワードを記載しないことを推奨
   - 実行時に入力する方法を推奨

2. **ネットワークセキュリティ**
   - 信頼できるネットワーク内でのみ使用
   - インターネット経由の場合はVPNを使用
   - WinRM版でインターネット経由の場合はHTTPS（ポート5986）を使用

3. **.gitignore設定**
   - `config.ini`, `config_winrm.ini`, `config_ps.ini` はGitにコミットしないように設定済み

## 📝 ライセンス

このツールはMITライセンスの下で公開されています。
個人・商用問わず自由に使用・改変できます。

## 🔗 関連リンク

- [タスクスケジューラ版の詳細ドキュメント](schtasks/README.md)
- [WinRM版（バッチ）の詳細ドキュメント](winrm/README.md)
- [PowerShell Remoting版の詳細ドキュメント](powershell-remoting/README.md)

---

**作成日:** 2025-12-01
**バージョン:** 1.0
