# JP1ジョブネット起動ツール - リモート実行版

PowerShell Remotingを使用して、リモートのJP1サーバ上で`ajsentry`コマンドを実行するツールです。

---

## 📋 概要

このツールは、ローカルPCにJP1をインストールせずに、PowerShell Remotingを使って
リモートサーバ上の`ajsentry`コマンドを実行し、ジョブネットを起動します。

---

## ✅ 必要な環境

### ローカルPC（実行する側）

- Windows OS
- PowerShell 5.1以降
- **JP1のインストール不要**

### JP1サーバ（リモート側）

- Windows Server
- JP1/AJS3 - Manager が稼働中
- **PowerShell Remotingが有効** (`Enable-PSRemoting`済み)
- WinRMサービスが起動中
- ファイアウォールでポート5985（HTTP）または5986（HTTPS）が開放

---

## 🔧 事前準備

### JP1サーバ側の設定

JP1サーバで**PowerShell Remotingを有効化**する必要があります。

#### 1. 管理者権限でPowerShellを起動

JP1サーバにリモートデスクトップ接続し、管理者権限でPowerShellを起動。

#### 2. PowerShell Remotingを有効化

```powershell
Enable-PSRemoting -Force
```

#### 3. WinRMサービスを確認

```powershell
Get-Service WinRM
```

**出力例**:
```
Status   Name               DisplayName
------   ----               -----------
Running  WinRM              Windows Remote Management (WS-Manag…
```

#### 4. ファイアウォール設定確認

```powershell
Get-NetFirewallRule -Name "WINRM-HTTP-In-TCP" | Select-Object DisplayName,Enabled
```

**出力例**:
```
DisplayName                           Enabled
-----------                           -------
Windows Remote Management (HTTP-In)  True
```

---

## 🚀 使い方

### 1. バッチファイルを編集

`jp1_remote_start.bat`をテキストエディタで開き、以下の設定項目を編集します：

```batch
rem JP1/AJS3が稼働しているリモートサーバ
set JP1_SERVER=192.168.1.100

rem リモートサーバのユーザー名（Windowsログインユーザー）
set REMOTE_USER=Administrator

rem JP1ユーザー名
set JP1_USER=jp1admin

rem JP1パスワード（空の場合は実行時に入力を求めます）
set JP1_PASSWORD=

rem 起動するジョブネットのフルパス
set JOBNET_PATH=/main_unit/jobgroup1/daily_batch

rem ajsentryコマンドのパス（リモートサーバ上）
set AJSENTRY_PATH=C:\Program Files\HITACHI\JP1AJS3\bin\ajsentry.exe
```

### 2. 実行

`jp1_remote_start.bat`をダブルクリックまたはコマンドプロンプトから実行：

```cmd
jp1_remote_start.bat
```

### 3. 認証情報の入力

実行時に以下の認証情報を入力：

1. **JP1パスワード**（JP1_PASSWORDが空の場合）
2. **リモートサーバのWindowsログイン情報**（GUIダイアログで入力）

---

## 📖 実行例

### 成功時の出力

```
========================================
JP1ジョブネット起動ツール
（リモート実行版）
========================================

JP1サーバ      : 192.168.1.100
リモートユーザー: Administrator
JP1ユーザー    : jp1admin
ジョブネットパス: /main_unit/jobgroup1/daily_batch

ジョブネットを起動しますか？
実行する場合はYを押してください [Y,N]?Y

========================================
リモート接続してジョブネット起動中...
========================================

リモートサーバの認証情報を入力してください。

[Windows認証ダイアログが表示される]

リモートサーバに接続中...
ajsentryコマンドを実行中...

========================================
ジョブネットの起動に成功しました
========================================

実行結果:
KAVS1820-I ajsentryコマンドが正常終了しました。

ジョブネット: /main_unit/jobgroup1/daily_batch
サーバ      : 192.168.1.100

続行するには何かキーを押してください . . .
```

---

## ⚙️ 設定項目の詳細

| 設定項目 | 説明 | 例 |
|---------|------|---|
| JP1_SERVER | JP1サーバのホスト名またはIPアドレス | `192.168.1.100` |
| REMOTE_USER | リモートサーバのWindowsユーザー名 | `Administrator` |
| JP1_USER | JP1ユーザー名 | `jp1admin` |
| JP1_PASSWORD | JP1パスワード（空の場合は実行時入力） | ` `（空推奨） |
| JOBNET_PATH | ジョブネットのフルパス | `/main_unit/jobgroup1/daily_batch` |
| AJSENTRY_PATH | リモートサーバ上のajsentryパス | `C:\Program Files\HITACHI\JP1AJS3\bin\ajsentry.exe` |

---

## 🔐 認証の仕組み

このツールでは**2段階の認証**が行われます：

### 1. Windows認証（PowerShell Remoting）

リモートサーバへのアクセスに使用

- **ユーザー**: REMOTE_USER（Windowsログインユーザー）
- **パスワード**: 実行時にGUIダイアログで入力
- **目的**: リモートサーバでコマンド実行権限を得る

### 2. JP1認証（ajsentryコマンド）

JP1ジョブネットの実行に使用

- **ユーザー**: JP1_USER（JP1ユーザー）
- **パスワード**: JP1_PASSWORD（実行時入力も可）
- **目的**: ジョブネットの実行権限を得る

---

## ⚠️ 注意事項

### セキュリティ

- **パスワード**: バッチファイルにパスワードを記載しない
- **権限**: リモートサーバのAdministrator権限が必要
- **ファイアウォール**: WinRMポートの開放が必要

### PowerShell Remoting

- **信頼されたホスト**: 必要に応じてTrustedHostsに追加
  ```powershell
  Set-Item WSMan:\localhost\Client\TrustedHosts -Value "192.168.1.100" -Force
  ```

### JP1環境

- **ajsentryパス**: JP1のインストールパスを確認
- **実行権限**: JP1ユーザーにジョブネット実行権限が必要
- **本番環境**: 本番実行前に必ずテスト環境で確認

---

## 🐛 トラブルシューティング

### エラー: 「リモート実行に失敗しました」

**原因1**: PowerShell Remotingが有効化されていない

**対処法**:
```powershell
# JP1サーバで実行
Enable-PSRemoting -Force
```

---

**原因2**: WinRMサービスが停止している

**対処法**:
```powershell
# JP1サーバで確認
Get-Service WinRM

# 停止している場合は起動
Start-Service WinRM
Set-Service WinRM -StartupType Automatic
```

---

**原因3**: ファイアウォールでブロックされている

**対処法**:
```powershell
# JP1サーバで確認
Get-NetFirewallRule -Name "WINRM-HTTP-In-TCP"

# 無効の場合は有効化
Enable-NetFirewallRule -Name "WINRM-HTTP-In-TCP"
```

---

**原因4**: TrustedHostsの設定が必要

**対処法**（ローカルPCで実行）:
```powershell
# 信頼されたホストに追加
Set-Item WSMan:\localhost\Client\TrustedHosts -Value "192.168.1.100" -Force

# 確認
Get-Item WSMan:\localhost\Client\TrustedHosts
```

---

### エラー: 「認証情報の入力がキャンセルされました」

**原因**:
- 認証ダイアログでキャンセルボタンを押した

**対処法**:
- 再実行して正しい認証情報を入力

---

### エラー: 「ajsentryコマンドが見つかりません」

**原因**:
- AJSENTRY_PATHの設定が間違っている

**対処法**:
1. JP1サーバにリモートデスクトップ接続
2. ajsentry.exeの実際のパスを確認
3. バッチファイルのAJSENTRY_PATHを修正

**デフォルトパス例**:
```
C:\Program Files\HITACHI\JP1AJS3\bin\ajsentry.exe
C:\Program Files\Hitachi\JP1AJS2\bin\ajsentry.exe
```

---

### エラー: 「ジョブネットの起動に失敗しました」

**原因**:
- JP1ユーザー名、パスワードが間違っている
- ジョブネットパスが間違っている
- JP1サービスが停止している

**対処法**:
1. JP1ユーザー名、パスワードを確認
2. ジョブネットパスを確認
3. JP1サーバでJP1サービスの状態を確認
   ```cmd
   net start | findstr JP1
   ```

---

## 💡 応用例

### 設定ファイル版に改造

config.iniから設定を読み込むように改造可能：

```batch
rem config.iniから読み込み
for /f "usebackq tokens=1,* delims==" %%a in ("config.ini") do (
    set "%%a=%%b"
)
```

### 複数サーバに対応

サーバごとにバッチファイルを作成：

```
jp1_remote_server1.bat  → JP1_SERVER=192.168.1.100
jp1_remote_server2.bat  → JP1_SERVER=192.168.1.101
jp1_remote_server3.bat  → JP1_SERVER=192.168.1.102
```

---

## 🔗 関連ツール

このリポジトリの他のツールも参照：

- [batch/remote-exec/](../../remote-exec/) - リモートバッチ実行ツール（PowerShell Remoting）
- [batch/jp1-job-executor/rest-api/](../rest-api/) - JP1ジョブネット起動ツール（REST API版）

---

## 📚 参考資料

- [PowerShell Remotingについて](https://learn.microsoft.com/ja-jp/powershell/scripting/learn/ps101/08-powershell-remoting)
- [Enable-PSRemoting](https://learn.microsoft.com/ja-jp/powershell/module/microsoft.powershell.core/enable-psremoting)
- [JP1/AJS3 ajsentryコマンド](https://www.hitachi.co.jp/Prod/comp/soft1/manual/pc/d3K2211/AJSK01/EU210278.HTM)

---

**更新日**: 2025-12-02
