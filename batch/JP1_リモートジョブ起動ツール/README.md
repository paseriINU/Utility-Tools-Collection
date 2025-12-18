# JP1_リモートジョブ起動ツール

PowerShell Remotingを使用して、リモートのJP1サーバ上で`ajsentry`コマンドを実行するツールです。

---

## 📋 概要

このツールは、ローカルPCにJP1をインストールせずに、PowerShell Remotingを使って
リモートサーバ上の`ajsentry`コマンドを実行し、ジョブネットを起動します。

### 主な機能

- **ジョブネット起動**: `ajsentry`でジョブネットを起動
- **完了待ち機能**: `ajsstatus`でジョブの完了を監視（正常終了/異常終了を判定）
- **詳細メッセージ取得**: `ajsshow`でジョブの実行詳細を取得
- **WinRM自動設定**: TrustedHostsを自動で設定・復元

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

`JP1_リモートジョブ起動ツール.bat`をテキストエディタで開き、設定セクションを編集します：

```powershell
$Config = @{
    # JP1/AJS3が稼働しているリモートサーバ
    JP1Server = "192.168.1.100"

    # リモートサーバのユーザー名（Windowsログインユーザー）
    RemoteUser = "Administrator"

    # リモートサーバのパスワード（空の場合は実行時に入力）
    RemotePassword = ""

    # JP1ユーザー名
    JP1User = "jp1admin"

    # JP1パスワード（空の場合は実行時に入力）
    JP1Password = ""

    # 起動するジョブネットのフルパス
    JobnetPath = "/main_unit/jobgroup1/daily_batch"

    # ajsentryコマンドのパス（リモートサーバ上）
    AjsentryPath = "C:\Program Files\HITACHI\JP1AJS3\bin\ajsentry.exe"

    # ajsstatusコマンドのパス（リモートサーバ上）
    AjsstatusPath = "C:\Program Files\HITACHI\JP1AJS3\bin\ajsstatus.exe"

    # ajsshowコマンドのパス（リモートサーバ上）
    AjsshowPath = "C:\Program Files\HITACHI\JP1AJS3\bin\ajsshow.exe"

    # HTTPS接続を使用する場合は $true
    UseSSL = $false

    # ジョブ完了を待つ場合は $true（起動のみの場合は $false）
    WaitForCompletion = $true

    # 完了待ちの最大時間（秒）。0の場合は無制限
    WaitTimeoutSeconds = 3600

    # 状態確認の間隔（秒）
    PollingIntervalSeconds = 10
}
```

### 2. 実行

`JP1_リモートジョブ起動ツール.bat`をダブルクリックまたはコマンドプロンプトから実行：

```cmd
JP1_リモートジョブ起動ツール.bat
```

### 3. 認証情報の入力

実行時に以下の認証情報を入力：

1. **JP1パスワード**（JP1_PASSWORDが空の場合）
2. **リモートサーバのWindowsログイン情報**（GUIダイアログで入力）

---

## 📖 実行例

### 成功時の出力（完了待ち有効）

```
========================================
JP1ジョブネット起動ツール
（リモート実行版）
========================================

JP1サーバ      : 192.168.1.100
リモートユーザー: Administrator
JP1ユーザー    : jp1admin
ジョブネットパス: /main_unit/jobgroup1/daily_batch
完了待ち       : 有効
タイムアウト   : 3600秒

ジョブネットを起動しますか？ (y/n)
y

========================================
リモート接続してジョブネット起動中...
========================================

リモートサーバに接続中...
[OK] 接続成功

ajsentryコマンドを実行中...

========================================
ジョブネットの起動に成功しました
========================================

ajsentry出力:
  KAVS1820-I ajsentryコマンドが正常終了しました。

========================================
ジョブ完了を待機中...
========================================

  状態: 実行中... (経過時間: 02:35)

========================================
ジョブネット実行結果: 正常終了
========================================

========================================
ジョブ詳細情報を取得中...
========================================

詳細情報 (ajsshow -E):
----------------------------------------
  UNIT-NAME       : /main_unit/jobgroup1/daily_batch
  STATUS          : ENDED NORMALLY
  START-TIME      : 2025/12/17 10:30:00
  END-TIME        : 2025/12/17 10:32:35
  RETURN-CODE     : 0
----------------------------------------

========================================
処理サマリー
========================================

  ジョブネット: /main_unit/jobgroup1/daily_batch
  サーバ      : 192.168.1.100
  起動結果    : 成功
  実行結果    : 正常終了

続行するには何かキーを押してください . . .
```

### 異常終了時の出力

```
========================================
ジョブネット実行結果: 異常終了
========================================

========================================
ジョブ詳細情報を取得中...
========================================

詳細情報 (ajsshow -E):
----------------------------------------
  UNIT-NAME       : /main_unit/jobgroup1/daily_batch
  STATUS          : ENDED ABNORMALLY
  START-TIME      : 2025/12/17 10:30:00
  END-TIME        : 2025/12/17 10:31:15
  RETURN-CODE     : 1
  MESSAGE         : KAVS0221-E ジョブが異常終了しました
----------------------------------------

========================================
処理サマリー
========================================

  ジョブネット: /main_unit/jobgroup1/daily_batch
  サーバ      : 192.168.1.100
  起動結果    : 成功
  実行結果    : 異常終了
```

---

## ⚙️ 設定項目の詳細

| 設定項目 | 説明 | 例 |
|---------|------|---|
| JP1Server | JP1サーバのホスト名またはIPアドレス | `192.168.1.100` |
| RemoteUser | リモートサーバのWindowsユーザー名 | `Administrator` |
| RemotePassword | リモートサーバのパスワード（空の場合は実行時入力） | ` `（空推奨） |
| JP1User | JP1ユーザー名 | `jp1admin` |
| JP1Password | JP1パスワード（空の場合は実行時入力） | ` `（空推奨） |
| JobnetPath | ジョブネットのフルパス | `/main_unit/jobgroup1/daily_batch` |
| AjsentryPath | リモートサーバ上のajsentryパス | `C:\Program Files\HITACHI\JP1AJS3\bin\ajsentry.exe` |
| AjsstatusPath | リモートサーバ上のajsstatusパス | `C:\Program Files\HITACHI\JP1AJS3\bin\ajsstatus.exe` |
| AjsshowPath | リモートサーバ上のajsshowパス | `C:\Program Files\HITACHI\JP1AJS3\bin\ajsshow.exe` |
| UseSSL | HTTPS接続を使用するか | `$false` |
| WaitForCompletion | ジョブ完了を待つか | `$true` |
| WaitTimeoutSeconds | 完了待ちのタイムアウト（秒）。0で無制限 | `3600` |
| PollingIntervalSeconds | 状態確認の間隔（秒） | `10` |

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

- [リモートバッチ実行ツール](../リモートバッチ実行ツール/) - リモートバッチ実行ツール（PowerShell Remoting）

---

## 📘 JP1/AJS3 コマンドリファレンス

このツールで使用するJP1/AJS3コマンドの一覧と説明です。

> **詳細版**: すべてのJP1/AJS3コマンドの詳細は [JP1_AJS3_COMMANDS.md](./JP1_AJS3_COMMANDS.md) を参照してください。

### コマンド一覧

| コマンド | 説明 | 用途 |
|----------|------|------|
| `ajsentry` | ジョブネットを即時実行する | ジョブネットの起動 |
| `ajsstatus` | ジョブネットの実行状態を取得する | 完了待ち・状態監視 |
| `ajsshow` | ジョブネットの詳細情報を取得する | 実行結果の詳細取得 |

### ajsentry（ジョブネット起動）

ジョブネットを即時実行するコマンドです。

```cmd
ajsentry.exe -h <ホスト名> -u <JP1ユーザー> -p <パスワード> -F <ジョブネットパス>
```

**主なオプション**:

| オプション | 説明 |
|-----------|------|
| `-h` | 接続先のJP1/AJS3マネージャーのホスト名 |
| `-u` | JP1ユーザー名 |
| `-p` | JP1パスワード |
| `-F` | 起動するジョブネットのフルパス |

**実行例**:
```cmd
ajsentry.exe -h localhost -u jp1admin -p password -F /main_unit/jobgroup1/daily_batch
```

**正常終了時の出力**:
```
KAVS1820-I ajsentryコマンドが正常終了しました。
```

---

### ajsstatus（状態確認）

ジョブネットの実行状態を取得するコマンドです。

```cmd
ajsstatus.exe -h <ホスト名> -u <JP1ユーザー> -p <パスワード> -F <ジョブネットパス>
```

**主なオプション**:

| オプション | 説明 |
|-----------|------|
| `-h` | 接続先のJP1/AJS3マネージャーのホスト名 |
| `-u` | JP1ユーザー名 |
| `-p` | JP1パスワード |
| `-F` | 状態を確認するジョブネットのフルパス |

**実行例**:
```cmd
ajsstatus.exe -h localhost -u jp1admin -p password -F /main_unit/jobgroup1/daily_batch
```

**状態の種類**:

| 状態 | 説明 |
|------|------|
| `now running` / `running` | 実行中 |
| `wait` / `queued` | 待機中・キュー待ち |
| `ended normally` / `normal end` | 正常終了 |
| `ended abnormally` / `abnormal end` | 異常終了 |
| `killed` / `interrupted` | 強制終了・中断 |

---

### ajsshow（詳細情報取得）

ジョブネットの詳細情報を取得するコマンドです。

```cmd
ajsshow.exe -h <ホスト名> -u <JP1ユーザー> -p <パスワード> -F <ジョブネットパス> -E
```

**主なオプション**:

| オプション | 説明 |
|-----------|------|
| `-h` | 接続先のJP1/AJS3マネージャーのホスト名 |
| `-u` | JP1ユーザー名 |
| `-p` | JP1パスワード |
| `-F` | 情報を取得するジョブネットのフルパス |
| `-E` | 実行結果の詳細情報を取得 |
| `-i` | ユニット定義情報を取得 |

**実行例**:
```cmd
ajsshow.exe -h localhost -u jp1admin -p password -F /main_unit/jobgroup1/daily_batch -E
```

**出力例**:
```
UNIT-NAME       : /main_unit/jobgroup1/daily_batch
STATUS          : ENDED NORMALLY
START-TIME      : 2025/12/17 10:30:00
END-TIME        : 2025/12/17 10:32:35
RETURN-CODE     : 0
```

---

### コマンドの配置場所

JP1/AJS3コマンドは通常、以下のパスにインストールされています：

```
C:\Program Files\HITACHI\JP1AJS3\bin\
C:\Program Files\Hitachi\JP1AJS2\bin\
```

---

## 📚 参考資料

- [PowerShell Remotingについて](https://learn.microsoft.com/ja-jp/powershell/scripting/learn/ps101/08-powershell-remoting)
- [Enable-PSRemoting](https://learn.microsoft.com/ja-jp/powershell/module/microsoft.powershell.core/enable-psremoting)
- [JP1/AJS3 ajsentryコマンド](https://www.hitachi.co.jp/Prod/comp/soft1/manual/pc/d3K2211/AJSK01/EU210278.HTM)
- [JP1/AJS3 コマンドリファレンス](https://www.hitachi.co.jp/Prod/comp/soft1/manual/pc/d3K2211/AJSK01/EU210000.HTM)

---

**更新日**: 2025-12-18
