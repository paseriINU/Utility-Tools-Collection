# WinRM Client for Linux

Linux（Red Hat等）からWindows Server 2022へWinRM（Windows Remote Management）で接続してバッチファイルやコマンドを実行するツールです。

## 特徴

✅ **追加パッケージ不要** - Python標準ライブラリのみで動作（pywinrm不要）
✅ **IT制限環境対応** - インターネット接続やpipインストールが制限されている環境でも使用可能
✅ **3つの実装** - Python版・Bash版・C言語版を提供
✅ **完全なWinRM実装** - WinRMプロトコルをフルサポート

## 概要

このツールは、LinuxからWindows Serverへのリモート管理を可能にします。IT制限環境でも使用できるよう、標準ライブラリのみで実装されています。

### 提供されるスクリプト

- **winrm_exec.py** - Python版（推奨）
  - **Python標準ライブラリのみ使用**（追加パッケージ不要）
  - フル機能対応
  - バッチファイル実行
  - コマンド実行
  - 詳細なログ出力
  - WinRMプロトコル完全実装

- **winrm_exec.sh** - Bashシェルスクリプト版
  - **標準コマンドのみ使用**（curl、base64、grepのみ）
  - フル機能対応（Python版と同等）
  - バッチファイル実行
  - コマンド実行
  - WinRMプロトコル完全実装

- **winrm_exec.c** - C言語版
  - **標準Cライブラリ + libcurl**のみ使用
  - 高速・軽量な実装
  - バッチファイル実行
  - コマンド実行
  - 組み込み環境やスクリプト言語が使用できない環境向け

## 必要な環境

### Linux側（クライアント）

#### Python版を使用する場合
- Python 3.6以降（**標準ライブラリのみ、追加パッケージ不要**）

```bash
# Pythonバージョン確認
python3 --version

# Python 3がない場合のみインストール
# Red Hat系（RHEL, CentOS, Rocky Linux等）
sudo yum install python3

# Debian/Ubuntu系
sudo apt install python3
```

**重要**: `pip install`は**不要**です。Python標準ライブラリのみで動作します。

**動作確認済み**:
- ✅ Python 3.6.8（RHEL 7 / CentOS 7標準）
- ✅ Python 3.7以降

#### Bashシェルスクリプト版を使用する場合
- bash（標準でインストール済み）
- curl（標準でインストール済み、または簡単にインストール可能）
- base64（標準でインストール済み）
- grep（標準でインストール済み）

```bash
# curlがない場合のみインストール
# Red Hat系
sudo yum install curl

# Debian/Ubuntu系
sudo apt install curl
```

#### C言語版を使用する場合
- GCCコンパイラ（標準でインストール済み、またはインストール可能）
- 標準Cライブラリ（glibc）

```bash
# GCCがない場合のみインストール
# Red Hat系
sudo yum install gcc

# Debian/Ubuntu系
sudo apt install gcc
```

**動作確認済み**:
- ✅ GCC 4.8以降（RHEL 7 / CentOS 7標準）
- ✅ GCC 8以降

### Windows側（サーバ）

- Windows Server 2022（または Windows 10/11）
- PowerShell 5.1以降
- 管理者権限

## Windows側のセットアップ

Windows Server 2022でWinRMを有効化し、Linuxからの接続を受け付ける設定を行います。

### 1. PowerShellを管理者として起動

スタートメニュー → "PowerShell" → 右クリック → "管理者として実行"

### 2. WinRMサービスの有効化と設定

```powershell
# WinRMサービスの有効化（クイック設定）
winrm quickconfig -force

# サービスの自動起動設定
Set-Service -Name WinRM -StartupType Automatic

# サービスの起動確認
Get-Service -Name WinRM
```

### 3. HTTP接続の許可（基本認証）

```powershell
# HTTPリスナーの確認
winrm enumerate winrm/config/listener

# HTTPリスナーがない場合は作成
New-Item -Path WSMan:\localhost\Listener -Transport HTTP -Address * -Force

# 基本認証の有効化（開発環境のみ推奨）
Set-Item WSMan:\localhost\Service\Auth\Basic -Value $true

# NTLM認証の有効化（推奨）
Set-Item WSMan:\localhost\Service\Auth\Negotiate -Value $true
```

### 4. ファイアウォール設定

```powershell
# WinRM HTTP（ポート5985）の許可
New-NetFirewallRule -Name "WinRM-HTTP-In" `
    -DisplayName "Windows Remote Management (HTTP-In)" `
    -Protocol TCP -LocalPort 5985 -Action Allow -Enabled True

# ファイアウォールルールの確認
Get-NetFirewallRule -Name "WinRM-HTTP-In"
```

### 5. TrustedHosts設定（Linuxクライアントを許可）

```powershell
# すべてのホストからの接続を許可（開発環境のみ推奨）
Set-Item WSMan:\localhost\Client\TrustedHosts -Value "*" -Force

# または特定のLinuxサーバのIPのみ許可（本番環境推奨）
Set-Item WSMan:\localhost\Client\TrustedHosts -Value "192.168.1.10" -Force

# 複数のIPを許可する場合
Set-Item WSMan:\localhost\Client\TrustedHosts -Value "192.168.1.10,192.168.1.11" -Force

# 設定確認
Get-Item WSMan:\localhost\Client\TrustedHosts
```

### 6. WinRM設定の確認

```powershell
# 全体設定の確認
winrm get winrm/config

# サービス設定の確認
winrm get winrm/config/service
winrm get winrm/config/service/auth

# クライアント設定の確認
winrm get winrm/config/client
```

### 7. 接続テスト（Windows側から自己テスト）

```powershell
# ローカルホストへの接続テスト
Test-WSMan -ComputerName localhost

# 認証テスト
$cred = Get-Credential
Test-WSMan -ComputerName localhost -Credential $cred -Authentication Negotiate
```

### セキュリティ強化（本番環境向け）

#### HTTPS接続の設定（推奨）

```powershell
# 自己署名証明書の作成
$cert = New-SelfSignedCertificate -DnsName "your-server.example.com" `
    -CertStoreLocation "Cert:\LocalMachine\My"

# HTTPSリスナーの作成
New-Item -Path WSMan:\localhost\Listener -Transport HTTPS `
    -Address * -CertificateThumbPrint $cert.Thumbprint -Force

# ファイアウォールでHTTPS（ポート5986）を許可
New-NetFirewallRule -Name "WinRM-HTTPS-In" `
    -DisplayName "Windows Remote Management (HTTPS-In)" `
    -Protocol TCP -LocalPort 5986 -Action Allow -Enabled True
```

#### 基本認証の無効化（NTLM/Kerberosのみ使用）

```powershell
# 基本認証を無効化
Set-Item WSMan:\localhost\Service\Auth\Basic -Value $false

# NTLM認証のみ有効
Set-Item WSMan:\localhost\Service\Auth\Negotiate -Value $true
```

## Linux側の使い方

### Python版の使い方（推奨）

**重要**: このスクリプトは**Python標準ライブラリのみ**を使用します。`pip install`は不要です。

#### 1. スクリプト内の設定を編集

`winrm_exec.py` を開き、設定セクションを編集します：

```python
# Windows接続情報
WINDOWS_HOST = "192.168.1.100"      # Windows ServerのIPアドレス
WINDOWS_PORT = 5985                  # WinRMポート（HTTP: 5985, HTTPS: 5986）
WINDOWS_USER = "Administrator"       # Windowsユーザー名
WINDOWS_PASSWORD = "YourPassword"    # Windowsパスワード

# 利用可能な環境のリスト
ENVIRONMENTS = ["TST1T", "TST2T"]

# 実行するバッチファイル（{ENV}は環境名に置換されます）
BATCH_FILE_PATH = r"C:\Scripts\{ENV}\test.bat"
```

#### 2. 実行

```bash
# 環境を引数で指定して実行（必須）
python3 winrm_exec.py TST1T
python3 winrm_exec.py TST2T

# コマンドライン引数で設定を上書き
python3 winrm_exec.py TST1T --host 192.168.1.100 --user Administrator --password Pass123

# バッチファイルを指定して実行
python3 winrm_exec.py TST1T --batch "C:\Scripts\backup.bat"

# 直接コマンドを実行
python3 winrm_exec.py TST1T --command "echo Hello from Linux"

# ログレベルを指定
python3 winrm_exec.py TST1T --log-level DEBUG
```

#### オプション一覧

```
ENV             環境名（TST1T, TST2T など）※必須
--host          Windows ServerのIPアドレスまたはホスト名
--port          WinRMポート（デフォルト: 5985）
--user          Windowsユーザー名
--password      Windowsパスワード
--domain        ドメイン名（ローカル認証の場合は空）
--batch         実行するバッチファイル（Windows側のパス）
--command       直接実行するコマンド
--timeout       タイムアウト（秒）（デフォルト: 300）
--log-level     ログレベル（DEBUG, INFO, WARNING, ERROR）
```

### C言語版の使い方

**注意**: このプログラムは**標準Cライブラリのみ**を使用します。外部ライブラリ不要でNTLM認証を自前実装しています。

#### 1. ソースファイル内の設定を編集

`winrm_exec.c` の設定セクションを編集します：

```c
/* --- Windows接続情報 --- */
#define DEFAULT_HOST "192.168.1.100"     /* Windows ServerのIPアドレス */
#define DEFAULT_USER "Administrator"      /* Windowsユーザー名 */
#define DEFAULT_PASS "YourPassword"       /* Windowsパスワード */
#define DEFAULT_DOMAIN ""                 /* ドメイン名（空 = ローカル認証） */
#define DEFAULT_PORT 5985                 /* WinRMポート */
```

#### 2. コンパイル

```bash
# 基本コンパイル
gcc -o winrm_exec winrm_exec.c

# 警告を確認する場合
gcc -Wall -o winrm_exec winrm_exec.c
```

#### 3. 実行

```bash
# 環境を引数で指定して実行（必須）
./winrm_exec TST1T
./winrm_exec TST2T

# 環境変数で設定を上書き
WINRM_HOST=192.168.1.100 WINRM_USER=Admin WINRM_PASS=Pass123 ./winrm_exec TST1T
```

#### C言語版の特徴

- **NTLM v2認証を自前実装** - MD4、MD5、HMAC-MD5を含む完全実装
- **Windows側の設定変更不要** - デフォルトのNTLM認証を使用
- **高速・軽量** - スクリプト言語より高速に動作
- **組み込み環境向け** - Python/Bashが使用できない環境でも動作

### Bashシェルスクリプト版の使い方

**注意**: このスクリプトは**標準コマンドのみ**（curl、base64、grep）を使用します。追加インストールはほぼ不要です。

#### 1. スクリプト内の設定を編集

`winrm_exec.sh` を開き、設定セクションを編集します：

```bash
# Windows接続情報
_DEFAULT_HOST='192.168.1.100'
_DEFAULT_USER='Administrator'
_DEFAULT_PASS='YourPassword'
WINRM_HOST="${WINRM_HOST:-$_DEFAULT_HOST}"
WINRM_PORT="${WINRM_PORT:-5985}"
WINRM_USER="${WINRM_USER:-$_DEFAULT_USER}"
WINRM_PASS="${WINRM_PASS:-$_DEFAULT_PASS}"

# 利用可能な環境のリスト
ENVIRONMENTS=("TST1T" "TST2T")

# 実行するバッチファイル（{ENV}は環境名に置換されます）
_DEFAULT_BATCH_PATH='C:\Scripts\{ENV}\test.bat'
BATCH_FILE_PATH="${BATCH_FILE_PATH:-$_DEFAULT_BATCH_PATH}"
```

**環境選択機能**:
- 環境名を引数で指定します（必須）
- `{ENV}` プレースホルダーが指定した環境名に置換されます
- 例: TST1T を指定すると `C:\Scripts\TST1T\test.bat` が実行されます

#### 2. 実行権限の付与

```bash
chmod +x winrm_exec.sh
```

#### 3. 実行

```bash
# 環境を引数で指定して実行（必須）
./winrm_exec.sh TST1T
./winrm_exec.sh TST2T

# 環境変数で設定を上書き
WINRM_HOST=192.168.1.100 WINRM_USER=Admin WINRM_PASS=Pass123 ./winrm_exec.sh TST1T

# デバッグモードで実行
DEBUG=true ./winrm_exec.sh TST1T
```

## 実行例

### Python版の実行例

```bash
$ python3 winrm_exec.py TST1T

2025-01-15 10:30:45 [INFO] === WinRM Remote Batch Executor (標準ライブラリ版) ===
2025-01-15 10:30:45 [INFO] 接続先: 192.168.1.100:5985
2025-01-15 10:30:45 [INFO] ユーザー: Administrator
2025-01-15 10:30:45 [INFO] WinRMエンドポイント: http://192.168.1.100:5985/wsman
2025-01-15 10:30:45 [INFO] バッチファイル実行: C:\Scripts\TST1T\test.bat
2025-01-15 10:30:45 [INFO] シェル作成中...
2025-01-15 10:30:46 [INFO] シェル作成成功: xxx-xxx-xxx
2025-01-15 10:30:46 [INFO] コマンド実行中...
2025-01-15 10:30:47 [INFO] コマンド実行開始: yyy-yyy-yyy
2025-01-15 10:30:47 [INFO] コマンド出力取得中...
2025-01-15 10:30:48 [INFO] コマンド完了: 終了コード=0
2025-01-15 10:30:48 [INFO] シェル削除中...
2025-01-15 10:30:48 [INFO] シェル削除完了

============================================================
実行結果
============================================================

[標準出力]
Hello from Windows Server!
Current time: 2025-01-15 10:30:47

終了コード: 0
============================================================
```

## トラブルシューティング

### Bashスクリプト版のエラー

#### エラー詳細表示

Bashスクリプト版では、コマンド実行時に発生したエラーの詳細が表示されます：

**curlエラーコード別の対処法**:

- **エラーコード 6**: ホスト名の解決に失敗
  - ホスト名またはIPアドレスを確認してください

- **エラーコード 7**: 接続に失敗
  - ホストが起動しているか、ファイアウォール設定を確認してください
  - WinRMサービスが起動しているか確認してください

- **エラーコード 28**: タイムアウト
  - TIMEOUT値を増やすか、ネットワーク接続を確認してください

- **エラーコード 52**: サーバーから応答なし
  - WinRMサービスが起動しているか確認してください

**HTTPエラー別の対処法**:

- **HTTP 401**: 認証失敗
  - ユーザー名とパスワードを確認してください

- **HTTP 500**: サーバー内部エラー
  - WinRM設定またはコマンド内容を確認してください

**デバッグモード**:

詳細なSOAP通信ログを確認するには、DEBUG=true を設定してください：

```bash
DEBUG=true ./winrm_exec.sh
```

### 接続エラーが発生する場合

#### エラー: "Connection refused" または "Connection timeout"

**原因**: WinRMサービスが起動していない、またはファイアウォールでブロックされている

**解決策**:
```powershell
# Windows側でWinRMサービスを確認
Get-Service -Name WinRM

# サービスが停止している場合は起動
Start-Service -Name WinRM

# ファイアウォールルールを確認
Get-NetFirewallRule -Name "WinRM-HTTP-In" | Format-List

# ポート5985がリスニング状態か確認
netstat -an | findstr :5985
```

#### エラー: "401 Unauthorized" または認証エラー

**原因**: 認証情報が正しくない、またはNTLM/Negotiate認証が無効

**解決策**:
```powershell
# Windows側でNegotiate認証が有効か確認
Get-Item WSMan:\localhost\Service\Auth\Negotiate

# 無効の場合は有効化（本ツールはNTLM認証を使用）
Set-Item WSMan:\localhost\Service\Auth\Negotiate -Value $true
```

**注意**: 本ツールのすべての実装（Python版・C言語版・Bash版）は**NTLM認証**を使用しています。Windows側でBasic認証を有効化する必要はありません。

#### エラー: "403 Forbidden"

**原因**: TrustedHostsに接続元が登録されていない

**解決策**:
```powershell
# Windows側でTrustedHostsを確認
Get-Item WSMan:\localhost\Client\TrustedHosts

# Linuxクライアントを追加
Set-Item WSMan:\localhost\Client\TrustedHosts -Value "192.168.1.10" -Force
```

### Python版のエラー

#### エラー: "ImportError: No module named 'urllib.request'"

**原因**: Python 2系を使用している

**解決策**:
```bash
# Python 3を明示的に使用
python3 winrm_exec.py

# Python 3のバージョン確認
python3 --version

# Python 3がインストールされていない場合
sudo yum install python3  # Red Hat系
sudo apt install python3  # Debian/Ubuntu系
```

**注意**: このツールは**Python標準ライブラリのみ**を使用します。`pip install`は不要です。

### ネットワーク関連の問題

#### ポート疎通確認（Linux側から）

```bash
# telnetで疎通確認
telnet 192.168.1.100 5985

# ncで疎通確認
nc -zv 192.168.1.100 5985

# curlで接続テスト
curl -v http://192.168.1.100:5985/wsman
```

#### Windows側でリスナー確認

```powershell
# WinRMリスナーの確認
winrm enumerate winrm/config/listener

# 出力例:
# Listener
#     Address = *
#     Transport = HTTP
#     Port = 5985
```

### パフォーマンスの問題

#### タイムアウトが発生する場合

```bash
# タイムアウト時間を延長
python3 winrm_exec.py --timeout 600 --batch "C:\Scripts\long-running-task.bat"
```

```powershell
# Windows側でタイムアウト設定を延長
Set-Item WSMan:\localhost\MaxTimeoutms -Value 600000
```

## 既知の制限事項・環境依存の問題

### サーバー側TrustedHosts設定によるNTLM認証の有効化

LinuxからのNTLM認証が401エラーで失敗する場合、**サーバー側のTrustedHosts設定**が必要です。

#### 設定手順（PowerShell）

```powershell
# 1. 現在の設定を確認（復元用に控える）
Get-Item WSMan:\localhost\Client\TrustedHosts

# 2. LinuxクライアントのIPアドレスを追加
Set-Item WSMan:\localhost\Client\TrustedHosts -Value "LinuxのIPアドレス" -Force

# または全ホスト許可（テスト用）
Set-Item WSMan:\localhost\Client\TrustedHosts -Value "*" -Force

# 3. 設定確認
Get-Item WSMan:\localhost\Client\TrustedHosts
```

#### 元に戻す方法（PowerShell）

```powershell
# 空に戻す（元々空だった場合）
Set-Item WSMan:\localhost\Client\TrustedHosts -Value "" -Force

# または元の値に戻す
Set-Item WSMan:\localhost\Client\TrustedHosts -Value "元の値" -Force

# 設定確認
Get-Item WSMan:\localhost\Client\TrustedHosts
```

#### GUI での設定方法

**方法1: グループポリシーエディター（gpedit.msc）**

1. `Win + R` → `gpedit.msc` を実行
2. 以下のパスに移動:
   ```
   コンピューターの構成
   └─ 管理用テンプレート
       └─ Windowsコンポーネント
           └─ Windows リモート管理 (WinRM)
               └─ WinRM クライアント
   ```
3. 「信頼されたホスト」をダブルクリック
4. 「有効」を選択し、TrustedHostsList に `*` または Linux の IP を入力
5. 「OK」をクリック

**方法2: レジストリエディター（regedit）**

1. `Win + R` → `regedit` を実行
2. 以下のキーに移動:
   ```
   HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WSMAN\Client
   ```
3. 「TrustedHosts」を右クリック → 「修正」
4. 値に `*` または Linux の IP アドレスを入力
5. 「OK」をクリック

**元に戻す場合**: 同じ手順で値を空にするか、元の値に戻します。

---

### NTLM認証が失敗する環境

以下の環境では、LinuxからのNTLM認証が失敗する場合があります。

#### 1. Active Directoryドメイン環境でKerberosが優先される場合

**現象**: 401 Unauthorized（エラーコード 0xC000006D: STATUS_LOGON_FAILURE）

**原因**:
- サーバーがActive Directoryドメインに参加している
- `winrm get winrm/config/service/auth` で `Kerberos = true`、`Negotiate = true` だが `NTLM` の項目がない
- WindowsクライアントはKerberosで認証成功するが、LinuxからのNTLM認証は拒否される

**確認コマンド**:
```powershell
# サーバー側の認証設定確認
winrm get winrm/config/service/auth

# 期待される出力にNTLM = trueがあるか確認
# Basic = false
# Kerberos = true
# Negotiate = true
# NTLM = true  ← これがない場合、NTLM認証不可
```

**解決策**:
1. **TrustedHostsにLinuxクライアントを追加**（推奨）
   ```powershell
   Set-Item WSMan:\localhost\Client\TrustedHosts -Value "LinuxのIPアドレス" -Force
   # または全ホスト許可（テスト用）
   Set-Item WSMan:\localhost\Client\TrustedHosts -Value "*" -Force
   ```

2. **LinuxにKerberosクライアントをインストール**（IT制限環境では困難）
   ```bash
   # RHEL/CentOS
   sudo yum install krb5-workstation krb5-libs
   # Ubuntu/Debian
   sudo apt install krb5-user libkrb5-dev
   ```

#### 2. TrustedHostsが空の場合

**現象**: WindowsクライアントからはKerberosで接続可能だが、LinuxからのNTLM接続は拒否される

**確認コマンド**:
```powershell
winrm get winrm/config/client
# TrustedHosts が空の場合、リモートNTLM認証が制限される
```

**解決策**:
```powershell
# 元に戻す場合は空文字列を設定
Set-Item WSMan:\localhost\Client\TrustedHosts -Value "" -Force
```

#### 3. IT制限環境での制約

以下の制限がある環境では、LinuxからのWinRM接続が困難な場合があります：

| 制限事項 | 影響 |
|---------|------|
| krb5パッケージインストール不可 | Kerberos認証が使用できない |
| サーバー設定変更不可 | TrustedHosts設定ができない |
| AllowUnencrypted = false | HTTP接続で暗号化が必要 |
| HTTPS (5986) 無効 | 暗号化接続が使用できない |

**回避策**:
- IT部門にTrustedHostsの設定変更を依頼
- Windowsの踏み台サーバーを経由して接続
- SSH for Windowsが有効な場合はSSHを使用

### NTLMv2計算の検証結果

本プログラムのNTLMv2実装は、MS-NLMPのテストベクトル（セクション4.2.4）で検証済みです：
- NT Hash計算: OK
- NTLMv2 Hash計算: OK
- HMAC-MD5、MD4、MD5: 全て正常動作

認証が失敗する場合、上記の環境設定の問題を確認してください。

## セキュリティに関する注意事項

### 本番環境での推奨設定

1. **HTTPS接続を使用**: HTTP（ポート5985）ではなくHTTPS（ポート5986）を使用
2. **基本認証を無効化**: NTLM認証またはKerberos認証を使用
3. **TrustedHostsを限定**: すべてのホスト（`*`）ではなく、特定のIPのみ許可
4. **ファイアウォールで接続元を制限**: 特定のLinuxサーバのIPのみ許可
5. **強力なパスワードを使用**: 管理者アカウントに複雑なパスワードを設定
6. **最小権限の原則**: WinRM用に専用の管理者アカウントを作成

### パスワード管理

スクリプト内にパスワードを直接記載せず、以下の方法を推奨します：

```bash
# 環境変数で指定
export WINRM_PASSWORD="SecretPassword"
python3 winrm_exec.py

# 外部ファイルから読み込み（ファイルは適切にパーミッション設定）
echo "SecretPassword" > /secure/path/.winrm_pass
chmod 600 /secure/path/.winrm_pass
export WINRM_PASSWORD=$(cat /secure/path/.winrm_pass)
```

## ライセンス

このツールはMITライセンスの下で提供されています。詳細は[LICENSE](../LICENSE)を参照してください。

## 参考資料

- [Microsoft WinRM Documentation](https://learn.microsoft.com/en-us/windows/win32/winrm/portal)
- [pywinrm Documentation](https://github.com/diyan/pywinrm)
- [WS-Management Protocol](https://www.dmtf.org/standards/ws-man)
