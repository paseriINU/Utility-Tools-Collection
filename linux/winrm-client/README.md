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
WINDOWS_PORT = 5985                  # WinRMポート（HTTP）
WINDOWS_USER = "Administrator"       # Windowsユーザー名
WINDOWS_PASSWORD = "YourPassword"    # Windowsパスワード

# 実行するバッチファイル
BATCH_FILE_PATH = r"C:\Scripts\test.bat"
```

#### 2. 実行

```bash
# スクリプト内の設定で実行（pip installは不要）
python3 winrm_exec.py

# またはコマンドライン引数で設定を上書き
python3 winrm_exec.py --host 192.168.1.100 --user Administrator --password Pass123

# バッチファイルを指定して実行
python3 winrm_exec.py --host 192.168.1.100 --user Admin --password Pass123 \
    --batch "C:\Scripts\backup.bat"

# 直接コマンドを実行
python3 winrm_exec.py --host 192.168.1.100 --user Admin --password Pass123 \
    --command "echo Hello from Linux"

# HTTPS接続を使用
python3 winrm_exec.py --host 192.168.1.100 --user Admin --password Pass123 \
    --https --port 5986 --batch "C:\Scripts\test.bat"

# ログレベルを指定
python3 winrm_exec.py --log-level DEBUG --batch "C:\Scripts\test.bat"
```

#### 3. 環境変数で設定

```bash
# 環境変数で設定（パスワードをコマンドラインに残さない）
export WINRM_HOST=192.168.1.100
export WINRM_USER=Administrator
export WINRM_PASSWORD=SecretPassword

python3 winrm_exec.py --batch "C:\Scripts\test.bat"
```

#### オプション一覧

```
--host          Windows ServerのIPアドレスまたはホスト名
--port          WinRMポート（デフォルト: 5985）
--user          Windowsユーザー名
--password      Windowsパスワード
--batch         実行するバッチファイル（Windows側のパス）
--command       直接実行するコマンド
--https         HTTPS接続を使用
--no-cert-check 証明書検証を無効化（自己署名証明書の場合）
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
WINRM_HOST="${WINRM_HOST:-192.168.1.100}"
WINRM_PORT="${WINRM_PORT:-5985}"
WINRM_USER="${WINRM_USER:-Administrator}"
WINRM_PASS="${WINRM_PASS:-YourPassword}"

# 環境フォルダ名（実行時に選択可能）
ENV_FOLDER="${ENV_FOLDER:-TST1T}"

# 実行するバッチファイル（{ENV}は選択した環境に置換されます）
BATCH_FILE_PATH="${BATCH_FILE_PATH:-C:\\Scripts\\{ENV}\\test.bat}"
```

**環境選択機能**:
- スクリプト実行時に TST1T または TST2T を選択できます
- `{ENV}` プレースホルダーが選択した環境名に置換されます
- 例: TST1T を選択すると `C:\Scripts\TST1T\test.bat` が実行されます

#### 2. 実行権限の付与

```bash
chmod +x winrm_exec.sh
```

#### 3. 実行

```bash
# スクリプト内の設定で実行（環境選択メニューが表示されます）
./winrm_exec.sh

# 実行時の環境選択例:
# ======================================
#   環境選択
# ======================================
# 1. TST1T
# 2. TST2T
# ======================================
# 環境を選択してください (1 または 2) [デフォルト: 1]: 1

# または環境変数で設定を上書き
WINRM_HOST=192.168.1.100 WINRM_USER=Admin WINRM_PASS=Pass123 ./winrm_exec.sh

# デバッグモードで実行
DEBUG=true ./winrm_exec.sh
```

## 実行例

### Python版の実行例

```bash
$ python3 winrm_exec.py --host 192.168.1.100 --user Administrator --password Pass123 \
    --batch "C:\Scripts\hello.bat"

2025-01-15 10:30:45 [INFO] === WinRM Remote Batch Executor (標準ライブラリ版) ===
2025-01-15 10:30:45 [INFO] 接続先: 192.168.1.100:5985
2025-01-15 10:30:45 [INFO] ユーザー: Administrator
2025-01-15 10:30:45 [INFO] WinRMエンドポイント: http://192.168.1.100:5985/wsman
2025-01-15 10:30:45 [INFO] バッチファイル実行: C:\Scripts\hello.bat
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

**原因**: 認証情報が正しくない、またはBasic認証が無効

**解決策**:
```powershell
# Windows側でBasic認証が有効か確認
Get-Item WSMan:\localhost\Service\Auth\Basic

# 無効の場合は有効化（現在の実装ではBasic認証を使用）
Set-Item WSMan:\localhost\Service\Auth\Basic -Value $true
```

**注意**: 現在の実装では**Basic認証**を使用しています。セキュリティを考慮してHTTPS接続を使用することを推奨します。

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
