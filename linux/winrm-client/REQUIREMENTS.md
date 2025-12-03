# 依存関係とシステム要件

このドキュメントでは、WinRM Client for Linuxの実行に必要な環境と依存関係を説明します。

## システム要件

### Linux側（クライアント）

#### 対応OS
- Red Hat Enterprise Linux (RHEL) 7以降
- CentOS 7以降
- Rocky Linux 8以降
- AlmaLinux 8以降
- Fedora 30以降
- Debian 9以降
- Ubuntu 18.04 LTS以降

#### 必要なコンポーネント

すべて**標準でインストール済み**のコンポーネントのみを使用します。

##### Python版（winrm_exec.py）

| コンポーネント | 用途 | 確認方法 |
|-------------|------|---------|
| Python 3.6以降 | スクリプト実行環境 | `python3 --version` |

**動作確認済みバージョン**:
- ✅ Python 3.6.8（RHEL 7 / CentOS 7標準）
- ✅ Python 3.7以降

**使用する標準ライブラリ**:
- `sys` - システム機能
- `argparse` - コマンドライン引数パース
- `logging` - ログ出力
- `base64` - Base64エンコード/デコード
- `uuid` - UUID生成
- `socket` - ソケット通信
- `ssl` - SSL/TLS接続
- `urllib.request` - HTTP通信
- `urllib.error` - HTTPエラー処理
- `xml.etree.ElementTree` - XMLパース

**追加パッケージは不要です。**

##### Bash版（winrm_exec.sh）

| コンポーネント | 用途 | 確認方法 | インストール方法 |
|-------------|------|---------|----------------|
| bash | シェルスクリプト実行 | `bash --version` | 標準でインストール済み |
| curl | HTTP通信 | `curl --version` | `sudo yum install curl` |
| base64 | Base64エンコード/デコード | `base64 --version` | 標準でインストール済み |
| date | タイムスタンプ生成 | `date --version` | 標準でインストール済み |
| grep | テキスト検索 | `grep --version` | 標準でインストール済み |

**注意**: `curl`は標準インストールされていない環境もあります。その場合は以下でインストール:

```bash
# Red Hat系
sudo yum install curl

# Debian/Ubuntu系
sudo apt install curl
```

### Windows側（サーバ）

- Windows Server 2022（またはWindows 10/11 Pro以上）
- PowerShell 5.1以降
- .NET Framework 4.5以降
- 管理者権限

## 依存関係の詳細

### 外部パッケージは不要

このツールは**追加のPythonパッケージやライブラリのインストールが不要**です。

以下のような外部パッケージは**使用していません**:
- ❌ `pywinrm` - 不要（WinRMプロトコルを標準ライブラリで実装）
- ❌ `requests` - 不要（`urllib`を使用）
- ❌ `xmltodict` - 不要（`xml.etree.ElementTree`を使用）
- ❌ `ntlm-auth` - 不要（Basic認証を使用）

### IT制限環境での使用

このツールは以下の制限がある環境で動作します:

✅ **インターネット接続不可の環境**
- すべての依存関係が標準ライブラリに含まれているため、pip等でのインストールが不要

✅ **パッケージ管理ツールの使用が制限されている環境**
- Python標準ライブラリのみ使用
- `pip install`が不要

✅ **管理者権限がない環境**
- Linuxクライアント側では一般ユーザー権限で実行可能
- （ただし、Windows側の設定には管理者権限が必要）

## 環境確認コマンド

### Python版の実行環境確認

```bash
# Pythonバージョン確認（3.6以降が必要）
python3 --version

# 標準ライブラリの確認（エラーが出なければOK）
python3 -c "import sys, argparse, logging, base64, uuid, socket, ssl, urllib.request, urllib.error, xml.etree.ElementTree"

# 実行可能か確認
python3 winrm_exec.py --help
```

### Bash版の実行環境確認

```bash
# 必要なコマンドの確認
bash --version
curl --version
base64 --version
date --version
grep --version

# すべてのコマンドが利用可能か一括確認
for cmd in bash curl base64 date grep; do
    if command -v $cmd &> /dev/null; then
        echo "✓ $cmd: OK"
    else
        echo "✗ $cmd: NG - インストールが必要"
    fi
done

# 実行権限の付与
chmod +x winrm_exec.sh

# ヘルプ表示テスト
./winrm_exec.sh
```

## Python 3がインストールされていない場合

多くのLinuxディストリビューションではPython 3が標準インストールされていますが、古いシステムでは以下でインストールできます:

```bash
# Red Hat系（RHEL 7/CentOS 7）
sudo yum install python3

# Red Hat系（RHEL 8/9, Rocky Linux, AlmaLinux）
sudo dnf install python3

# Debian/Ubuntu系
sudo apt update
sudo apt install python3
```

## curlがインストールされていない場合

```bash
# Red Hat系（RHEL 7/CentOS 7）
sudo yum install curl

# Red Hat系（RHEL 8/9, Rocky Linux, AlmaLinux）
sudo dnf install curl

# Debian/Ubuntu系
sudo apt update
sudo apt install curl
```

## ネットワーク要件

### ポート開放

Linux → Windows方向で以下のポートへの通信が必要です:

| プロトコル | ポート | 用途 |
|-----------|--------|------|
| HTTP | 5985 | WinRM (基本認証、開発環境向け) |
| HTTPS | 5986 | WinRM over SSL/TLS（推奨、本番環境向け） |

### ファイアウォール設定

Windows側でWinRMポートを開放する必要があります（詳細は`WINDOWS_SETUP.md`を参照）。

Linux側では通常、アウトバウンド通信は制限されていないため、ファイアウォール設定は不要です。

## セキュリティに関する注意

### Basic認証について

現在の実装では**Basic認証**を使用しています。

**重要な制限事項**:
- Basic認証はパスワードがBase64エンコードされるだけで、暗号化されません
- HTTP（ポート5985）では平文で送信されるため、**ネットワーク盗聴に対して脆弱**です

**推奨事項**:
1. **HTTPS（ポート5986）を使用**すること
   - 自己署名証明書でも可（`--no-cert-check`オプション使用）
2. **信頼できるネットワーク内でのみ使用**すること
   - 社内LAN、VPN経由など
3. 本番環境では**NTLM認証またはKerberos認証**の実装を検討すること
   - 現在の標準ライブラリ版ではBasic認証のみサポート

## トラブルシューティング

### Python版: ImportError

```bash
$ python3 winrm_exec.py
Traceback (most recent call last):
  File "winrm_exec.py", line 25, in <module>
    from urllib.request import Request, urlopen
ImportError: No module named urllib.request
```

**原因**: Python 2系を使用している

**解決策**:
```bash
# Python 3を明示的に使用
python3 winrm_exec.py

# Python 3がインストールされているか確認
which python3
python3 --version
```

### Bash版: curl: command not found

```bash
$ ./winrm_exec.sh
[ERROR] curlコマンドが見つかりません
```

**原因**: curlがインストールされていない

**解決策**:
```bash
sudo yum install curl  # Red Hat系
sudo apt install curl  # Debian/Ubuntu系
```

### Bash版: grep: invalid option -- 'P'

一部の古いシステムでは`grep -P`（Perl正規表現）がサポートされていない場合があります。

**原因**: 古いバージョンのgrep

**解決策**: Python版の使用を推奨

## パフォーマンスとリソース使用量

### Python版

- **メモリ使用量**: 約10-20MB
- **CPU使用率**: 低（ネットワーク待機がメイン）
- **実行速度**: 高速（ネイティブ実装）

### Bash版

- **メモリ使用量**: 約5-10MB
- **CPU使用率**: 低
- **実行速度**: Python版とほぼ同等

## まとめ

このツールは**IT制限環境でも使用できる**よう設計されています:

✅ 追加パッケージのインストール不要（Python標準ライブラリのみ）
✅ インターネット接続不要（依存関係なし）
✅ 一般ユーザー権限で実行可能（Linux側）
✅ 最小限のシステム要件（Python 3.6以降、またはbash + curl）

ご不明な点があれば`README.md`または`WINDOWS_SETUP.md`を参照してください。
