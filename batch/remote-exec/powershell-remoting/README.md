# PowerShell Remoting リモートバッチ実行ツール

## 概要

リモートWindowsサーバ上でバッチファイルを実行するためのハイブリッドスクリプトです。
.batファイル単体で動作し、PowerShell Remotingを使用してリモート実行を行います。

## 主な機能

- ✅ **管理者権限の自動昇格** - 必要に応じて自動的に管理者として再起動
- ✅ **WinRM自動設定** - TrustedHostsとWinRMサービスを自動設定・復元
- ✅ **環境選択機能** - tst1t/tst2t環境を実行時に選択可能
- ✅ **日時付きログ自動生成** - 実行結果を自動的にログファイルに保存
- ✅ **エラー時の自動復元** - 途中でエラーが発生しても設定を元に戻す
- ✅ **SSL/HTTPS対応** - セキュアな通信にも対応

## 必要な環境

### ローカル環境（実行元）
- Windows 7 以降
- PowerShell 5.1 以降
- 管理者権限（スクリプトが自動的に要求します）

### リモート環境（実行先）
- Windows Server 2012 以降（または Windows 7 以降）
- WinRM が有効化されていること
- リモート実行を許可する設定

## 使い方

### 1. 設定の編集

`Invoke-RemoteBatch-Hybrid.bat` をテキストエディタで開き、以下の設定を編集します：

```powershell
$Config = @{
    # リモートサーバのIPアドレスまたはホスト名
    ComputerName = "192.168.1.100"

    # 認証情報
    UserName = "Administrator"
    Password = ""  # 空の場合は実行時に入力を求められます

    # 実行するバッチファイルのパス（リモートサーバ上のパス）
    # {env}の部分が実行時に選択した環境（tst1t/tst2t）に置換されます
    BatchPath = "C:\Scripts\{env}\test.bat"

    # バッチファイルに渡す引数（オプション）
    Arguments = ""

    # SSL/HTTPS接続を使用する場合は $true
    UseSSL = $false
}
```

### 2. 実行

バッチファイルをダブルクリックするだけです。

```batch
Invoke-RemoteBatch-Hybrid.bat
```

### 3. 実行フロー

1. 管理者権限チェック → 必要に応じてUACプロンプト表示
2. 環境選択（BatchPathに{env}が含まれる場合）
   ```
   実行環境を選択してください:
     1. tst1t
     2. tst2t
   選択 (1-2):
   ```
3. WinRM設定の自動構成
4. パスワード入力（Passwordが空の場合）
5. リモートサーバに接続してバッチ実行
6. 実行結果の表示とログ保存
7. WinRM設定の復元

## ログファイル

実行結果は自動的にログファイルに保存されます：

- **保存場所**: バッチファイルと同じディレクトリ
- **ファイル名**: `RemoteBatch_YYYYMMDD_HHMMSS.log`
- **例**: `RemoteBatch_20251203_153045.log`

## 設定のカスタマイズ

### 環境選択機能を使用しない場合

BatchPathを固定値に設定します：

```powershell
BatchPath = "C:\Scripts\production\daily_batch.bat"
```

### パスワードを設定ファイルに記述する場合

**⚠️ セキュリティ注意**: パスワードを平文で保存することになります

```powershell
UserName = "Administrator"
Password = "YourPassword123"
```

### SSL/HTTPS接続を使用する場合

```powershell
UseSSL = $true
```

ポート 5986 が使用されます（通常は 5985）。

## WinRM設定について

このスクリプトは以下のWinRM設定を自動的に行います：

1. **WinRMサービスの起動**
   - 停止している場合のみ起動
   - 終了時に元の状態（停止）に戻す

2. **TrustedHostsの設定**
   - 接続先をTrustedHostsに追加
   - 終了時に元の設定に復元

### 手動でWinRMを設定する場合

スクリプトを使用せず、手動で設定する場合：

```powershell
# WinRMサービスの起動
Start-Service WinRM

# TrustedHostsの設定
Set-Item WSMan:\localhost\Client\TrustedHosts -Value "192.168.1.100" -Force

# 確認
Get-Item WSMan:\localhost\Client\TrustedHosts
```

## リモートサーバの準備

リモートサーバ側でWinRMを有効化する必要があります：

```powershell
# 管理者権限のPowerShellで実行
Enable-PSRemoting -Force

# ファイアウォールの設定を確認
Get-NetFirewallRule -Name "WINRM-HTTP-In-TCP" | Format-List
```

## トラブルシューティング

### エラー: WinRM設定の自動構成に失敗しました

**原因**: 管理者権限がない、またはWinRMサービスの起動に失敗

**解決方法**:
- スクリプトを右クリック → 「管理者として実行」
- WinRMサービスが無効化されていないか確認

### エラー: リモートサーバに接続できません

**原因**: ネットワーク接続の問題、またはリモートサーバのWinRM設定

**解決方法**:
1. ネットワーク接続を確認
   ```powershell
   Test-Connection -ComputerName 192.168.1.100
   ```

2. WinRM接続テスト
   ```powershell
   Test-WSMan -ComputerName 192.168.1.100
   ```

3. リモートサーバでWinRMが有効か確認
   ```powershell
   # リモートサーバで実行
   Get-Service WinRM
   ```

### エラー: 認証に失敗しました

**原因**: ユーザー名またはパスワードが正しくない

**解決方法**:
- ユーザー名とパスワードを確認
- ドメイン環境の場合は `ドメイン名\ユーザー名` 形式で指定

### TrustedHostsの設定が復元されない

**原因**: スクリプトが異常終了した

**解決方法**:
手動で設定を確認・復元
```powershell
# 現在の設定を確認
Get-Item WSMan:\localhost\Client\TrustedHosts

# 空に戻す
Set-Item WSMan:\localhost\Client\TrustedHosts -Value "" -Force
```

## セキュリティに関する注意

1. **パスワードの保存**
   - パスワードを設定ファイルに平文で保存しないことを推奨
   - 実行時入力（Password = ""）を使用することを推奨

2. **TrustedHosts設定**
   - スクリプトは実行後に自動的に設定を復元します
   - 手動で設定する場合は、必要最小限のホストのみを追加

3. **管理者権限**
   - このスクリプトは管理者権限で実行されます
   - 信頼できるスクリプトのみを実行してください

4. **SSL/HTTPS接続**
   - 本番環境ではSSL/HTTPS接続の使用を推奨
   - 証明書の設定が必要です

## ライセンス

このスクリプトはMITライセンスの下で公開されています。

## 作成日

2025-12-03

## バージョン履歴

- **v2.0** (2025-12-03)
  - ハイブリッド.bat形式に統合
  - 管理者権限の自動昇格機能を追加
  - WinRM自動設定・復元機能を追加
  - 環境選択機能を追加
  - エラー時の自動復元機能を追加
