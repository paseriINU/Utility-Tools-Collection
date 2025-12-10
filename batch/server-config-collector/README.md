# サーバ構成情報収集ツール（完全版）

Windowsサーバの構成情報を包括的に収集してExcelファイルに出力するツールです。

## 収集項目（14シート）

| シート名 | 内容 |
|----------|------|
| **01_OS情報** | OS名、バージョン、ドメイン参加状況、シリアル番号 |
| **02_ハードウェア** | CPU詳細、メモリ容量・スロット情報 |
| **03_ディスク** | ドライブごとの容量・使用率・空き容量 |
| **04_ネットワーク設定** | IP、サブネット、ゲートウェイ、DNS、DHCP |
| **05_ネットワークプロファイル** | パブリック/プライベート/ドメイン |
| **06_WinRM設定** | サービス状態、TrustedHosts、認証設定、リスナー |
| **07_ファイアウォール** | プロファイル状態、WinRM関連ルール |
| **08_開いているポート** | LISTENING状態のTCPポート一覧 |
| **09_レジストリ設定** | UAC、リモートUAC、WinRM関連設定 |
| **10_Windowsサービス** | 全サービス一覧と状態 |
| **11_ユーザー_グループ** | ローカルユーザー、Administrators、Remote Desktop Users |
| **12_インストール済ソフト** | インストール済みプログラム一覧 |
| **13_共有フォルダ** | ネットワーク共有一覧 |
| **14_タスクスケジューラ** | 登録されたタスク一覧 |

## 使い方

1. `collect-server-config.bat` をダブルクリック
2. 収集完了後、Excelファイルがデスクトップに出力される
3. 「出力ファイルを開きますか？」で `y` を入力するとExcelが開く

## 設定オプション

設定セクションで各収集項目の有効/無効を切り替え可能:

```powershell
$Config = @{
    CollectOSInfo = $true
    CollectHardwareInfo = $true
    CollectNetworkConfig = $true
    CollectWinRMConfig = $true
    CollectFirewallRules = $true
    CollectOpenPorts = $true
    CollectRegistrySettings = $true
    CollectAllServices = $true
    CollectLocalUsers = $true
    CollectInstalledSoftware = $true
    CollectSharedFolders = $true
    CollectScheduledTasks = $true
}
```

## WinRM関連の確認ポイント

### サービス状態
- WinRMサービスが「Running」であること
- スタートアップ種別が「Automatic」であること

### ファイアウォール
- WinRM関連ルールが有効であること
- ポート5985（HTTP）または5986（HTTPS）が開いていること

### 認証設定
- TrustedHostsに接続元が含まれていること（ワークグループ環境）
- 必要な認証方式が有効であること

### レジストリ
- `LocalAccountTokenFilterPolicy` = 1（リモートUAC無効化、ワークグループ環境で必要）

## 動作環境

- Windows Server 2012 R2以降 / Windows 10以降
- PowerShell 5.1以降
- Excel（インストールされていない場合はCSV出力）

## 注意事項

- 管理者権限は不要ですが、一部の情報取得に制限がある場合があります
- Excelがインストールされていない環境ではCSV形式で出力されます
- 全サービス収集時は数百件のデータが出力されます
