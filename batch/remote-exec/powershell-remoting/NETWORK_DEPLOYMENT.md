# サーバ上のPowerShellスクリプトをローカルから実行する方法

## 概要

PowerShellスクリプト（.ps1）をサーバ上に配置し、バッチファイルやVBScriptをローカルPCに配置して実行する方法を説明します。

---

## 📁 配置例

```
【サーバ側】\\192.168.1.100\Share\Scripts\
└── Invoke-RemoteBatch.ps1       ← PowerShellスクリプト本体

【ローカルPC】C:\Tools\
├── RemoteExec_NetworkPath.vbs   ← ネットワークパス版VBS
├── RemoteExec_CopyAndRun.vbs    ← コピー実行版VBS
└── remote_exec_network.bat      ← ネットワークパス版バッチ
```

---

## ✅ 方法1: ネットワークパスを直接指定（最も簡単）

### 仕組み

```
ローカルPC
  ↓ ダブルクリック
RemoteExec_NetworkPath.vbs
  ↓ ネットワーク経由で実行
\\Server\Share\Invoke-RemoteBatch.ps1
  ↓ リモート実行
リモートサーバのバッチファイル
```

### 使い方

#### ステップ1: サーバにスクリプトを配置

```
\\192.168.1.100\Share\Scripts\Invoke-RemoteBatch.ps1
```

または共有フォルダ：

```
\\FileServer\公開\Scripts\Invoke-RemoteBatch.ps1
```

#### ステップ2: ローカルでVBSを編集

`RemoteExec_NetworkPath.vbs` を編集：

```vbscript
' ネットワークパス上のPowerShellスクリプト
networkScriptPath = "\\192.168.1.100\Share\Scripts\Invoke-RemoteBatch.ps1"

' リモート実行の設定
computerName = "192.168.1.100"
userName = "Administrator"
batchPath = "C:\Scripts\backup.bat"
```

#### ステップ3: ダブルクリックで実行

`RemoteExec_NetworkPath.vbs` をダブルクリック

### メリット

- ✅ スクリプトを一元管理できる
- ✅ スクリプト更新時、サーバ側だけ変更すればOK
- ✅ 複数のPCから同じスクリプトを使用可能

### デメリット

- ❌ ネットワーク接続が必須
- ❌ 共有フォルダへのアクセス権限が必要

---

## ✅ 方法2: 一時コピーして実行（より安全）

### 仕組み

```
ローカルPC
  ↓ ダブルクリック
RemoteExec_CopyAndRun.vbs
  ↓ ① ネットワーク経由でコピー
\\Server\Share\Invoke-RemoteBatch.ps1
  ↓ ② ローカル一時フォルダにコピー
C:\Users\xxx\AppData\Local\Temp\Invoke-RemoteBatch.ps1
  ↓ ③ ローカルで実行
  ↓ ④ リモート実行
リモートサーバのバッチファイル
  ↓ ⑤ 一時ファイル削除
```

### 使い方

`RemoteExec_CopyAndRun.vbs` を使用（編集方法は方法1と同じ）

### メリット

- ✅ 実行中にネットワークが切断されても影響なし
- ✅ スクリプトの一元管理
- ✅ 一時ファイルは自動削除

### デメリット

- ❌ 毎回コピーが発生（わずかに遅い）

---

## ✅ 方法3: バッチファイル版

### 使い方

`remote_exec_network.bat` を編集：

```batch
rem ネットワークパス上のPowerShellスクリプト
set NETWORK_PS_SCRIPT=\\192.168.1.100\Share\Scripts\Invoke-RemoteBatch.ps1

rem リモートサーバの設定
set REMOTE_SERVER=192.168.1.100
set REMOTE_USER=Administrator
set REMOTE_BATCH_PATH=C:\Scripts\backup.bat
```

ダブルクリックで実行

---

## 🔒 必要な設定

### 1. サーバ側：共有フォルダの設定

#### 方法A: 既存の共有フォルダを使用

```
\\192.168.1.100\Share\Scripts\
\\FileServer\公開\Scripts\
```

#### 方法B: 新しい共有フォルダを作成

1. サーバでフォルダを作成（例: `C:\Scripts`）
2. フォルダを右クリック → 「プロパティ」
3. 「共有」タブ → 「詳細な共有」
4. 「このフォルダーを共有する」をチェック
5. 共有名を設定（例: `Scripts`）
6. 「アクセス許可」で適切なユーザーを追加
   - 読み取り権限のみでOK

**結果**: `\\サーバ名\Scripts\` でアクセス可能

### 2. ローカルPC側：アクセス権限

#### 確認方法

```cmd
# ネットワークパスにアクセスできるか確認
dir \\192.168.1.100\Share\Scripts\
```

#### アクセスできない場合

##### 方法A: ネットワークドライブをマップ

1. エクスプローラーを開く
2. 「PC」を右クリック → 「ネットワークドライブの割り当て」
3. ドライブ文字を選択（例: Z:）
4. フォルダー: `\\192.168.1.100\Share`
5. ユーザー名・パスワードを入力

**結果**: `Z:\Scripts\Invoke-RemoteBatch.ps1` としてアクセス可能

VBSの設定を変更：
```vbscript
networkScriptPath = "Z:\Scripts\Invoke-RemoteBatch.ps1"
```

##### 方法B: net useコマンドで接続

```cmd
net use \\192.168.1.100\Share /user:Administrator password
```

---

## 📊 3つの方法の比較

| 方法 | スクリプト配置 | ネットワーク依存 | 速度 | 推奨度 |
|-----|--------------|----------------|------|--------|
| **ネットワークパス直接** | サーバ | 高（実行中も必要） | ⭐⭐⭐ | ✅ 推奨 |
| **一時コピー実行** | サーバ | 低（開始時のみ） | ⭐⭐ | ✅ 推奨 |
| **バッチファイル版** | サーバ | 高（実行中も必要） | ⭐⭐⭐ | ✅ 推奨 |

---

## 🎯 使い分けガイド

### ケース1: スクリプトを一元管理したい

**推奨**: ネットワークパス直接実行

```
サーバ: \\FileServer\Scripts\Invoke-RemoteBatch.ps1
ローカル: RemoteExec_NetworkPath.vbs（各PC）
```

**メリット**:
- スクリプト更新時、サーバ側のみ変更
- 複数PCで同じバージョンを使用

---

### ケース2: ネットワークが不安定

**推奨**: 一時コピー実行

```
ローカル: RemoteExec_CopyAndRun.vbs
```

**メリット**:
- 実行開始後はネットワーク不要
- より安定

---

### ケース3: 完全ローカルで実行したい

**推奨**: 従来の方法（スクリプトもローカルに配置）

```
ローカル:
  ├── Invoke-RemoteBatch.ps1
  └── RemoteExec_GUI.vbs
```

---

## 🔧 トラブルシューティング

### エラー: "PowerShellスクリプトが見つかりません"

**原因**: ネットワークパスにアクセスできない

**対処法**:

1. パスが正しいか確認
   ```cmd
   dir \\192.168.1.100\Share\Scripts\
   ```

2. アクセス権限を確認
   ```cmd
   net use \\192.168.1.100\Share
   ```

3. ネットワークドライブをマップ
   ```
   Z:\Scripts\Invoke-RemoteBatch.ps1
   ```

---

### エラー: "アクセスが拒否されました"

**原因**: 共有フォルダへのアクセス権限がない

**対処法**:

1. サーバ側で共有フォルダの権限を確認
   - フォルダを右クリック → プロパティ → 共有 → 詳細な共有 → アクセス許可

2. 明示的に認証情報を指定
   ```cmd
   net use \\192.168.1.100\Share /user:Administrator password
   ```

---

### エラー: "スクリプトのコピーに失敗しました"（コピー実行版）

**原因**: ネットワーク接続の問題または一時フォルダの権限

**対処法**:

1. ネットワーク接続を確認
2. 一時フォルダの権限を確認
   ```cmd
   echo %TEMP%
   dir %TEMP%
   ```

---

## 💡 実践例

### 例1: 複数PCから同じスクリプトを使用

**サーバ配置**:
```
\\FileServer\Scripts\
└── Invoke-RemoteBatch.ps1
```

**各PCに配置**（カスタマイズ可能）:
```
PC1: RemoteExec_Server01.vbs  ← サーバ01用の設定
PC2: RemoteExec_Server02.vbs  ← サーバ02用の設定
PC3: RemoteExec_Backup.vbs    ← バックアップ用の設定
```

すべて同じ `Invoke-RemoteBatch.ps1` を参照

---

### 例2: スクリプトのバージョン管理

**サーバ配置**:
```
\\FileServer\Scripts\
├── v1.0\
│   └── Invoke-RemoteBatch.ps1
└── v2.0\
    └── Invoke-RemoteBatch.ps1
```

**VBSで切り替え**:
```vbscript
' v1.0を使用
networkScriptPath = "\\FileServer\Scripts\v1.0\Invoke-RemoteBatch.ps1"

' v2.0にアップグレード（VBSの1行を変更するだけ）
networkScriptPath = "\\FileServer\Scripts\v2.0\Invoke-RemoteBatch.ps1"
```

---

## まとめ

### 結論

**サーバ上の.ps1をローカルのVBS/Batから実行することは可能です！**

### 推奨構成

```
【サーバ】\\192.168.1.100\Share\Scripts\
└── Invoke-RemoteBatch.ps1          ← 一元管理

【ローカルPC】C:\RemoteTools\
├── RemoteExec_NetworkPath.vbs      ← ダブルクリックで実行
├── RemoteExec_CopyAndRun.vbs       ← より安全な方法
└── remote_exec_network.bat         ← バッチ版
```

### メリット

- ✅ スクリプトの一元管理
- ✅ 更新時はサーバ側のみ変更
- ✅ 複数PCから同じスクリプトを使用
- ✅ ローカルには小さなVBS/Batのみ

### 注意点

- ⚠️ ネットワーク接続が必要
- ⚠️ 共有フォルダへのアクセス権限が必要
- ⚠️ UNCパス（`\\サーバ\共有\`）が使える環境が必要

この方法を使えば、スクリプトの管理が格段に楽になります！
