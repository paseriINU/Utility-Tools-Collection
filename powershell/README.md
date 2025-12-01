# PowerShell Scripts

PowerShellで作成したWindows管理・自動化スクリプト集です。

## 📂 サブフォルダ

### scripts/
汎用的なPowerShellスクリプトを格納します。

- システム管理の自動化
- ファイル・フォルダ操作
- Active Directory管理
- ネットワーク設定
- タスクスケジューラ連携

## 必要な環境

- Windows 10/11
- PowerShell 5.1以降（Windows標準で搭載）
- PowerShell 7.x（推奨）

## 使い方

### 実行方法

1. PowerShellを管理者権限で起動（必要に応じて）
2. スクリプトのディレクトリに移動
3. 実行ポリシーを確認
   ```powershell
   Get-ExecutionPolicy
   ```
4. スクリプトを実行
   ```powershell
   .\script_name.ps1
   ```

### 実行ポリシーについて

初めて実行する場合、実行ポリシーの変更が必要な場合があります:

```powershell
# 現在のユーザーのみ実行を許可
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

または、一時的に実行:

```powershell
powershell -ExecutionPolicy Bypass -File .\script_name.ps1
```

## 注意事項

- スクリプトによっては管理者権限が必要です
- 実行前にスクリプトの内容を必ず確認してください
- システム設定を変更するスクリプトは慎重に使用してください

## 開発環境

- Visual Studio Code + PowerShell拡張機能を推奨
- Windows PowerShell ISEも利用可能

---

*現在、ツールは準備中です。今後追加予定。*
