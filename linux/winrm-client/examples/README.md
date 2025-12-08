# WinRM Executor 呼び出し例

このディレクトリには、`winrm_exec.sh` (Bash版)、`winrm_exec.py` (Python版)、`winrm_exec` (C言語版) を他のスクリプトから呼び出す例が含まれています。

## ファイル一覧

### caller_example.sh
基本的な呼び出し例です。Bash版、Python版、C言語版を使って、異なる環境を指定して実行します。

**実行方法:**
```bash
cd /path/to/linux/winrm-client/examples
./caller_example.sh
```

**動作内容:**
1. Bash版で TST1T 環境を指定して実行
2. Bash版で TST2T 環境を指定して実行
3. Python版で TST1T 環境を指定して実行
4. Python版で TST2T 環境を指定して実行
5. C言語版で TST1T 環境を指定して実行
6. C言語版で TST2T 環境を指定して実行

## 使用例

### シンプルな呼び出し
```bash
#!/bin/bash
# 単一環境を実行

# Bash版
/path/to/winrm_exec.sh TST1T

# Python版
python3 /path/to/winrm_exec.py TST1T

# C言語版
/path/to/winrm_exec TST1T
```

### ループ処理
```bash
#!/bin/bash
# 複数環境をループ実行
for env in TST1T TST2T TST3T; do
    echo "環境 $env を実行中..."
    /path/to/winrm_exec.sh "$env"
done
```

### 条件分岐
```bash
#!/bin/bash
# 環境を条件で選択
if [ "$(date +%u)" -le 5 ]; then
    # 平日は TST1T
    ENV="TST1T"
else
    # 週末は TST2T
    ENV="TST2T"
fi

/path/to/winrm_exec.sh "$ENV"
```

### 並列実行
```bash
#!/bin/bash
# 複数環境を並列実行（バックグラウンド実行）
/path/to/winrm_exec.sh TST1T &
/path/to/winrm_exec.sh TST2T &

# すべてのバックグラウンドジョブの完了を待つ
wait

echo "すべての環境の実行が完了しました"
```

## 注意事項

1. **実行前の設定確認**
   - 各スクリプトの設定セクションを事前に編集してください
   - Windows接続情報（ホスト、ユーザー、パスワード）を正しく設定してください

2. **環境リストの整合性**
   - 呼び出し元で指定する環境名は、`ENVIRONMENTS` リストに定義されている必要があります
   - 未定義の環境を指定するとエラーになります

3. **各版の使用条件**
   - **Bash版**: curl が必要
   - **Python版**: Python 3.6以降が必要
   - **C言語版**: 事前にコンパイルが必要（`gcc -o winrm_exec winrm_exec.c`）

4. **エラーハンドリング**
   - 本番環境で使用する場合は、適切なエラーハンドリングとログ記録を追加してください
   - 失敗時の通知（メール、Slack など）を実装することを推奨します

## トラブルシューティング

### 実行権限エラー
```bash
chmod +x caller_example.sh
```

### パスエラー
スクリプトのパスが正しいか確認してください：
```bash
ls -la ../winrm_exec.sh
ls -la ../winrm_exec.py
ls -la ../winrm_exec      # C言語版（コンパイル後）
```

### 環境指定エラー
利用可能な環境を確認してください：
```bash
# Bash版
../winrm_exec.sh --help

# Python版
python3 ../winrm_exec.py --help

# C言語版（引数なしで実行するとヘルプ表示）
../winrm_exec
```
