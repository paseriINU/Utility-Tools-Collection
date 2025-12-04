# WinRM Executor 呼び出し例

このディレクトリには、`winrm_exec.sh` (Bash版) と `winrm_exec.py` (Python版) を他のスクリプトから呼び出す例が含まれています。

## ファイル一覧

### 1. caller_example.sh
基本的な呼び出し例です。Bash版とPython版の両方を使って、異なる環境を指定して実行します。

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

### 2. caller_multi_env.sh
複数の環境をループで順次実行する例です。実行結果をサマリー形式で表示します。

**実行方法:**
```bash
cd /path/to/linux/winrm-client/examples
./caller_multi_env.sh
```

**動作内容:**
- ENVIRONMENTS 配列で定義された環境を順次実行
- 各環境の実行結果（成功/失敗）を記録
- 最後にサマリーを表示

**スクリプトの切り替え:**
ファイル内の以下の行を編集することで、Bash版とPython版を切り替えられます：
```bash
# Bash版を使用する場合
EXECUTOR="${PARENT_DIR}/winrm_exec.sh"

# Python版を使用する場合
EXECUTOR="${PARENT_DIR}/winrm_exec.py"
```

## カスタマイズ方法

### 環境リストの変更
複数環境実行スクリプトで実行する環境を変更する場合は、`ENVIRONMENTS` 配列を編集します：

```bash
# 例: 3つの環境を実行
ENVIRONMENTS=("TST1T" "TST2T" "TST3T")
```

### 実行間隔の調整
環境間の待機時間を調整する場合は、`sleep` コマンドの値を変更します：

```bash
# 5秒待機する場合
sleep 5
```

### エラー処理のカスタマイズ
失敗時の動作をカスタマイズする場合は、実行結果のチェック部分を編集します：

```bash
if eval "$EXECUTOR_CMD -e $env"; then
    # 成功時の処理
    echo "成功しました"
else
    # 失敗時の処理
    echo "失敗しました"
    # 必要に応じて中断
    # exit 1
fi
```

## 使用例

### シンプルな呼び出し
```bash
#!/bin/bash
# 単一環境を実行
/path/to/winrm_exec.sh TST1T
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
   - `winrm_exec.sh` または `winrm_exec.py` の設定セクションを事前に編集してください
   - Windows接続情報（ホスト、ユーザー、パスワード）を正しく設定してください

2. **環境リストの整合性**
   - 呼び出し元で指定する環境名は、`ENVIRONMENTS` リストに定義されている必要があります
   - 未定義の環境を指定するとエラーになります

3. **Python版の使用**
   - Python版を使用する場合は、Python 3.6以降がインストールされている必要があります
   - `python3` コマンドでPythonを実行できることを確認してください

4. **エラーハンドリング**
   - 本番環境で使用する場合は、適切なエラーハンドリングとログ記録を追加してください
   - 失敗時の通知（メール、Slack など）を実装することを推奨します

## トラブルシューティング

### 実行権限エラー
```bash
chmod +x caller_example.sh
chmod +x caller_multi_env.sh
```

### パスエラー
スクリプトのパスが正しいか確認してください：
```bash
ls -la ../winrm_exec.sh
ls -la ../winrm_exec.py
```

### 環境指定エラー
利用可能な環境を確認してください：
```bash
# Bash版
../winrm_exec.sh --help

# Python版
python3 ../winrm_exec.py --help
```
