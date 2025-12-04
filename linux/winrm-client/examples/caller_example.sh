#!/bin/bash
# -*- coding: utf-8 -*-
#
# WinRM Executor 呼び出し例
# winrm_exec.sh / winrm_exec.py を引数で環境を指定して呼び出す
#

# スクリプトのディレクトリを取得
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PARENT_DIR="$(dirname "$SCRIPT_DIR")"

# 色付き出力用
GREEN='\033[0;32m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

echo "======================================"
echo "  WinRM Executor 呼び出し例"
echo "======================================"
echo

# 使用例1: Bash版を使って TST1T 環境で実行
echo -e "${BLUE}[例1] Bash版で TST1T 環境を指定して実行${NC}"
"${PARENT_DIR}/winrm_exec.sh" -e TST1T

# 実行結果をチェック
if [ $? -eq 0 ]; then
    echo -e "${GREEN}✓ Bash版 TST1T 環境の実行が完了しました${NC}"
else
    echo -e "${RED}✗ Bash版 TST1T 環境の実行が失敗しました${NC}"
fi
echo

# 使用例2: Bash版を使って TST2T 環境で実行
echo -e "${BLUE}[例2] Bash版で TST2T 環境を指定して実行${NC}"
"${PARENT_DIR}/winrm_exec.sh" --env TST2T

if [ $? -eq 0 ]; then
    echo -e "${GREEN}✓ Bash版 TST2T 環境の実行が完了しました${NC}"
else
    echo -e "${RED}✗ Bash版 TST2T 環境の実行が失敗しました${NC}"
fi
echo

# 使用例3: Python版を使って TST1T 環境で実行
echo -e "${BLUE}[例3] Python版で TST1T 環境を指定して実行${NC}"
python3 "${PARENT_DIR}/winrm_exec.py" -e TST1T

if [ $? -eq 0 ]; then
    echo -e "${GREEN}✓ Python版 TST1T 環境の実行が完了しました${NC}"
else
    echo -e "${RED}✗ Python版 TST1T 環境の実行が失敗しました${NC}"
fi
echo

# 使用例4: Python版を使って TST2T 環境で実行
echo -e "${BLUE}[例4] Python版で TST2T 環境を指定して実行${NC}"
python3 "${PARENT_DIR}/winrm_exec.py" --env TST2T

if [ $? -eq 0 ]; then
    echo -e "${GREEN}✓ Python版 TST2T 環境の実行が完了しました${NC}"
else
    echo -e "${RED}✗ Python版 TST2T 環境の実行が失敗しました${NC}"
fi
echo

echo "======================================"
echo "  すべての実行が完了しました"
echo "======================================"
