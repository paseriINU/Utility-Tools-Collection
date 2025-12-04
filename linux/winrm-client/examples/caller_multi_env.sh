#!/bin/bash
# -*- coding: utf-8 -*-
#
# WinRM Executor 複数環境実行例
# 複数の環境をループで順次実行する
#

# スクリプトのディレクトリを取得
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PARENT_DIR="$(dirname "$SCRIPT_DIR")"

# 色付き出力用
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# 実行する環境のリスト
ENVIRONMENTS=("TST1T" "TST2T")

# 使用するスクリプト（Bash版またはPython版）
# EXECUTOR="${PARENT_DIR}/winrm_exec.sh"      # Bash版を使用する場合
EXECUTOR="${PARENT_DIR}/winrm_exec.py"       # Python版を使用する場合

# Python版の場合はpython3コマンドを追加
if [[ "$EXECUTOR" == *.py ]]; then
    EXECUTOR_CMD="python3 $EXECUTOR"
else
    EXECUTOR_CMD="$EXECUTOR"
fi

echo "======================================"
echo "  WinRM Executor 複数環境実行"
echo "======================================"
echo -e "${BLUE}実行スクリプト: $(basename $EXECUTOR)${NC}"
echo -e "${BLUE}対象環境: ${ENVIRONMENTS[*]}${NC}"
echo

# 実行結果を記録する変数
declare -A results
total=0
success=0
failed=0

# 各環境で順次実行
for env in "${ENVIRONMENTS[@]}"; do
    total=$((total + 1))

    echo "======================================"
    echo -e "${YELLOW}[${total}/${#ENVIRONMENTS[@]}] 環境: ${env}${NC}"
    echo "======================================"

    # スクリプト実行
    if eval "$EXECUTOR_CMD -e $env"; then
        results[$env]="成功"
        success=$((success + 1))
        echo -e "${GREEN}✓ ${env} 環境の実行が完了しました${NC}"
    else
        results[$env]="失敗"
        failed=$((failed + 1))
        echo -e "${RED}✗ ${env} 環境の実行が失敗しました${NC}"
    fi

    echo

    # 次の環境実行前に少し待機（オプション）
    if [ $total -lt ${#ENVIRONMENTS[@]} ]; then
        sleep 2
    fi
done

# 実行結果のサマリーを表示
echo "======================================"
echo "  実行結果サマリー"
echo "======================================"
echo -e "${BLUE}総実行数: ${total}${NC}"
echo -e "${GREEN}成功: ${success}${NC}"
echo -e "${RED}失敗: ${failed}${NC}"
echo

echo "環境別の結果:"
for env in "${ENVIRONMENTS[@]}"; do
    if [ "${results[$env]}" = "成功" ]; then
        echo -e "  ${env}: ${GREEN}${results[$env]}${NC}"
    else
        echo -e "  ${env}: ${RED}${results[$env]}${NC}"
    fi
done

echo "======================================"

# 終了コード（すべて成功なら0、1つでも失敗があれば1）
if [ $failed -gt 0 ]; then
    exit 1
else
    exit 0
fi
