#!/bin/bash
#
# Git Hooks インストールスクリプト
# リモートリポジトリにフックをインストールします
#

set -e  # エラー時に即座に終了

# 色付き出力用
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

echo "======================================"
echo "  Git Hooks インストーラー"
echo "======================================"
echo

# 引数チェック
if [ $# -eq 0 ]; then
    echo -e "${RED}エラー: リモートリポジトリのパスを指定してください${NC}"
    echo ""
    echo "使い方: $0 REMOTE_REPO_PATH [OPTIONS]"
    echo ""
    echo "オプション:"
    echo "  --all              すべてのフックをインストール（デフォルト）"
    echo "  --update-only      updateフックのみインストール"
    echo "  --pre-receive-only pre-receiveフックのみインストール"
    echo "  --commit-msg-only  commit-msgフックのみインストール"
    echo ""
    echo "例:"
    echo "  $0 /path/to/remote/repo.git"
    echo "  $0 /srv/git/myproject.git --all"
    echo "  $0 /path/to/repo.git --update-only"
    exit 1
fi

REMOTE_REPO="$1"
INSTALL_MODE="all"

# オプション解析
if [ $# -ge 2 ]; then
    case "$2" in
        --all)
            INSTALL_MODE="all"
            ;;
        --update-only)
            INSTALL_MODE="update"
            ;;
        --pre-receive-only)
            INSTALL_MODE="pre-receive"
            ;;
        --commit-msg-only)
            INSTALL_MODE="commit-msg"
            ;;
        *)
            echo -e "${RED}エラー: 不明なオプション: $2${NC}"
            exit 1
            ;;
    esac
fi

# リモートリポジトリの存在確認
if [ ! -d "$REMOTE_REPO" ]; then
    echo -e "${RED}エラー: リモートリポジトリが見つかりません: $REMOTE_REPO${NC}"
    exit 1
fi

# Gitリポジトリかチェック
if [ ! -d "$REMOTE_REPO/.git" ] && [ ! -f "$REMOTE_REPO/HEAD" ]; then
    echo -e "${RED}エラー: 指定されたパスはGitリポジトリではありません${NC}"
    exit 1
fi

# フックディレクトリの決定（bare repositoryの場合とそうでない場合）
if [ -f "$REMOTE_REPO/HEAD" ] && [ -d "$REMOTE_REPO/refs" ]; then
    # Bare repository
    HOOKS_DIR="$REMOTE_REPO/hooks"
else
    # 通常のリポジトリ
    HOOKS_DIR="$REMOTE_REPO/.git/hooks"
fi

echo -e "${BLUE}リモートリポジトリ:${NC} $REMOTE_REPO"
echo -e "${BLUE}フックディレクトリ:${NC} $HOOKS_DIR"
echo -e "${BLUE}インストールモード:${NC} $INSTALL_MODE"
echo

# フックディレクトリの存在確認
if [ ! -d "$HOOKS_DIR" ]; then
    echo -e "${YELLOW}フックディレクトリが存在しません。作成します...${NC}"
    mkdir -p "$HOOKS_DIR"
fi

# スクリプトのディレクトリを取得
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# フックインストール関数
install_hook() {
    local hook_name="$1"
    local hook_number="$2"
    local total_hooks="$3"

    echo -e "${BLUE}[$hook_number/$total_hooks] ${hook_name}フックをインストール中...${NC}"

    if [ ! -f "$SCRIPT_DIR/$hook_name" ]; then
        echo -e "${YELLOW}${hook_name}フックが見つかりません。スキップします。${NC}"
        return
    fi

    if [ -f "$HOOKS_DIR/$hook_name" ]; then
        echo -e "${YELLOW}既存の${hook_name}フックが見つかりました${NC}"
        echo -n "バックアップを作成してインストールしますか？ (y/n): "
        read -r answer
        if [ "$answer" != "y" ] && [ "$answer" != "Y" ]; then
            echo -e "${YELLOW}インストールをスキップしました${NC}"
            return
        fi

        # バックアップ作成
        backup_file="$HOOKS_DIR/${hook_name}.backup.$(date +%Y%m%d_%H%M%S)"
        cp "$HOOKS_DIR/$hook_name" "$backup_file"
        echo -e "${GREEN}バックアップを作成しました: $backup_file${NC}"
    fi

    # フックをコピー
    cp "$SCRIPT_DIR/$hook_name" "$HOOKS_DIR/$hook_name"
    chmod +x "$HOOKS_DIR/$hook_name"
    echo -e "${GREEN}✓ ${hook_name}フックをインストールしました${NC}"
    echo
}

# インストール実行
case "$INSTALL_MODE" in
    all)
        install_hook "update" 1 3
        install_hook "pre-receive" 2 3
        install_hook "commit-msg" 3 3
        ;;
    update)
        install_hook "update" 1 1
        ;;
    pre-receive)
        install_hook "pre-receive" 1 1
        ;;
    commit-msg)
        install_hook "commit-msg" 1 1
        ;;
esac

echo
echo "======================================"
echo -e "${GREEN}インストール完了！${NC}"
echo "======================================"
echo

if [ "$INSTALL_MODE" = "all" ] || [ "$INSTALL_MODE" = "update" ]; then
    echo "【updateフック】"
    echo "  • master/mainブランチの削除防止"
    echo "  • master/mainブランチへの直接プッシュ防止"
    echo
fi

if [ "$INSTALL_MODE" = "all" ] || [ "$INSTALL_MODE" = "pre-receive" ]; then
    echo "【pre-receiveフック】"
    echo "  • 機密情報の検出（パスワード、APIキー等）"
    echo "  • 大きなファイルの防止（10MB以上）"
    echo "  • 禁止ファイルタイプの検出（.env, credentials.json等）"
    echo
fi

if [ "$INSTALL_MODE" = "all" ] || [ "$INSTALL_MODE" = "commit-msg" ]; then
    echo "【commit-msgフック】"
    echo "  • コミットメッセージの最小文字数チェック"
    echo "  • Conventional Commits形式の検証（オプション）"
    echo "  • チケット番号の必須化（オプション）"
    echo "  注: クライアント側でも設定が必要"
    echo
fi

echo "動作確認:"
echo "  1. 別のマシンからリポジトリをクローン"
echo "  2. 各種制限をテスト"
echo "  3. エラーメッセージが表示されることを確認"
echo
echo "フックの無効化:"
echo "  rm $HOOKS_DIR/{update,pre-receive,commit-msg}"
echo
