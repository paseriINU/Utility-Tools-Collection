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
    echo "使い方: $0 REMOTE_REPO_PATH"
    echo ""
    echo "例:"
    echo "  $0 /path/to/remote/repo.git"
    echo "  $0 /srv/git/myproject.git"
    exit 1
fi

REMOTE_REPO="$1"

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
echo

# フックディレクトリの存在確認
if [ ! -d "$HOOKS_DIR" ]; then
    echo -e "${YELLOW}フックディレクトリが存在しません。作成します...${NC}"
    mkdir -p "$HOOKS_DIR"
fi

# スクリプトのディレクトリを取得
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# updateフックのインストール
echo -e "${BLUE}[1/2] updateフックをインストール中...${NC}"

if [ -f "$HOOKS_DIR/update" ]; then
    echo -e "${YELLOW}既存のupdateフックが見つかりました${NC}"
    echo -n "バックアップを作成してインストールしますか？ (y/n): "
    read -r answer
    if [ "$answer" != "y" ] && [ "$answer" != "Y" ]; then
        echo -e "${YELLOW}インストールをスキップしました${NC}"
        exit 0
    fi

    # バックアップ作成
    backup_file="$HOOKS_DIR/update.backup.$(date +%Y%m%d_%H%M%S)"
    cp "$HOOKS_DIR/update" "$backup_file"
    echo -e "${GREEN}バックアップを作成しました: $backup_file${NC}"
fi

# フックをコピー
cp "$SCRIPT_DIR/update" "$HOOKS_DIR/update"
chmod +x "$HOOKS_DIR/update"
echo -e "${GREEN}✓ updateフックをインストールしました${NC}"

echo
echo "======================================"
echo -e "${GREEN}インストール完了！${NC}"
echo "======================================"
echo
echo "有効化された保護:"
echo "  • master/mainブランチの削除防止"
echo "  • master/mainブランチへの直接プッシュ防止"
echo
echo "動作確認:"
echo "  1. 別のマシンからリポジトリをクローン"
echo "  2. masterブランチで変更を試みてプッシュ"
echo "  3. エラーメッセージが表示されることを確認"
echo
echo "フックの無効化:"
echo "  rm $HOOKS_DIR/update"
echo
