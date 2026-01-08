#!/bin/bash
# Web版Claude Code専用のセットアップスクリプト
# ローカル環境では実行されません

# Web版かどうかを判定
if [ "$CLAUDE_CODE_REMOTE" != "true" ]; then
  exit 0  # ローカル環境では何もしない
fi

echo "Web版Claude Code環境を検出しました。セットアップを実行します..."

# gh CLIがインストールされているか確認
if ! command -v gh &> /dev/null; then
  echo "GitHub CLI (gh) をインストールしています..."
  apt-get update && apt-get install -y gh

  if command -v gh &> /dev/null; then
    echo "GitHub CLI のインストールが完了しました: $(gh --version | head -1)"
  else
    echo "GitHub CLI のインストールに失敗しました"
    exit 1
  fi
else
  echo "GitHub CLI は既にインストールされています: $(gh --version | head -1)"
fi

echo "セットアップが完了しました"
