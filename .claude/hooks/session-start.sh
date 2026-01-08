#!/bin/bash
# Web版Claude Code専用のセットアップスクリプト
# ローカル環境では実行されません

log() {
  echo "[gh-setup] $1"
}

# Web版かどうかを判定
if [ "$CLAUDE_CODE_REMOTE" != "true" ]; then
  log "Not a remote session, skipping gh setup"
  exit 0
fi

log "Web版Claude Code環境を検出しました。セットアップを実行します..."

# gh CLIがインストールされているか確認
if command -v gh &> /dev/null; then
  log "GitHub CLI は既にインストールされています: $(gh --version | head -1)"
  exit 0
fi

log "GitHub CLI (gh) をインストールしています..."
apt-get update && apt-get install -y gh

if command -v gh &> /dev/null; then
  log "GitHub CLI のインストールが完了しました: $(gh --version | head -1)"

  # CLAUDE_ENV_FILEにPATHを永続化
  if [ -n "$CLAUDE_ENV_FILE" ]; then
    echo "export PATH=\"/usr/bin:\$PATH\"" >> "$CLAUDE_ENV_FILE"
    log "PATH persisted to CLAUDE_ENV_FILE"
  fi
else
  log "GitHub CLI のインストールに失敗しました"
  exit 1
fi

log "セットアップが完了しました"
