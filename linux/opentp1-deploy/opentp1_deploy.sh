#!/bin/bash
#==============================================================================
# OpenTP1 デプロイ自動化ツール
#
# 概要:
#   OpenTP1を停止 → ソース配置 → OpenTP1を起動
#   の一連の流れを自動化するシェルスクリプト
#
# 使用方法:
#   1. 設定セクションを環境に合わせて編集
#   2. chmod +x opentp1_deploy.sh
#   3. ./opentp1_deploy.sh
#
#==============================================================================

#==============================================================================
# 設定セクション（環境に合わせて編集してください）
#==============================================================================

# OpenTP1のインストールパス
OPENTP1_HOME="/opt/OpenTP1"

# OpenTP1コマンドのパス（通常はOPENTP1_HOME/bin）
OPENTP1_BIN="${OPENTP1_HOME}/bin"

# コピー元ファイル（フルパスで指定）
SOURCE_FILE="/home/user/src/myprogram"

# 配置先ディレクトリ
DEPLOY_DIR="/opt/OpenTP1/aplib"

# バックアップを作成するか（true/false）
CREATE_BACKUP=true

# 停止待機時間（秒）
STOP_WAIT_TIME=10

# 起動待機時間（秒）
START_WAIT_TIME=10

#==============================================================================
# 以下は編集不要
#==============================================================================

# 色定義
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
CYAN='\033[0;36m'
NC='\033[0m' # No Color

# ログファイル
LOG_FILE="opentp1_deploy_$(date +%Y%m%d_%H%M%S).log"

# ファイル名を抽出
FILE_NAME=$(basename "${SOURCE_FILE}")

#------------------------------------------------------------------------------
# ログ出力関数
#------------------------------------------------------------------------------
log_info() {
    local message="[$(date '+%Y-%m-%d %H:%M:%S')] [INFO] $1"
    echo -e "${GREEN}${message}${NC}"
    echo "${message}" >> "${LOG_FILE}"
}

log_warn() {
    local message="[$(date '+%Y-%m-%d %H:%M:%S')] [WARN] $1"
    echo -e "${YELLOW}${message}${NC}"
    echo "${message}" >> "${LOG_FILE}"
}

log_error() {
    local message="[$(date '+%Y-%m-%d %H:%M:%S')] [ERROR] $1"
    echo -e "${RED}${message}${NC}"
    echo "${message}" >> "${LOG_FILE}"
}

log_step() {
    local message="$1"
    echo ""
    echo -e "${CYAN}======================================${NC}"
    echo -e "${CYAN}  ${message}${NC}"
    echo -e "${CYAN}======================================${NC}"
    echo ""
    echo "=====================================" >> "${LOG_FILE}"
    echo "  ${message}" >> "${LOG_FILE}"
    echo "=====================================" >> "${LOG_FILE}"
}

#------------------------------------------------------------------------------
# ヘッダー表示
#------------------------------------------------------------------------------
show_header() {
    echo ""
    echo -e "${CYAN}================================================================${NC}"
    echo -e "${CYAN}  OpenTP1 デプロイ自動化ツール${NC}"
    echo -e "${CYAN}================================================================${NC}"
    echo ""
    echo "実行日時    : $(date '+%Y-%m-%d %H:%M:%S')"
    echo "コピー元    : ${SOURCE_FILE}"
    echo "配置先      : ${DEPLOY_DIR}/${FILE_NAME}"
    echo "ログファイル: ${LOG_FILE}"
    echo ""
}

#------------------------------------------------------------------------------
# 事前チェック
#------------------------------------------------------------------------------
pre_check() {
    log_step "事前チェック"

    local error_count=0

    # OpenTP1コマンドの存在確認
    if [ ! -d "${OPENTP1_BIN}" ]; then
        log_error "OpenTP1のbinディレクトリが見つかりません: ${OPENTP1_BIN}"
        ((error_count++))
    else
        log_info "OpenTP1 bin: ${OPENTP1_BIN} [OK]"
    fi

    # コピー元ファイルの確認
    if [ ! -f "${SOURCE_FILE}" ]; then
        log_error "コピー元ファイルが見つかりません: ${SOURCE_FILE}"
        ((error_count++))
    else
        log_info "コピー元: ${SOURCE_FILE} [OK]"
    fi

    # 配置先ディレクトリの確認
    if [ ! -d "${DEPLOY_DIR}" ]; then
        log_warn "配置先ディレクトリが見つかりません: ${DEPLOY_DIR}"
        log_warn "配置時に作成を試みます"
    else
        log_info "配置先Dir: ${DEPLOY_DIR} [OK]"
    fi

    if [ ${error_count} -gt 0 ]; then
        log_error "事前チェックで ${error_count} 件のエラーがあります"
        return 1
    fi

    log_info "事前チェック完了"
    return 0
}

#------------------------------------------------------------------------------
# OpenTP1の状態確認
#------------------------------------------------------------------------------
check_opentp1_status() {
    if [ -x "${OPENTP1_BIN}/dcstatus" ]; then
        "${OPENTP1_BIN}/dcstatus" 2>/dev/null
        return $?
    else
        log_warn "dcstatusコマンドが見つかりません"
        return 1
    fi
}

#------------------------------------------------------------------------------
# OpenTP1の停止
#------------------------------------------------------------------------------
stop_opentp1() {
    log_step "OpenTP1 停止"

    # 現在の状態を確認
    log_info "OpenTP1の状態を確認中..."
    if check_opentp1_status | grep -q "ONLINE"; then
        log_info "OpenTP1は稼働中です。停止します..."
    else
        log_warn "OpenTP1は既に停止しているか、状態を確認できません"
        log_warn "停止コマンドを実行します..."
    fi

    # 停止コマンド実行
    log_info "dcstop -f を実行中..."
    if [ -x "${OPENTP1_BIN}/dcstop" ]; then
        "${OPENTP1_BIN}/dcstop" -f >> "${LOG_FILE}" 2>&1
        local exit_code=$?

        if [ ${exit_code} -ne 0 ]; then
            log_warn "dcstop の終了コード: ${exit_code}"
        fi
    else
        log_error "dcstopコマンドが見つかりません: ${OPENTP1_BIN}/dcstop"
        return 1
    fi

    # 停止待機
    log_info "${STOP_WAIT_TIME}秒待機中..."
    sleep ${STOP_WAIT_TIME}

    # 停止確認
    log_info "停止を確認中..."
    if check_opentp1_status 2>/dev/null | grep -q "OFFLINE\|停止"; then
        log_info "OpenTP1の停止を確認しました"
    else
        log_warn "OpenTP1の停止状態を確認できませんでした（処理を続行します）"
    fi

    return 0
}

#------------------------------------------------------------------------------
# デプロイ（ファイル配置）
#------------------------------------------------------------------------------
deploy_files() {
    log_step "ソース配置"

    # 配置先ディレクトリの作成（存在しない場合）
    if [ ! -d "${DEPLOY_DIR}" ]; then
        log_info "配置先ディレクトリを作成: ${DEPLOY_DIR}"
        mkdir -p "${DEPLOY_DIR}" || {
            log_error "ディレクトリの作成に失敗しました"
            return 1
        }
    fi

    # バックアップ作成
    if [ "${CREATE_BACKUP}" = true ] && [ -f "${DEPLOY_DIR}/${FILE_NAME}" ]; then
        local backup_name="${FILE_NAME}.bak.$(date +%Y%m%d_%H%M%S)"
        log_info "既存ファイルをバックアップ: ${backup_name}"
        cp "${DEPLOY_DIR}/${FILE_NAME}" "${DEPLOY_DIR}/${backup_name}" || {
            log_warn "バックアップの作成に失敗しました（処理を続行）"
        }
    fi

    # ファイルコピー
    log_info "ファイルをコピー: ${SOURCE_FILE} → ${DEPLOY_DIR}/"
    cp "${SOURCE_FILE}" "${DEPLOY_DIR}/" || {
        log_error "ファイルのコピーに失敗しました"
        return 1
    }

    # 実行権限の付与
    chmod +x "${DEPLOY_DIR}/${FILE_NAME}" || {
        log_warn "実行権限の付与に失敗しました"
    }

    # 配置結果の確認
    log_info "配置完了:"
    ls -la "${DEPLOY_DIR}/${FILE_NAME}" | tee -a "${LOG_FILE}"

    return 0
}

#------------------------------------------------------------------------------
# OpenTP1の起動
#------------------------------------------------------------------------------
start_opentp1() {
    log_step "OpenTP1 起動"

    # 起動コマンド実行
    log_info "dcstart を実行中..."
    if [ -x "${OPENTP1_BIN}/dcstart" ]; then
        "${OPENTP1_BIN}/dcstart" >> "${LOG_FILE}" 2>&1
        local exit_code=$?

        if [ ${exit_code} -ne 0 ]; then
            log_error "dcstart の終了コード: ${exit_code}"
            return 1
        fi
    else
        log_error "dcstartコマンドが見つかりません: ${OPENTP1_BIN}/dcstart"
        return 1
    fi

    # 起動待機
    log_info "${START_WAIT_TIME}秒待機中..."
    sleep ${START_WAIT_TIME}

    # 起動確認
    log_info "起動を確認中..."
    if check_opentp1_status 2>/dev/null | grep -q "ONLINE\|稼働"; then
        log_info "OpenTP1の起動を確認しました"
    else
        log_warn "OpenTP1の起動状態を確認できませんでした"
        log_warn "手動で確認してください: ${OPENTP1_BIN}/dcstatus"
    fi

    return 0
}

#------------------------------------------------------------------------------
# 結果サマリー表示
#------------------------------------------------------------------------------
show_summary() {
    local result=$1

    echo ""
    echo -e "${CYAN}================================================================${NC}"
    if [ ${result} -eq 0 ]; then
        echo -e "${GREEN}  デプロイ完了${NC}"
    else
        echo -e "${RED}  デプロイ失敗${NC}"
    fi
    echo -e "${CYAN}================================================================${NC}"
    echo ""
    echo "実行結果    : $([ ${result} -eq 0 ] && echo '成功' || echo '失敗')"
    echo "終了日時    : $(date '+%Y-%m-%d %H:%M:%S')"
    echo "ログファイル: ${LOG_FILE}"
    echo ""

    if [ ${result} -ne 0 ]; then
        echo -e "${YELLOW}詳細はログファイルを確認してください${NC}"
    fi
}

#------------------------------------------------------------------------------
# 確認プロンプト
#------------------------------------------------------------------------------
confirm_execution() {
    echo ""
    echo "以下の処理を実行します:"
    echo "  1. OpenTP1 停止 (dcstop -f)"
    echo "  2. ソース配置"
    echo "  3. OpenTP1 起動 (dcstart)"
    echo ""
    read -p "実行しますか? (y/n): " answer

    case ${answer} in
        [Yy]* )
            return 0
            ;;
        * )
            log_info "キャンセルされました"
            exit 0
            ;;
    esac
}

#------------------------------------------------------------------------------
# メイン処理
#------------------------------------------------------------------------------
main() {
    local exit_code=0

    # ヘッダー表示
    show_header

    # 確認プロンプト
    confirm_execution

    # 事前チェック
    if ! pre_check; then
        log_error "事前チェックに失敗しました"
        show_summary 1
        exit 1
    fi

    # OpenTP1停止
    if ! stop_opentp1; then
        log_error "OpenTP1の停止に失敗しました"
        show_summary 1
        exit 1
    fi

    # デプロイ
    if ! deploy_files; then
        log_error "デプロイに失敗しました"
        log_warn "OpenTP1を再起動します..."
        start_opentp1
        show_summary 1
        exit 1
    fi

    # OpenTP1起動
    if ! start_opentp1; then
        log_error "OpenTP1の起動に失敗しました"
        show_summary 1
        exit 1
    fi

    # 成功
    show_summary 0
    exit 0
}

# スクリプト実行
main "$@"
