#!/bin/bash
# -*- coding: utf-8 -*-
#
# WinRM Remote Batch Executor for Linux (Shell Script版 - 標準コマンドのみ)
# Linux（Red Hat等）からWindows Server 2022へWinRM接続してバッチを実行
#
# 必要なツール（すべて標準でインストール済み）:
#   - bash
#   - curl
#   - date
#   - base64
#
# 使い方:
#   1. このスクリプト内の設定セクションを編集
#   2. 実行権限を付与: chmod +x winrm_exec.sh
#   3. 実行: ./winrm_exec.sh ENV
#
#   環境を引数で指定（必須）:
#   ./winrm_exec.sh TST1T
#   ./winrm_exec.sh TST2T
#
#   または環境変数で設定を上書き:
#   WINRM_HOST=192.168.1.100 WINRM_USER=Admin WINRM_PASS=Pass123 ./winrm_exec.sh TST1T
#

set -u  # 未定義変数の使用をエラーとする

# 引数チェック
if [ $# -eq 0 ]; then
    echo "エラー: 環境を指定してください"
    echo ""
    echo "使い方: $0 ENV"
    echo ""
    echo "利用可能な環境:"
    echo "  TST1T"
    echo "  TST2T"
    echo ""
    echo "例:"
    echo "  $0 TST1T"
    echo "  $0 TST2T"
    exit 1
fi

if [ "$1" = "-h" ] || [ "$1" = "--help" ]; then
    echo "使い方: $0 ENV"
    echo ""
    echo "引数:"
    echo "  ENV    環境名 (TST1T, TST2T など)"
    echo ""
    echo "例:"
    echo "  $0 TST1T"
    echo "  $0 TST2T"
    exit 0
fi

# 環境を引数から取得
ENV_ARG="$1"

# ==================== 設定セクション ====================
# ここを編集して使用してください
# 注意: バックスラッシュ(\)を含む値はシングルクォートで囲むこと
#       例: _DEFAULT_USER='DOMAIN\username'

# Windows接続情報
# ドメインユーザーの場合: 'DOMAIN\username' または 'username@domain.local'
_DEFAULT_HOST='192.168.1.100'
_DEFAULT_USER='Administrator'
_DEFAULT_PASS='YourPassword'
WINRM_HOST="${WINRM_HOST:-$_DEFAULT_HOST}"        # Windows ServerのIPアドレスまたはホスト名
WINRM_PORT="${WINRM_PORT:-5985}"                  # WinRMポート（HTTP: 5985, HTTPS: 5986）
WINRM_USER="${WINRM_USER:-$_DEFAULT_USER}"        # Windowsユーザー名
WINRM_PASS="${WINRM_PASS:-$_DEFAULT_PASS}"        # Windowsパスワード

# 環境フォルダ名のリスト（実行時に選択可能）
# 新しい環境を追加する場合は、この配列に追加してください
ENVIRONMENTS=("TST1T" "TST2T")                   # 利用可能な環境のリスト
ENV_FOLDER="${ENV_FOLDER:-}"                     # デフォルト環境なし（実行時に必ず選択）

# 実行するバッチファイル（Windows側のパス）
# {ENV} は選択した環境フォルダ名に置換されます
# 同じパス内で複数回使用可能:
#   例: C:\Scripts\{ENV}\{ENV}_deploy.bat → C:\Scripts\TST1T\TST1T_deploy.bat
# 注意: シングルクォートで囲むこと（ダブルクォートだと{ENV}がBashで展開される）
_DEFAULT_BATCH_PATH='C:\Scripts\{ENV}\test.bat'
BATCH_FILE_PATH="${BATCH_FILE_PATH:-$_DEFAULT_BATCH_PATH}"

# または直接コマンドを指定（シングルクォートで{ENV}を含める）
_DEFAULT_DIRECT_CMD=''
DIRECT_COMMAND="${DIRECT_COMMAND:-$_DEFAULT_DIRECT_CMD}"  # 例: 'echo {ENV} environment'

# HTTPS接続を使用する場合は"true"に設定
USE_HTTPS="${USE_HTTPS:-false}"

# 証明書検証を無効にする場合は"true"（自己署名証明書の場合）
DISABLE_CERT_VALIDATION="${DISABLE_CERT_VALIDATION:-true}"

# タイムアウト（秒）
TIMEOUT="${TIMEOUT:-300}"

# デバッグモード（trueにするとXML送受信を表示）
DEBUG="${DEBUG:-false}"

# =========================================================

# 色付き出力用
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# ログ関数（stderrに出力して関数の戻り値キャプチャに影響しないようにする）
log_info() {
    printf "${BLUE}[INFO]${NC} %s\n" "$1" >&2
}

log_success() {
    printf "${GREEN}[SUCCESS]${NC} %s\n" "$1" >&2
}

log_warn() {
    printf "${YELLOW}[WARN]${NC} %s\n" "$1" >&2
}

log_error() {
    printf "${RED}[ERROR]${NC} %s\n" "$1" >&2
}


# WinRMエンドポイントURL生成
generate_endpoint() {
    local protocol="http"
    if [ "$USE_HTTPS" = "true" ]; then
        protocol="https"
    fi
    echo "${protocol}://${WINRM_HOST}:${WINRM_PORT}/wsman"
}

# UUIDの生成（標準コマンドのみ）
generate_uuid() {
    # /proc/sys/kernel/random/uuidが利用可能な場合
    if [ -r /proc/sys/kernel/random/uuid ]; then
        cat /proc/sys/kernel/random/uuid
        return
    fi

    # dateコマンドで簡易UUID生成
    local timestamp=$(date +%s%N 2>/dev/null || date +%s)
    local random1=$RANDOM
    local random2=$RANDOM
    local random3=$RANDOM
    echo "$(printf '%08x' $timestamp)-$(printf '%04x' $random1)-4$(printf '%03x' $random2)-$(printf '%04x' $random3)-$(printf '%012x' $timestamp)"
}

# XML特殊文字のエスケープ
xml_escape() {
    local string="$1"
    string="${string//&/&amp;}"
    string="${string//</&lt;}"
    string="${string//>/&gt;}"
    string="${string//\"/&quot;}"
    string="${string//\'/&apos;}"
    echo "$string"
}

# SOAP リクエストの送信
send_soap_request() {
    local soap_envelope="$1"
    local endpoint=$(generate_endpoint)

    if [ "$DEBUG" = "true" ]; then
        log_info "送信XML:"
        echo "$soap_envelope"
        echo ""
    fi

    # curl オプションの構築
    local curl_opts=(-s -S --max-time "$TIMEOUT")

    if [ "$DISABLE_CERT_VALIDATION" = "true" ]; then
        curl_opts+=(-k)  # 証明書検証を無効化
    fi

    # Basic認証
    curl_opts+=(--user "${WINRM_USER}:${WINRM_PASS}")

    # HTTPヘッダー
    curl_opts+=(-H "Content-Type: application/soap+xml;charset=UTF-8")

    # データ送信
    curl_opts+=(--data-binary "$soap_envelope")

    # リクエスト送信
    local response
    response=$(curl "${curl_opts[@]}" "$endpoint" 2>&1)
    local exit_code=$?

    if [ $exit_code -ne 0 ]; then
        log_error "curlコマンドが失敗しました (終了コード: $exit_code)"

        # エラー内容の詳細表示
        case $exit_code in
            6)
                log_error "ホスト名の解決に失敗しました: $WINRM_HOST"
                log_error "ホスト名またはIPアドレスを確認してください"
                ;;
            7)
                log_error "接続に失敗しました: ${WINRM_HOST}:${WINRM_PORT}"
                log_error "ホストが起動しているか、ファイアウォール設定を確認してください"
                ;;
            28)
                log_error "タイムアウトしました (${TIMEOUT}秒)"
                log_error "TIMEOUT値を増やすか、ネットワーク接続を確認してください"
                ;;
            52)
                log_error "サーバーから応答がありませんでした"
                log_error "WinRMサービスが起動しているか確認してください"
                ;;
            *)
                log_error "curl エラー詳細: $response"
                ;;
        esac
        return 1
    fi

    # HTTPエラーレスポンスのチェック
    if echo "$response" | grep -q "HTTP/1.1 401"; then
        log_error "認証に失敗しました"
        log_error "ユーザー名とパスワードを確認してください"
        return 1
    fi

    if echo "$response" | grep -q "HTTP/1.1 500"; then
        log_error "サーバー内部エラーが発生しました"
        log_error "WinRM設定またはコマンド内容を確認してください"
        return 1
    fi

    if [ "$DEBUG" = "true" ]; then
        log_info "受信XML:"
        echo "$response"
        echo ""
    fi

    echo "$response"
}

# XMLからタグの値を抽出（シンプルなgrep/sed実装）
extract_xml_value() {
    local xml="$1"
    local tag="$2"

    # <tag>値</tag> の形式から値を抽出
    echo "$xml" | grep -oP "(?<=<${tag}>)[^<]+" | head -1
}

# Base64デコード
base64_decode() {
    echo "$1" | base64 -d 2>/dev/null || echo ""
}

# シェルの作成
create_shell() {
    local endpoint=$(generate_endpoint)
    local uuid=$(generate_uuid)

    local soap_envelope="<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<s:Envelope xmlns:s=\"http://www.w3.org/2003/05/soap-envelope\"
            xmlns:a=\"http://schemas.xmlsoap.org/ws/2004/08/addressing\"
            xmlns:w=\"http://schemas.dmtf.org/wbem/wsman/1/wsman.xsd\"
            xmlns:rsp=\"http://schemas.microsoft.com/wbem/wsman/1/windows/shell\">
  <s:Header>
    <a:To>${endpoint}</a:To>
    <a:ReplyTo>
      <a:Address s:mustUnderstand=\"true\">http://schemas.xmlsoap.org/ws/2004/08/addressing/role/anonymous</a:Address>
    </a:ReplyTo>
    <a:Action s:mustUnderstand=\"true\">http://schemas.xmlsoap.org/ws/2004/09/transfer/Create</a:Action>
    <w:MaxEnvelopeSize s:mustUnderstand=\"true\">153600</w:MaxEnvelopeSize>
    <a:MessageID>uuid:${uuid}</a:MessageID>
    <w:Locale xml:lang=\"ja-JP\" s:mustUnderstand=\"false\"/>
    <w:OperationTimeout>PT${TIMEOUT}S</w:OperationTimeout>
    <w:ResourceURI s:mustUnderstand=\"true\">http://schemas.microsoft.com/wbem/wsman/1/windows/shell/cmd</w:ResourceURI>
    <w:OptionSet>
      <w:Option Name=\"WINRS_NOPROFILE\">FALSE</w:Option>
      <w:Option Name=\"WINRS_CODEPAGE\">65001</w:Option>
    </w:OptionSet>
  </s:Header>
  <s:Body>
    <rsp:Shell>
      <rsp:InputStreams>stdin</rsp:InputStreams>
      <rsp:OutputStreams>stdout stderr</rsp:OutputStreams>
    </rsp:Shell>
  </s:Body>
</s:Envelope>"

    log_info "シェル作成中..."
    local response=$(send_soap_request "$soap_envelope")

    if [ $? -ne 0 ]; then
        log_error "シェル作成に失敗しました"
        log_error "WinRM接続設定を確認してください"
        return 1
    fi

    # ShellIdを抽出
    local shell_id=$(extract_xml_value "$response" "rsp:ShellId")

    if [ -z "$shell_id" ]; then
        log_error "ShellIDの取得に失敗しました"
        log_error "サーバーからの応答が不正です"
        if [ "$DEBUG" != "true" ]; then
            log_error "詳細を確認するには DEBUG=true を設定してください"
        fi
        return 1
    fi

    printf "${GREEN}[SUCCESS]${NC} シェル作成成功: %s\n" "$shell_id" >&2
    echo "$shell_id"
}

# コマンドの実行
run_command() {
    local shell_id="$1"
    local command="$2"
    local endpoint=$(generate_endpoint)
    local uuid=$(generate_uuid)

    # コマンドをXMLエスケープ
    local command_escaped=$(xml_escape "$command")

    local soap_envelope="<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<s:Envelope xmlns:s=\"http://www.w3.org/2003/05/soap-envelope\"
            xmlns:a=\"http://schemas.xmlsoap.org/ws/2004/08/addressing\"
            xmlns:w=\"http://schemas.dmtf.org/wbem/wsman/1/wsman.xsd\"
            xmlns:rsp=\"http://schemas.microsoft.com/wbem/wsman/1/windows/shell\">
  <s:Header>
    <a:To>${endpoint}</a:To>
    <a:ReplyTo>
      <a:Address s:mustUnderstand=\"true\">http://schemas.xmlsoap.org/ws/2004/08/addressing/role/anonymous</a:Address>
    </a:ReplyTo>
    <a:Action s:mustUnderstand=\"true\">http://schemas.microsoft.com/wbem/wsman/1/windows/shell/Command</a:Action>
    <w:MaxEnvelopeSize s:mustUnderstand=\"true\">153600</w:MaxEnvelopeSize>
    <a:MessageID>uuid:${uuid}</a:MessageID>
    <w:Locale xml:lang=\"ja-JP\" s:mustUnderstand=\"false\"/>
    <w:OperationTimeout>PT${TIMEOUT}S</w:OperationTimeout>
    <w:ResourceURI s:mustUnderstand=\"true\">http://schemas.microsoft.com/wbem/wsman/1/windows/shell/cmd</w:ResourceURI>
    <w:SelectorSet>
      <w:Selector Name=\"ShellId\">${shell_id}</w:Selector>
    </w:SelectorSet>
  </s:Header>
  <s:Body>
    <rsp:CommandLine>
      <rsp:Command>${command_escaped}</rsp:Command>
    </rsp:CommandLine>
  </s:Body>
</s:Envelope>"

    log_info "コマンド実行中..."
    local response=$(send_soap_request "$soap_envelope")

    if [ $? -ne 0 ]; then
        log_error "コマンド実行に失敗しました"
        log_error "実行コマンド: $command"
        return 1
    fi

    # CommandIdを抽出
    local command_id=$(extract_xml_value "$response" "rsp:CommandId")

    if [ -z "$command_id" ]; then
        log_error "CommandIDの取得に失敗しました"
        log_error "コマンドの構文が正しいか確認してください"
        if [ "$DEBUG" != "true" ]; then
            log_error "詳細を確認するには DEBUG=true を設定してください"
        fi
        return 1
    fi

    printf "${GREEN}[SUCCESS]${NC} コマンド実行開始: %s\n" "$command_id" >&2
    echo "$command_id"
}

# コマンド出力の取得
get_command_output() {
    local shell_id="$1"
    local command_id="$2"
    local endpoint=$(generate_endpoint)

    local stdout_all=""
    local stderr_all=""
    local exit_code=0
    local command_done=false
    local max_attempts=$((TIMEOUT * 2))  # TIMEOUT秒待機（0.5秒ごとにチェック）
    local attempt=0

    log_info "コマンド出力取得中...（最大${TIMEOUT}秒待機）"

    while [ "$command_done" = "false" ] && [ $attempt -lt $max_attempts ]; do
        local uuid=$(generate_uuid)

        local soap_envelope="<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<s:Envelope xmlns:s=\"http://www.w3.org/2003/05/soap-envelope\"
            xmlns:a=\"http://schemas.xmlsoap.org/ws/2004/08/addressing\"
            xmlns:w=\"http://schemas.dmtf.org/wbem/wsman/1/wsman.xsd\"
            xmlns:rsp=\"http://schemas.microsoft.com/wbem/wsman/1/windows/shell\">
  <s:Header>
    <a:To>${endpoint}</a:To>
    <a:ReplyTo>
      <a:Address s:mustUnderstand=\"true\">http://schemas.xmlsoap.org/ws/2004/08/addressing/role/anonymous</a:Address>
    </a:ReplyTo>
    <a:Action s:mustUnderstand=\"true\">http://schemas.microsoft.com/wbem/wsman/1/windows/shell/Receive</a:Action>
    <w:MaxEnvelopeSize s:mustUnderstand=\"true\">153600</w:MaxEnvelopeSize>
    <a:MessageID>uuid:${uuid}</a:MessageID>
    <w:Locale xml:lang=\"ja-JP\" s:mustUnderstand=\"false\"/>
    <w:OperationTimeout>PT${TIMEOUT}S</w:OperationTimeout>
    <w:ResourceURI s:mustUnderstand=\"true\">http://schemas.microsoft.com/wbem/wsman/1/windows/shell/cmd</w:ResourceURI>
    <w:SelectorSet>
      <w:Selector Name=\"ShellId\">${shell_id}</w:Selector>
    </w:SelectorSet>
  </s:Header>
  <s:Body>
    <rsp:Receive>
      <rsp:DesiredStream CommandId=\"${command_id}\">stdout stderr</rsp:DesiredStream>
    </rsp:Receive>
  </s:Body>
</s:Envelope>"

        local response=$(send_soap_request "$soap_envelope")

        if [ $? -ne 0 ]; then
            log_error "出力取得に失敗しました"
            log_error "コマンド実行中にエラーが発生した可能性があります"
            return 1
        fi

        # stdout抽出（Base64デコード）
        local stdout_b64=$(echo "$response" | grep -oP '(?<=<rsp:Stream Name="stdout">)[^<]+' | head -1)
        if [ -n "$stdout_b64" ]; then
            local stdout_decoded=$(echo "$stdout_b64" | base64 -d 2>/dev/null || echo "")
            stdout_all="${stdout_all}${stdout_decoded}"
        fi

        # stderr抽出（Base64デコード）
        local stderr_b64=$(echo "$response" | grep -oP '(?<=<rsp:Stream Name="stderr">)[^<]+' | head -1)
        if [ -n "$stderr_b64" ]; then
            local stderr_decoded=$(echo "$stderr_b64" | base64 -d 2>/dev/null || echo "")
            stderr_all="${stderr_all}${stderr_decoded}"
        fi

        # コマンド完了チェック
        if echo "$response" | grep -q "CommandState/Done"; then
            command_done=true
            exit_code=$(extract_xml_value "$response" "rsp:ExitCode")
            [ -z "$exit_code" ] && exit_code=0
        fi

        attempt=$((attempt + 1))
        sleep 0.5  # 0.5秒待機
    done

    if [ "$command_done" = "false" ]; then
        log_warn "コマンド完了待機がタイムアウトしました"
    fi

    printf "${GREEN}[SUCCESS]${NC} コマンド完了 (終了コード: %s)\n" "$exit_code" >&2

    # 出力を一時ファイルに保存（改行を保持）
    echo "$stdout_all" > /tmp/winrm_stdout_$$
    echo "$stderr_all" > /tmp/winrm_stderr_$$
    echo "$exit_code" > /tmp/winrm_exitcode_$$
}

# シェルの削除
delete_shell() {
    local shell_id="$1"
    local endpoint=$(generate_endpoint)
    local uuid=$(generate_uuid)

    local soap_envelope="<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<s:Envelope xmlns:s=\"http://www.w3.org/2003/05/soap-envelope\"
            xmlns:a=\"http://schemas.xmlsoap.org/ws/2004/08/addressing\"
            xmlns:w=\"http://schemas.dmtf.org/wbem/wsman/1/wsman.xsd\">
  <s:Header>
    <a:To>${endpoint}</a:To>
    <a:ReplyTo>
      <a:Address s:mustUnderstand=\"true\">http://schemas.xmlsoap.org/ws/2004/08/addressing/role/anonymous</a:Address>
    </a:ReplyTo>
    <a:Action s:mustUnderstand=\"true\">http://schemas.xmlsoap.org/ws/2004/09/transfer/Delete</a:Action>
    <w:MaxEnvelopeSize s:mustUnderstand=\"true\">153600</w:MaxEnvelopeSize>
    <a:MessageID>uuid:${uuid}</a:MessageID>
    <w:Locale xml:lang=\"ja-JP\" s:mustUnderstand=\"false\"/>
    <w:OperationTimeout>PT${TIMEOUT}S</w:OperationTimeout>
    <w:ResourceURI s:mustUnderstand=\"true\">http://schemas.microsoft.com/wbem/wsman/1/windows/shell/cmd</w:ResourceURI>
    <w:SelectorSet>
      <w:Selector Name=\"ShellId\">${shell_id}</w:Selector>
    </w:SelectorSet>
  </s:Header>
  <s:Body/>
</s:Envelope>"

    log_info "シェル削除中..."
    send_soap_request "$soap_envelope" > /dev/null
    printf "${GREEN}[SUCCESS]${NC} シェル削除完了\n" >&2
}

# メイン処理
main() {
    echo ""
    echo "========================================================================"
    echo "  WinRM Remote Batch Executor (Bash)"
    echo "  標準コマンドのみ版"
    echo "========================================================================"
    echo ""

    # 環境の有効性チェック
    local valid=false
    for env in "${ENVIRONMENTS[@]}"; do
        if [ "$env" = "$ENV_ARG" ]; then
            valid=true
            break
        fi
    done

    if [ "$valid" = "false" ]; then
        log_error "無効な環境が指定されました: $ENV_ARG"
        log_error "利用可能な環境: ${ENVIRONMENTS[*]}"
        exit 1
    fi

    ENV_FOLDER="$ENV_ARG"
    log_success "指定された環境: $ENV_FOLDER"
    echo

    log_info "接続先: $(generate_endpoint)"
    log_info "ユーザー: $WINRM_USER"

    # 実行するコマンドの決定
    local command=""
    if [ -n "$DIRECT_COMMAND" ]; then
        # 直接コマンドの場合も {ENV} を置換
        command="${DIRECT_COMMAND//\{ENV\}/$ENV_FOLDER}"
        log_info "バッチファイル実行: ${DIRECT_COMMAND//\{ENV\}/$ENV_FOLDER}"
    elif [ -n "$BATCH_FILE_PATH" ]; then
        # バッチファイルパスの {ENV} を選択した環境に置換
        local batch_path="${BATCH_FILE_PATH//\{ENV\}/$ENV_FOLDER}"
        command="cmd.exe /c \"$batch_path\""
        log_info "バッチファイル実行: $batch_path"
    else
        log_error "実行するコマンドまたはバッチファイルが指定されていません"
        log_error "スクリプト内のDIRECT_COMMANDまたはBATCH_FILE_PATHを設定してください"
        exit 1
    fi
    echo

    # 一時ファイルのクリーンアップ
    trap "rm -f /tmp/winrm_stdout_$$ /tmp/winrm_stderr_$$ /tmp/winrm_exitcode_$$" EXIT

    # シェル作成
    local shell_id=$(create_shell)
    if [ $? -ne 0 ] || [ -z "$shell_id" ]; then
        log_error "処理を中断します"
        exit 1
    fi
    echo

    # コマンド実行
    local command_id=$(run_command "$shell_id" "$command")
    if [ $? -ne 0 ] || [ -z "$command_id" ]; then
        delete_shell "$shell_id"
        log_error "処理を中断します"
        exit 1
    fi
    echo

    # 出力取得
    get_command_output "$shell_id" "$command_id"
    local get_output_result=$?
    echo

    # シェル削除
    delete_shell "$shell_id"

    if [ $get_output_result -ne 0 ]; then
        log_error "処理を中断します"
        exit 1
    fi

    # 結果の表示
    echo
    echo "============================================================"
    echo "実行結果"
    echo "============================================================"

    if [ -f /tmp/winrm_stdout_$$ ]; then
        local stdout_content=$(cat /tmp/winrm_stdout_$$)
        if [ -n "$stdout_content" ]; then
            echo
            echo "[標準出力]"
            echo "$stdout_content"
        fi
    fi

    if [ -f /tmp/winrm_stderr_$$ ]; then
        local stderr_content=$(cat /tmp/winrm_stderr_$$)
        if [ -n "$stderr_content" ]; then
            echo
            echo "[標準エラー出力]"
            echo "$stderr_content"
        fi
    fi

    local exit_code=0
    if [ -f /tmp/winrm_exitcode_$$ ]; then
        exit_code=$(cat /tmp/winrm_exitcode_$$)
    fi

    echo
    echo "終了コード: $exit_code"
    echo "============================================================"

    if [ $exit_code -eq 0 ]; then
        log_success "完了"
    else
        log_error "コマンドが失敗しました (終了コード: $exit_code)"
    fi

    exit $exit_code
}

# スクリプト実行
main "$@"
