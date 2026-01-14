#!/bin/bash
#==============================================================================
# build_wrapper.sh - ビルドシェル実行ラッパー
#==============================================================================
#
# 機能:
#   - 疑似端末(PTY)を必要とするビルドシェルに対して入力を自動送信
#   - scriptコマンドを使用してTTY環境をエミュレート
#   - プロンプト待機（オプション）: 指定したプロンプトが出力されるまで待機
#   - マルチビルドモード: 複数の業務IDを連続でビルド
#   - フルコンパイルモード: フルコンパイル実行
#
# 使用方法:
#   ./build_wrapper.sh <build_script> <input1> [input2] [input3] ...
#   ./build_wrapper.sh -w <build_script> <prompt1>:<input1> [prompt2:input2] ...
#   ./build_wrapper.sh -m <build_script> <env_prompt>:<env> <option_prompt>:<opt> \
#                      <gyomu_prompt>:<id1,id2,...> <confirm_prompt>:<y> \
#                      <option_prompt>:<exit> [error_pattern]
#   ./build_wrapper.sh -f <build_script> <env_prompt>:<env> <option_prompt>:<opt> \
#                      <confirm_prompt>:<y> <option_prompt>:<exit> [error_pattern] [timeout]
#
# 引数:
#   build_script : 実行するビルドシェルのパス
#   input1...    : 順番に送信する入力値
#
# オプション:
#   -w           : プロンプト待機モード（プロンプト:入力値 の形式で指定）
#   -m           : マルチビルドモード（複数業務IDを連続ビルド）
#   -f           : フルコンパイルモード
#   -d <秒>      : 入力間の待機秒数（デフォルト: 1秒）
#   -t <秒>      : プロンプト待機のタイムアウト秒数（デフォルト: 30秒）
#
# 例:
#   # シンプルモード（入力を順番に送信）
#   ./build_wrapper.sh /opt/build/build.sh 1 2 3
#
#   # プロンプト待機モード
#   ./build_wrapper.sh -w /opt/build/build.sh "環境を選択:1" "オプション:2"
#
#   # マルチビルドモード（複数業務IDを連続ビルド）
#   ./build_wrapper.sh -m /opt/build/build.sh "環境を選択:1" "オプション:1" \
#                      "業務ID:1,2,3" "(y/n):y" "オプション:99" "エラー"
#
#   # フルコンパイルモード
#   ./build_wrapper.sh -f /opt/build/build.sh "環境を選択:1" "オプション:2" \
#                      "(y/n):y" "オプション:99" "エラー" 400
#
#==============================================================================

set -e

# デフォルト値
WAIT_MODE=false
MULTI_MODE=false
FULL_MODE=false
DELAY=1
TIMEOUT=30

# 使用方法を表示
usage() {
    echo "使用方法: $0 [-w|-m|-f] [-d 秒] [-t 秒] <build_script> <args...>"
    echo ""
    echo "オプション:"
    echo "  -w         プロンプト待機モード"
    echo "  -m         マルチビルドモード（複数業務IDを連続ビルド）"
    echo "  -f         フルコンパイルモード"
    echo "  -d <秒>    入力間の待機秒数（デフォルト: 1秒）"
    echo "  -t <秒>    プロンプト待機のタイムアウト秒数（デフォルト: 30秒）"
    echo ""
    echo "例:"
    echo "  $0 /opt/build/build.sh 1 2 3"
    echo "  $0 -w /opt/build/build.sh '環境を選択:1' 'オプション:2'"
    echo "  $0 -m /opt/build/build.sh '環境を選択:1' 'オプション:1' '業務ID:1,2' '(y/n):y' 'オプション:99' 'エラー'"
    echo "  $0 -f /opt/build/build.sh '環境を選択:1' 'オプション:2' '(y/n):y' 'オプション:99' 'エラー' 400"
    exit 1
}

# 引数解析
while getopts "wmfd:t:h" opt; do
    case $opt in
        w)
            WAIT_MODE=true
            ;;
        m)
            MULTI_MODE=true
            ;;
        f)
            FULL_MODE=true
            ;;
        d)
            DELAY=$OPTARG
            ;;
        t)
            TIMEOUT=$OPTARG
            ;;
        h)
            usage
            ;;
        *)
            usage
            ;;
    esac
done
shift $((OPTIND-1))

# 引数チェック
if [ $# -lt 2 ]; then
    echo "[エラー] 引数が不足しています"
    usage
fi

BUILD_SCRIPT="$1"
shift

# ビルドシェルの存在確認
if [ ! -f "$BUILD_SCRIPT" ]; then
    echo "[エラー] ビルドシェルが見つかりません: $BUILD_SCRIPT"
    exit 1
fi

if [ ! -x "$BUILD_SCRIPT" ]; then
    echo "[エラー] ビルドシェルに実行権限がありません: $BUILD_SCRIPT"
    exit 1
fi

echo "[情報] ビルドシェル: $BUILD_SCRIPT"
if $FULL_MODE; then
    echo "[情報] 待機モード: フルコンパイル"
elif $MULTI_MODE; then
    echo "[情報] 待機モード: マルチビルド"
elif $WAIT_MODE; then
    echo "[情報] 待機モード: プロンプト待機"
else
    echo "[情報] 待機モード: シンプル"
fi
echo "[情報] 入力間隔: ${DELAY}秒"
echo ""

#==============================================================================
# プロンプト待機関数（ファイル位置ベース）
# 引数: $1=プロンプト, $2=出力ファイル, $3=タイムアウト秒, $4=検索開始位置（バイト、省略時は0）
# 戻り値: 0=成功, 1=タイムアウト, 2=プロセス終了
# グローバル変数: LAST_FOUND_POS に見つかった位置を設定
#==============================================================================
wait_for_prompt() {
    local prompt="$1"
    local output_file="$2"
    local timeout_sec="$3"
    local start_pos="${4:-0}"

    local start_time=$(date +%s)
    LAST_FOUND_POS=0

    while true; do
        local current_time=$(date +%s)
        local elapsed=$((current_time - start_time))

        if [ $elapsed -gt $timeout_sec ]; then
            return 1
        fi

        # 出力ファイルをチェック
        if [ -f "$output_file" ]; then
            local file_size=$(stat -c %s "$output_file" 2>/dev/null || echo "0")

            # 新しい出力がある場合のみ検索
            if [ "$file_size" -gt "$start_pos" ]; then
                # start_pos以降の新しい内容からプロンプトを検索
                if tail -c "+$((start_pos + 1))" "$output_file" 2>/dev/null | grep -qF "$prompt" 2>/dev/null; then
                    LAST_FOUND_POS=$file_size
                    return 0
                fi
            fi
        fi

        # プロセスが終了していないかチェック
        if ! kill -0 $SCRIPT_PID 2>/dev/null; then
            return 2
        fi

        sleep 0.5
    done
}

#==============================================================================
# 現在のファイルサイズを取得
#==============================================================================
get_file_size() {
    local output_file="$1"
    if [ -f "$output_file" ]; then
        stat -c %s "$output_file" 2>/dev/null || echo "0"
    else
        echo "0"
    fi
}

#==============================================================================
# フルコンパイルモード
#==============================================================================
if $FULL_MODE; then
    # 引数: env_prompt:env option_prompt:opt confirm_prompt:y option_prompt:exit [error_pattern] [timeout]
    if [ $# -lt 4 ]; then
        echo "[エラー] フルコンパイルモードには少なくとも4つの引数が必要です"
        usage
    fi

    # 引数を解析
    ENV_PAIR="$1"
    OPTION_PAIR="$2"
    CONFIRM_PAIR="$3"
    EXIT_PAIR="$4"
    ERROR_PATTERN="${5:-}"
    BUILD_TIMEOUT="${6:-400}"  # デフォルト400秒（約7分）

    # 各ペアを分割
    ENV_PROMPT="${ENV_PAIR%:*}"
    ENV_INPUT="${ENV_PAIR##*:}"

    OPTION_PROMPT="${OPTION_PAIR%:*}"
    OPTION_INPUT="${OPTION_PAIR##*:}"

    CONFIRM_PROMPT="${CONFIRM_PAIR%:*}"
    CONFIRM_INPUT="${CONFIRM_PAIR##*:}"

    EXIT_PROMPT="${EXIT_PAIR%:*}"
    EXIT_INPUT="${EXIT_PAIR##*:}"

    echo "[設定] 環境選択: '$ENV_PROMPT' → '$ENV_INPUT'"
    echo "[設定] オプション: '$OPTION_PROMPT' → '$OPTION_INPUT'"
    echo "[設定] 確認: '$CONFIRM_PROMPT' → '$CONFIRM_INPUT'"
    echo "[設定] 終了: '$EXIT_PROMPT' → '$EXIT_INPUT'"
    if [ -n "$ERROR_PATTERN" ]; then
        echo "[設定] エラー検出: '$ERROR_PATTERN'"
    fi
    echo "[設定] ビルドタイムアウト: ${BUILD_TIMEOUT}秒"
    echo ""

    # 一時ファイル作成
    OUTPUT_FILE=$(mktemp)
    INPUT_FIFO=$(mktemp -u)
    mkfifo "$INPUT_FIFO"

    # クリーンアップ関数
    cleanup() {
        rm -f "$OUTPUT_FILE" "$INPUT_FIFO" 2>/dev/null
    }
    trap cleanup EXIT

    # scriptコマンドでビルドシェルを実行（バックグラウンド）
    script -qf -c "$BUILD_SCRIPT" "$OUTPUT_FILE" < "$INPUT_FIFO" &
    SCRIPT_PID=$!

    # 入力FIFOを開いておく
    exec 3>"$INPUT_FIFO"

    BUILD_ERROR=false
    SEARCH_POS=0

    # 1. 環境選択プロンプト
    echo "[待機] 環境選択プロンプト: '$ENV_PROMPT'"
    if ! wait_for_prompt "$ENV_PROMPT" "$OUTPUT_FILE" "$TIMEOUT" "$SEARCH_POS"; then
        echo "[エラー] 環境選択プロンプトがタイムアウトしました"
        kill $SCRIPT_PID 2>/dev/null
        exit 1
    fi
    SEARCH_POS=$LAST_FOUND_POS
    echo "[送信] '$ENV_INPUT'"
    sleep "$DELAY"
    echo "$ENV_INPUT" >&3

    # 2. ビルドオプションプロンプト（1回目）
    echo "[待機] オプションプロンプト: '$OPTION_PROMPT'"
    if ! wait_for_prompt "$OPTION_PROMPT" "$OUTPUT_FILE" "$TIMEOUT" "$SEARCH_POS"; then
        echo "[エラー] オプションプロンプトがタイムアウトしました"
        kill $SCRIPT_PID 2>/dev/null
        exit 1
    fi
    SEARCH_POS=$LAST_FOUND_POS
    echo "[送信] '$OPTION_INPUT'"
    sleep "$DELAY"
    echo "$OPTION_INPUT" >&3

    # 3. 確認プロンプト
    echo "[待機] 確認プロンプト: '$CONFIRM_PROMPT'"
    if ! wait_for_prompt "$CONFIRM_PROMPT" "$OUTPUT_FILE" "$TIMEOUT" "$SEARCH_POS"; then
        echo "[エラー] 確認プロンプトがタイムアウトしました"
        kill $SCRIPT_PID 2>/dev/null
        exit 1
    fi
    SEARCH_POS=$LAST_FOUND_POS
    echo "[送信] '$CONFIRM_INPUT'"
    sleep "$DELAY"
    echo "$CONFIRM_INPUT" >&3

    # 現在の出力サイズを保存してエラーチェック用に使用
    BEFORE_BUILD_SIZE=$(get_file_size "$OUTPUT_FILE")

    # 4. ビルド完了を待機（オプションプロンプトが再度出るまで）
    echo "[待機] フルコンパイル実行中（タイムアウト: ${BUILD_TIMEOUT}秒）..."

    if ! wait_for_prompt "$EXIT_PROMPT" "$OUTPUT_FILE" "$BUILD_TIMEOUT" "$SEARCH_POS"; then
        echo "[エラー] フルコンパイルがタイムアウトしました（${BUILD_TIMEOUT}秒）"
        kill $SCRIPT_PID 2>/dev/null
        exit 1
    fi
    SEARCH_POS=$LAST_FOUND_POS

    # エラーチェック（ビルド後の新しい出力からエラーパターンを検索）
    if [ -n "$ERROR_PATTERN" ]; then
        CURRENT_SIZE=$(get_file_size "$OUTPUT_FILE")
        if [ "$CURRENT_SIZE" -gt "$BEFORE_BUILD_SIZE" ]; then
            if tail -c "+$((BEFORE_BUILD_SIZE + 1))" "$OUTPUT_FILE" 2>/dev/null | grep -qF "$ERROR_PATTERN"; then
                echo "[エラー] フルコンパイルでエラーを検出しました"
                BUILD_ERROR=true
            else
                echo "[OK] フルコンパイル完了"
            fi
        else
            echo "[OK] フルコンパイル完了"
        fi
    else
        echo "[OK] フルコンパイル完了"
    fi

    # 5. 終了処理（99を送信）
    echo ""
    echo "[終了] ビルドシェルを終了します"
    echo "[送信] '$EXIT_INPUT'"
    sleep "$DELAY"
    echo "$EXIT_INPUT" >&3

    # 入力FIFOを閉じる
    exec 3>&-

    # プロセスの終了を待機
    wait $SCRIPT_PID 2>/dev/null || true
    EXIT_CODE=$?

    echo ""
    echo "========================================"
    if $BUILD_ERROR; then
        echo "[結果] *** フルコンパイルでエラーが発生しました ***"
        echo "========================================"
        echo ""
        echo "エラー内容を確認してください。"
        exit 1
    else
        echo "[結果] フルコンパイルが正常に完了しました"
        echo "========================================"
        exit 0
    fi

#==============================================================================
# マルチビルドモード
#==============================================================================
elif $MULTI_MODE; then
    # 引数: env_prompt:env option_prompt:opt gyomu_prompt:id1,id2 confirm_prompt:y option_prompt:exit [error_pattern] [gyomu_mapping]
    if [ $# -lt 5 ]; then
        echo "[エラー] マルチビルドモードには少なくとも5つの引数が必要です"
        usage
    fi

    # 引数を解析
    ENV_PAIR="$1"
    OPTION_PAIR="$2"
    GYOMU_PAIR="$3"
    CONFIRM_PAIR="$4"
    EXIT_PAIR="$5"
    ERROR_PATTERN="${6:-}"
    GYOMU_MAPPING="${7:-}"

    # 各ペアを分割
    ENV_PROMPT="${ENV_PAIR%:*}"
    ENV_INPUT="${ENV_PAIR##*:}"

    OPTION_PROMPT="${OPTION_PAIR%:*}"
    OPTION_INPUT="${OPTION_PAIR##*:}"

    GYOMU_PROMPT="${GYOMU_PAIR%:*}"
    GYOMU_IDS="${GYOMU_PAIR##*:}"

    CONFIRM_PROMPT="${CONFIRM_PAIR%:*}"
    CONFIRM_INPUT="${CONFIRM_PAIR##*:}"

    EXIT_PROMPT="${EXIT_PAIR%:*}"
    EXIT_INPUT="${EXIT_PAIR##*:}"

    # 業務IDをカンマで分割
    IFS=',' read -ra GYOMU_ARRAY <<< "$GYOMU_IDS"

    # 業務IDマッピングを連想配列に格納（ビルド番号→業務ID）
    declare -A GYOMU_ID_MAP
    if [ -n "$GYOMU_MAPPING" ]; then
        IFS=',' read -ra MAPPING_ARRAY <<< "$GYOMU_MAPPING"
        for mapping in "${MAPPING_ARRAY[@]}"; do
            build_num="${mapping%%=*}"
            gyomu_name="${mapping#*=}"
            GYOMU_ID_MAP[$build_num]="$gyomu_name"
        done
    fi

    echo "[設定] 環境選択: '$ENV_PROMPT' → '$ENV_INPUT'"
    echo "[設定] オプション: '$OPTION_PROMPT' → '$OPTION_INPUT'"
    echo "[設定] 業務ID: '$GYOMU_PROMPT' → ${GYOMU_ARRAY[*]}"
    echo "[設定] 確認: '$CONFIRM_PROMPT' → '$CONFIRM_INPUT'"
    echo "[設定] 終了: '$EXIT_PROMPT' → '$EXIT_INPUT'"
    if [ -n "$ERROR_PATTERN" ]; then
        echo "[設定] エラー検出: '$ERROR_PATTERN'"
    fi
    echo ""

    # 一時ファイル作成
    OUTPUT_FILE=$(mktemp)
    INPUT_FIFO=$(mktemp -u)
    mkfifo "$INPUT_FIFO"

    # クリーンアップ関数
    cleanup() {
        rm -f "$OUTPUT_FILE" "$INPUT_FIFO" 2>/dev/null
    }
    trap cleanup EXIT

    # scriptコマンドでビルドシェルを実行（バックグラウンド）
    script -qf -c "$BUILD_SCRIPT" "$OUTPUT_FILE" < "$INPUT_FIFO" &
    SCRIPT_PID=$!

    # 入力FIFOを開いておく
    exec 3>"$INPUT_FIFO"

    BUILD_ERRORS=0
    BUILD_ERROR_IDS=()
    SEARCH_POS=0

    # 1. 環境選択プロンプト
    echo "[待機] 環境選択プロンプト: '$ENV_PROMPT'"
    if ! wait_for_prompt "$ENV_PROMPT" "$OUTPUT_FILE" "$TIMEOUT" "$SEARCH_POS"; then
        echo "[エラー] 環境選択プロンプトがタイムアウトしました"
        kill $SCRIPT_PID 2>/dev/null
        exit 1
    fi
    SEARCH_POS=$LAST_FOUND_POS
    echo "[送信] '$ENV_INPUT'"
    sleep "$DELAY"
    echo "$ENV_INPUT" >&3

    # 2. ビルドオプションプロンプト
    echo "[待機] オプションプロンプト: '$OPTION_PROMPT'"
    if ! wait_for_prompt "$OPTION_PROMPT" "$OUTPUT_FILE" "$TIMEOUT" "$SEARCH_POS"; then
        echo "[エラー] オプションプロンプトがタイムアウトしました"
        kill $SCRIPT_PID 2>/dev/null
        exit 1
    fi
    SEARCH_POS=$LAST_FOUND_POS
    echo "[送信] '$OPTION_INPUT'"
    sleep "$DELAY"
    echo "$OPTION_INPUT" >&3

    # 3. 各業務IDをループ処理
    for gyomu_id in "${GYOMU_ARRAY[@]}"; do
        echo ""
        echo "========================================"
        echo "[ビルド] 業務ID: $gyomu_id"
        echo "========================================"

        # 業務ID選択プロンプト（新しい出力から検索）
        echo "[待機] 業務IDプロンプト: '$GYOMU_PROMPT'"
        if ! wait_for_prompt "$GYOMU_PROMPT" "$OUTPUT_FILE" "$TIMEOUT" "$SEARCH_POS"; then
            echo "[エラー] 業務IDプロンプトがタイムアウトしました"
            kill $SCRIPT_PID 2>/dev/null
            exit 1
        fi
        SEARCH_POS=$LAST_FOUND_POS
        echo "[送信] '$gyomu_id'"
        sleep "$DELAY"
        echo "$gyomu_id" >&3

        # 確認プロンプト（新しい出力から検索）
        echo "[待機] 確認プロンプト: '$CONFIRM_PROMPT'"
        if ! wait_for_prompt "$CONFIRM_PROMPT" "$OUTPUT_FILE" "$TIMEOUT" "$SEARCH_POS"; then
            echo "[エラー] 確認プロンプトがタイムアウトしました"
            kill $SCRIPT_PID 2>/dev/null
            exit 1
        fi
        SEARCH_POS=$LAST_FOUND_POS
        echo "[送信] '$CONFIRM_INPUT'"
        sleep "$DELAY"
        echo "$CONFIRM_INPUT" >&3

        # 現在の出力サイズを保存してエラーチェック用に使用
        BEFORE_BUILD_SIZE=$(get_file_size "$OUTPUT_FILE")

        # ビルド完了を待機（オプションプロンプトが再度出るまで）
        echo "[待機] ビルド完了（オプションプロンプト待機）: '$OPTION_PROMPT'"

        if ! wait_for_prompt "$OPTION_PROMPT" "$OUTPUT_FILE" 300 "$SEARCH_POS"; then
            echo "[エラー] ビルド完了待機がタイムアウトしました（5分）"
            kill $SCRIPT_PID 2>/dev/null
            exit 1
        fi
        SEARCH_POS=$LAST_FOUND_POS

        # エラーチェック（ビルド後の新しい出力からエラーパターンを検索）
        # ビルド番号から業務IDを逆引き（マッピングがない場合はビルド番号をそのまま使用）
        actual_gyomu_id="${GYOMU_ID_MAP[$gyomu_id]:-$gyomu_id}"

        if [ -n "$ERROR_PATTERN" ]; then
            # ビルド後の新しい出力を取得してエラーパターンを検索
            CURRENT_SIZE=$(get_file_size "$OUTPUT_FILE")
            if [ "$CURRENT_SIZE" -gt "$BEFORE_BUILD_SIZE" ]; then
                # tail -c +N は N バイト目から出力（1-indexed）
                if tail -c "+$((BEFORE_BUILD_SIZE + 1))" "$OUTPUT_FILE" 2>/dev/null | grep -qF "$ERROR_PATTERN"; then
                    echo "[エラー] ビルドエラーを検出しました: 業務ID $actual_gyomu_id"
                    BUILD_ERRORS=$((BUILD_ERRORS + 1))
                    BUILD_ERROR_IDS+=("$actual_gyomu_id")
                else
                    echo "[OK] 業務ID $actual_gyomu_id のビルド完了"
                fi
            else
                echo "[OK] 業務ID $actual_gyomu_id のビルド完了"
            fi
        else
            echo "[OK] 業務ID $actual_gyomu_id のビルド完了"
        fi
    done

    # 4. 終了処理（99を送信）
    echo ""
    echo "[終了] ビルドシェルを終了します"
    echo "[送信] '$EXIT_INPUT'"
    sleep "$DELAY"
    echo "$EXIT_INPUT" >&3

    # 入力FIFOを閉じる
    exec 3>&-

    # プロセスの終了を待機
    wait $SCRIPT_PID 2>/dev/null || true
    EXIT_CODE=$?

    echo ""
    echo "========================================"
    if [ $BUILD_ERRORS -gt 0 ]; then
        echo "[結果] *** ビルドエラー: $BUILD_ERRORS 件 ***"
        echo "========================================"
        echo ""
        echo "エラーが発生した業務ID:"
        for error_id in "${BUILD_ERROR_IDS[@]}"; do
            echo "  - $error_id"
        done
        echo ""
        echo "上記の業務IDを確認してください。"
        exit 1
    else
        echo "[結果] すべてのビルドが正常に完了しました"
        echo "========================================"
        exit 0
    fi

elif $WAIT_MODE; then
    #==========================================================================
    # プロンプト待機モード
    # プロンプトが出力されるまで待機してから入力を送信
    #==========================================================================

    # プロンプトと入力値のペアを解析
    PROMPTS=()
    INPUTS=()

    for pair in "$@"; do
        if [[ "$pair" != *":"* ]]; then
            echo "[エラー] 無効な形式: '$pair' (プロンプト:入力値 の形式で指定してください)"
            exit 1
        fi

        # 最後のコロンで分割
        prompt="${pair%:*}"
        input="${pair##*:}"

        PROMPTS+=("$prompt")
        INPUTS+=("$input")
        echo "  待機: '$prompt' → 送信: '$input'"
    done
    echo ""

    # 一時ファイル作成
    OUTPUT_FILE=$(mktemp)
    INPUT_FIFO=$(mktemp -u)
    mkfifo "$INPUT_FIFO"

    # クリーンアップ関数
    cleanup() {
        rm -f "$OUTPUT_FILE" "$INPUT_FIFO" 2>/dev/null
    }
    trap cleanup EXIT

    # scriptコマンドでビルドシェルを実行（バックグラウンド）
    # -f: 出力を即座にフラッシュ（バッファリングを無効化）
    script -qf -c "$BUILD_SCRIPT" "$OUTPUT_FILE" < "$INPUT_FIFO" &
    SCRIPT_PID=$!

    # 入力FIFOを開いておく（閉じるとプロセスが終了する）
    exec 3>"$INPUT_FIFO"

    # 各プロンプトを待機して入力を送信
    for i in "${!PROMPTS[@]}"; do
        prompt="${PROMPTS[$i]}"
        input="${INPUTS[$i]}"

        echo "[待機] プロンプト: '$prompt'"

        # プロンプト待機
        start_time=$(date +%s)
        found=false

        while true; do
            # タイムアウトチェック
            current_time=$(date +%s)
            elapsed=$((current_time - start_time))

            if [ $elapsed -gt $TIMEOUT ]; then
                echo "[エラー] タイムアウト: プロンプト '$prompt' が ${TIMEOUT}秒以内に見つかりませんでした"
                kill $SCRIPT_PID 2>/dev/null
                exit 1
            fi

            # 出力ファイルをチェック（-F: 固定文字列として検索）
            if [ -f "$OUTPUT_FILE" ] && grep -qF "$prompt" "$OUTPUT_FILE" 2>/dev/null; then
                found=true
                break
            fi

            # プロセスが終了していないかチェック
            if ! kill -0 $SCRIPT_PID 2>/dev/null; then
                echo "[エラー] ビルドシェルが予期せず終了しました"
                exit 1
            fi

            sleep 0.5
        done

        if $found; then
            echo "[送信] '$input'"
            sleep "$DELAY"
            echo "$input" >&3
        fi
    done

    # 入力FIFOを閉じる
    exec 3>&-

    # プロセスの終了を待機
    wait $SCRIPT_PID 2>/dev/null
    EXIT_CODE=$?

    echo ""
    echo "========================================"
    if [ $EXIT_CODE -eq 0 ]; then
        echo "[結果] プロンプト待機モードが正常に完了しました"
    else
        echo "[結果] *** プロンプト待機モードでエラーが発生しました ***"
    fi
    echo "========================================"
    exit $EXIT_CODE

else
    #==========================================================================
    # シンプルモード
    # 入力を順番に送信（スリープで間隔を空ける）
    #==========================================================================

    echo "[情報] 入力値:"
    for input in "$@"; do
        echo "  - $input"
    done
    echo ""

    # 入力値を一時ファイルに書き込み（スリープ付き）
    INPUT_FILE=$(mktemp)

    for input in "$@"; do
        echo "$input" >> "$INPUT_FILE"
    done

    # scriptコマンドで疑似端末を作成し、入力をリダイレクト
    # -q: 静音モード（開始/終了メッセージを抑制）
    # -c: 指定したコマンドを実行
    echo "[実行] ビルドシェルを開始します..."
    echo ""

    # 入力を遅延付きで送信
    (
        for input in "$@"; do
            sleep "$DELAY"
            echo "$input"
        done
        # 最後に少し待機してからEOFを送信
        sleep 2
    ) | script -q -c "$BUILD_SCRIPT" /dev/null

    EXIT_CODE=$?

    # 一時ファイルを削除
    rm -f "$INPUT_FILE"
fi

echo ""
if [ $EXIT_CODE -eq 0 ]; then
    echo "[完了] 正常終了 (終了コード: $EXIT_CODE)"
else
    echo "[エラー] 異常終了 (終了コード: $EXIT_CODE)"
fi
exit $EXIT_CODE
