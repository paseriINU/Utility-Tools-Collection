#!/bin/bash
#==============================================================================
# build_wrapper.sh - ビルドシェル実行ラッパー
#==============================================================================
#
# 機能:
#   - 疑似端末(PTY)を必要とするビルドシェルに対して入力を自動送信
#   - scriptコマンドを使用してTTY環境をエミュレート
#   - プロンプト待機（オプション）: 指定したプロンプトが出力されるまで待機
#
# 使用方法:
#   ./build_wrapper.sh <build_script> <input1> [input2] [input3] ...
#   ./build_wrapper.sh -w <build_script> <prompt1>:<input1> [prompt2:input2] ...
#
# 引数:
#   build_script : 実行するビルドシェルのパス
#   input1...    : 順番に送信する入力値
#
# オプション:
#   -w           : プロンプト待機モード（プロンプト:入力値 の形式で指定）
#   -d <秒>      : 入力間の待機秒数（デフォルト: 1秒）
#
# 例:
#   # シンプルモード（入力を順番に送信）
#   ./build_wrapper.sh /opt/build/build.sh 1 2 3
#
#   # プロンプト待機モード
#   ./build_wrapper.sh -w /opt/build/build.sh "環境を選択:1" "オプション:2"
#
#==============================================================================

set -e

# デフォルト値
WAIT_MODE=false
DELAY=1
TIMEOUT=30

# 使用方法を表示
usage() {
    echo "使用方法: $0 [-w] [-d 秒] <build_script> <input1|prompt1:input1> [...]"
    echo ""
    echo "オプション:"
    echo "  -w         プロンプト待機モード"
    echo "  -d <秒>    入力間の待機秒数（デフォルト: 1秒）"
    echo "  -t <秒>    プロンプト待機のタイムアウト秒数（デフォルト: 30秒）"
    echo ""
    echo "例:"
    echo "  $0 /opt/build/build.sh 1 2 3"
    echo "  $0 -w /opt/build/build.sh '環境を選択:1' 'オプション:2'"
    exit 1
}

# 引数解析
while getopts "wd:t:h" opt; do
    case $opt in
        w)
            WAIT_MODE=true
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
echo "[情報] 待機モード: $(if $WAIT_MODE; then echo 'プロンプト待機'; else echo 'シンプル'; fi)"
echo "[情報] 入力間隔: ${DELAY}秒"
echo ""

if $WAIT_MODE; then
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
    script -q -c "$BUILD_SCRIPT" "$OUTPUT_FILE" < "$INPUT_FIFO" &
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

    # 出力を表示
    echo ""
    echo "======== ビルド出力 ========"
    cat "$OUTPUT_FILE"
    echo "============================"

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
echo "[完了] 終了コード: $EXIT_CODE"
exit $EXIT_CODE
