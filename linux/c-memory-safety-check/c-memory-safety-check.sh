#!/bin/bash
#===============================================================================
# C言語 メモリ安全性・デッドロジック チェックツール
#
# 概要:
#   C言語ソースファイルのメモリ安全性とデッドロジックを検出するツール
#
# 機能:
#   1. GCC警告オプションによる静的チェック
#   2. cppcheckによる静的解析（インストールされている場合）
#   3. AddressSanitizerによるメモリ安全性チェック
#   4. Valgrindによるメモリリーク検出（インストールされている場合）
#
# 使用方法:
#   ./c-memory-safety-check.sh <source.c> [-m Makefile] [-- テスト引数...]
#   ./c-memory-safety-check.sh --help
#
# 必須環境:
#   - GCC (gcc)
#
# オプション環境:
#   - cppcheck (静的解析の強化)
#   - valgrind (動的メモリチェック)
#
#===============================================================================

set -u

#-------------------------------------------------------------------------------
# カラー定義
#-------------------------------------------------------------------------------
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[0;33m'
BLUE='\033[0;34m'
CYAN='\033[0;36m'
NC='\033[0m' # No Color
BOLD='\033[1m'

#-------------------------------------------------------------------------------
# グローバル変数
#-------------------------------------------------------------------------------
VERSION="1.2.0"
SCRIPT_NAME=$(basename "$0")
SOURCE_FILE=""
MAKEFILE=""
TEST_ARGS=()
OUTPUT_DIR=""
REPORT_FILE=""
TEMP_BINARY=""

# Makefileから抽出したオプション
EXTRACTED_CFLAGS=""
EXTRACTED_INCLUDES=""

# チェック結果カウンター
TOTAL_WARNINGS=0
TOTAL_ERRORS=0
MEMORY_ISSUES=0
DEAD_CODE_ISSUES=0

#-------------------------------------------------------------------------------
# ヘルプ表示
#-------------------------------------------------------------------------------
show_help() {
    cat << EOF
${BOLD}C言語 メモリ安全性・デッドロジック チェックツール v${VERSION}${NC}

${BOLD}使用方法:${NC}
  $SCRIPT_NAME <source.c> [-m Makefile] [-- テスト実行引数...]
  $SCRIPT_NAME --help

${BOLD}説明:${NC}
  C言語ソースファイルのメモリ安全性とデッドロジックを検出します。

${BOLD}チェック項目:${NC}
  ${CYAN}1. GCC警告チェック${NC}
     - 未初期化変数、未使用変数、到達不能コード等

  ${CYAN}2. cppcheck静的解析${NC} (インストールされている場合)
     - メモリリーク、バッファオーバーフロー、デッドコード等

  ${CYAN}3. AddressSanitizer${NC}
     - バッファオーバーフロー、解放済みメモリアクセス等

  ${CYAN}4. Valgrind${NC} (インストールされている場合)
     - メモリリーク、不正メモリアクセス等

${BOLD}引数:${NC}
  source.c          チェック対象のCソースファイル

${BOLD}オプション:${NC}
  -m <Makefile>     Makefileからコンパイルオプションを抽出
  --                以降の引数をテスト実行時の引数として扱う
  --help, -h        このヘルプを表示

${BOLD}必須環境:${NC}
  - GCC (gcc)

${BOLD}推奨環境:${NC}
  - cppcheck        静的解析の強化
  - valgrind        動的メモリチェック

${BOLD}例:${NC}
  # 基本的な使用方法
  $SCRIPT_NAME myprogram.c

  # Makefileからコンパイルオプションを抽出
  $SCRIPT_NAME main.c -m Makefile

  # テスト実行時に引数を渡す
  $SCRIPT_NAME myprogram.c -- arg1 arg2

  # Makefileとテスト引数の両方を指定
  $SCRIPT_NAME main.c -m Makefile -- input.txt

${BOLD}出力:${NC}
  チェック結果は以下に保存されます:
  - ./c-check-results/<ソース名>_report.txt

EOF
}

#-------------------------------------------------------------------------------
# ログ出力関数
#-------------------------------------------------------------------------------
log_info() {
    echo -e "${BLUE}[INFO]${NC} $1"
}

log_success() {
    echo -e "${GREEN}[OK]${NC} $1"
}

log_warning() {
    echo -e "${YELLOW}[WARNING]${NC} $1"
}

log_error() {
    echo -e "${RED}[ERROR]${NC} $1"
}

log_section() {
    echo ""
    echo -e "${BOLD}${CYAN}========================================${NC}"
    echo -e "${BOLD}${CYAN} $1${NC}"
    echo -e "${BOLD}${CYAN}========================================${NC}"
}

#-------------------------------------------------------------------------------
# Makefileからコンパイルオプションを抽出
#-------------------------------------------------------------------------------
extract_makefile_options() {
    if [ -z "$MAKEFILE" ]; then
        return
    fi

    if [ ! -f "$MAKEFILE" ]; then
        log_error "Makefileが見つかりません: $MAKEFILE"
        exit 1
    fi

    log_info "Makefileからコンパイルオプションを抽出中..."

    local makefile_dir=$(dirname "$MAKEFILE")

    # make -p でMakefileの変数を展開して取得
    local make_output
    make_output=$(make -p -f "$MAKEFILE" -C "$makefile_dir" 2>/dev/null | grep -E "^(CFLAGS|CPPFLAGS|INCLUDES|INCDIR|INC_PATH|INC) " || true)

    # CFLAGSを抽出
    EXTRACTED_CFLAGS=$(echo "$make_output" | grep "^CFLAGS" | head -1 | sed 's/^CFLAGS[[:space:]]*=[[:space:]]*//' || true)

    # インクルードパス関連を抽出
    local cppflags=$(echo "$make_output" | grep "^CPPFLAGS" | head -1 | sed 's/^CPPFLAGS[[:space:]]*=[[:space:]]*//' || true)
    local includes=$(echo "$make_output" | grep "^INCLUDES" | head -1 | sed 's/^INCLUDES[[:space:]]*=[[:space:]]*//' || true)
    local incdir=$(echo "$make_output" | grep "^INCDIR" | head -1 | sed 's/^INCDIR[[:space:]]*=[[:space:]]*//' || true)
    local inc_path=$(echo "$make_output" | grep "^INC_PATH" | head -1 | sed 's/^INC_PATH[[:space:]]*=[[:space:]]*//' || true)
    local inc=$(echo "$make_output" | grep "^INC " | head -1 | sed 's/^INC[[:space:]]*=[[:space:]]*//' || true)

    # インクルードパスを結合
    EXTRACTED_INCLUDES="$cppflags $includes $incdir $inc_path $inc"
    EXTRACTED_INCLUDES=$(echo "$EXTRACTED_INCLUDES" | tr -s ' ')

    # -Iオプションのみを抽出
    EXTRACTED_INCLUDES=$(echo "$EXTRACTED_INCLUDES" | grep -oE '\-I[^ ]+' | tr '\n' ' ' || true)

    if [ -n "$EXTRACTED_CFLAGS" ] || [ -n "$EXTRACTED_INCLUDES" ]; then
        log_success "Makefileからオプションを抽出しました"
        if [ -n "$EXTRACTED_CFLAGS" ]; then
            log_info "  CFLAGS: $EXTRACTED_CFLAGS"
        fi
        if [ -n "$EXTRACTED_INCLUDES" ]; then
            log_info "  INCLUDES: $EXTRACTED_INCLUDES"
        fi
    else
        log_warning "Makefileからコンパイルオプションを抽出できませんでした"
    fi
}

#-------------------------------------------------------------------------------
# 環境チェック
#-------------------------------------------------------------------------------
check_environment() {
    log_section "環境チェック"

    local missing_required=0

    # GCC（必須）
    if command -v gcc &> /dev/null; then
        local gcc_version=$(gcc --version | head -n1)
        log_success "GCC: $gcc_version"
    else
        log_error "GCC が見つかりません（必須）"
        missing_required=1
    fi

    # cppcheck（オプション）
    if command -v cppcheck &> /dev/null; then
        local cppcheck_version=$(cppcheck --version)
        log_success "cppcheck: $cppcheck_version"
    else
        log_warning "cppcheck が見つかりません（静的解析がスキップされます）"
    fi

    # valgrind（オプション）
    if command -v valgrind &> /dev/null; then
        local valgrind_version=$(valgrind --version)
        log_success "valgrind: $valgrind_version"
    else
        log_warning "valgrind が見つかりません（動的メモリチェックがスキップされます）"
    fi

    if [ $missing_required -ne 0 ]; then
        log_error "必須ツールが不足しています"
        exit 1
    fi
}

#-------------------------------------------------------------------------------
# 出力ディレクトリ準備
#-------------------------------------------------------------------------------
prepare_output_dir() {
    OUTPUT_DIR="./c-check-results"
    mkdir -p "$OUTPUT_DIR"

    local base_name=$(basename "$SOURCE_FILE" .c)
    REPORT_FILE="${OUTPUT_DIR}/${base_name}_report.txt"
    TEMP_BINARY="${OUTPUT_DIR}/${base_name}_test"

    # レポートファイル初期化
    cat > "$REPORT_FILE" << EOF
================================================================================
C言語 メモリ安全性・デッドロジック チェックレポート
================================================================================
ソースファイル: $SOURCE_FILE
Makefile: ${MAKEFILE:-なし}
チェック日時: $(date '+%Y-%m-%d %H:%M:%S')
================================================================================

EOF

    if [ -n "$EXTRACTED_CFLAGS" ] || [ -n "$EXTRACTED_INCLUDES" ]; then
        echo "抽出されたコンパイルオプション:" >> "$REPORT_FILE"
        echo "  CFLAGS: $EXTRACTED_CFLAGS" >> "$REPORT_FILE"
        echo "  INCLUDES: $EXTRACTED_INCLUDES" >> "$REPORT_FILE"
        echo "" >> "$REPORT_FILE"
    fi
}

#-------------------------------------------------------------------------------
# GCC警告チェック
#-------------------------------------------------------------------------------
run_gcc_warnings() {
    log_section "GCC警告チェック"

    echo "----------------------------------------" >> "$REPORT_FILE"
    echo "1. GCC警告チェック" >> "$REPORT_FILE"
    echo "----------------------------------------" >> "$REPORT_FILE"

    local gcc_output
    local gcc_exit_code

    # 厳格な警告オプションでコンパイル
    gcc_output=$(gcc -Wall -Wextra -Wpedantic \
        -Wuninitialized -Wmaybe-uninitialized \
        -Wunused -Wunused-parameter -Wunused-variable -Wunused-function \
        -Wunreachable-code -Wdead-code \
        -Wformat=2 -Wformat-overflow=2 -Wformat-truncation=2 \
        -Warray-bounds=2 \
        -Wstringop-overflow=4 \
        -Wconversion -Wsign-conversion \
        -Wnull-dereference \
        -Wdouble-promotion \
        -Wshadow \
        -Wpointer-arith \
        -Wcast-align \
        -Wstrict-prototypes \
        -Wmissing-prototypes \
        -fsyntax-only \
        $EXTRACTED_CFLAGS $EXTRACTED_INCLUDES \
        "$SOURCE_FILE" 2>&1)
    gcc_exit_code=$?

    if [ -z "$gcc_output" ]; then
        log_success "GCC警告: なし"
        echo "結果: 警告・エラーなし" >> "$REPORT_FILE"
    else
        local warning_count=$(echo "$gcc_output" | grep -c "warning:" || true)
        local error_count=$(echo "$gcc_output" | grep -c "error:" || true)

        TOTAL_WARNINGS=$((TOTAL_WARNINGS + warning_count))
        TOTAL_ERRORS=$((TOTAL_ERRORS + error_count))

        # デッドコード関連の警告をカウント
        local dead_code=$(echo "$gcc_output" | grep -c -E "(unreachable|dead|unused)" || true)
        DEAD_CODE_ISSUES=$((DEAD_CODE_ISSUES + dead_code))

        if [ $error_count -gt 0 ]; then
            log_error "GCC: エラー ${error_count}件、警告 ${warning_count}件"
        elif [ $warning_count -gt 0 ]; then
            log_warning "GCC: 警告 ${warning_count}件"
        fi

        echo "$gcc_output" >> "$REPORT_FILE"
    fi

    echo "" >> "$REPORT_FILE"
}

#-------------------------------------------------------------------------------
# cppcheckによる静的解析
#-------------------------------------------------------------------------------
run_cppcheck() {
    if ! command -v cppcheck &> /dev/null; then
        log_info "cppcheck: スキップ（インストールされていません）"
        echo "----------------------------------------" >> "$REPORT_FILE"
        echo "2. cppcheck静的解析" >> "$REPORT_FILE"
        echo "----------------------------------------" >> "$REPORT_FILE"
        echo "スキップ: cppcheckがインストールされていません" >> "$REPORT_FILE"
        echo "" >> "$REPORT_FILE"
        return
    fi

    log_section "cppcheck静的解析"

    echo "----------------------------------------" >> "$REPORT_FILE"
    echo "2. cppcheck静的解析" >> "$REPORT_FILE"
    echo "----------------------------------------" >> "$REPORT_FILE"

    local cppcheck_output

    # cppcheckで詳細な静的解析
    cppcheck_output=$(cppcheck --enable=all \
        --inconclusive \
        --force \
        --suppress=missingIncludeSystem \
        $EXTRACTED_INCLUDES \
        "$SOURCE_FILE" 2>&1)

    if echo "$cppcheck_output" | grep -q -E "(error|warning|style|performance|portability)"; then
        local issue_count=$(echo "$cppcheck_output" | grep -c -E "\(error\)|\(warning\)|\(style\)|\(performance\)|\(portability\)" || true)

        # メモリ関連の問題をカウント
        local memory_count=$(echo "$cppcheck_output" | grep -c -E "(memleak|nullPointer|uninitvar|bufferAccessOutOfBounds)" || true)
        MEMORY_ISSUES=$((MEMORY_ISSUES + memory_count))

        # デッドコード関連をカウント
        local dead_count=$(echo "$cppcheck_output" | grep -c -E "(unreachableCode|unusedFunction|unusedVariable)" || true)
        DEAD_CODE_ISSUES=$((DEAD_CODE_ISSUES + dead_count))

        TOTAL_WARNINGS=$((TOTAL_WARNINGS + issue_count))

        log_warning "cppcheck: ${issue_count}件の問題を検出"
        echo "$cppcheck_output" >> "$REPORT_FILE"
    else
        log_success "cppcheck: 問題なし"
        echo "結果: 問題なし" >> "$REPORT_FILE"
    fi

    echo "" >> "$REPORT_FILE"
}

#-------------------------------------------------------------------------------
# AddressSanitizerによるチェック
#-------------------------------------------------------------------------------
run_address_sanitizer() {
    log_section "AddressSanitizer チェック"

    echo "----------------------------------------" >> "$REPORT_FILE"
    echo "3. AddressSanitizer チェック" >> "$REPORT_FILE"
    echo "----------------------------------------" >> "$REPORT_FILE"

    # AddressSanitizerでビルド
    log_info "AddressSanitizerでビルド中..."

    local build_output
    build_output=$(gcc -fsanitize=address -fsanitize=undefined \
        -fno-omit-frame-pointer \
        -g -O1 \
        $EXTRACTED_CFLAGS $EXTRACTED_INCLUDES \
        -o "$TEMP_BINARY" \
        "$SOURCE_FILE" 2>&1)

    if [ $? -ne 0 ]; then
        log_error "ビルド失敗"
        echo "ビルド失敗:" >> "$REPORT_FILE"
        echo "$build_output" >> "$REPORT_FILE"
        echo "" >> "$REPORT_FILE"
        return 1
    fi

    log_success "ビルド成功"

    # テスト実行
    log_info "テスト実行中..."

    local asan_output
    local asan_exit_code

    # ASANの出力を取得
    export ASAN_OPTIONS="detect_leaks=1:halt_on_error=0:print_stats=1"

    if [ ${#TEST_ARGS[@]} -gt 0 ]; then
        asan_output=$("$TEMP_BINARY" "${TEST_ARGS[@]}" 2>&1) || true
    else
        asan_output=$("$TEMP_BINARY" 2>&1) || true
    fi
    asan_exit_code=$?

    if echo "$asan_output" | grep -q "ERROR: AddressSanitizer"; then
        local asan_errors=$(echo "$asan_output" | grep -c "ERROR: AddressSanitizer" || true)
        MEMORY_ISSUES=$((MEMORY_ISSUES + asan_errors))
        TOTAL_ERRORS=$((TOTAL_ERRORS + asan_errors))

        log_error "AddressSanitizer: ${asan_errors}件のメモリエラーを検出"
        echo "検出されたエラー:" >> "$REPORT_FILE"
        echo "$asan_output" >> "$REPORT_FILE"
    elif echo "$asan_output" | grep -q "LeakSanitizer"; then
        local leak_count=$(echo "$asan_output" | grep -c "Direct leak\|Indirect leak" || true)
        MEMORY_ISSUES=$((MEMORY_ISSUES + leak_count))
        TOTAL_WARNINGS=$((TOTAL_WARNINGS + leak_count))

        log_warning "AddressSanitizer: ${leak_count}件のメモリリークを検出"
        echo "$asan_output" >> "$REPORT_FILE"
    else
        log_success "AddressSanitizer: メモリエラーなし"
        echo "結果: メモリエラーなし" >> "$REPORT_FILE"
    fi

    echo "" >> "$REPORT_FILE"

    # テンポラリバイナリを削除
    rm -f "$TEMP_BINARY"
}

#-------------------------------------------------------------------------------
# Valgrindによるメモリチェック
#-------------------------------------------------------------------------------
run_valgrind() {
    if ! command -v valgrind &> /dev/null; then
        log_info "Valgrind: スキップ（インストールされていません）"
        echo "----------------------------------------" >> "$REPORT_FILE"
        echo "4. Valgrind メモリチェック" >> "$REPORT_FILE"
        echo "----------------------------------------" >> "$REPORT_FILE"
        echo "スキップ: valgrindがインストールされていません" >> "$REPORT_FILE"
        echo "" >> "$REPORT_FILE"
        return
    fi

    log_section "Valgrind メモリチェック"

    echo "----------------------------------------" >> "$REPORT_FILE"
    echo "4. Valgrind メモリチェック" >> "$REPORT_FILE"
    echo "----------------------------------------" >> "$REPORT_FILE"

    # デバッグ情報付きでビルド
    log_info "デバッグビルド中..."

    local build_output
    build_output=$(gcc -g -O0 $EXTRACTED_CFLAGS $EXTRACTED_INCLUDES -o "$TEMP_BINARY" "$SOURCE_FILE" 2>&1)

    if [ $? -ne 0 ]; then
        log_error "ビルド失敗"
        echo "ビルド失敗:" >> "$REPORT_FILE"
        echo "$build_output" >> "$REPORT_FILE"
        echo "" >> "$REPORT_FILE"
        return 1
    fi

    log_success "ビルド成功"

    # Valgrind実行
    log_info "Valgrind実行中..."

    local valgrind_output
    if [ ${#TEST_ARGS[@]} -gt 0 ]; then
        valgrind_output=$(valgrind --leak-check=full \
            --show-leak-kinds=all \
            --track-origins=yes \
            --verbose \
            "$TEMP_BINARY" "${TEST_ARGS[@]}" 2>&1)
    else
        valgrind_output=$(valgrind --leak-check=full \
            --show-leak-kinds=all \
            --track-origins=yes \
            --verbose \
            "$TEMP_BINARY" 2>&1)
    fi

    # エラーサマリーを解析
    local error_summary=$(echo "$valgrind_output" | grep "ERROR SUMMARY")
    local leak_summary=$(echo "$valgrind_output" | grep -A5 "LEAK SUMMARY")

    if echo "$error_summary" | grep -q "0 errors"; then
        log_success "Valgrind: エラーなし"
        echo "結果: メモリエラーなし" >> "$REPORT_FILE"
    else
        local error_count=$(echo "$error_summary" | grep -oP '\d+(?= errors)' || echo "0")
        MEMORY_ISSUES=$((MEMORY_ISSUES + error_count))
        TOTAL_ERRORS=$((TOTAL_ERRORS + error_count))

        log_error "Valgrind: ${error_count}件のエラーを検出"
        echo "$valgrind_output" >> "$REPORT_FILE"
    fi

    # リーク情報を追加
    if [ -n "$leak_summary" ]; then
        echo "" >> "$REPORT_FILE"
        echo "リークサマリー:" >> "$REPORT_FILE"
        echo "$leak_summary" >> "$REPORT_FILE"
    fi

    echo "" >> "$REPORT_FILE"

    # テンポラリバイナリを削除
    rm -f "$TEMP_BINARY"
}

#-------------------------------------------------------------------------------
# サマリー表示
#-------------------------------------------------------------------------------
show_summary() {
    log_section "チェック結果サマリー"

    echo "========================================" >> "$REPORT_FILE"
    echo "サマリー" >> "$REPORT_FILE"
    echo "========================================" >> "$REPORT_FILE"

    echo ""
    echo -e "${BOLD}チェック対象:${NC} $SOURCE_FILE"
    echo -e "${BOLD}レポート:${NC} $REPORT_FILE"
    echo ""

    # 結果表示
    if [ $TOTAL_ERRORS -gt 0 ]; then
        echo -e "${RED}${BOLD}エラー: ${TOTAL_ERRORS}件${NC}"
    else
        echo -e "${GREEN}${BOLD}エラー: 0件${NC}"
    fi

    if [ $TOTAL_WARNINGS -gt 0 ]; then
        echo -e "${YELLOW}${BOLD}警告: ${TOTAL_WARNINGS}件${NC}"
    else
        echo -e "${GREEN}${BOLD}警告: 0件${NC}"
    fi

    if [ $MEMORY_ISSUES -gt 0 ]; then
        echo -e "${RED}${BOLD}メモリ問題: ${MEMORY_ISSUES}件${NC}"
    else
        echo -e "${GREEN}${BOLD}メモリ問題: 0件${NC}"
    fi

    if [ $DEAD_CODE_ISSUES -gt 0 ]; then
        echo -e "${YELLOW}${BOLD}デッドコード: ${DEAD_CODE_ISSUES}件${NC}"
    else
        echo -e "${GREEN}${BOLD}デッドコード: 0件${NC}"
    fi

    echo ""
    echo "エラー: ${TOTAL_ERRORS}件" >> "$REPORT_FILE"
    echo "警告: ${TOTAL_WARNINGS}件" >> "$REPORT_FILE"
    echo "メモリ問題: ${MEMORY_ISSUES}件" >> "$REPORT_FILE"
    echo "デッドコード: ${DEAD_CODE_ISSUES}件" >> "$REPORT_FILE"

    # 総合判定
    echo ""
    if [ $TOTAL_ERRORS -eq 0 ] && [ $MEMORY_ISSUES -eq 0 ]; then
        if [ $TOTAL_WARNINGS -eq 0 ] && [ $DEAD_CODE_ISSUES -eq 0 ]; then
            echo -e "${GREEN}${BOLD}[結果] すべてのチェックに合格しました${NC}"
            echo "" >> "$REPORT_FILE"
            echo "[結果] すべてのチェックに合格しました" >> "$REPORT_FILE"
        else
            echo -e "${YELLOW}${BOLD}[結果] 警告がありますが、重大な問題は検出されませんでした${NC}"
            echo "" >> "$REPORT_FILE"
            echo "[結果] 警告がありますが、重大な問題は検出されませんでした" >> "$REPORT_FILE"
        fi
    else
        echo -e "${RED}${BOLD}[結果] 修正が必要な問題が検出されました${NC}"
        echo "" >> "$REPORT_FILE"
        echo "[結果] 修正が必要な問題が検出されました" >> "$REPORT_FILE"
    fi

    echo ""
    log_info "詳細は ${REPORT_FILE} を参照してください"
}

#-------------------------------------------------------------------------------
# メイン処理
#-------------------------------------------------------------------------------
main() {
    # ヘルプ表示
    if [ $# -eq 0 ] || [ "$1" = "--help" ] || [ "$1" = "-h" ]; then
        show_help
        exit 0
    fi

    # 引数解析
    SOURCE_FILE="$1"
    shift

    # オプションと -- 以降のテスト引数を解析
    while [ $# -gt 0 ]; do
        case "$1" in
            -m)
                if [ $# -lt 2 ]; then
                    log_error "-m オプションにはMakefileパスが必要です"
                    exit 1
                fi
                MAKEFILE="$2"
                shift 2
                ;;
            --)
                shift
                TEST_ARGS=("$@")
                break
                ;;
            *)
                # 不明なオプションはテスト引数として扱う
                TEST_ARGS=("$@")
                break
                ;;
        esac
    done

    # ソースファイル存在チェック
    if [ ! -f "$SOURCE_FILE" ]; then
        log_error "ファイルが見つかりません: $SOURCE_FILE"
        exit 1
    fi

    # 拡張子チェック
    if [[ ! "$SOURCE_FILE" =~ \.c$ ]]; then
        log_warning "Cソースファイル(.c)ではない可能性があります: $SOURCE_FILE"
    fi

    echo ""
    echo -e "${BOLD}${CYAN}============================================${NC}"
    echo -e "${BOLD}${CYAN} C言語 メモリ安全性チェックツール v${VERSION}${NC}"
    echo -e "${BOLD}${CYAN}============================================${NC}"
    echo ""
    echo -e "${BOLD}チェック対象:${NC} $SOURCE_FILE"
    if [ -n "$MAKEFILE" ]; then
        echo -e "${BOLD}Makefile:${NC} $MAKEFILE"
    fi
    if [ ${#TEST_ARGS[@]} -gt 0 ]; then
        echo -e "${BOLD}テスト引数:${NC} ${TEST_ARGS[*]}"
    fi

    # Makefileからオプション抽出
    extract_makefile_options

    # 環境チェック
    check_environment

    # 出力ディレクトリ準備
    prepare_output_dir

    # 各種チェック実行
    run_gcc_warnings
    run_cppcheck
    run_address_sanitizer
    run_valgrind

    # サマリー表示
    show_summary

    # 終了コード
    if [ $TOTAL_ERRORS -gt 0 ] || [ $MEMORY_ISSUES -gt 0 ]; then
        exit 1
    else
        exit 0
    fi
}

# スクリプト実行
main "$@"
