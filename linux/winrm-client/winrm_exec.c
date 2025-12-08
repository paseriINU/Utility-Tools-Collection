/*
 * ============================================================================
 * WinRM Remote Batch Executor for Linux (C言語版 - 標準ライブラリのみ)
 * ============================================================================
 *
 * 【概要】
 * このプログラムは、LinuxからWindows ServerへWinRM (Windows Remote Management)
 * プロトコルを使用してリモート接続し、バッチファイルを実行するツールです。
 *
 * 【特徴】
 * - 標準ライブラリのみ使用（外部ライブラリ不要）
 * - NTLM v2認証を自前実装（MD4, MD5, HMAC-MD5を含む）
 * - IT制限環境でも動作可能（pip/yum等のパッケージ管理不要）
 * - Windows側の設定変更不要（デフォルトのNTLM認証を使用）
 *
 * 【なぜ標準ライブラリのみで実装するのか】
 * 企業のIT制限環境では、外部ライブラリのインストールが禁止されていることが多い。
 * このプログラムは、そのような環境でも確実に動作するよう設計されている。
 *
 * 【NTLM認証の仕組み】
 * NTLM認証は3段階のハンドシェイクで構成される：
 *   1. Type 1 (Negotiate): クライアントが認証開始を宣言
 *   2. Type 2 (Challenge): サーバーがランダムなチャレンジを返送
 *   3. Type 3 (Authenticate): クライアントがチャレンジに対する応答を送信
 *
 * 【WinRMプロトコルの流れ】
 *   1. シェル作成 (Create): リモートシェルセッションを開始
 *   2. コマンド実行 (Command): バッチファイルを実行
 *   3. 出力取得 (Receive): 標準出力・標準エラーを取得
 *   4. シェル削除 (Delete): セッションをクリーンアップ
 *
 * 【必要な環境】
 * - Linux（Red Hat, CentOS, Ubuntu等）
 * - GCCコンパイラ
 * - ネットワーク接続（ポート5985/HTTP または 5986/HTTPS）
 *
 * 【コンパイル方法】
 *   gcc -o winrm_exec winrm_exec.c
 *   # 警告を確認する場合
 *   gcc -Wall -o winrm_exec winrm_exec.c
 *
 * 【使い方】
 *   1. このソースファイル内の設定セクションを編集
 *   2. コンパイル: gcc -o winrm_exec winrm_exec.c
 *   3. 実行: ./winrm_exec ENV
 *
 *   環境を引数で指定（必須）:
 *   ./winrm_exec TST1T
 *   ./winrm_exec TST2T
 *
 *   または環境変数で設定を上書き:
 *   WINRM_HOST=192.168.1.100 WINRM_USER=Admin WINRM_PASS=Pass123 ./winrm_exec TST1T
 *
 * 【セキュリティに関する注意】
 * - パスワードはソースコード内に記載するため、適切なファイル権限を設定すること
 * - 本番環境では環境変数での上書きを推奨
 *
 * ============================================================================
 */

/* ============================================================================
 * インクルードファイル
 * すべて標準ライブラリのみを使用（外部依存なし）
 * ============================================================================ */

#include <stdio.h>      /* 標準入出力: printf, fprintf, fopen等 */
#include <stdlib.h>     /* 標準ユーティリティ: malloc, free, getenv, atoi等 */
#include <string.h>     /* 文字列操作: strcpy, strcat, strlen, memcpy等 */
#include <stdbool.h>    /* ブール型: true, false */
#include <stdint.h>     /* 固定幅整数型: uint8_t, uint16_t, uint32_t, uint64_t */
#include <time.h>       /* 時間関連: time, srand */
#include <unistd.h>     /* POSIX API: close, usleep, read, write */
#include <sys/socket.h> /* ソケットAPI: socket, connect, send, recv */
#include <sys/time.h>   /* 時間構造体: gettimeofday, timeval */
#include <netinet/in.h> /* インターネットアドレス: sockaddr_in, htons */
#include <netdb.h>      /* ネットワークデータベース: gethostbyname, hostent */
#include <errno.h>      /* エラー番号: errno */
#include <fcntl.h>      /* ファイル制御: open, O_RDONLY等 */

/* ============================================================================
 * 設定セクション（ユーザー編集エリア）
 * ============================================================================
 *
 * 【使用方法】
 * 1. 以下の設定値を環境に合わせて編集してください
 * 2. コンパイルして実行: gcc -o winrm_exec winrm_exec.c && ./winrm_exec TST1T
 *
 * 【設定の優先順位】
 * 環境変数 > ソースコード内の設定
 * 例: WINRM_HOST=10.0.0.1 ./winrm_exec TST1T
 *     → 環境変数の値が優先される
 * ============================================================================ */

/* --- Windows接続情報 --- */
#define DEFAULT_HOST "192.168.1.100"     /* Windows ServerのIPアドレスまたはホスト名 */
#define DEFAULT_USER "Administrator"      /* Windowsのログインユーザー名 */
#define DEFAULT_PASS "YourPassword"       /* Windowsのログインパスワード */
#define DEFAULT_DOMAIN ""                 /* ドメイン名（空文字列 = ローカル認証） */
#define DEFAULT_PORT 5985                 /* WinRMポート: HTTP=5985, HTTPS=5986 */

/* --- 実行するバッチファイル ---
 * {ENV} プレースホルダは実行時に環境名（TST1T等）に置換されます
 * 例: "C:\\Scripts\\{ENV}\\test.bat" → "C:\\Scripts\\TST1T\\test.bat"
 * 注: {ENV}は複数箇所に使用可能（whileループで全て置換） */
#define DEFAULT_BATCH_PATH "C:\\Scripts\\{ENV}\\test.bat"

/* --- 利用可能な環境のリスト ---
 * コマンドライン引数で指定可能な環境名を定義
 * 新しい環境を追加する場合はこのリストに追加してください
 * 最後は必ずNULLで終端すること */
static const char *ENVIRONMENTS[] = {"TST1T", "TST2T", NULL};

/* --- タイムアウト設定 ---
 * コマンド実行の最大待機時間（秒）
 * バッチ処理が長時間かかる場合は増やしてください */
#define TIMEOUT 300

/* --- デバッグモード ---
 * 1: 送受信するSOAP XMLを表示（トラブルシューティング用）
 * 0: 通常モード（本番運用時はこちら） */
#define DEBUG 0

/* ============================================================================ */

/* ============================================================================
 * 定数定義
 * ============================================================================ */

/* --- コンソール色付き出力用ANSIエスケープシーケンス ---
 * Linuxターミナルでログメッセージを色分け表示するために使用 */
#define COLOR_RED     "\033[0;31m"  /* エラーメッセージ用（赤） */
#define COLOR_GREEN   "\033[0;32m"  /* 成功メッセージ用（緑） */
#define COLOR_YELLOW  "\033[1;33m"  /* 警告メッセージ用（黄） */
#define COLOR_BLUE    "\033[0;34m"  /* 情報メッセージ用（青） */
#define COLOR_RESET   "\033[0m"     /* 色をリセット */

/* --- バッファサイズ定義 ---
 * 各種データ格納用のバッファサイズを定義
 * 大きなXMLレスポンスを扱うため、十分なサイズを確保 */
#define MAX_BUFFER_SIZE 65536   /* HTTPレスポンス受信用バッファ（64KB） */
#define MAX_HEADER_SIZE 4096    /* HTTPヘッダー用バッファ（4KB） */
#define MAX_URL_SIZE 512        /* URL文字列用バッファ */
#define MAX_UUID_SIZE 64        /* UUID文字列用バッファ */
#define MAX_ENVELOPE_SIZE 8192  /* SOAP XMLエンベロープ用バッファ（8KB） */

/* ============================================================================
 * NTLM認証プロトコル定数
 * ============================================================================
 *
 * NTLM (NT LAN Manager) は、Microsoftが開発したチャレンジ・レスポンス認証プロトコル。
 * Windows環境でデフォルトで有効になっており、設定変更なしで使用可能。
 *
 * 【認証フロー】
 * クライアント                    サーバー
 *     |--- Type 1 (Negotiate) --->|  認証開始を宣言
 *     |<-- Type 2 (Challenge) ----|  8バイトのランダムチャレンジを送信
 *     |--- Type 3 (Authenticate)->|  チャレンジに対する暗号化応答
 *     |<-- 認証成功/失敗 ---------|
 * ============================================================================ */

/* NTLM署名: すべてのNTLMメッセージの先頭に配置される固定文字列 */
#define NTLM_SIGNATURE "NTLMSSP\0"

/* NTLMメッセージタイプ */
#define NTLM_TYPE1 1  /* Negotiate: クライアント→サーバー（認証開始） */
#define NTLM_TYPE2 2  /* Challenge: サーバー→クライアント（チャレンジ送信） */
#define NTLM_TYPE3 3  /* Authenticate: クライアント→サーバー（認証応答） */

/* NTLMネゴシエーションフラグ
 * クライアントとサーバー間で使用する機能を交渉するためのビットフラグ */
#define NTLMSSP_NEGOTIATE_UNICODE          0x00000001  /* Unicode文字列を使用 */
#define NTLMSSP_NEGOTIATE_OEM              0x00000002  /* OEM文字列を使用 */
#define NTLMSSP_REQUEST_TARGET             0x00000004  /* ターゲット情報を要求 */
#define NTLMSSP_NEGOTIATE_NTLM             0x00000200  /* NTLM認証を使用 */
#define NTLMSSP_NEGOTIATE_ALWAYS_SIGN      0x00008000  /* 常に署名を使用 */
#define NTLMSSP_NEGOTIATE_EXTENDED_SESSIONSECURITY 0x00080000  /* NTLMv2セッションセキュリティ */
#define NTLMSSP_NEGOTIATE_TARGET_INFO      0x00800000  /* TargetInfo構造体を含む */

/* ============================================================================
 * グローバル変数
 * ============================================================================
 * 実行時に設定される接続情報を保持する変数群
 * load_config()関数で初期化され、環境変数での上書きも可能
 * ============================================================================ */
static char g_host[256];        /* 接続先ホスト名/IPアドレス */
static char g_user[256];        /* 認証ユーザー名 */
static char g_pass[256];        /* 認証パスワード */
static char g_domain[256];      /* ドメイン名（ローカル認証時は空） */
static int g_port;              /* WinRMポート番号 */
static char g_batch_path[512];  /* 実行するバッチファイルのパス */
static char g_env_folder[64];   /* 選択された環境フォルダ名 */

/* ============================================================================
 * ログ出力関数
 * ============================================================================
 * コンソールに色付きでログメッセージを出力するユーティリティ関数群
 * 標準エラー出力(stderr)を使用することで、標準出力と分離可能
 * ============================================================================ */

/* 情報メッセージ（青）: 処理の進行状況等を通知 */
static void log_info(const char *msg) {
    fprintf(stderr, "%s[INFO]%s %s\n", COLOR_BLUE, COLOR_RESET, msg);
}

/* 成功メッセージ（緑）: 処理が正常に完了したことを通知 */
static void log_success(const char *msg) {
    fprintf(stderr, "%s[SUCCESS]%s %s\n", COLOR_GREEN, COLOR_RESET, msg);
}

/* 警告メッセージ（黄）: 注意が必要だが処理は継続可能 */
static void log_warn(const char *msg) {
    fprintf(stderr, "%s[WARN]%s %s\n", COLOR_YELLOW, COLOR_RESET, msg);
}

/* エラーメッセージ（赤）: 処理が失敗したことを通知 */
static void log_error(const char *msg) {
    fprintf(stderr, "%s[ERROR]%s %s\n", COLOR_RED, COLOR_RESET, msg);
}

/* ============================================================================
 * MD4ハッシュアルゴリズム実装
 * ============================================================================
 *
 * 【なぜMD4を自前実装するのか】
 * - MD4はNTLM認証でパスワードハッシュの計算に必須
 * - OpenSSLのMD4は多くのディストリビューションで非推奨/削除済み
 * - 標準ライブラリにはMD4が含まれていない
 * - 外部ライブラリに依存せずIT制限環境で動作させるため
 *
 * 【MD4アルゴリズムの概要】
 * - 入力: 任意長のデータ
 * - 出力: 128ビット（16バイト）のハッシュ値
 * - RFC 1320で定義
 * - 3ラウンド×16ステップ = 48ステップの処理
 *
 * 【セキュリティ上の注意】
 * MD4は暗号学的に破られており、新規開発には推奨されない。
 * ただし、NTLMプロトコルとの互換性のために使用が必要。
 * ============================================================================ */

/* --- MD4補助関数 ---
 * 各ラウンドで使用される非線形関数 */

/* Round 1用: 条件選択関数 - xが1のビットはyを、0のビットはzを選択 */
static uint32_t md4_F(uint32_t x, uint32_t y, uint32_t z) {
    return (x & y) | (~x & z);
}

/* Round 2用: 多数決関数 - 3つの入力のうち2つ以上が1なら1 */
static uint32_t md4_G(uint32_t x, uint32_t y, uint32_t z) {
    return (x & y) | (x & z) | (y & z);
}

/* Round 3用: パリティ関数 - XOR演算 */
static uint32_t md4_H(uint32_t x, uint32_t y, uint32_t z) {
    return x ^ y ^ z;
}

/* 左循環シフト（ローテート）: ビットを左にn個回転 */
static uint32_t md4_rotate_left(uint32_t x, int n) {
    return (x << n) | (x >> (32 - n));
}

/*
 * md4_hash - MD4ハッシュを計算
 *
 * @input:  ハッシュ対象のデータ
 * @len:    データの長さ（バイト）
 * @output: ハッシュ値の出力先（16バイト以上必要）
 *
 * 処理の流れ:
 * 1. メッセージのパディング（512ビットの倍数に調整）
 * 2. 64バイトブロックごとに処理
 * 3. 各ブロックで3ラウンド×16ステップの変換
 * 4. 最終的な状態を出力
 */
static void md4_hash(const uint8_t *input, size_t len, uint8_t *output) {
    uint32_t a0 = 0x67452301;
    uint32_t b0 = 0xefcdab89;
    uint32_t c0 = 0x98badcfe;
    uint32_t d0 = 0x10325476;

    /* パディング */
    size_t new_len = ((len + 8) / 64 + 1) * 64;
    uint8_t *msg = calloc(new_len, 1);
    memcpy(msg, input, len);
    msg[len] = 0x80;

    uint64_t bits_len = len * 8;
    memcpy(msg + new_len - 8, &bits_len, 8);

    /* 各ブロックを処理 */
    for (size_t offset = 0; offset < new_len; offset += 64) {
        uint32_t *M = (uint32_t *)(msg + offset);
        uint32_t A = a0, B = b0, C = c0, D = d0;

        /* Round 1 */
        A = md4_rotate_left(A + md4_F(B, C, D) + M[0], 3);
        D = md4_rotate_left(D + md4_F(A, B, C) + M[1], 7);
        C = md4_rotate_left(C + md4_F(D, A, B) + M[2], 11);
        B = md4_rotate_left(B + md4_F(C, D, A) + M[3], 19);
        A = md4_rotate_left(A + md4_F(B, C, D) + M[4], 3);
        D = md4_rotate_left(D + md4_F(A, B, C) + M[5], 7);
        C = md4_rotate_left(C + md4_F(D, A, B) + M[6], 11);
        B = md4_rotate_left(B + md4_F(C, D, A) + M[7], 19);
        A = md4_rotate_left(A + md4_F(B, C, D) + M[8], 3);
        D = md4_rotate_left(D + md4_F(A, B, C) + M[9], 7);
        C = md4_rotate_left(C + md4_F(D, A, B) + M[10], 11);
        B = md4_rotate_left(B + md4_F(C, D, A) + M[11], 19);
        A = md4_rotate_left(A + md4_F(B, C, D) + M[12], 3);
        D = md4_rotate_left(D + md4_F(A, B, C) + M[13], 7);
        C = md4_rotate_left(C + md4_F(D, A, B) + M[14], 11);
        B = md4_rotate_left(B + md4_F(C, D, A) + M[15], 19);

        /* Round 2 */
        A = md4_rotate_left(A + md4_G(B, C, D) + M[0] + 0x5a827999, 3);
        D = md4_rotate_left(D + md4_G(A, B, C) + M[4] + 0x5a827999, 5);
        C = md4_rotate_left(C + md4_G(D, A, B) + M[8] + 0x5a827999, 9);
        B = md4_rotate_left(B + md4_G(C, D, A) + M[12] + 0x5a827999, 13);
        A = md4_rotate_left(A + md4_G(B, C, D) + M[1] + 0x5a827999, 3);
        D = md4_rotate_left(D + md4_G(A, B, C) + M[5] + 0x5a827999, 5);
        C = md4_rotate_left(C + md4_G(D, A, B) + M[9] + 0x5a827999, 9);
        B = md4_rotate_left(B + md4_G(C, D, A) + M[13] + 0x5a827999, 13);
        A = md4_rotate_left(A + md4_G(B, C, D) + M[2] + 0x5a827999, 3);
        D = md4_rotate_left(D + md4_G(A, B, C) + M[6] + 0x5a827999, 5);
        C = md4_rotate_left(C + md4_G(D, A, B) + M[10] + 0x5a827999, 9);
        B = md4_rotate_left(B + md4_G(C, D, A) + M[14] + 0x5a827999, 13);
        A = md4_rotate_left(A + md4_G(B, C, D) + M[3] + 0x5a827999, 3);
        D = md4_rotate_left(D + md4_G(A, B, C) + M[7] + 0x5a827999, 5);
        C = md4_rotate_left(C + md4_G(D, A, B) + M[11] + 0x5a827999, 9);
        B = md4_rotate_left(B + md4_G(C, D, A) + M[15] + 0x5a827999, 13);

        /* Round 3 */
        A = md4_rotate_left(A + md4_H(B, C, D) + M[0] + 0x6ed9eba1, 3);
        D = md4_rotate_left(D + md4_H(A, B, C) + M[8] + 0x6ed9eba1, 9);
        C = md4_rotate_left(C + md4_H(D, A, B) + M[4] + 0x6ed9eba1, 11);
        B = md4_rotate_left(B + md4_H(C, D, A) + M[12] + 0x6ed9eba1, 15);
        A = md4_rotate_left(A + md4_H(B, C, D) + M[2] + 0x6ed9eba1, 3);
        D = md4_rotate_left(D + md4_H(A, B, C) + M[10] + 0x6ed9eba1, 9);
        C = md4_rotate_left(C + md4_H(D, A, B) + M[6] + 0x6ed9eba1, 11);
        B = md4_rotate_left(B + md4_H(C, D, A) + M[14] + 0x6ed9eba1, 15);
        A = md4_rotate_left(A + md4_H(B, C, D) + M[1] + 0x6ed9eba1, 3);
        D = md4_rotate_left(D + md4_H(A, B, C) + M[9] + 0x6ed9eba1, 9);
        C = md4_rotate_left(C + md4_H(D, A, B) + M[5] + 0x6ed9eba1, 11);
        B = md4_rotate_left(B + md4_H(C, D, A) + M[13] + 0x6ed9eba1, 15);
        A = md4_rotate_left(A + md4_H(B, C, D) + M[3] + 0x6ed9eba1, 3);
        D = md4_rotate_left(D + md4_H(A, B, C) + M[11] + 0x6ed9eba1, 9);
        C = md4_rotate_left(C + md4_H(D, A, B) + M[7] + 0x6ed9eba1, 11);
        B = md4_rotate_left(B + md4_H(C, D, A) + M[15] + 0x6ed9eba1, 15);

        a0 += A;
        b0 += B;
        c0 += C;
        d0 += D;
    }

    free(msg);

    memcpy(output, &a0, 4);
    memcpy(output + 4, &b0, 4);
    memcpy(output + 8, &c0, 4);
    memcpy(output + 12, &d0, 4);
}

/* ============================================================================
 * MD5ハッシュアルゴリズム実装
 * ============================================================================
 *
 * 【MD5の役割】
 * NTLMv2認証では、MD5がHMAC-MD5の内部アルゴリズムとして使用される。
 * - NTProofStr（認証応答）の計算
 * - セッションキーの導出
 *
 * 【MD5アルゴリズムの概要】
 * - 入力: 任意長のデータ
 * - 出力: 128ビット（16バイト）のハッシュ値
 * - RFC 1321で定義
 * - 4ラウンド×16ステップ = 64ステップの処理
 * ============================================================================ */

/* MD5定数テーブル: sin関数から導出された64個の定数 */
static const uint32_t md5_k[] = {
    0xd76aa478, 0xe8c7b756, 0x242070db, 0xc1bdceee,
    0xf57c0faf, 0x4787c62a, 0xa8304613, 0xfd469501,
    0x698098d8, 0x8b44f7af, 0xffff5bb1, 0x895cd7be,
    0x6b901122, 0xfd987193, 0xa679438e, 0x49b40821,
    0xf61e2562, 0xc040b340, 0x265e5a51, 0xe9b6c7aa,
    0xd62f105d, 0x02441453, 0xd8a1e681, 0xe7d3fbc8,
    0x21e1cde6, 0xc33707d6, 0xf4d50d87, 0x455a14ed,
    0xa9e3e905, 0xfcefa3f8, 0x676f02d9, 0x8d2a4c8a,
    0xfffa3942, 0x8771f681, 0x6d9d6122, 0xfde5380c,
    0xa4beea44, 0x4bdecfa9, 0xf6bb4b60, 0xbebfbc70,
    0x289b7ec6, 0xeaa127fa, 0xd4ef3085, 0x04881d05,
    0xd9d4d039, 0xe6db99e5, 0x1fa27cf8, 0xc4ac5665,
    0xf4292244, 0x432aff97, 0xab9423a7, 0xfc93a039,
    0x655b59c3, 0x8f0ccc92, 0xffeff47d, 0x85845dd1,
    0x6fa87e4f, 0xfe2ce6e0, 0xa3014314, 0x4e0811a1,
    0xf7537e82, 0xbd3af235, 0x2ad7d2bb, 0xeb86d391
};

/* MD5シフト量テーブル: 各ステップでの左循環シフト量 */
static const uint32_t md5_s[] = {
    7, 12, 17, 22, 7, 12, 17, 22, 7, 12, 17, 22, 7, 12, 17, 22,  /* Round 1 */
    5, 9, 14, 20, 5, 9, 14, 20, 5, 9, 14, 20, 5, 9, 14, 20,      /* Round 2 */
    4, 11, 16, 23, 4, 11, 16, 23, 4, 11, 16, 23, 4, 11, 16, 23,  /* Round 3 */
    6, 10, 15, 21, 6, 10, 15, 21, 6, 10, 15, 21, 6, 10, 15, 21   /* Round 4 */
};

/*
 * md5_hash - MD5ハッシュを計算
 *
 * @input:  ハッシュ対象のデータ
 * @len:    データの長さ（バイト）
 * @output: ハッシュ値の出力先（16バイト以上必要）
 */
static void md5_hash(const uint8_t *input, size_t len, uint8_t *output) {
    uint32_t a0 = 0x67452301;
    uint32_t b0 = 0xefcdab89;
    uint32_t c0 = 0x98badcfe;
    uint32_t d0 = 0x10325476;

    size_t new_len = ((len + 8) / 64 + 1) * 64;
    uint8_t *msg = calloc(new_len, 1);
    memcpy(msg, input, len);
    msg[len] = 0x80;

    uint64_t bits_len = len * 8;
    memcpy(msg + new_len - 8, &bits_len, 8);

    for (size_t offset = 0; offset < new_len; offset += 64) {
        uint32_t *M = (uint32_t *)(msg + offset);
        uint32_t A = a0, B = b0, C = c0, D = d0;

        for (int i = 0; i < 64; i++) {
            uint32_t F, g;
            if (i < 16) {
                F = (B & C) | (~B & D);
                g = i;
            } else if (i < 32) {
                F = (D & B) | (~D & C);
                g = (5 * i + 1) % 16;
            } else if (i < 48) {
                F = B ^ C ^ D;
                g = (3 * i + 5) % 16;
            } else {
                F = C ^ (B | ~D);
                g = (7 * i) % 16;
            }
            F = F + A + md5_k[i] + M[g];
            A = D;
            D = C;
            C = B;
            B = B + md4_rotate_left(F, md5_s[i]);
        }

        a0 += A;
        b0 += B;
        c0 += C;
        d0 += D;
    }

    free(msg);

    memcpy(output, &a0, 4);
    memcpy(output + 4, &b0, 4);
    memcpy(output + 8, &c0, 4);
    memcpy(output + 12, &d0, 4);
}

/* ============================================================================
 * HMAC-MD5実装
 * ============================================================================
 *
 * 【HMAC (Hash-based Message Authentication Code) とは】
 * 秘密鍵を使用してメッセージの認証コードを生成するアルゴリズム。
 * RFC 2104で定義。
 *
 * 【NTLMv2での用途】
 * - NTLMv2ハッシュの生成: HMAC-MD5(NT_Hash, Username.upper() + Domain)
 * - NTProofStrの生成: HMAC-MD5(NTLMv2Hash, ServerChallenge + Blob)
 * - セッションキーの生成: HMAC-MD5(NTLMv2Hash, NTProofStr)
 *
 * 【計算式】
 * HMAC(K, m) = H((K' XOR opad) || H((K' XOR ipad) || m))
 * - K': 鍵（64バイトにパディング）
 * - ipad: 0x36を64回繰り返し
 * - opad: 0x5cを64回繰り返し
 * - H: ハッシュ関数（ここではMD5）
 * ============================================================================ */

/*
 * hmac_md5 - HMAC-MD5を計算
 *
 * @key:      秘密鍵
 * @key_len:  秘密鍵の長さ
 * @data:     認証対象のデータ
 * @data_len: データの長さ
 * @output:   HMAC値の出力先（16バイト）
 */
static void hmac_md5(const uint8_t *key, size_t key_len,
                     const uint8_t *data, size_t data_len,
                     uint8_t *output) {
    uint8_t k[64] = {0};
    uint8_t o_key_pad[64];
    uint8_t i_key_pad[64];

    if (key_len > 64) {
        md5_hash(key, key_len, k);
    } else {
        memcpy(k, key, key_len);
    }

    for (int i = 0; i < 64; i++) {
        o_key_pad[i] = k[i] ^ 0x5c;
        i_key_pad[i] = k[i] ^ 0x36;
    }

    uint8_t *inner = malloc(64 + data_len);
    memcpy(inner, i_key_pad, 64);
    memcpy(inner + 64, data, data_len);

    uint8_t inner_hash[16];
    md5_hash(inner, 64 + data_len, inner_hash);
    free(inner);

    uint8_t outer[64 + 16];
    memcpy(outer, o_key_pad, 64);
    memcpy(outer + 64, inner_hash, 16);

    md5_hash(outer, 80, output);
}

/* ============================================================================
 * Base64エンコード/デコード実装
 * ============================================================================
 *
 * 【Base64とは】
 * バイナリデータをASCII文字列として表現するためのエンコード方式。
 * RFC 4648で定義。
 *
 * 【NTLMでの用途】
 * HTTPヘッダでNTLMメッセージをやり取りする際に使用。
 * バイナリのNTLMメッセージをテキストとして送受信可能にする。
 *
 * 例: Authorization: NTLM TlRMTVNTUAABAAAABzIAAAYABgAoAAAA...
 *
 * 【エンコード方式】
 * 3バイト（24ビット）を4文字（6ビット×4）に変換。
 * 64種類の文字（A-Z, a-z, 0-9, +, /）を使用。
 * パディングには'='を使用。
 * ============================================================================ */

/* Base64文字セット */
static const char base64_chars[] =
    "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

/*
 * base64_encode - バイナリデータをBase64文字列にエンコード
 *
 * @input:  エンコード対象のバイナリデータ
 * @len:    データの長さ
 * @output: 出力先バッファ（入力の約1.4倍のサイズが必要）
 * @return: 出力された文字数
 */
static size_t base64_encode(const uint8_t *input, size_t len, char *output) {
    size_t out_len = 0;

    for (size_t i = 0; i < len; i += 3) {
        uint32_t octet_a = i < len ? input[i] : 0;
        uint32_t octet_b = i + 1 < len ? input[i + 1] : 0;
        uint32_t octet_c = i + 2 < len ? input[i + 2] : 0;

        uint32_t triple = (octet_a << 16) | (octet_b << 8) | octet_c;

        output[out_len++] = base64_chars[(triple >> 18) & 0x3f];
        output[out_len++] = base64_chars[(triple >> 12) & 0x3f];
        output[out_len++] = i + 1 < len ? base64_chars[(triple >> 6) & 0x3f] : '=';
        output[out_len++] = i + 2 < len ? base64_chars[triple & 0x3f] : '=';
    }

    output[out_len] = '\0';
    return out_len;
}

/*
 * base64_decode - Base64文字列をバイナリデータにデコード
 *
 * @input:       デコード対象のBase64文字列
 * @output:      出力先バッファ
 * @output_size: 出力バッファのサイズ
 * @return:      デコードされたバイト数
 */
static size_t base64_decode(const char *input, uint8_t *output, size_t output_size) {
    size_t input_len = strlen(input);
    size_t output_len = 0;
    unsigned int buffer = 0;
    int bits = 0;

    for (size_t i = 0; i < input_len && output_len < output_size; i++) {
        char c = input[i];
        if (c == '=' || c == '\n' || c == '\r') continue;

        const char *p = strchr(base64_chars, c);
        if (!p) continue;

        buffer = (buffer << 6) | (p - base64_chars);
        bits += 6;

        if (bits >= 8) {
            bits -= 8;
            output[output_len++] = (buffer >> bits) & 0xff;
        }
    }
    return output_len;
}

/* ============================================================================
 * Unicode変換
 * ============================================================================
 *
 * 【なぜUTF-16LEが必要か】
 * WindowsはUTF-16LE（Little Endian）を内部文字コードとして使用。
 * NTLM認証でもパスワードやユーザー名はUTF-16LEでエンコードする必要がある。
 *
 * 【LinuxとWindowsの文字コードの違い】
 * - Linux: UTF-8が標準
 * - Windows: UTF-16LE（内部）、Shift_JIS（日本語環境）
 *
 * 【UTF-8からUTF-16LEへの変換ルール】
 * - 1バイト文字（ASCII）: 0x00-0x7F → 1バイト + 0x00
 * - 2バイト文字: 110xxxxx 10xxxxxx → 2バイトLE
 * - 3バイト文字: 1110xxxx 10xxxxxx 10xxxxxx → 2バイトLE
 * ============================================================================ */

/*
 * utf8_to_utf16le - UTF-8文字列をUTF-16LEに変換
 *
 * @utf8:       変換元のUTF-8文字列
 * @utf16:      変換先のUTF-16LEバッファ
 * @utf16_size: 出力バッファのサイズ
 * @return:     変換後のバイト数
 */
static size_t utf8_to_utf16le(const char *utf8, uint8_t *utf16, size_t utf16_size) {
    size_t out_len = 0;
    const uint8_t *p = (const uint8_t *)utf8;

    while (*p && out_len + 2 <= utf16_size) {
        uint32_t codepoint;

        if ((*p & 0x80) == 0) {
            codepoint = *p++;
        } else if ((*p & 0xe0) == 0xc0) {
            codepoint = (*p++ & 0x1f) << 6;
            codepoint |= (*p++ & 0x3f);
        } else if ((*p & 0xf0) == 0xe0) {
            codepoint = (*p++ & 0x0f) << 12;
            codepoint |= (*p++ & 0x3f) << 6;
            codepoint |= (*p++ & 0x3f);
        } else {
            p++;
            continue;
        }

        if (codepoint < 0x10000) {
            utf16[out_len++] = codepoint & 0xff;
            utf16[out_len++] = (codepoint >> 8) & 0xff;
        }
    }

    return out_len;
}

/* ============================================================================
 * NTLM認証実装
 * ============================================================================
 *
 * 【NTLMv2認証の流れ】
 *
 * 1. NT Hashの生成
 *    NT_Hash = MD4(UTF-16LE(Password))
 *
 * 2. NTLMv2 Hashの生成
 *    NTLMv2_Hash = HMAC-MD5(NT_Hash, UTF-16LE(Username.upper() + Domain))
 *
 * 3. Type 1メッセージ（Negotiate）の送信
 *    - 使用可能な認証方式を通知
 *    - サポートするフラグを送信
 *
 * 4. Type 2メッセージ（Challenge）の受信
 *    - サーバーから8バイトのチャレンジを受信
 *    - TargetInfo構造体を受信（オプション）
 *
 * 5. Type 3メッセージ（Authenticate）の送信
 *    - NTProofStr = HMAC-MD5(NTLMv2_Hash, ServerChallenge + Blob)
 *    - NTResponse = NTProofStr + Blob
 *    - Blobにはタイムスタンプ、クライアントチャレンジ等を含む
 *
 * 【セキュリティ考慮】
 * - NTLMv2は、NTLMv1より安全（リプレイ攻撃耐性あり）
 * - タイムスタンプにより時間制限付きの認証が可能
 * - クライアントチャレンジにより、サーバー側のなりすましを防止
 * ============================================================================ */

/*
 * ntlm_hash - NTハッシュを生成（パスワードのMD4ハッシュ）
 *
 * @password: パスワード（UTF-8）
 * @hash:     出力先（16バイト）
 *
 * 計算式: NT_Hash = MD4(UTF-16LE(Password))
 */
static void ntlm_hash(const char *password, uint8_t *hash) {
    uint8_t utf16_pass[512];
    size_t utf16_len = utf8_to_utf16le(password, utf16_pass, sizeof(utf16_pass));
    md4_hash(utf16_pass, utf16_len, hash);
}

/*
 * ntlmv2_hash - NTLMv2ハッシュを生成
 *
 * @password: パスワード（UTF-8）
 * @user:     ユーザー名（UTF-8）
 * @domain:   ドメイン名（UTF-8）
 * @hash:     出力先（16バイト）
 *
 * 計算式: NTLMv2_Hash = HMAC-MD5(NT_Hash, UTF-16LE(Username.upper() + Domain))
 *
 * 注: ユーザー名は大文字に変換されるが、ドメイン名はそのまま使用
 */
static void ntlmv2_hash(const char *password, const char *user, const char *domain,
                        uint8_t *hash) {
    uint8_t nt_hash[16];
    ntlm_hash(password, nt_hash);

    /* ユーザー名を大文字に変換してドメインと結合 */
    char user_domain[512];
    size_t i;
    for (i = 0; user[i]; i++) {
        user_domain[i] = (user[i] >= 'a' && user[i] <= 'z') ? user[i] - 32 : user[i];
    }
    strcpy(user_domain + i, domain);

    uint8_t utf16_ud[512];
    size_t utf16_len = utf8_to_utf16le(user_domain, utf16_ud, sizeof(utf16_ud));

    hmac_md5(nt_hash, 16, utf16_ud, utf16_len, hash);
}

/*
 * ntlm_create_type1 - Type 1メッセージ（Negotiate）を生成
 *
 * @buffer:      出力バッファ
 * @buffer_size: バッファサイズ
 * @return:      生成されたメッセージのサイズ
 *
 * Type 1メッセージ構造:
 * - Signature (8バイト): "NTLMSSP\0"
 * - MessageType (4バイト): 0x00000001
 * - NegotiateFlags (4バイト): サポートする機能フラグ
 * - DomainNameFields (8バイト): ドメイン名（空）
 * - WorkstationFields (8バイト): ワークステーション名（空）
 */
static size_t ntlm_create_type1(uint8_t *buffer, size_t buffer_size) {
    if (buffer_size < 32) return 0;

    memset(buffer, 0, 32);
    memcpy(buffer, NTLM_SIGNATURE, 8);

    uint32_t type = NTLM_TYPE1;
    memcpy(buffer + 8, &type, 4);

    uint32_t flags = NTLMSSP_NEGOTIATE_UNICODE |
                     NTLMSSP_NEGOTIATE_NTLM |
                     NTLMSSP_NEGOTIATE_ALWAYS_SIGN |
                     NTLMSSP_NEGOTIATE_EXTENDED_SESSIONSECURITY |
                     NTLMSSP_REQUEST_TARGET;
    memcpy(buffer + 12, &flags, 4);

    return 32;
}

/*
 * ntlm_parse_type2 - Type 2メッセージ（Challenge）を解析
 *
 * @buffer:          受信したメッセージ
 * @len:             メッセージ長
 * @challenge:       サーバーチャレンジの出力先（8バイト）
 * @flags:           ネゴシエートフラグの出力先
 * @target_info:     TargetInfo構造体の出力先
 * @target_info_len: TargetInfoの長さの出力先
 * @return:          成功時true
 *
 * Type 2メッセージ構造:
 * - Signature (8バイト): "NTLMSSP\0"
 * - MessageType (4バイト): 0x00000002
 * - TargetNameFields (8バイト)
 * - NegotiateFlags (4バイト)
 * - ServerChallenge (8バイト): ランダムなチャレンジ値 ← 重要
 * - Reserved (8バイト)
 * - TargetInfoFields (8バイト)
 * - [TargetInfo]: サーバー情報（オプション）
 */
static bool ntlm_parse_type2(const uint8_t *buffer, size_t len,
                             uint8_t *challenge, uint32_t *flags,
                             uint8_t *target_info, size_t *target_info_len) {
    if (len < 32) return false;
    if (memcmp(buffer, NTLM_SIGNATURE, 8) != 0) return false;

    uint32_t type;
    memcpy(&type, buffer + 8, 4);
    if (type != NTLM_TYPE2) return false;

    memcpy(challenge, buffer + 24, 8);
    memcpy(flags, buffer + 20, 4);

    /* TargetInfo取得 */
    if (len >= 48 && (*flags & NTLMSSP_NEGOTIATE_TARGET_INFO)) {
        uint16_t ti_len, ti_offset;
        memcpy(&ti_len, buffer + 40, 2);
        memcpy(&ti_offset, buffer + 44, 2);

        if (ti_offset + ti_len <= len && ti_len < 1024) {
            memcpy(target_info, buffer + ti_offset, ti_len);
            *target_info_len = ti_len;
        }
    }

    return true;
}

/*
 * ntlm_create_type3 - Type 3メッセージ（Authenticate）を生成
 *
 * @user:            ユーザー名
 * @password:        パスワード
 * @domain:          ドメイン名
 * @challenge:       サーバーから受信したチャレンジ（8バイト）
 * @target_info:     サーバーから受信したTargetInfo
 * @target_info_len: TargetInfoの長さ
 * @buffer:          出力バッファ
 * @buffer_size:     バッファサイズ
 * @return:          生成されたメッセージのサイズ
 *
 * Type 3メッセージ構造:
 * - Signature (8バイト): "NTLMSSP\0"
 * - MessageType (4バイト): 0x00000003
 * - LmChallengeResponseFields (8バイト): LMv2レスポンス（空）
 * - NtChallengeResponseFields (8バイト): NTLMv2レスポンス
 * - DomainNameFields (8バイト)
 * - UserNameFields (8バイト)
 * - WorkstationFields (8バイト)
 * - EncryptedRandomSessionKeyFields (8バイト)
 * - NegotiateFlags (4バイト)
 * - [データ部分]: 各フィールドの実データ
 *
 * NTLMv2レスポンスの構造:
 * - NTProofStr (16バイト): HMAC-MD5(NTLMv2Hash, ServerChallenge + Blob)
 * - Blob: タイムスタンプ、クライアントチャレンジ、TargetInfo等
 */
static size_t ntlm_create_type3(const char *user, const char *password,
                                const char *domain, const uint8_t *challenge,
                                const uint8_t *target_info, size_t target_info_len,
                                uint8_t *buffer, size_t buffer_size) {
    uint8_t ntlmv2_h[16];
    ntlmv2_hash(password, user, domain, ntlmv2_h);

    /* クライアントチャレンジ（ランダム8バイト） */
    uint8_t client_challenge[8];
    FILE *urandom = fopen("/dev/urandom", "r");
    if (urandom) {
        fread(client_challenge, 1, 8, urandom);
        fclose(urandom);
    } else {
        for (int i = 0; i < 8; i++) client_challenge[i] = rand() & 0xff;
    }

    /* タイムスタンプ（Windows FILETIME形式） */
    uint64_t timestamp;
    struct timeval tv;
    gettimeofday(&tv, NULL);
    timestamp = ((uint64_t)tv.tv_sec + 11644473600ULL) * 10000000ULL + tv.tv_usec * 10;

    /* NTLMv2 blob */
    size_t blob_len = 28 + target_info_len + 4;
    uint8_t *blob = calloc(blob_len, 1);

    blob[0] = 0x01; blob[1] = 0x01;  /* Blob signature */
    memcpy(blob + 8, &timestamp, 8);
    memcpy(blob + 16, client_challenge, 8);
    memcpy(blob + 28, target_info, target_info_len);

    /* NTProofStr = HMAC-MD5(NTLMv2Hash, ServerChallenge + Blob) */
    size_t concat_len = 8 + blob_len;
    uint8_t *concat = malloc(concat_len);
    memcpy(concat, challenge, 8);
    memcpy(concat + 8, blob, blob_len);

    uint8_t nt_proof_str[16];
    hmac_md5(ntlmv2_h, 16, concat, concat_len, nt_proof_str);
    free(concat);

    /* NTLMv2 Response = NTProofStr + Blob */
    size_t nt_response_len = 16 + blob_len;
    uint8_t *nt_response = malloc(nt_response_len);
    memcpy(nt_response, nt_proof_str, 16);
    memcpy(nt_response + 16, blob, blob_len);
    free(blob);

    /* セッションキー */
    uint8_t session_key[16];
    hmac_md5(ntlmv2_h, 16, nt_proof_str, 16, session_key);

    /* UTF-16LEに変換 */
    uint8_t domain_utf16[256], user_utf16[256];
    size_t domain_len = utf8_to_utf16le(domain, domain_utf16, sizeof(domain_utf16));
    size_t user_len = utf8_to_utf16le(user, user_utf16, sizeof(user_utf16));

    /* Type 3メッセージ構築 */
    size_t offset = 88;  /* ヘッダサイズ */

    memset(buffer, 0, buffer_size);
    memcpy(buffer, NTLM_SIGNATURE, 8);

    uint32_t type = NTLM_TYPE3;
    memcpy(buffer + 8, &type, 4);

    /* LM Response（空） */
    uint16_t lm_len = 0;
    memcpy(buffer + 12, &lm_len, 2);
    memcpy(buffer + 14, &lm_len, 2);

    /* NT Response */
    uint16_t nt_len = nt_response_len;
    uint32_t nt_offset = offset;
    memcpy(buffer + 20, &nt_len, 2);
    memcpy(buffer + 22, &nt_len, 2);
    memcpy(buffer + 24, &nt_offset, 4);
    memcpy(buffer + offset, nt_response, nt_response_len);
    offset += nt_response_len;
    free(nt_response);

    /* Domain */
    uint16_t dom_len = domain_len;
    uint32_t dom_offset = offset;
    memcpy(buffer + 28, &dom_len, 2);
    memcpy(buffer + 30, &dom_len, 2);
    memcpy(buffer + 32, &dom_offset, 4);
    memcpy(buffer + offset, domain_utf16, domain_len);
    offset += domain_len;

    /* User */
    uint16_t usr_len = user_len;
    uint32_t usr_offset = offset;
    memcpy(buffer + 36, &usr_len, 2);
    memcpy(buffer + 38, &usr_len, 2);
    memcpy(buffer + 40, &usr_offset, 4);
    memcpy(buffer + offset, user_utf16, user_len);
    offset += user_len;

    /* Workstation（空） */
    uint32_t ws_offset = offset;
    memcpy(buffer + 48, &ws_offset, 4);

    /* Encrypted Session Key（空） */
    uint32_t sk_offset = offset;
    memcpy(buffer + 52, &sk_offset, 4);

    /* Flags */
    uint32_t flags = NTLMSSP_NEGOTIATE_UNICODE |
                     NTLMSSP_NEGOTIATE_NTLM |
                     NTLMSSP_NEGOTIATE_ALWAYS_SIGN |
                     NTLMSSP_NEGOTIATE_EXTENDED_SESSIONSECURITY;
    memcpy(buffer + 60, &flags, 4);

    return offset;
}

/* ============================================================================
 * HTTP通信
 * ============================================================================
 *
 * 【WinRMのトランスポート】
 * WinRMはHTTP/HTTPSをトランスポートとして使用。
 * - ポート5985: HTTP（暗号化なし、ただしNTLM認証あり）
 * - ポート5986: HTTPS（SSL/TLS暗号化）
 *
 * 【HTTPリクエスト形式】
 * POST /wsman HTTP/1.1
 * Host: server:5985
 * Authorization: NTLM <Base64エンコードされたNTLMメッセージ>
 * Content-Type: application/soap+xml;charset=UTF-8
 * Content-Length: <本文の長さ>
 *
 * <SOAP XMLエンベロープ>
 *
 * 【NTLM認証のHTTPでの流れ】
 * 1. Type 1メッセージを含むリクエスト送信 → 401応答受信
 * 2. 401のWWW-AuthenticateヘッダからType 2メッセージ取得
 * 3. Type 3メッセージを含むリクエスト送信 → 200応答受信（認証成功）
 * ============================================================================ */

/*
 * connect_to_host - サーバーへのTCPソケット接続を確立
 *
 * @host: ホスト名またはIPアドレス
 * @port: ポート番号
 * @return: ソケットファイルディスクリプタ（エラー時は-1）
 */
static int connect_to_host(const char *host, int port) {
    struct hostent *he;
    struct sockaddr_in server_addr;
    int sock;

    he = gethostbyname(host);
    if (!he) {
        log_error("ホスト名の解決に失敗しました");
        return -1;
    }

    sock = socket(AF_INET, SOCK_STREAM, 0);
    if (sock < 0) {
        log_error("ソケット作成に失敗しました");
        return -1;
    }

    /* タイムアウト設定 */
    struct timeval tv;
    tv.tv_sec = TIMEOUT;
    tv.tv_usec = 0;
    setsockopt(sock, SOL_SOCKET, SO_RCVTIMEO, &tv, sizeof(tv));
    setsockopt(sock, SOL_SOCKET, SO_SNDTIMEO, &tv, sizeof(tv));

    memset(&server_addr, 0, sizeof(server_addr));
    server_addr.sin_family = AF_INET;
    server_addr.sin_port = htons(port);
    memcpy(&server_addr.sin_addr, he->h_addr, he->h_length);

    if (connect(sock, (struct sockaddr *)&server_addr, sizeof(server_addr)) < 0) {
        char msg[256];
        snprintf(msg, sizeof(msg), "接続に失敗しました: %s:%d", host, port);
        log_error(msg);
        close(sock);
        return -1;
    }

    return sock;
}

/*
 * recv_http_response - HTTPレスポンスを受信・解析
 *
 * @sock:        ソケットファイルディスクリプタ
 * @buffer:      受信データ格納バッファ
 * @buffer_size: バッファサイズ
 * @http_code:   HTTPステータスコード出力先
 * @auth_header: WWW-AuthenticateヘッダのNTLM部分出力先
 * @return:      受信したバイト数
 *
 * 処理内容:
 * 1. ヘッダ終端（\r\n\r\n）まで受信
 * 2. HTTPステータスコードを抽出
 * 3. WWW-AuthenticateヘッダからNTLMチャレンジを抽出
 * 4. Content-Lengthに基づいて本文を受信
 */
static size_t recv_http_response(int sock, char *buffer, size_t buffer_size,
                                 int *http_code, char *auth_header) {
    size_t total = 0;
    size_t header_end = 0;
    bool headers_done = false;
    size_t content_length = 0;

    *http_code = 0;
    auth_header[0] = '\0';

    while (total < buffer_size - 1) {
        ssize_t n = recv(sock, buffer + total, buffer_size - 1 - total, 0);
        if (n <= 0) break;
        total += n;
        buffer[total] = '\0';

        if (!headers_done) {
            char *end = strstr(buffer, "\r\n\r\n");
            if (end) {
                headers_done = true;
                header_end = end - buffer + 4;

                /* HTTPステータスコード取得 */
                if (sscanf(buffer, "HTTP/%*s %d", http_code) != 1) {
                    *http_code = 0;
                }

                /* WWW-Authenticateヘッダ取得 */
                char *auth = strcasestr(buffer, "WWW-Authenticate: NTLM ");
                if (auth) {
                    auth += strlen("WWW-Authenticate: NTLM ");
                    char *end_auth = strstr(auth, "\r\n");
                    if (end_auth) {
                        size_t len = end_auth - auth;
                        if (len < 1024) {
                            strncpy(auth_header, auth, len);
                            auth_header[len] = '\0';
                        }
                    }
                }

                /* Content-Length取得 */
                char *cl = strcasestr(buffer, "Content-Length:");
                if (cl) {
                    sscanf(cl + 15, "%zu", &content_length);
                }

                /* 本文を十分受信したか確認 */
                if (content_length > 0 && total >= header_end + content_length) {
                    break;
                }
            }
        } else {
            if (content_length > 0 && total >= header_end + content_length) {
                break;
            }
        }
    }

    return total;
}

/*
 * send_http_with_ntlm - NTLM認証を行いHTTP POSTリクエストを送信
 *
 * @host:          接続先ホスト
 * @port:          接続先ポート
 * @body:          リクエスト本文（SOAP XML）
 * @response:      レスポンス本文格納バッファ
 * @response_size: バッファサイズ
 * @return:        成功時true
 *
 * 処理フロー:
 * 1. Type 1メッセージを送信（認証開始）
 * 2. 401応答からType 2メッセージ（チャレンジ）を取得
 * 3. Type 2をデコード・解析
 * 4. Type 3メッセージ（認証応答）を生成・送信
 * 5. 200応答で認証成功、レスポンス本文を返却
 */
static bool send_http_with_ntlm(const char *host, int port, const char *body,
                                char *response, size_t response_size) {
    int sock;
    char request[MAX_BUFFER_SIZE];
    char recv_buffer[MAX_BUFFER_SIZE];
    int http_code;
    char auth_header[1024];

    /* Step 1: Type 1メッセージを送信 */
    sock = connect_to_host(host, port);
    if (sock < 0) return false;

    uint8_t type1[64];
    size_t type1_len = ntlm_create_type1(type1, sizeof(type1));

    char type1_b64[256];
    base64_encode(type1, type1_len, type1_b64);

    snprintf(request, sizeof(request),
             "POST /wsman HTTP/1.1\r\n"
             "Host: %s:%d\r\n"
             "Authorization: NTLM %s\r\n"
             "Content-Type: application/soap+xml;charset=UTF-8\r\n"
             "Content-Length: %zu\r\n"
             "Connection: keep-alive\r\n"
             "\r\n%s",
             host, port, type1_b64, strlen(body), body);

    if (DEBUG) {
        log_info("Type 1メッセージ送信中...");
    }

    send(sock, request, strlen(request), 0);
    recv_http_response(sock, recv_buffer, sizeof(recv_buffer), &http_code, auth_header);
    close(sock);

    if (http_code != 401 || auth_header[0] == '\0') {
        log_error("NTLM認証のType 2応答を受信できませんでした");
        return false;
    }

    /* Step 2: Type 2メッセージを解析 */
    uint8_t type2[2048];
    size_t type2_len = base64_decode(auth_header, type2, sizeof(type2));

    uint8_t challenge[8];
    uint32_t flags;
    uint8_t target_info[1024];
    size_t target_info_len = 0;

    if (!ntlm_parse_type2(type2, type2_len, challenge, &flags, target_info, &target_info_len)) {
        log_error("Type 2メッセージの解析に失敗しました");
        return false;
    }

    if (DEBUG) {
        log_info("Type 2メッセージ受信・解析成功");
    }

    /* Step 3: Type 3メッセージを送信 */
    sock = connect_to_host(host, port);
    if (sock < 0) return false;

    uint8_t type3[2048];
    size_t type3_len = ntlm_create_type3(g_user, g_pass, g_domain,
                                          challenge, target_info, target_info_len,
                                          type3, sizeof(type3));

    char type3_b64[4096];
    base64_encode(type3, type3_len, type3_b64);

    snprintf(request, sizeof(request),
             "POST /wsman HTTP/1.1\r\n"
             "Host: %s:%d\r\n"
             "Authorization: NTLM %s\r\n"
             "Content-Type: application/soap+xml;charset=UTF-8\r\n"
             "Content-Length: %zu\r\n"
             "Connection: close\r\n"
             "\r\n%s",
             host, port, type3_b64, strlen(body), body);

    if (DEBUG) {
        log_info("Type 3メッセージ送信中...");
    }

    send(sock, request, strlen(request), 0);
    size_t recv_len = recv_http_response(sock, recv_buffer, sizeof(recv_buffer),
                                          &http_code, auth_header);
    close(sock);

    if (DEBUG) {
        char msg[64];
        snprintf(msg, sizeof(msg), "HTTPステータスコード: %d", http_code);
        log_info(msg);
    }

    if (http_code == 401) {
        log_error("認証に失敗しました (HTTP 401)");
        log_error("ユーザー名とパスワードを確認してください");
        return false;
    } else if (http_code == 500) {
        log_error("サーバー内部エラーが発生しました (HTTP 500)");
        return false;
    } else if (http_code != 200) {
        char msg[64];
        snprintf(msg, sizeof(msg), "予期しないHTTPステータスコード: %d", http_code);
        log_warn(msg);
    }

    /* レスポンス本文を抽出 */
    char *body_start = strstr(recv_buffer, "\r\n\r\n");
    if (body_start) {
        body_start += 4;
        strncpy(response, body_start, response_size - 1);
        response[response_size - 1] = '\0';
    }

    if (DEBUG) {
        log_info("受信XML:");
        fprintf(stderr, "%s\n", response);
    }

    return true;
}

/* ============================================================================
 * ユーティリティ関数
 * ============================================================================ */

/*
 * generate_uuid - UUIDを生成
 *
 * @uuid: 出力バッファ
 * @size: バッファサイズ
 *
 * WinRMのSOAPメッセージには各リクエストにユニークなMessageIDが必要。
 * Linuxの /proc/sys/kernel/random/uuid を使用（利用可能な場合）。
 */
static void generate_uuid(char *uuid, size_t size) {
    FILE *fp = fopen("/proc/sys/kernel/random/uuid", "r");
    if (fp) {
        if (fgets(uuid, size, fp)) {
            char *newline = strchr(uuid, '\n');
            if (newline) *newline = '\0';
        }
        fclose(fp);
    } else {
        snprintf(uuid, size, "%08lx-%04x-4%03x-%04x-%012lx",
                 (unsigned long)time(NULL), rand() & 0xffff,
                 rand() & 0x0fff, rand() & 0xffff, (unsigned long)time(NULL));
    }
}

/*
 * xml_escape - XML特殊文字をエスケープ
 *
 * @src:      入力文字列
 * @dst:      出力バッファ
 * @dst_size: バッファサイズ
 *
 * SOAPメッセージに含めるコマンド文字列内の特殊文字を
 * XMLエンティティに変換（例: & → &amp;）
 */
static void xml_escape(const char *src, char *dst, size_t dst_size) {
    size_t j = 0;
    for (size_t i = 0; src[i] && j < dst_size - 6; i++) {
        switch (src[i]) {
            case '&':  j += snprintf(dst + j, dst_size - j, "&amp;"); break;
            case '<':  j += snprintf(dst + j, dst_size - j, "&lt;"); break;
            case '>':  j += snprintf(dst + j, dst_size - j, "&gt;"); break;
            case '"':  j += snprintf(dst + j, dst_size - j, "&quot;"); break;
            case '\'': j += snprintf(dst + j, dst_size - j, "&apos;"); break;
            default:   dst[j++] = src[i]; break;
        }
    }
    dst[j] = '\0';
}

/*
 * str_replace - 文字列内のすべての出現箇所を置換
 *
 * @str:     置換対象の文字列（直接変更される）
 * @old:     検索文字列
 * @new_str: 置換文字列
 *
 * 【重要】whileループを使用しているため、{ENV}が複数箇所にあっても
 * すべて置換される。
 *
 * 例: "C:\\{ENV}\\sub\\{ENV}\\test.bat" → "C:\\TST1T\\sub\\TST1T\\test.bat"
 */
static void str_replace(char *str, const char *old, const char *new_str) {
    char buffer[MAX_BUFFER_SIZE];
    char *pos;

    while ((pos = strstr(str, old)) != NULL) {
        size_t prefix_len = pos - str;
        strncpy(buffer, str, prefix_len);
        buffer[prefix_len] = '\0';
        strcat(buffer, new_str);
        strcat(buffer, pos + strlen(old));
        strcpy(str, buffer);
    }
}

/*
 * extract_xml_value - XMLから指定タグの値を抽出
 *
 * @xml:        XML文字列
 * @tag:        抽出するタグ名（例: "rsp:ShellId"）
 * @value:      値の出力バッファ
 * @value_size: バッファサイズ
 * @return:     成功時true
 *
 * 単純な文字列検索で実装（完全なXMLパーサーは使用しない）。
 * WinRMレスポンスからShellIdやCommandId等を抽出するために使用。
 */
static bool extract_xml_value(const char *xml, const char *tag, char *value, size_t value_size) {
    char open_tag[128], close_tag[128];
    snprintf(open_tag, sizeof(open_tag), "<%s>", tag);
    snprintf(close_tag, sizeof(close_tag), "</%s>", tag);

    const char *start = strstr(xml, open_tag);
    if (!start) return false;

    start += strlen(open_tag);
    const char *end = strstr(start, close_tag);
    if (!end) return false;

    size_t len = end - start;
    if (len >= value_size) len = value_size - 1;

    strncpy(value, start, len);
    value[len] = '\0';
    return true;
}

/* ============================================================================
 * WinRM操作
 * ============================================================================
 *
 * 【WinRM (Windows Remote Management) とは】
 * Microsoftが開発したリモート管理プロトコル。
 * WS-Management（Web Services for Management）仕様に基づく。
 *
 * 【WinRMセッションの流れ】
 * 1. Create: リモートシェル（cmd.exe）を作成
 *    → ShellIdを取得
 *
 * 2. Command: シェル上でコマンドを実行
 *    → CommandIdを取得
 *
 * 3. Receive: コマンドの出力（stdout/stderr）を取得
 *    → 出力はBase64エンコードされて返される
 *    → CommandState/Doneになるまでポーリング
 *
 * 4. Delete: シェルを削除（リソース解放）
 *
 * 【SOAPメッセージ形式】
 * <?xml version="1.0" encoding="UTF-8"?>
 * <s:Envelope xmlns:s="..." xmlns:a="..." ...>
 *   <s:Header>
 *     <a:Action>...</a:Action>
 *     <w:ResourceURI>...</w:ResourceURI>
 *     ...
 *   </s:Header>
 *   <s:Body>
 *     ... 操作固有のXML ...
 *   </s:Body>
 * </s:Envelope>
 * ============================================================================ */

/*
 * send_soap_request - SOAPリクエストを送信
 *
 * @soap_envelope: 送信するSOAP XMLエンベロープ
 * @response:      レスポンス格納バッファ
 * @response_size: バッファサイズ
 * @return:        成功時true
 */
static bool send_soap_request(const char *soap_envelope, char *response, size_t response_size) {
    if (DEBUG) {
        log_info("送信XML:");
        fprintf(stderr, "%s\n\n", soap_envelope);
        char msg[256];
        snprintf(msg, sizeof(msg), "接続先: http://%s:%d/wsman", g_host, g_port);
        log_info(msg);
        snprintf(msg, sizeof(msg), "ユーザー: %s", g_user);
        log_info(msg);
    }

    return send_http_with_ntlm(g_host, g_port, soap_envelope, response, response_size);
}

/*
 * create_shell - リモートシェル（cmd.exe）を作成
 *
 * @shell_id:      ShellIdの出力バッファ
 * @shell_id_size: バッファサイズ
 * @return:        成功時true
 *
 * WS-Transfer Createアクションを使用してリモートシェルを作成。
 * 成功すると、後続のコマンド実行に使用するShellIdが返される。
 */
static bool create_shell(char *shell_id, size_t shell_id_size) {
    char url[MAX_URL_SIZE];
    char uuid[MAX_UUID_SIZE];
    char envelope[MAX_ENVELOPE_SIZE];
    char response[MAX_BUFFER_SIZE];

    snprintf(url, sizeof(url), "http://%s:%d/wsman", g_host, g_port);
    generate_uuid(uuid, sizeof(uuid));

    snprintf(envelope, sizeof(envelope),
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
        "<s:Envelope xmlns:s=\"http://www.w3.org/2003/05/soap-envelope\"\n"
        "            xmlns:a=\"http://schemas.xmlsoap.org/ws/2004/08/addressing\"\n"
        "            xmlns:w=\"http://schemas.dmtf.org/wbem/wsman/1/wsman.xsd\"\n"
        "            xmlns:rsp=\"http://schemas.microsoft.com/wbem/wsman/1/windows/shell\">\n"
        "  <s:Header>\n"
        "    <a:To>%s</a:To>\n"
        "    <a:ReplyTo>\n"
        "      <a:Address s:mustUnderstand=\"true\">http://schemas.xmlsoap.org/ws/2004/08/addressing/role/anonymous</a:Address>\n"
        "    </a:ReplyTo>\n"
        "    <a:Action s:mustUnderstand=\"true\">http://schemas.xmlsoap.org/ws/2004/09/transfer/Create</a:Action>\n"
        "    <w:MaxEnvelopeSize s:mustUnderstand=\"true\">153600</w:MaxEnvelopeSize>\n"
        "    <a:MessageID>uuid:%s</a:MessageID>\n"
        "    <w:Locale xml:lang=\"ja-JP\" s:mustUnderstand=\"false\"/>\n"
        "    <w:OperationTimeout>PT%dS</w:OperationTimeout>\n"
        "    <w:ResourceURI s:mustUnderstand=\"true\">http://schemas.microsoft.com/wbem/wsman/1/windows/shell/cmd</w:ResourceURI>\n"
        "    <w:OptionSet>\n"
        "      <w:Option Name=\"WINRS_NOPROFILE\">FALSE</w:Option>\n"
        "      <w:Option Name=\"WINRS_CODEPAGE\">65001</w:Option>\n"
        "    </w:OptionSet>\n"
        "  </s:Header>\n"
        "  <s:Body>\n"
        "    <rsp:Shell>\n"
        "      <rsp:InputStreams>stdin</rsp:InputStreams>\n"
        "      <rsp:OutputStreams>stdout stderr</rsp:OutputStreams>\n"
        "    </rsp:Shell>\n"
        "  </s:Body>\n"
        "</s:Envelope>",
        url, uuid, TIMEOUT);

    log_info("シェル作成中...");

    if (!send_soap_request(envelope, response, sizeof(response))) {
        log_error("シェル作成に失敗しました");
        return false;
    }

    if (!extract_xml_value(response, "rsp:ShellId", shell_id, shell_id_size)) {
        log_error("ShellIDの取得に失敗しました");
        return false;
    }

    char msg[256];
    snprintf(msg, sizeof(msg), "シェル作成成功: %s", shell_id);
    log_success(msg);

    return true;
}

/*
 * run_command - シェル上でコマンドを実行
 *
 * @shell_id:        対象のShellId
 * @command:         実行するコマンド文字列
 * @command_id:      CommandIdの出力バッファ
 * @command_id_size: バッファサイズ
 * @return:          成功時true
 *
 * WinRS Commandアクションを使用してコマンドを実行。
 * 成功すると、出力取得に使用するCommandIdが返される。
 */
static bool run_command(const char *shell_id, const char *command,
                        char *command_id, size_t command_id_size) {
    char url[MAX_URL_SIZE];
    char uuid[MAX_UUID_SIZE];
    char envelope[MAX_ENVELOPE_SIZE];
    char command_escaped[1024];
    char response[MAX_BUFFER_SIZE];

    snprintf(url, sizeof(url), "http://%s:%d/wsman", g_host, g_port);
    generate_uuid(uuid, sizeof(uuid));
    xml_escape(command, command_escaped, sizeof(command_escaped));

    snprintf(envelope, sizeof(envelope),
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
        "<s:Envelope xmlns:s=\"http://www.w3.org/2003/05/soap-envelope\"\n"
        "            xmlns:a=\"http://schemas.xmlsoap.org/ws/2004/08/addressing\"\n"
        "            xmlns:w=\"http://schemas.dmtf.org/wbem/wsman/1/wsman.xsd\"\n"
        "            xmlns:rsp=\"http://schemas.microsoft.com/wbem/wsman/1/windows/shell\">\n"
        "  <s:Header>\n"
        "    <a:To>%s</a:To>\n"
        "    <a:ReplyTo>\n"
        "      <a:Address s:mustUnderstand=\"true\">http://schemas.xmlsoap.org/ws/2004/08/addressing/role/anonymous</a:Address>\n"
        "    </a:ReplyTo>\n"
        "    <a:Action s:mustUnderstand=\"true\">http://schemas.microsoft.com/wbem/wsman/1/windows/shell/Command</a:Action>\n"
        "    <w:MaxEnvelopeSize s:mustUnderstand=\"true\">153600</w:MaxEnvelopeSize>\n"
        "    <a:MessageID>uuid:%s</a:MessageID>\n"
        "    <w:Locale xml:lang=\"ja-JP\" s:mustUnderstand=\"false\"/>\n"
        "    <w:OperationTimeout>PT%dS</w:OperationTimeout>\n"
        "    <w:ResourceURI s:mustUnderstand=\"true\">http://schemas.microsoft.com/wbem/wsman/1/windows/shell/cmd</w:ResourceURI>\n"
        "    <w:SelectorSet>\n"
        "      <w:Selector Name=\"ShellId\">%s</w:Selector>\n"
        "    </w:SelectorSet>\n"
        "  </s:Header>\n"
        "  <s:Body>\n"
        "    <rsp:CommandLine>\n"
        "      <rsp:Command>%s</rsp:Command>\n"
        "    </rsp:CommandLine>\n"
        "  </s:Body>\n"
        "</s:Envelope>",
        url, uuid, TIMEOUT, shell_id, command_escaped);

    log_info("コマンド実行中...");

    if (!send_soap_request(envelope, response, sizeof(response))) {
        log_error("コマンド実行に失敗しました");
        return false;
    }

    if (!extract_xml_value(response, "rsp:CommandId", command_id, command_id_size)) {
        log_error("CommandIDの取得に失敗しました");
        return false;
    }

    char msg[256];
    snprintf(msg, sizeof(msg), "コマンド実行開始: %s", command_id);
    log_success(msg);

    return true;
}

/*
 * get_command_output - コマンドの出力を取得
 *
 * @shell_id:    対象のShellId
 * @command_id:  対象のCommandId
 * @stdout_buf:  標準出力の格納バッファ
 * @stdout_size: 標準出力バッファサイズ
 * @stderr_buf:  標準エラー出力の格納バッファ
 * @stderr_size: 標準エラーバッファサイズ
 * @exit_code:   終了コードの出力先
 * @return:      成功時true
 *
 * WinRS Receiveアクションを使用して出力を取得。
 * CommandState/Doneになるまでポーリングを繰り返す。
 * 出力はBase64エンコードされているためデコードが必要。
 */
static bool get_command_output(const char *shell_id, const char *command_id,
                               char *stdout_buf, size_t stdout_size,
                               char *stderr_buf, size_t stderr_size,
                               int *exit_code) {
    char url[MAX_URL_SIZE];
    char envelope[MAX_ENVELOPE_SIZE];
    char response[MAX_BUFFER_SIZE];
    bool command_done = false;
    int max_attempts = TIMEOUT * 2;

    snprintf(url, sizeof(url), "http://%s:%d/wsman", g_host, g_port);

    stdout_buf[0] = '\0';
    stderr_buf[0] = '\0';
    *exit_code = 0;

    char msg[128];
    snprintf(msg, sizeof(msg), "コマンド出力取得中...（最大%d秒待機）", TIMEOUT);
    log_info(msg);

    for (int attempt = 0; attempt < max_attempts && !command_done; attempt++) {
        char uuid[MAX_UUID_SIZE];
        generate_uuid(uuid, sizeof(uuid));

        snprintf(envelope, sizeof(envelope),
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
            "<s:Envelope xmlns:s=\"http://www.w3.org/2003/05/soap-envelope\"\n"
            "            xmlns:a=\"http://schemas.xmlsoap.org/ws/2004/08/addressing\"\n"
            "            xmlns:w=\"http://schemas.dmtf.org/wbem/wsman/1/wsman.xsd\"\n"
            "            xmlns:rsp=\"http://schemas.microsoft.com/wbem/wsman/1/windows/shell\">\n"
            "  <s:Header>\n"
            "    <a:To>%s</a:To>\n"
            "    <a:ReplyTo>\n"
            "      <a:Address s:mustUnderstand=\"true\">http://schemas.xmlsoap.org/ws/2004/08/addressing/role/anonymous</a:Address>\n"
            "    </a:ReplyTo>\n"
            "    <a:Action s:mustUnderstand=\"true\">http://schemas.microsoft.com/wbem/wsman/1/windows/shell/Receive</a:Action>\n"
            "    <w:MaxEnvelopeSize s:mustUnderstand=\"true\">153600</w:MaxEnvelopeSize>\n"
            "    <a:MessageID>uuid:%s</a:MessageID>\n"
            "    <w:Locale xml:lang=\"ja-JP\" s:mustUnderstand=\"false\"/>\n"
            "    <w:OperationTimeout>PT%dS</w:OperationTimeout>\n"
            "    <w:ResourceURI s:mustUnderstand=\"true\">http://schemas.microsoft.com/wbem/wsman/1/windows/shell/cmd</w:ResourceURI>\n"
            "    <w:SelectorSet>\n"
            "      <w:Selector Name=\"ShellId\">%s</w:Selector>\n"
            "    </w:SelectorSet>\n"
            "  </s:Header>\n"
            "  <s:Body>\n"
            "    <rsp:Receive>\n"
            "      <rsp:DesiredStream CommandId=\"%s\">stdout stderr</rsp:DesiredStream>\n"
            "    </rsp:Receive>\n"
            "  </s:Body>\n"
            "</s:Envelope>",
            url, uuid, TIMEOUT, shell_id, command_id);

        if (!send_soap_request(envelope, response, sizeof(response))) {
            log_error("出力取得に失敗しました");
            return false;
        }

        /* stdout抽出 */
        char *stdout_start = strstr(response, "<rsp:Stream Name=\"stdout\">");
        if (stdout_start) {
            stdout_start += strlen("<rsp:Stream Name=\"stdout\">");
            char *stdout_end = strstr(stdout_start, "</rsp:Stream>");
            if (stdout_end) {
                size_t b64_len = stdout_end - stdout_start;
                char *b64_buf = malloc(b64_len + 1);
                strncpy(b64_buf, stdout_start, b64_len);
                b64_buf[b64_len] = '\0';

                uint8_t decoded[MAX_BUFFER_SIZE];
                size_t decoded_len = base64_decode(b64_buf, decoded, sizeof(decoded));
                decoded[decoded_len] = '\0';
                strncat(stdout_buf, (char *)decoded, stdout_size - strlen(stdout_buf) - 1);
                free(b64_buf);
            }
        }

        /* stderr抽出 */
        char *stderr_start = strstr(response, "<rsp:Stream Name=\"stderr\">");
        if (stderr_start) {
            stderr_start += strlen("<rsp:Stream Name=\"stderr\">");
            char *stderr_end = strstr(stderr_start, "</rsp:Stream>");
            if (stderr_end) {
                size_t b64_len = stderr_end - stderr_start;
                char *b64_buf = malloc(b64_len + 1);
                strncpy(b64_buf, stderr_start, b64_len);
                b64_buf[b64_len] = '\0';

                uint8_t decoded[MAX_BUFFER_SIZE];
                size_t decoded_len = base64_decode(b64_buf, decoded, sizeof(decoded));
                decoded[decoded_len] = '\0';
                strncat(stderr_buf, (char *)decoded, stderr_size - strlen(stderr_buf) - 1);
                free(b64_buf);
            }
        }

        /* コマンド完了チェック */
        if (strstr(response, "CommandState/Done")) {
            command_done = true;
            char exit_code_str[16];
            if (extract_xml_value(response, "rsp:ExitCode", exit_code_str, sizeof(exit_code_str))) {
                *exit_code = atoi(exit_code_str);
            }
        }

        if (!command_done) {
            usleep(500000); /* 0.5秒 */
        }
    }

    if (!command_done) {
        log_warn("コマンド完了待機がタイムアウトしました");
    }

    snprintf(msg, sizeof(msg), "コマンド完了 (終了コード: %d)", *exit_code);
    log_success(msg);

    return true;
}

/*
 * delete_shell - リモートシェルを削除
 *
 * @shell_id: 削除対象のShellId
 *
 * WS-Transfer Deleteアクションを使用してシェルを削除。
 * リソース解放のため、コマンド完了後は必ず呼び出すこと。
 */
static void delete_shell(const char *shell_id) {
    char url[MAX_URL_SIZE];
    char uuid[MAX_UUID_SIZE];
    char envelope[MAX_ENVELOPE_SIZE];
    char response[MAX_BUFFER_SIZE];

    snprintf(url, sizeof(url), "http://%s:%d/wsman", g_host, g_port);
    generate_uuid(uuid, sizeof(uuid));

    snprintf(envelope, sizeof(envelope),
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
        "<s:Envelope xmlns:s=\"http://www.w3.org/2003/05/soap-envelope\"\n"
        "            xmlns:a=\"http://schemas.xmlsoap.org/ws/2004/08/addressing\"\n"
        "            xmlns:w=\"http://schemas.dmtf.org/wbem/wsman/1/wsman.xsd\">\n"
        "  <s:Header>\n"
        "    <a:To>%s</a:To>\n"
        "    <a:ReplyTo>\n"
        "      <a:Address s:mustUnderstand=\"true\">http://schemas.xmlsoap.org/ws/2004/08/addressing/role/anonymous</a:Address>\n"
        "    </a:ReplyTo>\n"
        "    <a:Action s:mustUnderstand=\"true\">http://schemas.xmlsoap.org/ws/2004/09/transfer/Delete</a:Action>\n"
        "    <w:MaxEnvelopeSize s:mustUnderstand=\"true\">153600</w:MaxEnvelopeSize>\n"
        "    <a:MessageID>uuid:%s</a:MessageID>\n"
        "    <w:Locale xml:lang=\"ja-JP\" s:mustUnderstand=\"false\"/>\n"
        "    <w:OperationTimeout>PT%dS</w:OperationTimeout>\n"
        "    <w:ResourceURI s:mustUnderstand=\"true\">http://schemas.microsoft.com/wbem/wsman/1/windows/shell/cmd</w:ResourceURI>\n"
        "    <w:SelectorSet>\n"
        "      <w:Selector Name=\"ShellId\">%s</w:Selector>\n"
        "    </w:SelectorSet>\n"
        "  </s:Header>\n"
        "  <s:Body/>\n"
        "</s:Envelope>",
        url, uuid, TIMEOUT, shell_id);

    log_info("シェル削除中...");
    send_soap_request(envelope, response, sizeof(response));
    log_success("シェル削除完了");
}

/* ============================================================================
 * メイン処理
 * ============================================================================ */

/*
 * load_config - 設定を読み込み
 *
 * デフォルト値を設定し、環境変数があれば上書き。
 * 環境変数: WINRM_HOST, WINRM_USER, WINRM_PASS, WINRM_DOMAIN, WINRM_PORT
 */
static void load_config(void) {
    const char *env;

    env = getenv("WINRM_HOST");
    strncpy(g_host, env ? env : DEFAULT_HOST, sizeof(g_host) - 1);

    env = getenv("WINRM_USER");
    strncpy(g_user, env ? env : DEFAULT_USER, sizeof(g_user) - 1);

    env = getenv("WINRM_PASS");
    strncpy(g_pass, env ? env : DEFAULT_PASS, sizeof(g_pass) - 1);

    env = getenv("WINRM_DOMAIN");
    strncpy(g_domain, env ? env : DEFAULT_DOMAIN, sizeof(g_domain) - 1);

    env = getenv("WINRM_PORT");
    g_port = env ? atoi(env) : DEFAULT_PORT;

    env = getenv("BATCH_FILE_PATH");
    strncpy(g_batch_path, env ? env : DEFAULT_BATCH_PATH, sizeof(g_batch_path) - 1);
}

/*
 * print_help - 使い方を表示
 *
 * @prog_name: プログラム名（argv[0]）
 */
static void print_help(const char *prog_name) {
    printf("使い方: %s ENV\n\n", prog_name);
    printf("引数:\n");
    printf("  ENV    環境名 (");
    for (int i = 0; ENVIRONMENTS[i]; i++) {
        if (i > 0) printf(", ");
        printf("%s", ENVIRONMENTS[i]);
    }
    printf(")\n\n");
    printf("例:\n");
    for (int i = 0; ENVIRONMENTS[i] && i < 2; i++) {
        printf("  %s %s\n", prog_name, ENVIRONMENTS[i]);
    }
    printf("\n環境変数で設定を上書き可能:\n");
    printf("  WINRM_HOST, WINRM_PORT, WINRM_USER, WINRM_PASS, WINRM_DOMAIN\n");
}

/*
 * main - メインエントリーポイント
 *
 * 処理フロー:
 * 1. 引数チェック（環境名の指定が必須）
 * 2. 設定読み込み（デフォルト値 + 環境変数）
 * 3. 環境名の有効性チェック
 * 4. バッチファイルパスの{ENV}プレースホルダ置換
 * 5. WinRM接続・コマンド実行
 * 6. 結果表示
 *
 * @argc: 引数の数
 * @argv: 引数配列
 * @return: 終了コード（成功時0、エラー時1以上）
 */
int main(int argc, char *argv[]) {
    /* 乱数シード初期化（クライアントチャレンジ生成用） */
    srand(time(NULL));

    /* 引数チェック */
    if (argc < 2) {
        fprintf(stderr, "エラー: 環境を指定してください\n\n");
        print_help(argv[0]);
        return 1;
    }

    if (strcmp(argv[1], "-h") == 0 || strcmp(argv[1], "--help") == 0) {
        print_help(argv[0]);
        return 0;
    }

    /* 設定読み込み */
    load_config();

    /* 環境の有効性チェック */
    bool valid = false;
    for (int i = 0; ENVIRONMENTS[i]; i++) {
        if (strcmp(ENVIRONMENTS[i], argv[1]) == 0) {
            valid = true;
            break;
        }
    }

    if (!valid) {
        char msg[256];
        snprintf(msg, sizeof(msg), "無効な環境が指定されました: %s", argv[1]);
        log_error(msg);
        fprintf(stderr, "利用可能な環境: ");
        for (int i = 0; ENVIRONMENTS[i]; i++) {
            if (i > 0) fprintf(stderr, ", ");
            fprintf(stderr, "%s", ENVIRONMENTS[i]);
        }
        fprintf(stderr, "\n");
        return 1;
    }

    strncpy(g_env_folder, argv[1], sizeof(g_env_folder) - 1);

    /* ヘッダー表示 */
    printf("\n");
    printf("========================================================================\n");
    printf("  WinRM Remote Batch Executor (C言語版)\n");
    printf("  標準ライブラリのみ - NTLM認証\n");
    printf("========================================================================\n");
    printf("\n");

    char msg[256];
    snprintf(msg, sizeof(msg), "指定された環境: %s", g_env_folder);
    log_success(msg);
    printf("\n");

    snprintf(msg, sizeof(msg), "接続先: http://%s:%d/wsman", g_host, g_port);
    log_info(msg);
    snprintf(msg, sizeof(msg), "ユーザー: %s", g_user);
    log_info(msg);

    /* バッチファイルパスの{ENV}を置換 */
    str_replace(g_batch_path, "{ENV}", g_env_folder);
    snprintf(msg, sizeof(msg), "バッチファイル実行: %s", g_batch_path);
    log_info(msg);
    printf("\n");

    /* コマンド構築 */
    char command[1024];
    snprintf(command, sizeof(command), "cmd.exe /c \"%s\"", g_batch_path);

    /* シェル作成 */
    char shell_id[128];
    if (!create_shell(shell_id, sizeof(shell_id))) {
        log_error("処理を中断します");
        return 1;
    }
    printf("\n");

    /* コマンド実行 */
    char command_id[128];
    if (!run_command(shell_id, command, command_id, sizeof(command_id))) {
        delete_shell(shell_id);
        log_error("処理を中断します");
        return 1;
    }
    printf("\n");

    /* 出力取得 */
    char stdout_buf[MAX_BUFFER_SIZE];
    char stderr_buf[MAX_BUFFER_SIZE];
    int exit_code = 0;

    if (!get_command_output(shell_id, command_id,
                            stdout_buf, sizeof(stdout_buf),
                            stderr_buf, sizeof(stderr_buf),
                            &exit_code)) {
        delete_shell(shell_id);
        log_error("処理を中断します");
        return 1;
    }
    printf("\n");

    /* シェル削除 */
    delete_shell(shell_id);

    /* 結果表示 */
    printf("\n");
    printf("============================================================\n");
    printf("実行結果\n");
    printf("============================================================\n");

    if (strlen(stdout_buf) > 0) {
        printf("\n[標準出力]\n%s", stdout_buf);
    }

    if (strlen(stderr_buf) > 0) {
        printf("\n[標準エラー出力]\n%s", stderr_buf);
    }

    printf("\n終了コード: %d\n", exit_code);
    printf("============================================================\n");

    if (exit_code == 0) {
        log_success("完了");
    } else {
        snprintf(msg, sizeof(msg), "コマンドが失敗しました (終了コード: %d)", exit_code);
        log_error(msg);
    }

    return exit_code;
}
