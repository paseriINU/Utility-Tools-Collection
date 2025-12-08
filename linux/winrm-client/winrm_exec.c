/*
 * WinRM Remote Batch Executor for Linux (C言語版 - libcurl使用)
 * Linux（Red Hat等）からWindows Server 2022へWinRM接続してバッチを実行
 *
 * 必要なライブラリ:
 *   - libcurl (多くのLinux環境で標準インストール済み)
 *
 * コンパイル:
 *   gcc -o winrm_exec winrm_exec.c -lcurl
 *
 * 使い方:
 *   1. このソースファイル内の設定セクションを編集
 *   2. コンパイル: gcc -o winrm_exec winrm_exec.c -lcurl
 *   3. 実行: ./winrm_exec ENV
 *
 *   環境を引数で指定（必須）:
 *   ./winrm_exec TST1T
 *   ./winrm_exec TST2T
 *
 *   または環境変数で設定を上書き:
 *   WINRM_HOST=192.168.1.100 WINRM_USER=Admin WINRM_PASS=Pass123 ./winrm_exec TST1T
 */

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <stdbool.h>
#include <time.h>
#include <curl/curl.h>

/* ==================== 設定セクション ====================
 * ここを編集して使用してください
 */

/* Windows接続情報 */
#define DEFAULT_HOST "192.168.1.100"
#define DEFAULT_USER "Administrator"
#define DEFAULT_PASS "YourPassword"
#define DEFAULT_PORT 5985                /* WinRMポート（HTTP: 5985, HTTPS: 5986） */

/* 実行するバッチファイル（Windows側のパス）
 * {ENV} は選択した環境フォルダ名に置換されます */
#define DEFAULT_BATCH_PATH "C:\\Scripts\\{ENV}\\test.bat"

/* 利用可能な環境のリスト */
static const char *ENVIRONMENTS[] = {"TST1T", "TST2T", NULL};

/* タイムアウト（秒） */
#define TIMEOUT 300

/* デバッグモード（1にするとXML送受信を表示） */
#define DEBUG 0

/* ========================================================= */

/* 色付き出力用 */
#define COLOR_RED     "\033[0;31m"
#define COLOR_GREEN   "\033[0;32m"
#define COLOR_YELLOW  "\033[1;33m"
#define COLOR_BLUE    "\033[0;34m"
#define COLOR_RESET   "\033[0m"

/* バッファサイズ */
#define MAX_BUFFER_SIZE 65536
#define MAX_URL_SIZE 512
#define MAX_UUID_SIZE 64
#define MAX_ENVELOPE_SIZE 8192

/* グローバル設定（環境変数で上書き可能） */
static char g_host[256];
static char g_user[256];
static char g_pass[256];
static int g_port;
static bool g_use_https;
static char g_batch_path[512];
static char g_env_folder[64];

/* レスポンス格納用構造体 */
typedef struct {
    char *data;
    size_t size;
} ResponseBuffer;

/* ログ関数 */
static void log_info(const char *msg) {
    fprintf(stderr, "%s[INFO]%s %s\n", COLOR_BLUE, COLOR_RESET, msg);
}

static void log_success(const char *msg) {
    fprintf(stderr, "%s[SUCCESS]%s %s\n", COLOR_GREEN, COLOR_RESET, msg);
}

static void log_warn(const char *msg) {
    fprintf(stderr, "%s[WARN]%s %s\n", COLOR_YELLOW, COLOR_RESET, msg);
}

static void log_error(const char *msg) {
    fprintf(stderr, "%s[ERROR]%s %s\n", COLOR_RED, COLOR_RESET, msg);
}

/* UUID生成（簡易版） */
static void generate_uuid(char *uuid, size_t size) {
    FILE *fp = fopen("/proc/sys/kernel/random/uuid", "r");
    if (fp) {
        if (fgets(uuid, size, fp)) {
            /* 改行を削除 */
            char *newline = strchr(uuid, '\n');
            if (newline) *newline = '\0';
        }
        fclose(fp);
    } else {
        /* フォールバック: 時刻ベースの簡易UUID */
        snprintf(uuid, size, "%08lx-%04x-4%03x-%04x-%012lx",
                 (unsigned long)time(NULL),
                 rand() & 0xffff,
                 rand() & 0x0fff,
                 rand() & 0xffff,
                 (unsigned long)time(NULL));
    }
}

/* XML特殊文字のエスケープ */
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

/* 文字列置換 */
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

/* XMLからタグの値を抽出 */
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

/* Base64デコード */
static size_t base64_decode(const char *input, char *output, size_t output_size) {
    static const char base64_chars[] =
        "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

    size_t input_len = strlen(input);
    size_t output_len = 0;
    unsigned int buffer = 0;
    int bits = 0;

    for (size_t i = 0; i < input_len && output_len < output_size - 1; i++) {
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
    output[output_len] = '\0';
    return output_len;
}

/* curlレスポンスコールバック */
static size_t write_callback(void *contents, size_t size, size_t nmemb, void *userp) {
    size_t realsize = size * nmemb;
    ResponseBuffer *buf = (ResponseBuffer *)userp;

    char *ptr = realloc(buf->data, buf->size + realsize + 1);
    if (!ptr) {
        log_error("メモリ割り当てに失敗しました");
        return 0;
    }

    buf->data = ptr;
    memcpy(&(buf->data[buf->size]), contents, realsize);
    buf->size += realsize;
    buf->data[buf->size] = '\0';

    return realsize;
}

/* エンドポイントURL生成 */
static void generate_endpoint(char *url, size_t size) {
    const char *protocol = g_use_https ? "https" : "http";
    snprintf(url, size, "%s://%s:%d/wsman", protocol, g_host, g_port);
}

/* SOAPリクエスト送信 */
static bool send_soap_request(const char *soap_envelope, ResponseBuffer *response) {
    CURL *curl;
    CURLcode res;
    char url[MAX_URL_SIZE];
    char userpwd[512];
    long http_code = 0;

    generate_endpoint(url, sizeof(url));
    snprintf(userpwd, sizeof(userpwd), "%s:%s", g_user, g_pass);

    if (DEBUG) {
        log_info("送信XML:");
        fprintf(stderr, "%s\n\n", soap_envelope);
        fprintf(stderr, "%s[INFO]%s 接続先: %s\n", COLOR_BLUE, COLOR_RESET, url);
        fprintf(stderr, "%s[INFO]%s ユーザー: %s\n", COLOR_BLUE, COLOR_RESET, g_user);
    }

    curl = curl_easy_init();
    if (!curl) {
        log_error("curlの初期化に失敗しました");
        return false;
    }

    response->data = malloc(1);
    response->size = 0;
    response->data[0] = '\0';

    struct curl_slist *headers = NULL;
    headers = curl_slist_append(headers, "Content-Type: application/soap+xml;charset=UTF-8");

    curl_easy_setopt(curl, CURLOPT_URL, url);
    curl_easy_setopt(curl, CURLOPT_HTTPHEADER, headers);
    curl_easy_setopt(curl, CURLOPT_POSTFIELDS, soap_envelope);
    curl_easy_setopt(curl, CURLOPT_WRITEFUNCTION, write_callback);
    curl_easy_setopt(curl, CURLOPT_WRITEDATA, (void *)response);
    curl_easy_setopt(curl, CURLOPT_TIMEOUT, (long)TIMEOUT);
    curl_easy_setopt(curl, CURLOPT_HTTPAUTH, CURLAUTH_NTLM);
    curl_easy_setopt(curl, CURLOPT_USERPWD, userpwd);

    if (g_use_https) {
        /* 自己署名証明書を許可 */
        curl_easy_setopt(curl, CURLOPT_SSL_VERIFYPEER, 0L);
        curl_easy_setopt(curl, CURLOPT_SSL_VERIFYHOST, 0L);
    }

    if (DEBUG) {
        curl_easy_setopt(curl, CURLOPT_VERBOSE, 1L);
    }

    res = curl_easy_perform(curl);
    curl_easy_getinfo(curl, CURLINFO_RESPONSE_CODE, &http_code);

    curl_slist_free_all(headers);
    curl_easy_cleanup(curl);

    if (res != CURLE_OK) {
        char msg[256];
        snprintf(msg, sizeof(msg), "curlエラー: %s", curl_easy_strerror(res));
        log_error(msg);

        switch (res) {
            case CURLE_COULDNT_RESOLVE_HOST:
                log_error("ホスト名の解決に失敗しました");
                break;
            case CURLE_COULDNT_CONNECT:
                log_error("接続に失敗しました。ホストが起動しているか確認してください");
                break;
            case CURLE_OPERATION_TIMEDOUT:
                log_error("タイムアウトしました");
                break;
            default:
                break;
        }
        return false;
    }

    if (DEBUG) {
        char msg[64];
        snprintf(msg, sizeof(msg), "HTTPステータスコード: %ld", http_code);
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
        snprintf(msg, sizeof(msg), "予期しないHTTPステータスコード: %ld", http_code);
        log_warn(msg);
    }

    if (DEBUG) {
        log_info("受信XML:");
        fprintf(stderr, "%s\n\n", response->data);
    }

    return true;
}

/* シェル作成 */
static bool create_shell(char *shell_id, size_t shell_id_size) {
    char url[MAX_URL_SIZE];
    char uuid[MAX_UUID_SIZE];
    char envelope[MAX_ENVELOPE_SIZE];
    ResponseBuffer response = {NULL, 0};

    generate_endpoint(url, sizeof(url));
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

    if (!send_soap_request(envelope, &response)) {
        log_error("シェル作成に失敗しました");
        if (response.data) free(response.data);
        return false;
    }

    if (!extract_xml_value(response.data, "rsp:ShellId", shell_id, shell_id_size)) {
        log_error("ShellIDの取得に失敗しました");
        if (response.data) free(response.data);
        return false;
    }

    char msg[256];
    snprintf(msg, sizeof(msg), "シェル作成成功: %s", shell_id);
    log_success(msg);

    free(response.data);
    return true;
}

/* コマンド実行 */
static bool run_command(const char *shell_id, const char *command, char *command_id, size_t command_id_size) {
    char url[MAX_URL_SIZE];
    char uuid[MAX_UUID_SIZE];
    char envelope[MAX_ENVELOPE_SIZE];
    char command_escaped[1024];
    ResponseBuffer response = {NULL, 0};

    generate_endpoint(url, sizeof(url));
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

    if (!send_soap_request(envelope, &response)) {
        log_error("コマンド実行に失敗しました");
        if (response.data) free(response.data);
        return false;
    }

    if (!extract_xml_value(response.data, "rsp:CommandId", command_id, command_id_size)) {
        log_error("CommandIDの取得に失敗しました");
        if (response.data) free(response.data);
        return false;
    }

    char msg[256];
    snprintf(msg, sizeof(msg), "コマンド実行開始: %s", command_id);
    log_success(msg);

    free(response.data);
    return true;
}

/* コマンド出力取得 */
static bool get_command_output(const char *shell_id, const char *command_id,
                               char *stdout_buf, size_t stdout_size,
                               char *stderr_buf, size_t stderr_size,
                               int *exit_code) {
    char url[MAX_URL_SIZE];
    char envelope[MAX_ENVELOPE_SIZE];
    bool command_done = false;
    int max_attempts = TIMEOUT * 2;

    generate_endpoint(url, sizeof(url));

    stdout_buf[0] = '\0';
    stderr_buf[0] = '\0';
    *exit_code = 0;

    char msg[128];
    snprintf(msg, sizeof(msg), "コマンド出力取得中...（最大%d秒待機）", TIMEOUT);
    log_info(msg);

    for (int attempt = 0; attempt < max_attempts && !command_done; attempt++) {
        char uuid[MAX_UUID_SIZE];
        ResponseBuffer response = {NULL, 0};

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

        if (!send_soap_request(envelope, &response)) {
            log_error("出力取得に失敗しました");
            if (response.data) free(response.data);
            return false;
        }

        /* stdout抽出 */
        char *stdout_start = strstr(response.data, "<rsp:Stream Name=\"stdout\">");
        if (stdout_start) {
            stdout_start += strlen("<rsp:Stream Name=\"stdout\">");
            char *stdout_end = strstr(stdout_start, "</rsp:Stream>");
            if (stdout_end) {
                size_t b64_len = stdout_end - stdout_start;
                char *b64_buf = malloc(b64_len + 1);
                strncpy(b64_buf, stdout_start, b64_len);
                b64_buf[b64_len] = '\0';

                char decoded[MAX_BUFFER_SIZE];
                base64_decode(b64_buf, decoded, sizeof(decoded));
                strncat(stdout_buf, decoded, stdout_size - strlen(stdout_buf) - 1);
                free(b64_buf);
            }
        }

        /* stderr抽出 */
        char *stderr_start = strstr(response.data, "<rsp:Stream Name=\"stderr\">");
        if (stderr_start) {
            stderr_start += strlen("<rsp:Stream Name=\"stderr\">");
            char *stderr_end = strstr(stderr_start, "</rsp:Stream>");
            if (stderr_end) {
                size_t b64_len = stderr_end - stderr_start;
                char *b64_buf = malloc(b64_len + 1);
                strncpy(b64_buf, stderr_start, b64_len);
                b64_buf[b64_len] = '\0';

                char decoded[MAX_BUFFER_SIZE];
                base64_decode(b64_buf, decoded, sizeof(decoded));
                strncat(stderr_buf, decoded, stderr_size - strlen(stderr_buf) - 1);
                free(b64_buf);
            }
        }

        /* コマンド完了チェック */
        if (strstr(response.data, "CommandState/Done")) {
            command_done = true;
            char exit_code_str[16];
            if (extract_xml_value(response.data, "rsp:ExitCode", exit_code_str, sizeof(exit_code_str))) {
                *exit_code = atoi(exit_code_str);
            }
        }

        free(response.data);

        if (!command_done) {
            struct timespec ts = {0, 500000000}; /* 0.5秒 */
            nanosleep(&ts, NULL);
        }
    }

    if (!command_done) {
        log_warn("コマンド完了待機がタイムアウトしました");
    }

    snprintf(msg, sizeof(msg), "コマンド完了 (終了コード: %d)", *exit_code);
    log_success(msg);

    return true;
}

/* シェル削除 */
static void delete_shell(const char *shell_id) {
    char url[MAX_URL_SIZE];
    char uuid[MAX_UUID_SIZE];
    char envelope[MAX_ENVELOPE_SIZE];
    ResponseBuffer response = {NULL, 0};

    generate_endpoint(url, sizeof(url));
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
    send_soap_request(envelope, &response);
    log_success("シェル削除完了");

    if (response.data) free(response.data);
}

/* 環境変数から設定を読み込み */
static void load_config(void) {
    const char *env;

    env = getenv("WINRM_HOST");
    strncpy(g_host, env ? env : DEFAULT_HOST, sizeof(g_host) - 1);

    env = getenv("WINRM_USER");
    strncpy(g_user, env ? env : DEFAULT_USER, sizeof(g_user) - 1);

    env = getenv("WINRM_PASS");
    strncpy(g_pass, env ? env : DEFAULT_PASS, sizeof(g_pass) - 1);

    env = getenv("WINRM_PORT");
    g_port = env ? atoi(env) : DEFAULT_PORT;

    /* ポート番号に応じてHTTPS/HTTPを自動判定 */
    env = getenv("USE_HTTPS");
    if (env) {
        g_use_https = (strcmp(env, "true") == 0 || strcmp(env, "1") == 0);
    } else {
        g_use_https = (g_port != 5985);
    }

    env = getenv("BATCH_FILE_PATH");
    strncpy(g_batch_path, env ? env : DEFAULT_BATCH_PATH, sizeof(g_batch_path) - 1);
}

/* ヘルプ表示 */
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
    printf("  WINRM_HOST, WINRM_PORT, WINRM_USER, WINRM_PASS, USE_HTTPS\n");
}

/* メイン処理 */
int main(int argc, char *argv[]) {
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
    printf("  libcurl使用\n");
    printf("========================================================================\n");
    printf("\n");

    char msg[256];
    snprintf(msg, sizeof(msg), "指定された環境: %s", g_env_folder);
    log_success(msg);
    printf("\n");

    char url[MAX_URL_SIZE];
    generate_endpoint(url, sizeof(url));
    snprintf(msg, sizeof(msg), "接続先: %s", url);
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

    /* curl初期化 */
    curl_global_init(CURL_GLOBAL_ALL);

    /* シェル作成 */
    char shell_id[128];
    if (!create_shell(shell_id, sizeof(shell_id))) {
        log_error("処理を中断します");
        curl_global_cleanup();
        return 1;
    }
    printf("\n");

    /* コマンド実行 */
    char command_id[128];
    if (!run_command(shell_id, command, command_id, sizeof(command_id))) {
        delete_shell(shell_id);
        log_error("処理を中断します");
        curl_global_cleanup();
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
        curl_global_cleanup();
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

    curl_global_cleanup();
    return exit_code;
}
