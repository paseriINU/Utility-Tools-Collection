<# :
@echo off
chcp 65001 >nul
title JP1 REST API ジョブ情報取得ツール
setlocal

powershell -NoProfile -ExecutionPolicy Bypass -Command "$scriptDir=('%~dp0' -replace '\\$',''); iex ((gc '%~f0' -Encoding UTF8) -join \"`n\")"
set EXITCODE=%ERRORLEVEL%
pause
exit /b %EXITCODE%
: #>

# ==============================================================================
# JP1 REST API ジョブ情報取得ツール
#
# 説明:
#   JP1/AJS3 Web Console REST APIを使用して、ジョブ/ジョブネットの
#   状態情報を取得します。（ajsshow相当の情報をREST APIで取得）
#   ※ JP1/AJS3 - Web Consoleが必要です
#
# 使い方:
#   1. 下記の「設定セクション」を編集
#   2. このファイルをダブルクリックで実行
#
# 参考:
#   https://itpfdoc.hitachi.co.jp/manuals/3021/30213b1920/AJSO0280.HTM
# ==============================================================================

# ==============================================================================
# ■ 設定セクション（ここを編集してください）
# ==============================================================================

# Web Consoleサーバーのホスト名またはIPアドレス
$webConsoleHost = "localhost"

# Web Consoleのポート番号（HTTP: 22252, HTTPS: 22253）
$webConsolePort = "22252"

# HTTPSを使用する場合は $true に設定
$useHttps = $false

# JP1/AJS3 Managerのホスト名
$managerHost = "localhost"

# スケジューラーサービス名
$schedulerService = "AJSROOT1"

# JP1ユーザー名
$jp1User = "jp1admin"

# JP1パスワード（★★★ ここにパスワードを入力 ★★★）
$jp1Password = "password"

# 取得対象のユニットパス（ジョブネットまたはジョブ）
# 例: "/main_unit/jobgroup1/daily_batch"
$unitPath = "/main_unit/jobgroup1/daily_batch"

# 実行ID（execResultDetails API用）
# 例: "@A100", "@A101" など（ジョブネットの実行登録番号）
# ※ 実行登録時に割り当てられるID。Viewで確認可能
$execId = "@A100"

# デバッグモード（$true でレスポンス詳細を表示）
$debugMode = $true

# 試すAPIエンドポイント（1〜4を選択）
# 1: statuses（実行登録中のユニット状態）
# 2: definitions（ユニット定義情報）
# 3: results（実行結果 - 存在する場合）
# 4: すべてのAPIを順番に試す
$apiMode = 4

# ==============================================================================
# ■ メイン処理（以下は編集不要）
# ==============================================================================

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  JP1 REST API ジョブ情報取得ツール" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "設定内容:"
Write-Host "  Web Consoleサーバー : ${webConsoleHost}:${webConsolePort}"
Write-Host "  Managerホスト       : $managerHost"
Write-Host "  スケジューラー      : $schedulerService"
Write-Host "  JP1ユーザー         : $jp1User"
Write-Host "  ユニットパス        : $unitPath"
Write-Host "  実行ID              : $execId"
Write-Host ""

# プロトコル設定
$protocol = if ($useHttps) { "https" } else { "http" }

# 認証情報の作成（Base64エンコード）
$authString = "${jp1User}:${jp1Password}"
$authBytes = [System.Text.Encoding]::ASCII.GetBytes($authString)
$authBase64 = [System.Convert]::ToBase64String($authBytes)

Write-Host "[DEBUG] 認証文字列: ${jp1User}:***" -ForegroundColor Gray
Write-Host "[DEBUG] Base64: $($authBase64.Substring(0,10))..." -ForegroundColor Gray
Write-Host ""

# 共通ヘッダー
$headers = @{
    "Content-Type" = "application/json"
    "Accept" = "application/json"
    "Accept-Language" = "ja"
    "X-AJS-Authorization" = $authBase64
}

# SSL証明書検証をスキップ（自己署名証明書対応）
if ($useHttps) {
    Add-Type @"
        using System.Net;
        using System.Security.Cryptography.X509Certificates;
        public class TrustAllCertsPolicy : ICertificatePolicy {
            public bool CheckValidationResult(
                ServicePoint srvPoint, X509Certificate certificate,
                WebRequest request, int certificateProblem) {
                return true;
            }
        }
"@
    [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
}

# ========================================
# API呼び出し関数
# ========================================
function Call-JP1Api {
    param(
        [string]$ApiName,
        [string]$ApiUrl
    )

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "API: $ApiName" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "リクエストURL:" -ForegroundColor Cyan
    Write-Host "  $ApiUrl"
    Write-Host ""

    try {
        $webResponse = Invoke-WebRequest -Uri $ApiUrl -Method GET -Headers $headers -TimeoutSec 30 -UseBasicParsing
        Write-Host "[OK] HTTPステータス: $($webResponse.StatusCode)" -ForegroundColor Green
        Write-Host ""
        Write-Host "レスポンスボディ:" -ForegroundColor Cyan
        Write-Host $webResponse.Content
        Write-Host ""
        return $true
    } catch {
        $errMsg = $_.Exception.Message
        Write-Host "[エラー] $errMsg" -ForegroundColor Red
        if ($_.Exception.Response) {
            $statusCode = [int]$_.Exception.Response.StatusCode
            Write-Host "HTTPステータス: $statusCode" -ForegroundColor Red
        }
        Write-Host ""
        return $false
    }
}

# ========================================
# 試すAPIエンドポイント一覧
# ========================================
$baseUrl = "${protocol}://${webConsoleHost}:${webConsolePort}"

$apiEndpoints = @(
    @{
        Name = "1. statuses（実行登録中ユニット状態）"
        Url = "${baseUrl}/ajs/api/v1/objects/statuses?manager=${managerHost}&serviceName=${schedulerService}&location=${unitPath}&mode=search"
    },
    @{
        Name = "2. definitions（ユニット定義情報）"
        Url = "${baseUrl}/ajs/api/v1/objects/definitions?manager=${managerHost}&serviceName=${schedulerService}&location=${unitPath}"
    },
    @{
        Name = "3. results（実行結果詳細）"
        Url = "${baseUrl}/ajs/api/v1/objects/results?manager=${managerHost}&serviceName=${schedulerService}&location=${unitPath}"
    },
    @{
        Name = "4. executions（実行履歴）"
        Url = "${baseUrl}/ajs/api/v1/executions?manager=${managerHost}&serviceName=${schedulerService}&location=${unitPath}"
    },
    @{
        Name = "5. statuses（mode無し）"
        Url = "${baseUrl}/ajs/api/v1/objects/statuses?manager=${managerHost}&serviceName=${schedulerService}&location=${unitPath}"
    },
    @{
        Name = "6. execResultDetails（実行結果詳細 - 標準エラー出力）"
        Url = "${baseUrl}/ajs/api/v1/objects/statuses/${unitPath}:${execId}/actions/execResultDetails/invoke?manager=${managerHost}&serviceName=${schedulerService}"
    }
)

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "APIエンドポイントを順番に試します..." -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan

$successCount = 0
foreach ($api in $apiEndpoints) {
    $result = Call-JP1Api -ApiName $api.Name -ApiUrl $api.Url
    if ($result) {
        $successCount++
    }
    Write-Host "----------------------------------------------------------------"
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "テスト完了: $successCount / $($apiEndpoints.Count) 件成功" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green

Write-Host ""
Write-Host "注意:" -ForegroundColor Yellow
Write-Host "  - 「statuses」APIは実行登録中のジョブのみ対象です"
Write-Host "  - 「execResultDetails」APIは標準エラー出力を取得します（標準出力ではありません）"
Write-Host "  - execResultDetails APIを使用するには実行ID（@A100など）が必要です"
Write-Host "  - 標準出力の取得には ajsshow コマンド（WinRM経由）が必要です"
Write-Host ""

exit 0
