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

# デバッグモード（$true でレスポンス詳細を表示）
$debugMode = $true

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
# メイン処理: 2段階でAPIを呼び出し
# ========================================
$baseUrl = "${protocol}://${webConsoleHost}:${webConsolePort}"

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "STEP 1: ユニット一覧取得API（execID取得）" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan

# Step 1: statuses API でユニット一覧と execID を取得
$statusesUrl = "${baseUrl}/ajs/api/v1/objects/statuses?manager=${managerHost}&serviceName=${schedulerService}&location=${unitPath}&mode=search"

Write-Host ""
Write-Host "リクエストURL:" -ForegroundColor Cyan
Write-Host "  $statusesUrl"
Write-Host ""

$execIdList = @()

try {
    $response = Invoke-WebRequest -Uri $statusesUrl -Method GET -Headers $headers -TimeoutSec 30 -UseBasicParsing
    Write-Host "[OK] HTTPステータス: $($response.StatusCode)" -ForegroundColor Green

    $jsonData = $response.Content | ConvertFrom-Json

    if ($debugMode) {
        Write-Host ""
        Write-Host "レスポンス:" -ForegroundColor Gray
        Write-Host $response.Content
    }

    # statuses配列からexecIDを抽出
    # レスポンス構造: statuses[].definition.unitName, statuses[].unitStatus.execID, statuses[].unitStatus.status
    if ($jsonData.statuses -and $jsonData.statuses.Count -gt 0) {
        Write-Host ""
        Write-Host "取得したユニット一覧:" -ForegroundColor Green
        foreach ($unit in $jsonData.statuses) {
            # ドキュメントに基づく正しいパス取得
            $path = $unit.definition.unitName
            $execId = $unit.unitStatus.execID
            $status = $unit.unitStatus.status
            Write-Host "  パス: $path | execID: $execId | 状態: $status"
            if ($execId) {
                $execIdList += @{ Path = $path; ExecId = $execId }
            }
        }
    } else {
        Write-Host ""
        Write-Host "[警告] 該当するユニットがありません" -ForegroundColor Yellow
    }
} catch {
    Write-Host "[エラー] $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "STEP 2: 実行結果詳細取得API（execResultDetails）" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan

if ($execIdList.Count -eq 0) {
    Write-Host ""
    Write-Host "[スキップ] execIDが取得できなかったため、実行結果詳細は取得できません" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "ヒント:" -ForegroundColor Cyan
    Write-Host "  - ユニットパスを確認してください"
    Write-Host "  - 参照権限があるか確認してください"
} else {
    foreach ($item in $execIdList) {
        $targetPath = $item.Path
        $targetExecId = $item.ExecId

        Write-Host ""
        Write-Host "----------------------------------------" -ForegroundColor Yellow
        Write-Host "対象: $targetPath (execID: $targetExecId)" -ForegroundColor Yellow
        Write-Host "----------------------------------------" -ForegroundColor Yellow

        # execResultDetails API を呼び出し
        $execResultUrl = "${baseUrl}/ajs/api/v1/objects/statuses/${targetPath}:${targetExecId}/actions/execResultDetails/invoke?manager=${managerHost}&serviceName=${schedulerService}"

        Write-Host ""
        Write-Host "リクエストURL:" -ForegroundColor Cyan
        Write-Host "  $execResultUrl"
        Write-Host ""

        try {
            $resultResponse = Invoke-WebRequest -Uri $execResultUrl -Method GET -Headers $headers -TimeoutSec 30 -UseBasicParsing
            Write-Host "[OK] HTTPステータス: $($resultResponse.StatusCode)" -ForegroundColor Green

            $resultJson = $resultResponse.Content | ConvertFrom-Json

            Write-Host ""
            Write-Host "実行結果詳細（標準エラー出力）:" -ForegroundColor Green
            Write-Host "----------------------------------------"
            if ($resultJson.execResultDetails) {
                Write-Host $resultJson.execResultDetails
            } else {
                Write-Host "(出力なし)"
            }
            Write-Host "----------------------------------------"
        } catch {
            Write-Host "[エラー] $($_.Exception.Message)" -ForegroundColor Red
            if ($_.Exception.Response) {
                $statusCode = [int]$_.Exception.Response.StatusCode
                Write-Host "HTTPステータス: $statusCode" -ForegroundColor Red
            }
        }
    }
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Green
Write-Host "処理完了" -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Green

Write-Host ""
Write-Host "注意:" -ForegroundColor Yellow
Write-Host "  - execResultDetails API は実行結果詳細（標準エラー出力相当）を取得します"
Write-Host "  - 標準出力の取得には ajsshow コマンド（WinRM経由）が必要です"
Write-Host ""

exit 0
