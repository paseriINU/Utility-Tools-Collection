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
#   状態情報と実行結果詳細を取得します。
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

# 取得対象のユニットパス（ジョブネット）
# 例: "/JobGroup/Jobnet"
$unitPath = "/JobGroup/Jobnet"

# デバッグモード（$true でレスポンス詳細を表示）
$debugMode = $true

# 世代指定（RESULT: 直近終了世代, STATUS: 最新世代, PERIOD: 期間指定）
# ※ RESULT を指定すると終了済みジョブの直近終了世代を取得
$generation = "RESULT"

# 期間指定（generation=PERIOD の場合に使用）
# 形式: YYYY-MM-DDThh:mm
$periodBegin = "2025-12-01T00:00"
$periodEnd = "2025-12-25T23:59"

# ステータスフィルタ（空欄で全件、ABNORMAL: 異常終了のみ、等）
# 指定可能値: ABNORMAL, NORMAL, RUNNING, WAITING, etc.
$statusFilter = ""

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
Write-Host "  世代                : $generation"
if ($generation -eq "PERIOD") {
    Write-Host "  期間                : $periodBegin ～ $periodEnd"
}
if ($statusFilter) {
    Write-Host "  ステータスフィルタ  : $statusFilter"
}
Write-Host ""

# プロトコル設定
$protocol = if ($useHttps) { "https" } else { "http" }

# 認証情報の作成（Base64エンコード）
$authString = "${jp1User}:${jp1Password}"
$authBytes = [System.Text.Encoding]::UTF8.GetBytes($authString)
$authBase64 = [System.Convert]::ToBase64String($authBytes)

Write-Host "[DEBUG] 認証文字列: ${jp1User}:***" -ForegroundColor Gray
Write-Host "[DEBUG] Base64: $($authBase64.Substring(0,10))..." -ForegroundColor Gray
Write-Host ""

# 共通ヘッダー（Pythonサンプルに合わせてシンプルに）
$headers = @{
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
# メイン処理: 2段階でAPIを呼び出し
# ========================================
$baseUrl = "${protocol}://${webConsoleHost}:${webConsolePort}/ajs/api/v1"

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "STEP 1: ユニット一覧取得API（execID取得）" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan

# URLエンコード（Pythonの urllib.parse.quote(LOCATION, safe='') と同等）
$encodedLocation = [System.Uri]::EscapeDataString($unitPath)

Write-Host ""
Write-Host "パス解析結果:" -ForegroundColor Cyan
Write-Host "  ユニットパス (location)     : $unitPath"
Write-Host "  エンコード後 location       : $encodedLocation"
Write-Host ""

# Step 1: statuses API でユニット一覧と execID を取得
# Pythonサンプルに準拠したURL構築
$statusUrl = "${baseUrl}/objects/statuses?mode=search"
$statusUrl += "&manager=${managerHost}"
$statusUrl += "&serviceName=${schedulerService}"
$statusUrl += "&location=${encodedLocation}"
$statusUrl += "&generation=${generation}"

# 期間指定の場合
if ($generation -eq "PERIOD") {
    $statusUrl += "&periodBegin=${periodBegin}"
    $statusUrl += "&periodEnd=${periodEnd}"
}

# ステータスフィルタがある場合
if ($statusFilter) {
    $statusUrl += "&status=${statusFilter}"
}

Write-Host "[DEBUG] リクエストヘッダー:" -ForegroundColor Gray
Write-Host "  X-AJS-Authorization: $($authBase64.Substring(0,10))..." -ForegroundColor Gray
Write-Host "  Accept-Language: ja" -ForegroundColor Gray

Write-Host ""
Write-Host "リクエストURL:" -ForegroundColor Cyan
Write-Host "  $statusUrl"
Write-Host ""

$execIdList = @()

try {
    $response = Invoke-WebRequest -Uri $statusUrl -Method GET -Headers $headers -TimeoutSec 30 -UseBasicParsing
    Write-Host "[OK] HTTPステータス: $($response.StatusCode)" -ForegroundColor Green

    # UTF-8文字化け対策: RawContentStreamからUTF-8としてデコード
    $responseBytes = $response.RawContentStream.ToArray()
    $responseText = [System.Text.Encoding]::UTF8.GetString($responseBytes)
    $jsonData = $responseText | ConvertFrom-Json

    if ($debugMode) {
        Write-Host ""
        Write-Host "レスポンス:" -ForegroundColor Gray
        Write-Host $responseText
    }

    # statuses配列からexecIDを抽出
    if ($jsonData.statuses -and $jsonData.statuses.Count -gt 0) {
        Write-Host ""
        Write-Host "取得したユニット一覧:" -ForegroundColor Green
        foreach ($unit in $jsonData.statuses) {
            $unitName = $unit.definition.unitName
            $unitType = $unit.definition.unitType
            $unitStatus = $unit.unitStatus
            $execId = if ($unitStatus) { $unitStatus.execID } else { $null }
            $status = if ($unitStatus) { $unitStatus.status } else { "N/A" }

            Write-Host "  ユニット: $unitName | 種別: $unitType | execID: $execId | 状態: $status"

            # ジョブ（PCJOB, UNIXJOB, QUEUEJOB等）でexecIDがある場合のみリストに追加
            if ($execId -and $unitType -match "JOB") {
                $execIdList += @{
                    Path = $unitName
                    ExecId = $execId
                    Status = $status
                    UnitType = $unitType
                }
            }
        }
        Write-Host ""
        Write-Host "ジョブ件数（execIDあり）: $($execIdList.Count)" -ForegroundColor Cyan
    } else {
        Write-Host ""
        Write-Host "[警告] 該当するユニットがありません" -ForegroundColor Yellow
    }
} catch {
    Write-Host "[エラー] $($_.Exception.Message)" -ForegroundColor Red
    if ($_.Exception.Response) {
        $statusCode = [int]$_.Exception.Response.StatusCode
        Write-Host "HTTPステータス: $statusCode" -ForegroundColor Red
    }
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
    Write-Host "  - 期間設定（periodBegin/periodEnd）を確認してください"
    Write-Host "  - 参照権限があるか確認してください"
} else {
    foreach ($item in $execIdList) {
        $targetPath = $item.Path
        $targetExecId = $item.ExecId
        $targetStatus = $item.Status

        Write-Host ""
        Write-Host "========================================" -ForegroundColor Yellow
        Write-Host "ユニット: $targetPath" -ForegroundColor Yellow
        Write-Host "実行ID  : $targetExecId" -ForegroundColor Yellow
        Write-Host "状態    : $targetStatus" -ForegroundColor Yellow
        Write-Host "========================================" -ForegroundColor Yellow

        # URLエンコード（Pythonの urllib.parse.quote と同等）
        $encodedPath = [System.Uri]::EscapeDataString($targetPath)

        # execResultDetails API を呼び出し
        # Pythonサンプル: f"{base_url}/objects/statuses/{unit_encoded}:{exec_id}/actions/execResultDetails/invoke"
        $detailUrl = "${baseUrl}/objects/statuses/${encodedPath}:${targetExecId}/actions/execResultDetails/invoke"
        $detailUrl += "?manager=${managerHost}&serviceName=${schedulerService}"

        Write-Host ""
        Write-Host "リクエストURL:" -ForegroundColor Cyan
        Write-Host "  $detailUrl"
        Write-Host ""

        try {
            $resultResponse = Invoke-WebRequest -Uri $detailUrl -Method GET -Headers $headers -TimeoutSec 30 -UseBasicParsing
            Write-Host "[OK] HTTPステータス: $($resultResponse.StatusCode)" -ForegroundColor Green

            # UTF-8文字化け対策
            $resultBytes = $resultResponse.RawContentStream.ToArray()
            $resultText = [System.Text.Encoding]::UTF8.GetString($resultBytes)
            $resultJson = $resultText | ConvertFrom-Json

            Write-Host ""
            Write-Host "実行結果詳細:" -ForegroundColor Green
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
