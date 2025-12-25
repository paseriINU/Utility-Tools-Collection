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

# 画面クリア
Clear-Host

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  JP1 REST API ジョブ情報取得ツール" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  対象: $unitPath"
Write-Host "  世代: $generation"
if ($statusFilter) {
    Write-Host "  フィルタ: $statusFilter"
}
Write-Host ""

# プロトコル設定
$protocol = if ($useHttps) { "https" } else { "http" }

# 認証情報の作成（Base64エンコード）
$authString = "${jp1User}:${jp1Password}"
$authBytes = [System.Text.Encoding]::UTF8.GetBytes($authString)
$authBase64 = [System.Convert]::ToBase64String($authBytes)

# 共通ヘッダー
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
# メイン処理
# ========================================
$baseUrl = "${protocol}://${webConsoleHost}:${webConsolePort}/ajs/api/v1"

Write-Host "ユニット一覧を取得中..." -ForegroundColor Cyan

# URLエンコード
$encodedLocation = [System.Uri]::EscapeDataString($unitPath)

# statuses API でユニット一覧と execID を取得
$statusUrl = "${baseUrl}/objects/statuses?mode=search"
$statusUrl += "&manager=${managerHost}"
$statusUrl += "&serviceName=${schedulerService}"
$statusUrl += "&location=${encodedLocation}"
$statusUrl += "&generation=${generation}"

if ($generation -eq "PERIOD") {
    $statusUrl += "&periodBegin=${periodBegin}"
    $statusUrl += "&periodEnd=${periodEnd}"
}

if ($statusFilter) {
    $statusUrl += "&status=${statusFilter}"
}

$execIdList = @()

try {
    $response = Invoke-WebRequest -Uri $statusUrl -Method GET -Headers $headers -TimeoutSec 30 -UseBasicParsing

    # UTF-8文字化け対策
    $responseBytes = $response.RawContentStream.ToArray()
    $responseText = [System.Text.Encoding]::UTF8.GetString($responseBytes)
    $jsonData = $responseText | ConvertFrom-Json

    if ($jsonData.statuses -and $jsonData.statuses.Count -gt 0) {
        Write-Host ""
        Write-Host "取得したユニット一覧:" -ForegroundColor Green
        Write-Host "----------------------------------------"
        foreach ($unit in $jsonData.statuses) {
            $unitName = $unit.definition.unitName
            $unitType = $unit.definition.unitType
            $unitStatus = $unit.unitStatus
            $execId = if ($unitStatus) { $unitStatus.execID } else { $null }
            $status = if ($unitStatus) { $unitStatus.status } else { "N/A" }

            Write-Host "  $unitName [$unitType] - $status"

            # ジョブでexecIDがある場合のみリストに追加
            if ($execId -and $unitType -match "JOB") {
                $execIdList += @{
                    Path = $unitName
                    ExecId = $execId
                    Status = $status
                    UnitType = $unitType
                }
            }
        }
        Write-Host "----------------------------------------"
        Write-Host "ジョブ件数: $($execIdList.Count)" -ForegroundColor Cyan
    } else {
        Write-Host ""
        Write-Host "[警告] 該当するユニットがありません" -ForegroundColor Yellow
    }
} catch {
    Write-Host "[エラー] $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# ========================================
# 実行結果詳細の取得
# ========================================
if ($execIdList.Count -gt 0) {
    Write-Host ""
    Write-Host "実行結果詳細を取得中..." -ForegroundColor Cyan
    Write-Host ""

    $jobIndex = 0
    foreach ($item in $execIdList) {
        $jobIndex++
        $targetPath = $item.Path
        $targetExecId = $item.ExecId
        $targetStatus = $item.Status

        # 見やすいヘッダー
        Write-Host "[$jobIndex/$($execIdList.Count)] $targetPath" -ForegroundColor Yellow
        Write-Host "  実行ID: $targetExecId | 状態: $targetStatus" -ForegroundColor Gray

        # URLエンコード
        $encodedPath = [System.Uri]::EscapeDataString($targetPath)

        # execResultDetails API
        $detailUrl = "${baseUrl}/objects/statuses/${encodedPath}:${targetExecId}/actions/execResultDetails/invoke"
        $detailUrl += "?manager=${managerHost}&serviceName=${schedulerService}"

        try {
            $resultResponse = Invoke-WebRequest -Uri $detailUrl -Method GET -Headers $headers -TimeoutSec 30 -UseBasicParsing

            # UTF-8文字化け対策
            $resultBytes = $resultResponse.RawContentStream.ToArray()
            $resultText = [System.Text.Encoding]::UTF8.GetString($resultBytes)
            $resultJson = $resultText | ConvertFrom-Json

            if ($resultJson.execResultDetails) {
                Write-Host "  ----------------------------------------" -ForegroundColor DarkGray
                # 各行にインデントを付けて表示
                $resultJson.execResultDetails -split "`n" | ForEach-Object {
                    Write-Host "  $_"
                }
                Write-Host "  ----------------------------------------" -ForegroundColor DarkGray
            } else {
                Write-Host "  (出力なし)" -ForegroundColor DarkGray
            }
        } catch {
            Write-Host "  [エラー] 詳細取得失敗" -ForegroundColor Red
        }
        Write-Host ""
    }
}

Write-Host "================================================================" -ForegroundColor Green
Write-Host "処理完了" -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Green

exit 0
