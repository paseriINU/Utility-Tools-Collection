<# :
@echo off
setlocal

rem 引数チェック
if "%~1"=="" (
    echo ERROR: ユニットパスを指定してください
    echo 使い方: %~nx0 "/JobGroup/Jobnet"
    exit /b 1
)

rem 引数を環境変数に設定
set "JP1_UNIT_PATH=%~1"

powershell -NoProfile -ExecutionPolicy Bypass -Command "$scriptDir=('%~dp0' -replace '\\$',''); iex ((gc '%~f0' -Encoding UTF8) -join \"`n\")"
exit /b %ERRORLEVEL%
: #>

# ==============================================================================
# JP1 REST API ジョブ情報取得ツール
#
# 説明:
#   JP1/AJS3 Web Console REST APIを使用して、ジョブ/ジョブネットの
#   実行結果詳細を取得します。
#   ※ JP1/AJS3 - Web Consoleが必要です
#
# 使い方:
#   JP1_REST_ジョブ情報取得ツール.bat "/JobGroup/Jobnet"
#
# 参考:
#   https://itpfdoc.hitachi.co.jp/manuals/3021/30213b1920/AJSO0280.HTM
# ==============================================================================

# 出力をShift-JISに設定
[Console]::OutputEncoding = [System.Text.Encoding]::GetEncoding(932)

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

# ユニットパスを環境変数から取得
$unitPath = $env:JP1_UNIT_PATH

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
        foreach ($unit in $jsonData.statuses) {
            $unitName = $unit.definition.unitName
            $unitType = $unit.definition.unitType
            $unitStatus = $unit.unitStatus
            $execId = if ($unitStatus) { $unitStatus.execID } else { $null }
            $status = if ($unitStatus) { $unitStatus.status } else { "N/A" }

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
    }
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
    exit 1
}

# ========================================
# 実行結果詳細の取得
# ========================================
if ($execIdList.Count -gt 0) {
    foreach ($item in $execIdList) {
        $targetPath = $item.Path
        $targetExecId = $item.ExecId

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

            # all が false の場合はエラー（5MB超過で切り捨て）
            if ($resultJson.all -eq $false) {
                Write-Output "ERROR: Result truncated (exceeded 5MB limit) for $targetPath"
                exit 1
            }

            if ($resultJson.execResultDetails) {
                Write-Output $resultJson.execResultDetails
            }
        } catch {
            Write-Output "ERROR: Failed to get details for $targetPath"
        }
    }
}

exit 0
