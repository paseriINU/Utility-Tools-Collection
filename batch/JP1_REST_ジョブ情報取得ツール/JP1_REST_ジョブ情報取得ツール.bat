<# :
@echo off
chcp 65001 >nul
title JP1 REST API ジョブ情報取得ツール
setlocal

rem 引数チェック
if "%~1"=="" (
    echo.
    echo [エラー] ユニットパスを指定してください
    echo.
    echo 使い方:
    echo   %~nx0 "/JobGroup/Jobnet"
    echo.
    pause
    exit /b 1
)

rem 引数を環境変数に設定
set "JP1_UNIT_PATH=%~1"

rem 第2引数があればサイレントモード（他バッチからの呼び出し用）
if not "%~2"=="" (
    set "JP1_SILENT_MODE=1"
)

powershell -NoProfile -ExecutionPolicy Bypass -Command "$scriptDir=('%~dp0' -replace '\\$',''); iex ((gc '%~f0' -Encoding UTF8) -join \"`n\")"
set EXITCODE=%ERRORLEVEL%

rem サイレントモードでなければpause
if not defined JP1_SILENT_MODE pause
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
#   JP1_REST_ジョブ情報取得ツール.bat "/JobGroup/Jobnet"
#   JP1_REST_ジョブ情報取得ツール.bat "/JobGroup/Jobnet" silent  ← サイレントモード
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

# サイレントモード判定
$silentMode = $false
if ($env:JP1_SILENT_MODE) {
    $silentMode = $true
}

# 対話モードの場合のみヘッダー表示
if (-not $silentMode) {
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
}

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

if (-not $silentMode) {
    Write-Host "ユニット一覧を取得中..." -ForegroundColor Cyan
}

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
        if (-not $silentMode) {
            Write-Host ""
            Write-Host "取得したユニット一覧:" -ForegroundColor Green
            Write-Host "----------------------------------------"
        }
        foreach ($unit in $jsonData.statuses) {
            $unitName = $unit.definition.unitName
            $unitType = $unit.definition.unitType
            $unitStatus = $unit.unitStatus
            $execId = if ($unitStatus) { $unitStatus.execID } else { $null }
            $status = if ($unitStatus) { $unitStatus.status } else { "N/A" }

            if (-not $silentMode) {
                Write-Host "  $unitName [$unitType] - $status"
            }

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
        if (-not $silentMode) {
            Write-Host "----------------------------------------"
            Write-Host "ジョブ件数: $($execIdList.Count)" -ForegroundColor Cyan
        }
    } else {
        if (-not $silentMode) {
            Write-Host ""
            Write-Host "[警告] 該当するユニットがありません" -ForegroundColor Yellow
        }
    }
} catch {
    if ($silentMode) {
        Write-Output "ERROR: $($_.Exception.Message)"
    } else {
        Write-Host "[エラー] $($_.Exception.Message)" -ForegroundColor Red
    }
    exit 1
}

# ========================================
# 実行結果詳細の取得
# ========================================
if ($execIdList.Count -gt 0) {
    if (-not $silentMode) {
        Write-Host ""
        Write-Host "実行結果詳細を取得中..." -ForegroundColor Cyan
        Write-Host ""
    }

    $jobIndex = 0
    foreach ($item in $execIdList) {
        $jobIndex++
        $targetPath = $item.Path
        $targetExecId = $item.ExecId
        $targetStatus = $item.Status

        if (-not $silentMode) {
            Write-Host "[$jobIndex/$($execIdList.Count)] $targetPath" -ForegroundColor Yellow
            Write-Host "  実行ID: $targetExecId | 状態: $targetStatus" -ForegroundColor Gray
        }

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
                if ($silentMode) {
                    # サイレントモード: 結果のみ標準出力
                    Write-Output $resultJson.execResultDetails
                } else {
                    # 対話モード: 装飾付きで表示
                    Write-Host "  ----------------------------------------" -ForegroundColor DarkGray
                    $resultJson.execResultDetails -split "`n" | ForEach-Object {
                        Write-Host "  $_"
                    }
                    Write-Host "  ----------------------------------------" -ForegroundColor DarkGray
                }
            } else {
                if (-not $silentMode) {
                    Write-Host "  (出力なし)" -ForegroundColor DarkGray
                }
            }
        } catch {
            if ($silentMode) {
                Write-Output "ERROR: Failed to get details for $targetPath"
            } else {
                Write-Host "  [エラー] 詳細取得失敗" -ForegroundColor Red
            }
        }
        if (-not $silentMode) {
            Write-Host ""
        }
    }
}

if (-not $silentMode) {
    Write-Host "================================================================" -ForegroundColor Green
    Write-Host "処理完了" -ForegroundColor Green
    Write-Host "================================================================" -ForegroundColor Green
}

exit 0
