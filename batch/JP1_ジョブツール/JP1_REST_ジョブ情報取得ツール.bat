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
# ユニット状態情報の取得
# ========================================
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "ユニット状態情報を取得中..." -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

try {
    # ユニット状態取得API
    $statusUri = "${protocol}://${webConsoleHost}:${webConsolePort}/ajs/api/v1/objects/statuses"
    $statusUri += "?manager=$managerHost"
    $statusUri += "&serviceName=$schedulerService"
    $statusUri += "&location=$unitPath"
    $statusUri += "&mode=search"

    Write-Host "リクエストURL: $statusUri"
    Write-Host ""

    $response = Invoke-RestMethod -Uri $statusUri -Method GET -Headers $headers -TimeoutSec 30

    if ($response.units -and $response.units.Count -gt 0) {
        Write-Host "[OK] ユニット情報を取得しました" -ForegroundColor Green
        Write-Host ""

        foreach ($unit in $response.units) {
            Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan
            Write-Host "ユニット名    : $($unit.definition.unitName)" -ForegroundColor White
            Write-Host "ユニットタイプ: $($unit.definition.unitType)"
            Write-Host "パス          : $($unit.definition.path)"
            Write-Host ""

            if ($unit.unitStatus) {
                $status = $unit.unitStatus

                # 状態の日本語変換
                $statusJp = switch ($status.status) {
                    "NORMAL"          { "正常終了" }
                    "ABNORMAL"        { "異常終了" }
                    "RUNNING"         { "実行中" }
                    "WARNING"         { "警告終了" }
                    "WAITING"         { "待機中" }
                    "HOLDING"         { "保留中" }
                    "NOT_REGISTERED"  { "未登録" }
                    "SKIPPED"         { "スキップ" }
                    "KILLED"          { "強制終了" }
                    "NOT_SCHEDULED"   { "未予定" }
                    default           { $status.status }
                }

                Write-Host "【状態情報】" -ForegroundColor Yellow
                Write-Host "  状態        : $statusJp ($($status.status))"

                if ($status.startTime) {
                    Write-Host "  開始時刻    : $($status.startTime)"
                }
                if ($status.endTime) {
                    Write-Host "  終了時刻    : $($status.endTime)"
                }
                if ($status.returnCode -ne $null) {
                    Write-Host "  終了コード  : $($status.returnCode)"
                }
                if ($status.holdAttr) {
                    Write-Host "  保留属性    : $($status.holdAttr)"
                }
                if ($status.execID) {
                    Write-Host "  実行ID      : $($status.execID)"
                }
            }
            Write-Host ""
        }

        Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan
        Write-Host "取得件数: $($response.units.Count) 件" -ForegroundColor Green

    } else {
        Write-Host "[情報] ユニットが見つかりませんでした" -ForegroundColor Yellow
        Write-Host "  パスを確認してください: $unitPath"
    }

} catch {
    Write-Host "[エラー] API呼び出しに失敗しました" -ForegroundColor Red
    Write-Host ""
    Write-Host "エラー詳細:" -ForegroundColor Red
    Write-Host $_.Exception.Message
    Write-Host ""

    # HTTPステータスコード別のヒント
    if ($_.Exception.Response) {
        $statusCode = [int]$_.Exception.Response.StatusCode
        Write-Host "HTTPステータスコード: $statusCode"

        switch ($statusCode) {
            401 { Write-Host "→ 認証エラー: JP1ユーザー名またはパスワードを確認してください" }
            403 { Write-Host "→ 権限エラー: JP1ユーザーに必要な権限があるか確認してください" }
            404 { Write-Host "→ 指定したユニットパスが見つかりません" }
            500 { Write-Host "→ サーバー内部エラー: Web Consoleのログを確認してください" }
        }
    }

    Write-Host ""
    Write-Host "確認事項:" -ForegroundColor Yellow
    Write-Host "  - Web Consoleが起動しているか"
    Write-Host "    → http://${webConsoleHost}:${webConsolePort}/ajs/login.html にアクセス"
    Write-Host "  - ホスト名・ポート番号が正しいか"
    Write-Host "  - JP1ユーザー名・パスワードが正しいか"
    Write-Host "  - ユニットパスが正しいか"

    exit 1
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "[完了] 処理が正常に終了しました" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green

exit 0
