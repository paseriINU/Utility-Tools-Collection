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
# ユニット状態情報の取得
# ========================================
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "ユニット状態情報を取得中..." -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

try {
    # ユニット状態取得API（パスはエンコードせずそのまま使用）
    $baseUri = "${protocol}://${webConsoleHost}:${webConsolePort}/ajs/api/v1/objects/statuses"

    # URLを構築（パスはそのまま使用）
    $statusUri = "${baseUri}?manager=${managerHost}&serviceName=${schedulerService}&location=${unitPath}&mode=search"

    Write-Host "リクエストURL:" -ForegroundColor Cyan
    Write-Host "  $statusUri"
    Write-Host ""
    Write-Host "リクエストヘッダー:" -ForegroundColor Cyan
    Write-Host "  Content-Type: application/json"
    Write-Host "  Accept: application/json"
    Write-Host "  Accept-Language: ja"
    Write-Host "  X-AJS-Authorization: (Base64認証情報)"
    Write-Host ""

    # Invoke-WebRequestを使用して詳細なレスポンスを取得
    $webResponse = Invoke-WebRequest -Uri $statusUri -Method GET -Headers $headers -TimeoutSec 30 -UseBasicParsing

    Write-Host "[OK] HTTPステータス: $($webResponse.StatusCode)" -ForegroundColor Green
    Write-Host ""

    if ($debugMode) {
        Write-Host "[DEBUG] レスポンスヘッダー:" -ForegroundColor Gray
        foreach ($key in $webResponse.Headers.Keys) {
            $val = $webResponse.Headers[$key]
            Write-Host "  $key = $val" -ForegroundColor Gray
        }
        Write-Host ""
        Write-Host "[DEBUG] レスポンスボディ（生データ）:" -ForegroundColor Gray
        Write-Host $webResponse.Content -ForegroundColor Gray
        Write-Host ""
    }

    # JSONパース
    $response = $webResponse.Content | ConvertFrom-Json

    # レスポンス構造を確認
    Write-Host "[DEBUG] レスポンス構造:" -ForegroundColor Gray
    foreach ($prop in ($response | Get-Member -MemberType NoteProperty)) {
        Write-Host "  - $($prop.Name)" -ForegroundColor Gray
    }
    Write-Host ""

    # statuses配列またはunits配列の確認（バージョンによって異なる）
    $dataArray = $null
    $arrayName = ""

    if ($response.PSObject.Properties.Name -contains "statuses") {
        $dataArray = $response.statuses
        $arrayName = "statuses"
    } elseif ($response.PSObject.Properties.Name -contains "units") {
        $dataArray = $response.units
        $arrayName = "units"
    }

    if ($dataArray -ne $null) {
        if ($dataArray.Count -gt 0) {
            Write-Host "[OK] ユニット情報を取得しました（$($dataArray.Count) 件）" -ForegroundColor Green
            Write-Host ""

            foreach ($item in $dataArray) {
                Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan

                # 各プロパティを表示
                $itemJson = $item | ConvertTo-Json -Depth 3
                Write-Host $itemJson
                Write-Host ""
            }

            Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan
            Write-Host "取得件数: $($dataArray.Count) 件" -ForegroundColor Green

        } else {
            Write-Host "[情報] ${arrayName} 配列が空です（0件）" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "考えられる原因:" -ForegroundColor Yellow
            Write-Host "  1. ユニットパスが正しくない"
            Write-Host "     現在の設定: $unitPath"
            Write-Host "  2. ジョブネットが実行登録されていない"
            Write-Host "  3. JP1ユーザーに参照権限がない"
            Write-Host ""
            Write-Host "試してみてください:" -ForegroundColor Cyan
            Write-Host "  - パスの最後のスラッシュを削除/追加してみる"
            Write-Host "  - 親のジョブグループパスを指定してみる"
            Write-Host "  - Web Console画面で同じパスが表示されるか確認"
            Write-Host ""
            Write-Host "パス形式の例:" -ForegroundColor Cyan
            Write-Host "  /JOBGROUP/JOBNET"
            Write-Host "  /グループ名/ジョブネット名"
        }
    } else {
        Write-Host "[情報] レスポンスに statuses/units が含まれていません" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "レスポンス全体:" -ForegroundColor Gray
        $jsonOutput = $response | ConvertTo-Json -Depth 5
        Write-Host $jsonOutput
        Write-Host ""
        Write-Host "APIが異なるレスポンス形式を返しています。" -ForegroundColor Yellow
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
        Write-Host "HTTPステータスコード: $statusCode" -ForegroundColor Red

        # レスポンスボディを取得
        try {
            $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
            $errorBody = $reader.ReadToEnd()
            $reader.Close()
            Write-Host ""
            Write-Host "エラーレスポンス:" -ForegroundColor Red
            Write-Host $errorBody
        } catch {}

        Write-Host ""
        switch ($statusCode) {
            400 { Write-Host "→ リクエスト形式エラー: パラメータを確認してください" -ForegroundColor Yellow }
            401 { Write-Host "→ 認証エラー: JP1ユーザー名またはパスワードを確認してください" -ForegroundColor Yellow }
            403 { Write-Host "→ 権限エラー: JP1ユーザーに必要な権限があるか確認してください" -ForegroundColor Yellow }
            404 { Write-Host "→ 指定したユニットパスが見つかりません" -ForegroundColor Yellow }
            500 { Write-Host "→ サーバー内部エラー: Web Consoleのログを確認してください" -ForegroundColor Yellow }
        }
    }

    Write-Host ""
    Write-Host "確認事項:" -ForegroundColor Yellow
    Write-Host "  - Web Consoleが起動しているか"
    Write-Host "    → ブラウザで http://${webConsoleHost}:${webConsolePort}/ajs/login.html にアクセス"
    Write-Host "  - ホスト名・ポート番号が正しいか"
    Write-Host "  - JP1ユーザー名・パスワードが正しいか"
    Write-Host "  - ユニットパスが正しいか（/で始まる）"

    exit 1
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "[完了] 処理が正常に終了しました" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green

exit 0
