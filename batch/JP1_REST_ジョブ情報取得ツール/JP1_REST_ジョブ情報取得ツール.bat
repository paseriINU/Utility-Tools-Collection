<# :
@echo off
setlocal
chcp 932 >nul

rem ============================================================================
rem バッチファイル部分（PowerShellを起動するためのラッパー）
rem ============================================================================
rem このファイルはバッチファイルとPowerShellスクリプトの両方として動作します。
rem ダブルクリックまたはコマンドラインから実行すると、以下の処理を行います。
rem ============================================================================

rem 引数チェック: 引数が空の場合はエラーコード1で終了
if "%~1"=="" exit /b 1

rem 第1引数（ジョブパス）を環境変数に設定（PowerShellに渡すため）
set "JP1_UNIT_PATH=%~1"

rem PowerShellを起動し、このファイル自体をスクリプトとして実行
rem -NoProfile: プロファイルを読み込まない（高速化）
rem -ExecutionPolicy Bypass: 実行ポリシーを回避
rem gc '%~f0': このファイル自体を読み込む
rem iex: 読み込んだ内容をPowerShellスクリプトとして実行
powershell -NoProfile -ExecutionPolicy Bypass -Command "$scriptDir=('%~dp0' -replace '\\$',''); iex ((gc '%~f0' -Encoding Default) -join \"`n\")"

rem PowerShellの終了コードをそのまま返す
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
#   ジョブのパスを指定して実行します:
#     JP1_REST_ジョブ情報取得ツール.bat "/JobGroup/Jobnet/Job1"
#
#   ※ 親パス（ジョブネット）を検索対象とし、ジョブ名でフィルタします
#
# 処理フロー:
#   STEP 1: DEFINITION で存在確認・ユニット種別確認
#   STEP 2: DEFINITION_AND_STATUS で execID 取得
#   STEP 3: 実行結果詳細取得
#
# 終了コード（実行順）:
#   0: 正常終了
#   1: 引数エラー（ユニットパスが指定されていません）
#   2: ユニット未検出（STEP 1: 指定したユニットが存在しません）
#   3: ユニット種別エラー（STEP 1: 指定したユニットがジョブではありません）
#   4: 実行世代なし（STEP 2: 実行履歴が存在しません）
#   5: 5MB超過エラー（STEP 3: 実行結果が切り捨てられました）
#   6: 詳細取得エラー（STEP 3: 実行結果詳細の取得に失敗）
#   9: API接続エラー（各STEPでの接続失敗）
#
# 参考:
#   https://itpfdoc.hitachi.co.jp/manuals/3021/30213b1920/AJSO0280.HTM
# ==============================================================================

# ------------------------------------------------------------------------------
# 出力エンコーディング設定
# ------------------------------------------------------------------------------
# 出力をShift-JIS（コードページ932）に設定します。
# これにより、日本語Windowsのコマンドプロンプトで正しく表示されます。
[Console]::OutputEncoding = [System.Text.Encoding]::GetEncoding(932)

# ==============================================================================
# ■ 接続設定セクション
# ==============================================================================
# このセクションでは、JP1/AJS3 Web Consoleへの接続情報を設定します。

# ------------------------------------------------------------------------------
# Web Consoleサーバー設定
# ------------------------------------------------------------------------------
# Web Consoleサーバーのホスト名またはIPアドレスを指定します。
# 例: "localhost", "192.168.1.100", "jp1server.example.com"
$webConsoleHost = "localhost"

# Web Consoleのポート番号を指定します。
# ・HTTP接続の場合: 22252（デフォルト）
# ・HTTPS接続の場合: 22253（デフォルト）
$webConsolePort = "22252"

# HTTPS（暗号化通信）を使用する場合は $true に設定します。
# ・$false: HTTP接続（暗号化なし、社内ネットワーク向け）
# ・$true:  HTTPS接続（暗号化あり、セキュリティ重視）
$useHttps = $false

# ------------------------------------------------------------------------------
# JP1/AJS3 Manager設定
# ------------------------------------------------------------------------------
# JP1/AJS3 Managerのホスト名またはIPアドレスを指定します。
# Web Consoleサーバーと同じマシンの場合は "localhost" でOKです。
$managerHost = "localhost"

# スケジューラーサービス名を指定します。
# デフォルトは "AJSROOT1" です。複数サービスがある場合は適宜変更してください。
# 例: "AJSROOT1", "AJSROOT2", "SCHEDULE_SERVICE"
$schedulerService = "AJSROOT1"

# ------------------------------------------------------------------------------
# 認証設定
# ------------------------------------------------------------------------------
# JP1ユーザー名を指定します。
# このユーザーには、対象ユニットへの参照権限が必要です。
$jp1User = "jp1admin"

# JP1パスワードを指定します。
# ★★★ セキュリティ注意 ★★★
# パスワードは平文で保存されます。本番環境では取り扱いに注意してください。
$jp1Password = "password"

# ==============================================================================
# ■ 検索条件設定セクション
# ==============================================================================
# このセクションでは、ユニット一覧取得APIの検索条件を設定します。

# ------------------------------------------------------------------------------
# (1) SearchTargetType - 取得情報範囲
# ------------------------------------------------------------------------------
# 取得する情報の範囲を指定します。
#
# 指定可能な値:
#   "DEFINITION"            - ユニットの定義情報のみを取得します
#                             （実行状態は取得しない、高速）
#   "DEFINITION_AND_STATUS" - ユニットの定義情報と実行状態を両方取得します
#                             （execIDを取得するにはこちらが必要）
#
# ★ 実行結果詳細を取得する場合は "DEFINITION_AND_STATUS" が必要です
$searchTarget = "DEFINITION_AND_STATUS"

# ------------------------------------------------------------------------------
# (2) GenerationType - 世代指定
# ------------------------------------------------------------------------------
# 取得するユニットの世代を指定します。
#
# 指定可能な値:
#   "NO"     - 世代を検索条件にしません
#   "STATUS" - 最新状態の世代を取得します
#              （VIEWSTATUSRANGEの設定値に従う）
#   "RESULT" - 最新結果の世代を取得します（★推奨★）
#              （終了済みジョブの直近終了世代を取得）
#   "PERIOD" - 指定した期間に存在する世代を取得します
#              （periodBegin/periodEnd の設定が必要）
#   "EXECID" - 指定した実行IDの世代を取得します
#              （execID パラメータの設定が必要）
#
# ★ 通常は "RESULT" を使用することで、終了済みジョブの結果を取得できます
$generation = "RESULT"

# 期間指定（generation="PERIOD" の場合に使用）
# 形式: YYYY-MM-DDThh:mm（ISO 8601形式）
# 例: "2025-01-01T00:00" ～ "2025-01-31T23:59"
$periodBegin = "2025-12-01T00:00"
$periodEnd = "2025-12-25T23:59"

# 実行ID指定（generation="EXECID" の場合に使用）
# 形式: @[mmmm]{A～Z}nnnn（例: @A100, @10A200）
$execID = ""

# ------------------------------------------------------------------------------
# (3) UnitStatus - ユニット状態フィルタ
# ------------------------------------------------------------------------------
# 取得するユニットの状態を指定します。
#
# 【個別状態】
#   "NO"             - ユニット状態を検索条件にしません（すべて取得）
#   "UNREGISTERED"   - 未登録
#   "NOPLAN"         - 未計画
#   "UNEXEC"         - 未実行終了
#   "BYPASS"         - 計画未実行
#   "EXECDEFFER"     - 繰越未実行
#   "SHUTDOWN"       - 閉塞
#   "TIMEWAIT"       - 開始時刻待ち
#   "TERMWAIT"       - 先行終了待ち
#   "EXECWAIT"       - 実行待ち
#   "QUEUING"        - キューイング
#   "CONDITIONWAIT"  - 起動条件待ち
#   "HOLDING"        - 保留中
#   "RUNNING"        - 実行中
#   "WACONT"         - 警告検出実行中
#   "ABCONT"         - 異常検出実行中
#   "MONITORING"     - 監視中
#   "ABNORMAL"       - 異常検出終了（★エラー調査時に便利★）
#   "INVALIDSEQ"     - 順序不正
#   "INTERRUPT"      - 中断
#   "KILL"           - 強制終了
#   "FAIL"           - 起動失敗
#   "UNKNOWN"        - 終了状態不正
#   "MONITORCLOSE"   - 監視打ち切り終了
#   "WARNING"        - 警告検出終了
#   "NORMAL"         - 正常終了
#   "NORMALFALSE"    - 正常終了-偽
#   "UNEXECMONITOR"  - 監視未起動終了
#   "MONITORINTRPT"  - 監視中断
#   "MONITORNORMAL"  - 監視正常終了
#
# 【グループ状態】（複数の状態をまとめて指定）
#   "GRP_WAIT"     - 待ち状態（開始時刻待ち、先行終了待ち、実行待ち、キューイング、起動条件待ち）
#   "GRP_RUN"      - 実行中状態（実行中、警告検出実行中、異常検出実行中、監視中）
#   "GRP_ABNORMAL" - 異常終了状態（異常検出終了、順序不正、中断、強制終了、起動失敗、終了状態不明、監視打ち切り終了）
#   "GRP_NORMAL"   - 正常終了状態（正常終了、正常終了-偽、監視未起動終了、監視中断、監視正常終了）
#
# ★ 空欄または "NO" で全件取得
# ★ エラー調査時は "ABNORMAL" や "GRP_ABNORMAL" が便利
$statusFilter = "NO"

# ------------------------------------------------------------------------------
# (4) DelayType - 遅延状態フィルタ
# ------------------------------------------------------------------------------
# 開始遅延または終了遅延の有無でフィルタします。
#
# 指定可能な値:
#   "NO"    - 遅延状態を検索条件にしません（すべて取得）
#   "START" - 開始遅延のあるユニットのみ取得
#   "END"   - 終了遅延のあるユニットのみ取得
#   "YES"   - 開始遅延または終了遅延のあるユニットを取得
$delayStatus = "NO"

# ------------------------------------------------------------------------------
# (5) HoldPlan - 保留予定フィルタ
# ------------------------------------------------------------------------------
# 保留予定の有無でフィルタします。
#
# 指定可能な値:
#   "NO"        - 保留予定を検索条件にしません（すべて取得）
#   "PLAN_NONE" - 保留予定のないユニットのみ取得
#   "PLAN_YES"  - 保留予定のあるユニットのみ取得
$holdPlan = "NO"

# ==============================================================================
# ■ メイン処理（以下は通常編集不要）
# ==============================================================================

# ------------------------------------------------------------------------------
# 環境変数からユニットパスを取得
# ------------------------------------------------------------------------------
# バッチファイルから渡された引数（ユニットパス）を取得します
$unitPath = $env:JP1_UNIT_PATH

# ------------------------------------------------------------------------------
# プロトコル設定
# ------------------------------------------------------------------------------
# HTTPS使用フラグに基づいて、接続プロトコルを決定します
$protocol = if ($useHttps) { "https" } else { "http" }

# ------------------------------------------------------------------------------
# 認証情報の作成
# ------------------------------------------------------------------------------
# JP1/AJS3 Web Console REST APIは、X-AJS-Authorizationヘッダーで認証します。
# 形式: Base64エンコードした "{ユーザー名}:{パスワード}"
$authString = "${jp1User}:${jp1Password}"
$authBytes = [System.Text.Encoding]::UTF8.GetBytes($authString)
$authBase64 = [System.Convert]::ToBase64String($authBytes)

# ------------------------------------------------------------------------------
# HTTPリクエストヘッダーの設定
# ------------------------------------------------------------------------------
# Accept-Language: 日本語でレスポンスを受け取る
# X-AJS-Authorization: 認証情報（Base64エンコード）
$headers = @{
    "Accept-Language" = "ja"
    "X-AJS-Authorization" = $authBase64
}

# ------------------------------------------------------------------------------
# SSL証明書検証の設定（HTTPS使用時）
# ------------------------------------------------------------------------------
# 自己署名証明書を使用している環境では、証明書検証をスキップします。
# ★ 本番環境では正規の証明書を使用することを推奨します
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

# ベースURLの構築
$baseUrl = "${protocol}://${webConsoleHost}:${webConsolePort}/ajs/api/v1"

# ------------------------------------------------------------------------------
# ジョブパスの解析
# ------------------------------------------------------------------------------
# 指定されたジョブパスから親パス（ジョブネット）とジョブ名を分離
# 例: "/JobGroup/Jobnet/Job1" → 親: "/JobGroup/Jobnet", ジョブ名: "Job1"
$lastSlashIndex = $unitPath.LastIndexOf("/")
if ($lastSlashIndex -le 0) {
    exit 1  # パス形式エラー（スラッシュがない、またはルートのみ）
}
$parentPath = $unitPath.Substring(0, $lastSlashIndex)
$jobName = $unitPath.Substring($lastSlashIndex + 1)

if (-not $jobName) {
    exit 1  # ジョブ名が空
}

# パスとジョブ名をURLエンコード
$encodedParentPath = [System.Uri]::EscapeDataString($parentPath)
$encodedJobName = [System.Uri]::EscapeDataString($jobName)

# ==============================================================================
# STEP 1: ユニット存在確認・種別確認（DEFINITION）
# ==============================================================================
# 最初に DEFINITION のみで呼び出し、以下を確認します：
#   - 指定したユニットが存在するか
#   - 指定したユニットがジョブ（JOB系）かどうか
#   - 親ユニット（ジョブネット）の情報

$defUrl = "${baseUrl}/objects/statuses?mode=search"
$defUrl += "&manager=${managerHost}"
$defUrl += "&serviceName=${schedulerService}"
$defUrl += "&location=${encodedParentPath}"
$defUrl += "&searchLowerUnits=NO"
$defUrl += "&searchTarget=DEFINITION"
$defUrl += "&unitName=${encodedJobName}"
$defUrl += "&unitNameMatchMethods=EQ"

$unitTypeValue = $null
$unitFullName = $null

try {
    $defResponse = Invoke-WebRequest -Uri $defUrl -Method GET -Headers $headers -TimeoutSec 30 -UseBasicParsing

    # UTF-8文字化け対策
    $defBytes = $defResponse.RawContentStream.ToArray()
    $defText = [System.Text.Encoding]::UTF8.GetString($defBytes)
    $defJson = $defText | ConvertFrom-Json

    # ユニット存在確認
    if (-not $defJson.statuses -or $defJson.statuses.Count -eq 0) {
        exit 2  # ユニット未検出
    }

    # 最初のユニットの情報を取得
    $defUnit = $defJson.statuses[0]
    $unitFullName = $defUnit.definition.unitName
    $unitTypeValue = $defUnit.definition.unitType

    # ユニット種別確認（JOB系かどうか）
    # JOB系: JOB, PJOB, QJOB, EVWJB, FLWJB, MLWJB, MSWJB, LFWJB, TMWJB,
    #        EVSJB, MLSJB, MSSJB, PWLJB, PWRJB, CJOB, HTPJOB, CPJOB, FXJOB, CUSTOM, JDJOB, ORJOB
    if ($unitTypeValue -notmatch "JOB") {
        exit 3  # ユニット種別エラー（ジョブではない）
    }

} catch {
    exit 9  # API接続エラー（存在確認）
}

# ==============================================================================
# STEP 2: 実行状態・execID取得（DEFINITION_AND_STATUS）
# ==============================================================================
# 存在確認・種別確認が成功したら、DEFINITION_AND_STATUS で execID を取得します。

$statusUrl = "${baseUrl}/objects/statuses?mode=search"
$statusUrl += "&manager=${managerHost}"
$statusUrl += "&serviceName=${schedulerService}"
$statusUrl += "&location=${encodedParentPath}"
$statusUrl += "&searchLowerUnits=NO"
$statusUrl += "&searchTarget=${searchTarget}"
$statusUrl += "&unitName=${encodedJobName}"
$statusUrl += "&unitNameMatchMethods=EQ"

# 世代指定
$statusUrl += "&generation=${generation}"

# 期間指定（generation=PERIOD の場合）
if ($generation -eq "PERIOD") {
    $statusUrl += "&periodBegin=${periodBegin}"
    $statusUrl += "&periodEnd=${periodEnd}"
}

# 実行ID指定（generation=EXECID の場合）
if ($generation -eq "EXECID" -and $execID) {
    $statusUrl += "&execID=${execID}"
}

# ステータスフィルタ
if ($statusFilter -and $statusFilter -ne "NO") {
    $statusUrl += "&status=${statusFilter}"
}

# 遅延状態フィルタ
if ($delayStatus -and $delayStatus -ne "NO") {
    $statusUrl += "&delayStatus=${delayStatus}"
}

# 保留予定フィルタ
if ($holdPlan -and $holdPlan -ne "NO") {
    $statusUrl += "&holdPlan=${holdPlan}"
}

$execIdList = @()

try {
    $response = Invoke-WebRequest -Uri $statusUrl -Method GET -Headers $headers -TimeoutSec 30 -UseBasicParsing

    # UTF-8文字化け対策
    $responseBytes = $response.RawContentStream.ToArray()
    $responseText = [System.Text.Encoding]::UTF8.GetString($responseBytes)
    $jsonData = $responseText | ConvertFrom-Json

    # レスポンスからユニット情報を抽出
    if ($jsonData.statuses -and $jsonData.statuses.Count -gt 0) {
        foreach ($unit in $jsonData.statuses) {
            # ユニット定義情報
            $unitFullName = $unit.definition.unitName
            $unitTypeValue = $unit.definition.unitType

            # ユニット状態情報
            $unitStatus = $unit.unitStatus
            $execIdValue = if ($unitStatus) { $unitStatus.execID } else { $null }
            $statusValue = if ($unitStatus) { $unitStatus.status } else { "N/A" }

            # execIDがある場合のみリストに追加
            if ($execIdValue) {
                $execIdList += @{
                    Path = $unitFullName
                    ExecId = $execIdValue
                    Status = $statusValue
                    UnitType = $unitTypeValue
                }
            }
        }
    }

    # 実行世代が存在しない場合
    if ($execIdList.Count -eq 0) {
        exit 4  # 実行世代なし
    }

} catch {
    exit 9  # API接続エラー（状態取得）
}

# ==============================================================================
# STEP 3: 実行結果詳細取得API
# ==============================================================================
# STEP 2で取得した各ジョブについて、実行結果詳細を取得します。
# 実行結果詳細には、標準出力・標準エラー出力の内容が含まれます。

if ($execIdList.Count -gt 0) {
    foreach ($item in $execIdList) {
        $targetPath = $item.Path
        $targetExecId = $item.ExecId

        # ユニットパスをURLエンコード
        $encodedPath = [System.Uri]::EscapeDataString($targetPath)

        # 実行結果詳細取得APIのURL構築
        # 形式: /objects/statuses/{unitName}:{execID}/actions/execResultDetails/invoke
        $detailUrl = "${baseUrl}/objects/statuses/${encodedPath}:${targetExecId}/actions/execResultDetails/invoke"
        $detailUrl += "?manager=${managerHost}&serviceName=${schedulerService}"

        try {
            # APIリクエストを送信
            $resultResponse = Invoke-WebRequest -Uri $detailUrl -Method GET -Headers $headers -TimeoutSec 30 -UseBasicParsing

            # UTF-8文字化け対策
            $resultBytes = $resultResponse.RawContentStream.ToArray()
            $resultText = [System.Text.Encoding]::UTF8.GetString($resultBytes)
            $resultJson = $resultText | ConvertFrom-Json

            # all フラグのチェック（falseの場合は5MB超過で切り捨て）
            if ($resultJson.all -eq $false) { exit 5 }  # 5MB超過エラー

            # 実行結果詳細を出力
            if ($resultJson.execResultDetails) {
                [Console]::WriteLine($resultJson.execResultDetails)
            }
        } catch {
            exit 6  # 詳細取得エラー
        }
    }
}

# 正常終了
exit 0
