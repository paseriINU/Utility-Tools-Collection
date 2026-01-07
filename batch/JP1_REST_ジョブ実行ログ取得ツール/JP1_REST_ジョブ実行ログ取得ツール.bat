<# :
@echo off
setlocal
chcp 932 >nul

rem ============================================================================
rem バッチファイル部分（PowerShellを起動するためのラッパー）
rem ============================================================================
rem このファイルはバッチファイルとPowerShellスクリプトの両方として動作します。
rem 引数にユニットパスを指定して実行してください。
rem ============================================================================

rem 引数チェック: 引数が空の場合はエラーコード1で終了
if "%~1"=="" exit /b 1

rem 第1引数（ジョブパス）を環境変数に設定（PowerShellに渡すため）
set "JP1_UNIT_PATH=%~1"

rem UNCパス対応（PushD/PopDで自動マッピング）
pushd "%~dp0"

rem PowerShellを起動し、このファイル自体をスクリプトとして実行
rem -NoProfile: プロファイルを読み込まない（高速化）
rem -ExecutionPolicy Bypass: 実行ポリシーを回避
rem gc '%~f0': このファイル自体を読み込む
rem iex: 読み込んだ内容をPowerShellスクリプトとして実行
powershell -NoProfile -ExecutionPolicy Bypass -Command "$scriptDir=('%~dp0' -replace '\\$',''); try { iex ((gc '%~f0' -Encoding Default) -join \"`n\") } finally { Set-Location C:\ }"
set "EXITCODE=%ERRORLEVEL%"

popd

rem PowerShellの終了コードをそのまま返す
exit /b %EXITCODE%
: #>

# ==============================================================================
# JP1 REST API ジョブ実行ログ取得ツール
#
# 説明:
#   JP1/AJS3 Web Console REST APIを使用して、ジョブネットを即時実行し、
#   完了後にジョブの実行結果詳細を取得します。
#   ※ JP1/AJS3 - Web Consoleが必要です
#
# 使い方:
#   ジョブのパスを指定して実行します:
#     JP1_REST_ジョブ実行ログ取得ツール.bat "/JobGroup/Jobnet/Job1"
#
#   ※ ルートジョブネットを自動特定して即時実行し、指定ジョブの完了を待ちます
#
# 処理フロー:
#   STEP 1: DEFINITION で存在確認・ユニット種別確認
#   STEP 2: ルートジョブネットを特定
#   STEP 3: 即時実行登録API でルートジョブネットを実行
#   STEP 4: 状態監視（指定ジョブの完了待ち）
#   STEP 5: 実行結果詳細取得
#
# 終了コード（実行順）:
#   0: 正常終了
#   1: 引数エラー（ユニットパスが指定されていません）
#   2: ユニット未検出（STEP 1: 指定したユニットが存在しません）
#   3: ユニット種別エラー（STEP 1: 指定したユニットがジョブではありません）
#   4: ルートジョブネット特定エラー（STEP 2）
#   5: 即時実行登録エラー（STEP 3: API呼び出し失敗）
#   6: タイムアウト（STEP 4: 指定時間内に完了しませんでした）
#   7: 5MB超過エラー（STEP 5: 実行結果が切り捨てられました）
#   8: 詳細取得エラー（STEP 5: 実行結果詳細の取得に失敗）
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
# このユーザーには、対象ユニットへの参照権限と実行権限が必要です。
# ★ 空欄の場合は資格情報マネージャーから取得します
$jp1User = ""

# JP1パスワードを指定します。
# ★ 空欄の場合は資格情報マネージャー → 入力プロンプトの順で取得します
# ★ セキュリティのため、空欄にしてWindows資格情報マネージャーの使用を推奨
$jp1Password = ""

# Windows資格情報マネージャーのターゲット名
# 事前に以下のコマンドで登録してください:
#   cmdkey /generic:JP1_WebConsole /user:jp1admin /pass:yourpassword
# または「資格情報マネージャー」（コントロールパネル）からGUIで登録
$credentialTarget = "JP1_WebConsole"

# ==============================================================================
# ■ 実行設定セクション
# ==============================================================================

# ------------------------------------------------------------------------------
# タイムアウト設定
# ------------------------------------------------------------------------------
# ジョブ完了待ちのタイムアウト時間（秒）
# ★ デフォルト: 3600秒（60分）
$timeoutSeconds = 3600

# ステータス確認間隔（秒）
# ★ デフォルト: 5秒
$pollIntervalSeconds = 5

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
# Windows資格情報マネージャーからの認証情報取得
# ------------------------------------------------------------------------------
# ユーザー名またはパスワードが空の場合、資格情報マネージャーから取得を試みます

if (-not $jp1User -or -not $jp1Password) {
    # Windows API (CredRead) を使用して資格情報を取得
    Add-Type -TypeDefinition @"
        using System;
        using System.Runtime.InteropServices;
        public class CredManager {
            [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
            public static extern bool CredRead(string target, int type, int reservedFlag, out IntPtr credentialPtr);
            [DllImport("advapi32.dll", SetLastError = true)]
            public static extern bool CredFree(IntPtr credential);
            [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
            public struct CREDENTIAL {
                public int Flags;
                public int Type;
                public string TargetName;
                public string Comment;
                public System.Runtime.InteropServices.ComTypes.FILETIME LastWritten;
                public int CredentialBlobSize;
                public IntPtr CredentialBlob;
                public int Persist;
                public int AttributeCount;
                public IntPtr Attributes;
                public string TargetAlias;
                public string UserName;
            }
            public static string[] GetCredential(string target) {
                IntPtr credPtr;
                if (CredRead(target, 1, 0, out credPtr)) {
                    CREDENTIAL cred = (CREDENTIAL)Marshal.PtrToStructure(credPtr, typeof(CREDENTIAL));
                    string password = cred.CredentialBlobSize > 0 ? Marshal.PtrToStringUni(cred.CredentialBlob, cred.CredentialBlobSize / 2) : "";
                    string userName = cred.UserName ?? "";
                    CredFree(credPtr);
                    return new string[] { userName, password };
                }
                return null;
            }
        }
"@
    $storedCred = [CredManager]::GetCredential($credentialTarget)
    if ($storedCred) {
        if (-not $jp1User) { $jp1User = $storedCred[0] }
        if (-not $jp1Password) { $jp1Password = $storedCred[1] }
    }
}

# それでもパスワードが空の場合は入力プロンプトを表示
if (-not $jp1User) {
    $jp1User = Read-Host "JP1ユーザー名を入力してください"
}
if (-not $jp1Password) {
    $securePass = Read-Host "JP1パスワードを入力してください" -AsSecureString
    $jp1Password = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePass))
}

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
# Content-Type: POSTリクエスト用（即時実行登録APIで必要）
$headers = @{
    "Accept-Language" = "ja"
    "X-AJS-Authorization" = $authBase64
    "Content-Type" = "application/json"
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

# 親ジョブネット名を取得（コメント取得用）
$grandParentSlashIndex = $parentPath.LastIndexOf("/")
$grandParentPath = if ($grandParentSlashIndex -gt 0) { $parentPath.Substring(0, $grandParentSlashIndex) } else { "/" }
$jobnetName = if ($grandParentSlashIndex -ge 0) { $parentPath.Substring($grandParentSlashIndex + 1) } else { $parentPath.TrimStart("/") }

# パスとジョブ名をURLエンコード
$encodedParentPath = [System.Uri]::EscapeDataString($parentPath)
$encodedJobName = [System.Uri]::EscapeDataString($jobName)

# ==============================================================================
# STEP 1: ユニット存在確認・種別確認（DEFINITION）
# ==============================================================================
# 最初に DEFINITION のみで呼び出し、以下を確認します：
#   - 指定したユニットが存在するか
#   - 指定したユニットがジョブ（JOB系）かどうか

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
$rootJobnetName = $null
$jobnetComment = ""

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
    $rootJobnetName = $defUnit.definition.rootJobnetName

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
# STEP 2: ルートジョブネット特定
# ==============================================================================
# STEP 1で取得した rootJobnetName を使用してルートジョブネットを特定します。

if (-not $rootJobnetName) {
    exit 4  # ルートジョブネット特定エラー
}

# ==============================================================================
# STEP 2.5: 親ジョブネットのコメント取得
# ==============================================================================
# 親ジョブネットの定義を取得し、コメント（cm属性）を取得します。

$encodedGrandParentPath = [System.Uri]::EscapeDataString($grandParentPath)
$encodedJobnetName = [System.Uri]::EscapeDataString($jobnetName)

$jobnetUrl = "${baseUrl}/objects/statuses?mode=search"
$jobnetUrl += "&manager=${managerHost}"
$jobnetUrl += "&serviceName=${schedulerService}"
$jobnetUrl += "&location=${encodedGrandParentPath}"
$jobnetUrl += "&searchLowerUnits=NO"
$jobnetUrl += "&searchTarget=DEFINITION"
$jobnetUrl += "&unitName=${encodedJobnetName}"
$jobnetUrl += "&unitNameMatchMethods=EQ"

try {
    $jobnetResponse = Invoke-WebRequest -Uri $jobnetUrl -Method GET -Headers $headers -TimeoutSec 30 -UseBasicParsing

    # UTF-8文字化け対策
    $jobnetBytes = $jobnetResponse.RawContentStream.ToArray()
    $jobnetText = [System.Text.Encoding]::UTF8.GetString($jobnetBytes)
    $jobnetJson = $jobnetText | ConvertFrom-Json

    # ジョブネットのコメントを取得
    if ($jobnetJson.statuses -and $jobnetJson.statuses.Count -gt 0) {
        $jobnetDef = $jobnetJson.statuses[0].definition
        # unitComment フィールドを確認（JP1 REST APIのフィールド名）
        if ($jobnetDef.unitComment) {
            $jobnetComment = $jobnetDef.unitComment
        }
    }
} catch {
    # コメント取得失敗は無視して続行（必須ではない）
    $jobnetComment = ""
}

# ==============================================================================
# STEP 3: 即時実行登録API
# ==============================================================================
# ルートジョブネットを即時実行します。
# API: POST /ajs/api/v1/objects/definitions/{unitName}/actions/registerImmediateExec/invoke

$encodedRootJobnet = [System.Uri]::EscapeDataString($rootJobnetName)

# URLにはクエリパラメータを含めない（パラメータはボディに指定）
$execUrl = "${baseUrl}/objects/definitions/${encodedRootJobnet}/actions/registerImmediateExec/invoke"

# リクエストボディ（parametersオブジェクト内にmanager/serviceNameを指定）
$execBody = @{
    parameters = @{
        manager = $managerHost
        serviceName = $schedulerService
    }
} | ConvertTo-Json -Depth 3

$execIdFromRegister = $null

try {
    # POSTリクエスト（パラメータをボディに含める）
    $execResponse = Invoke-WebRequest -Uri $execUrl -Method POST -Headers $headers -Body $execBody -TimeoutSec 30 -UseBasicParsing

    # UTF-8文字化け対策
    $execBytes = $execResponse.RawContentStream.ToArray()
    $execText = [System.Text.Encoding]::UTF8.GetString($execBytes)
    $execJson = $execText | ConvertFrom-Json

    # execIDを取得
    $execIdFromRegister = $execJson.execID

    if (-not $execIdFromRegister) {
        exit 5  # 即時実行登録エラー（execIDが取得できない）
    }

} catch {
    exit 5  # 即時実行登録エラー
}

# ==============================================================================
# STEP 4: 状態監視（指定ジョブの完了待ち）
# ==============================================================================
# 指定したジョブの状態を監視し、終了を待ちます。

$statusUrl = "${baseUrl}/objects/statuses?mode=search"
$statusUrl += "&manager=${managerHost}"
$statusUrl += "&serviceName=${schedulerService}"
$statusUrl += "&location=${encodedParentPath}"
$statusUrl += "&searchLowerUnits=NO"
$statusUrl += "&searchTarget=DEFINITION_AND_STATUS"
$statusUrl += "&unitName=${encodedJobName}"
$statusUrl += "&unitNameMatchMethods=EQ"
$statusUrl += "&generation=EXECID"
$statusUrl += "&execID=${execIdFromRegister}"

$startTime = Get-Date
$endTime = $startTime.AddSeconds($timeoutSeconds)
$jobStatus = $null
$jobExecId = $null
$jobStartTime = $null

# 終了状態の定義
$finishedStatuses = @(
    "NORMAL", "NORMALFALSE", "WARNING", "ABNORMAL",
    "INTERRUPT", "KILL", "FAIL", "UNKNOWN",
    "MONITORCLOSE", "UNEXECMONITOR", "MONITORINTRPT", "MONITORNORMAL",
    "UNEXEC", "BYPASS", "EXECDEFFER", "INVALIDSEQ", "SHUTDOWN"
)

while ((Get-Date) -lt $endTime) {
    try {
        $pollResponse = Invoke-WebRequest -Uri $statusUrl -Method GET -Headers $headers -TimeoutSec 30 -UseBasicParsing

        # UTF-8文字化け対策
        $pollBytes = $pollResponse.RawContentStream.ToArray()
        $pollText = [System.Text.Encoding]::UTF8.GetString($pollBytes)
        $pollJson = $pollText | ConvertFrom-Json

        # レスポンスからジョブ情報を抽出
        if ($pollJson.statuses -and $pollJson.statuses.Count -gt 0) {
            $unit = $pollJson.statuses[0]
            $unitStatus = $unit.unitStatus

            if ($unitStatus) {
                $jobStatus = $unitStatus.status
                $jobExecId = $unitStatus.execID
                $jobStartTime = $unitStatus.startTime

                # 終了状態かどうかをチェック
                if ($finishedStatuses -contains $jobStatus) {
                    break  # 終了状態になったのでループを抜ける
                }
            }
        }

    } catch {
        # API呼び出しエラーは無視して再試行
    }

    Start-Sleep -Seconds $pollIntervalSeconds
}

# タイムアウトチェック
if ((Get-Date) -ge $endTime) {
    if (-not ($finishedStatuses -contains $jobStatus)) {
        exit 6  # タイムアウト
    }
}

# 開始日時をファイル名用フォーマットに変換（yyyyMMdd_HHmmss）
$startTimeForFileName = ""
if ($jobStartTime) {
    try {
        $dt = [DateTime]::Parse($jobStartTime)
        $startTimeForFileName = $dt.ToString("yyyyMMdd_HHmmss")
    } catch {
        $startTimeForFileName = ""
    }
}

# ==============================================================================
# STEP 5: 実行結果詳細取得API
# ==============================================================================
# ジョブの実行結果詳細を取得します。

$encodedPath = [System.Uri]::EscapeDataString($unitPath)

# 実行結果詳細取得APIのURL構築
$detailUrl = "${baseUrl}/objects/statuses/${encodedPath}:${jobExecId}/actions/execResultDetails/invoke"
$detailUrl += "?manager=${managerHost}&serviceName=${schedulerService}"

try {
    # APIリクエストを送信
    $resultResponse = Invoke-WebRequest -Uri $detailUrl -Method GET -Headers $headers -TimeoutSec 30 -UseBasicParsing

    # UTF-8文字化け対策
    $resultBytes = $resultResponse.RawContentStream.ToArray()
    $resultText = [System.Text.Encoding]::UTF8.GetString($resultBytes)
    $resultJson = $resultText | ConvertFrom-Json

    # all フラグのチェック（falseの場合は5MB超過で切り捨て）
    if ($resultJson.all -eq $false) { exit 7 }  # 5MB超過エラー

    # 実行結果の内容を取得
    $execResultContent = ""
    if ($resultJson.execResultDetails) {
        $execResultContent = $resultJson.execResultDetails
    }

    # 開始日時を最初の行に出力（ファイル名用フォーマット）
    [Console]::WriteLine("START_TIME:$startTimeForFileName")

    # ジョブネット名を出力（ファイル名用）
    [Console]::WriteLine("JOBNET_NAME:$jobnetName")

    # ジョブネットコメントを出力（ファイル名用）
    [Console]::WriteLine("JOBNET_COMMENT:$jobnetComment")

    # ジョブ終了ステータスを出力
    [Console]::WriteLine("JOB_STATUS:$jobStatus")

    # 実行結果詳細を出力
    [Console]::WriteLine($execResultContent)
} catch {
    exit 8  # 詳細取得エラー
}

# 正常終了
exit 0
