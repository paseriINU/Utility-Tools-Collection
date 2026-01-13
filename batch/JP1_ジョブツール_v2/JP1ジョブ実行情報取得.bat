<# :
@echo off
setlocal
chcp 932 >nul

rem ============================================================================
rem バッチファイル部分（PowerShellを起動するためのラッパー）
rem ============================================================================
rem このファイルはバッチファイルとPowerShellスクリプトの両方として動作します。
rem 引数にユニットパスを指定して実行してください。
rem
rem 使い方:
rem   JP1ジョブ実行情報取得.bat "ジョブパス"
rem
rem 出力オプション（環境変数 JP1_OUTPUT_MODE で指定、必須）:
rem   /NOTEPAD  - メモ帳で開く
rem   /EXCEL    - Excelに貼り付け
rem   /WINMERGE - WinMergeで比較
rem ============================================================================

rem 引数チェック: 引数が空の場合はエラーコード1で終了
if "%~1"=="" exit /b 1

rem 第1引数（ジョブパス）を環境変数に設定（PowerShellに渡すため）
set "JP1_UNIT_PATH=%~1"

rem 出力オプションは環境変数 JP1_OUTPUT_MODE から取得（必須）
if "%JP1_OUTPUT_MODE%"=="" exit /b 1

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
#     JP1ジョブ実行情報取得.bat "/JobGroup/Jobnet/Job1"
#
#   ※ ルートジョブネットを自動特定して即時実行し、指定ジョブの完了を待ちます
#
# 処理フロー:
#   STEP 1: DEFINITION で存在確認・ユニット種別確認
#   STEP 2: ルートジョブネットを特定
#   STEP 3: 親ジョブネットのコメント取得
#   STEP 4: 即時実行登録API でルートジョブネットを実行
#   STEP 5: 状態監視（指定ジョブの完了待ち）
#   STEP 6: 実行結果詳細取得
#
# 終了コード（実行順）:
#   0: 正常終了
#   1: 引数エラー（ユニットパスが指定されていません）
#   2: ユニット未検出（STEP 1: 指定したユニットが存在しません）
#   3: ユニット種別エラー（STEP 1: 指定したユニットがジョブではありません）
#   4: ルートジョブネット特定エラー（STEP 2）
#   5: 即時実行登録エラー（STEP 4: API呼び出し失敗）
#   6: タイムアウト（STEP 5: 指定時間内に完了しませんでした）
#   7: 5MB超過エラー（STEP 6: 実行結果が切り捨てられました）
#   8: 詳細取得エラー（STEP 6: 実行結果詳細の取得に失敗）
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

# コンソールに直接出力する関数（ファイルリダイレクトの影響を受けない）
function Write-Console {
    param([string]$Message)
    [Console]::WriteLine($Message)
}

# タイトル表示
Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  JP1ジョブ実行情報取得ツール" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

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

# ------------------------------------------------------------------------------
# Excel貼り付け設定（/EXCEL オプション使用時）
# ------------------------------------------------------------------------------
# 雛形フォルダ名（スクリプトと同じフォルダに配置）
# このフォルダがyyyymmddフォルダにコピーされます
$templateFolderName = "【雛形】【コピーして使うこと！】ツール・手順書"

# 出力先フォルダ名（親フォルダの02_outputに出力）
$outputFolderName = "..\02_output"

# ジョブパスとExcelファイルの紐づけ設定
# ★ キー: ジョブのフルパス（完全一致で検索）
# ★ 値: Excelファイル名、シート名、貼り付けセルをカンマ区切りで指定
# 例: "/JobGroup/Jobnet/Job1" = "ファイル名.xlsx,Sheet1,A1"
$jobExcelMapping = @{
    # === ジョブパスとExcelファイルのマッピング ===
    # 以下に「ジョブのフルパス」=「Excelファイル名,シート名,セル」の形式で記載してください
    # 完全一致で検索されるため、ジョブの正確なフルパスを指定してください
    #
    # 例:
    # "/AJSROOT1/TIA/Jobnet/週単位ジョブ" = "TIA解析(自習当初)_週単位.xls,Sheet1,A1"
    # "/AJSROOT1/TIA/Jobnet/年単位ジョブ" = "TIA解析(自習当初)_年単位.xls,Sheet1,A1"
    #
    # ★ 以下を編集してください ★
    "/AJSROOT1/サンプル/Jobnet/週単位ジョブ" = "TIA解析(自習当初)_週単位.xls,Sheet1,A1"
    "/AJSROOT1/サンプル/Jobnet/年単位ジョブ" = "TIA解析(自習当初)_年単位.xls,Sheet1,A1"
    # "/AJSROOT1/グループ/ネット/ジョブ3" = "Excelファイル3.xls,Sheet1,A1"
    # "/AJSROOT1/グループ/ネット/ジョブ4" = "Excelファイル4.xls,Sheet1,A1"
    # "/AJSROOT1/グループ/ネット/ジョブ5" = "Excelファイル5.xls,Sheet1,A1"
    # "/AJSROOT1/グループ/ネット/ジョブ6" = "Excelファイル6.xls,Sheet1,A1"
}

# ジョブパスとテキストファイル名の紐づけ設定（クリップボード保存用）
# ★ キー: ジョブのフルパス（完全一致で検索）
# ★ 値: 保存するテキストファイル名
# 例: "/AJSROOT1/JobGroup/Jobnet/Job1" = "runh_week.txt"
$jobTextFileMapping = @{
    # === ジョブパスとテキストファイルのマッピング ===
    # Excelに貼り付けた後、クリップボード内容を保存するファイル名を指定します
    # 完全一致で検索されるため、ジョブの正確なフルパスを指定してください
    #
    # ★ 以下を編集してください ★
    "/AJSROOT1/サンプル/Jobnet/週単位ジョブ" = "runh_week.txt"
    "/AJSROOT1/サンプル/Jobnet/年単位ジョブ" = "runh_year.txt"
    # "/AJSROOT1/グループ/ネット/ジョブ3" = "runh_file3.txt"
    # "/AJSROOT1/グループ/ネット/ジョブ4" = "runh_file4.txt"
    # "/AJSROOT1/グループ/ネット/ジョブ5" = "runh_file5.txt"
    # "/AJSROOT1/グループ/ネット/ジョブ6" = "runh_file6.txt"
}

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

Write-Console "[STEP 1] ユニット存在確認中..."

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

Write-Console "[STEP 2] ルートジョブネット特定中..."

if (-not $rootJobnetName) {
    exit 4  # ルートジョブネット特定エラー
}

# ==============================================================================
# STEP 3: 親ジョブネットのコメント取得
# ==============================================================================
# 親ジョブネットの定義を取得し、コメント（cm属性）を取得します。

Write-Console "[STEP 3] 親ジョブネットのコメント取得中..."

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
# STEP 4: 即時実行登録API
# ==============================================================================
# ルートジョブネットを即時実行します。
# API: POST /ajs/api/v1/objects/definitions/{unitName}/actions/registerImmediateExec/invoke

Write-Console "[STEP 4] 即時実行登録中..."

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
# STEP 5: 状態監視（指定ジョブの完了待ち）
# ==============================================================================
# 指定したジョブの状態を監視し、終了を待ちます。

Write-Console "[STEP 5] ジョブ完了待機中..."

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
# STEP 6: 実行結果詳細取得API
# ==============================================================================
# ジョブの実行結果詳細を取得します。

Write-Console "[STEP 6] 実行結果詳細取得中..."

# ステータス値を日本語に変換する関数
function Get-StatusDisplayName {
    param([string]$status)
    switch ($status) {
        "NORMAL"        { return "正常終了" }
        "WARNING"       { return "警告検出終了" }
        "ABNORMAL"      { return "異常検出終了" }
        "KILL"          { return "強制終了" }
        "INTERRUPT"     { return "中断" }
        "FAIL"          { return "起動失敗" }
        "UNKNOWN"       { return "終了状態不正" }
        "MONITORCLOSE"  { return "監視打ち切り終了" }
        "INVALIDSEQ"    { return "順序不正" }
        "NORMALFALSE"   { return "正常終了-偽" }
        "UNEXECMONITOR" { return "監視未起動終了" }
        "MONITORINTRPT" { return "監視中断" }
        "MONITORNORMAL" { return "監視正常終了" }
        "RUNNING"       { return "実行中" }
        "WACONT"        { return "警告検出実行中" }
        "ABCONT"        { return "異常検出実行中" }
        default         { return $status }
    }
}

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

    # 終了状態を取得（日本語変換済み）
    $endStatusDisplay = Get-StatusDisplayName -status $jobStatus

    # 出力オプションを環境変数から取得
    $outputMode = $env:JP1_OUTPUT_MODE
    if (-not $outputMode) { $outputMode = "/NOTEPAD" }

    # 出力ディレクトリを作成
    $outputDir = Join-Path $scriptDir "..\02.Output"
    if (-not (Test-Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    }

    # 出力ファイル名を生成
    $outputFileName = "【ジョブ実行結果】【${startTimeForFileName}実行分】【${endStatusDisplay}】${jobnetName}_${jobnetComment}.txt"
    $outputFilePath = Join-Path $outputDir $outputFileName

    # 実行結果詳細をファイルに出力
    $execResultContent | Out-File -FilePath $outputFilePath -Encoding Default

    # メタデータを標準出力に出力（後方互換性のため）
    [Console]::WriteLine("START_TIME:$startTimeForFileName")
    [Console]::WriteLine("END_STATUS:$endStatusDisplay")
    [Console]::WriteLine("JOBNET_NAME:$jobnetName")
    [Console]::WriteLine("JOBNET_COMMENT:$jobnetComment")
    [Console]::WriteLine("JOB_STATUS:$jobStatus")
    [Console]::WriteLine("OUTPUT_FILE:$outputFilePath")

    # 出力オプションに応じた後処理
    switch ($outputMode.ToUpper()) {
        "/NOTEPAD" {
            # メモ帳で開く
            Start-Process notepad $outputFilePath

            # スクロール位置の設定（環境変数から取得）
            $scrollToText = $env:JP1_SCROLL_TO_TEXT
            if ($scrollToText) {
                # ファイル内容を読み込んで検索
                $lines = Get-Content $outputFilePath -Encoding Default
                $scrollLineNum = 0
                for ($i = 0; $i -lt $lines.Count; $i++) {
                    if ($lines[$i] -match [regex]::Escape($scrollToText)) {
                        $scrollLineNum = $i + 1  # 1始まりの行番号
                        break
                    }
                }

                if ($scrollLineNum -gt 0) {
                    # メモ帳が起動するのを待つ
                    Start-Sleep -Milliseconds 500

                    # WScript.Shellを使用してキー入力を送信
                    $wshell = New-Object -ComObject WScript.Shell
                    # Ctrl+G（行ジャンプダイアログ）を送信
                    $wshell.SendKeys("^g")
                    Start-Sleep -Milliseconds 200
                    # 行番号を入力
                    $wshell.SendKeys($scrollLineNum.ToString())
                    Start-Sleep -Milliseconds 100
                    # Enterキーで確定
                    $wshell.SendKeys("{ENTER}")
                }
            }
        }
        "/EXCEL" {
            # Excelに貼り付け（雛形フォルダコピー + ジョブ別Excel選択）

            # --------------------------------------------------------------
            # STEP 1: ジョブパスからExcel設定を取得
            # --------------------------------------------------------------
            $excelFileName = $null
            $excelSheetName = $null
            $excelPasteCell = $null

            # ジョブパスに一致するマッピングを検索（完全一致）
            foreach ($key in $jobExcelMapping.Keys) {
                if ($unitPath -eq $key) {
                    $mappingValue = $jobExcelMapping[$key]
                    $parts = $mappingValue -split ","
                    if ($parts.Count -ge 3) {
                        $excelFileName = $parts[0].Trim()
                        $excelSheetName = $parts[1].Trim()
                        $excelPasteCell = $parts[2].Trim()
                    }
                    break
                }
            }

            # マッピングが見つからない場合はエラー
            if (-not $excelFileName) {
                Write-Console "[エラー] ジョブパス '$unitPath' に対応するExcel設定が見つかりません。"
                Write-Console "[エラー] 設定セクションの `$jobExcelMapping を確認してください。"
                exit 10  # Excel設定エラー
            }

            # ------------------------------------------------------------------
            # Excel処理開始ヘッダー
            # ------------------------------------------------------------------
            Write-Host ""
            Write-Host "================================================================" -ForegroundColor Yellow
            Write-Host "  Excel貼り付け処理" -ForegroundColor Yellow
            Write-Host "================================================================" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "  [設定情報]" -ForegroundColor Cyan
            Write-Host "    ジョブパス    : $unitPath"
            Write-Host "    Excelファイル : $excelFileName"
            Write-Host "    シート名      : $excelSheetName"
            Write-Host "    貼り付けセル  : $excelPasteCell"
            Write-Host ""

            # --------------------------------------------------------------
            # STEP 2: 02_output/yyyymmddフォルダを作成
            # --------------------------------------------------------------
            Write-Host "  [STEP 1] 出力フォルダ準備" -ForegroundColor Cyan
            $dateFolder = Get-Date -Format "yyyyMMdd"
            $outputBasePath = Join-Path $scriptDir $outputFolderName
            $dateFolderPath = Join-Path $outputBasePath $dateFolder

            # 02_outputフォルダが存在しない場合は作成
            if (-not (Test-Path $outputBasePath)) {
                New-Item -Path $outputBasePath -ItemType Directory -Force | Out-Null
                Write-Host "    出力フォルダ作成: $outputBasePath"
            }

            # yyyymmddフォルダが存在しない場合は作成
            if (-not (Test-Path $dateFolderPath)) {
                New-Item -Path $dateFolderPath -ItemType Directory -Force | Out-Null
                Write-Host "    日付フォルダ作成: $dateFolderPath"
            } else {
                Write-Host "    日付フォルダ    : $dateFolderPath (既存)"
            }
            Write-Host ""

            # --------------------------------------------------------------
            # STEP 3: 雛形フォルダの中身をコピー
            # --------------------------------------------------------------
            Write-Host "  [STEP 2] 雛形フォルダコピー" -ForegroundColor Cyan
            $templatePath = Join-Path $scriptDir $templateFolderName

            if (-not (Test-Path $templatePath)) {
                Write-Console "[エラー] 雛形フォルダが見つかりません: $templatePath"
                exit 13  # 雛形フォルダ未検出エラー
            }

            # 雛形フォルダの中身をyyyymmddフォルダに直接コピー
            # （雛形フォルダ自体はコピーせず、中のファイルのみ）
            $templateItems = Get-ChildItem -Path $templatePath
            $copiedCount = 0
            foreach ($item in $templateItems) {
                $destPath = Join-Path $dateFolderPath $item.Name
                if (-not (Test-Path $destPath)) {
                    Copy-Item -Path $item.FullName -Destination $destPath -Recurse -Force
                    $copiedCount++
                }
            }
            if ($copiedCount -gt 0) {
                Write-Host "    コピー完了: $copiedCount 件のファイル/フォルダ"
            } else {
                Write-Host "    スキップ: 既にコピー済み"
            }
            Write-Host ""

            # --------------------------------------------------------------
            # STEP 4: Excelファイルにログを貼り付け
            # --------------------------------------------------------------
            Write-Host "  [STEP 3] Excelに貼り付け" -ForegroundColor Cyan
            $excelPath = Join-Path $dateFolderPath $excelFileName

            if (-not (Test-Path $excelPath)) {
                Write-Console "[エラー] Excelファイルが見つかりません: $excelPath"
                exit 12  # Excelファイル未検出エラー
            }

            try {
                # ログファイルの内容を読み込み（-Raw: 全体を1つの文字列として）
                $logContent = Get-Content $outputFilePath -Encoding Default -Raw

                # クリップボードにコピー
                # Excelの通常の貼り付け動作を利用するため、クリップボード経由で貼り付けます
                # これにより改行区切りのテキストが各行別々のセルに入ります
                Set-Clipboard -Value $logContent

                $excel = New-Object -ComObject Excel.Application
                $excel.Visible = $true
                $workbook = $excel.Workbooks.Open($excelPath)
                $sheet = $workbook.Worksheets.Item($excelSheetName)

                # 貼り付け先のセルを選択してクリップボードから貼り付け
                # これにより改行で区切られた各行が A1, A2, A3... に配置されます
                $sheet.Range($excelPasteCell).Select()
                $sheet.Paste()

                $workbook.Save()
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) | Out-Null
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
                Write-Host "    貼り付け完了: $excelFileName"
                Write-Host ""

                # --------------------------------------------------------------
                # STEP 5: クリップボード内容をファイルに保存し、変換バッチを実行
                # --------------------------------------------------------------
                Write-Host "  [STEP 4] テキストファイル保存" -ForegroundColor Cyan
                # ジョブパスに対応するテキストファイル名を取得
                $textFileName = "runh_default.txt"  # デフォルト値
                foreach ($key in $jobTextFileMapping.Keys) {
                    if ($unitPath -eq $key) {
                        $textFileName = $jobTextFileMapping[$key]
                        break
                    }
                }
                # クリップボード内容は02_output/yyyymmddフォルダに保存
                $clipboardOutputFile = Join-Path $dateFolderPath $textFileName
                $convertBatchFile = Join-Path $scriptDir "【削除禁止】ConvertNS932Result.bat"

                # クリップボードの内容をファイルに保存
                Get-Clipboard | Out-File -FilePath $clipboardOutputFile -Encoding Default
                Write-Host "    保存完了: $textFileName"

                # 変換バッチファイルを実行（出力先フォルダを環境変数で渡す）
                if (Test-Path $convertBatchFile) {
                    Write-Host "    変換バッチ実行中..."
                    $env:OUTPUT_FOLDER = $dateFolderPath
                    & cmd /c "`"$convertBatchFile`""
                    $env:OUTPUT_FOLDER = $null
                }
                Write-Host ""

                # ------------------------------------------------------------------
                # 完了サマリー
                # ------------------------------------------------------------------
                Write-Host "================================================================" -ForegroundColor Green
                Write-Host "  Excel貼り付け完了" -ForegroundColor Green
                Write-Host "================================================================" -ForegroundColor Green
                Write-Host ""
                Write-Host "  出力先フォルダ: $dateFolderPath"
                Write-Host "  Excelファイル : $excelFileName"
                Write-Host "  テキストファイル: $textFileName"
                Write-Host ""
            } catch {
                Write-Console "[エラー] Excel貼り付けに失敗しました: $($_.Exception.Message)"
                exit 11  # Excel貼り付けエラー
            }
        }
        "/WINMERGE" {
            # ======================================================================
            # WinMerge比較処理（キーワード抽出機能付き）
            # ======================================================================
            # 【処理概要】
            # ログファイルから指定したキーワード間のテキストを抽出し、
            # WinMergeで比較表示します。
            # キーワードが見つからない場合は元のログファイルを開きます。
            # ======================================================================

            # WinMergeの実行ファイルパス（デフォルトのインストール先）
            $winMergePath = "C:\Program Files\WinMerge\WinMergeU.exe"

            # WinMergeが存在するか確認
            if (Test-Path $winMergePath) {
                # ----------------------------------------------------------
                # キーワード抽出設定
                # ----------------------------------------------------------
                # StartKeyword: 抽出開始キーワード（この行から抽出開始）
                # EndKeyword: 抽出終了キーワード（この行まで抽出）
                # OutputSuffix: 出力ファイル名のサフィックス
                $extractPairs = @(
                    @{
                        StartKeyword = "#パッチ適用前チェック"
                        EndKeyword = "#パッチ適用"
                        OutputSuffix = "_データパッチ前.txt"
                    },
                    @{
                        StartKeyword = "#パッチ適用後チェック"
                        EndKeyword = "#トランザクション処理"
                        OutputSuffix = "_データパッチ後.txt"
                    }
                )

                $extractedFiles = @()

                # ログファイルを読み込み
                $logLines = Get-Content -Path $outputFilePath -Encoding Default

                foreach ($pair in $extractPairs) {
                    $startKeyword = $pair.StartKeyword
                    $endKeyword = $pair.EndKeyword
                    $outputSuffix = $pair.OutputSuffix

                    # 出力ファイルパスを作成（元ファイル名 + サフィックス）
                    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($outputFilePath)
                    $extractedFile = Join-Path $outputFolder ($baseName + $outputSuffix)

                    # ----------------------------------------------------------
                    # キーワード間のテキスト抽出処理
                    # ----------------------------------------------------------
                    # 注意: 開始キーワードが終了キーワードを含む場合があるため
                    # （例: "#パッチ適用前チェック" は "#パッチ適用" を含む）
                    # 開始行では終了チェックをスキップする
                    $inSection = $false
                    $extractedContent = @()

                    foreach ($line in $logLines) {
                        # 開始キーワードを検出したら抽出開始
                        if ($line -like "*$startKeyword*") {
                            $inSection = $true
                            $extractedContent += $line
                            continue  # 開始行では終了チェックをスキップ
                        }
                        # 抽出中の処理
                        if ($inSection) {
                            # 終了キーワードを検出したら抽出終了（この行は含めない）
                            if ($line -like "*$endKeyword*") {
                                break
                            }
                            $extractedContent += $line
                        }
                    }

                    # 抽出結果をファイルに保存
                    if ($extractedContent.Count -gt 0) {
                        $extractedContent | Out-File -FilePath $extractedFile -Encoding Default
                        $extractedFiles += $extractedFile
                        Write-Console "[完了] キーワード抽出: $extractedFile"
                    }
                }

                # ----------------------------------------------------------
                # WinMergeで比較
                # ----------------------------------------------------------
                if ($extractedFiles.Count -eq 2) {
                    # 2つのファイルが作成された場合は比較モード
                    Start-Process $winMergePath -ArgumentList "`"$($extractedFiles[0])`" `"$($extractedFiles[1])`""
                    Write-Console "[完了] WinMergeで比較を開きました"
                    # 元のログファイルを削除
                    Remove-Item -Path $outputFilePath -Force
                    Write-Console "[完了] 元のログファイルを削除しました"
                } elseif ($extractedFiles.Count -eq 1) {
                    # 1つのファイルのみ作成された場合
                    Start-Process $winMergePath -ArgumentList "`"$($extractedFiles[0])`""
                    Write-Console "[完了] WinMergeでファイルを開きました"
                } else {
                    # キーワードが見つからない場合は元のログファイルを開く
                    Start-Process $winMergePath -ArgumentList "`"$outputFilePath`""
                    Write-Console "[情報] キーワードが見つからないため、元のログを開きました"
                }
            } else {
                Write-Console "[エラー] WinMergeが見つかりません: $winMergePath"
                Write-Console "        WinMergeをインストールするか、パスを確認してください。"
            }
        }
        default {
            # デフォルトはログファイル出力のみ
        }
    }

    Write-Console "[完了] 出力ファイル: $outputFilePath"
} catch {
    exit 8  # 詳細取得エラー
}

# 正常終了
exit 0
