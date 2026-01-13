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
rem 使い方（単一ジョブ）:
rem   JP1ジョブ実行情報取得.bat "ジョブパス"
rem
rem 使い方（複数ジョブ - 同一ジョブネット内）:
rem   JP1ジョブ実行情報取得.bat "ジョブパス1" "ジョブパス2"
rem
rem 出力オプション（環境変数 JP1_OUTPUT_MODE で指定、必須）:
rem   /NOTEPAD  - メモ帳で開く
rem   /EXCEL    - Excelに貼り付け
rem   /WINMERGE - WinMergeで比較
rem ============================================================================

rem 引数がない場合はエラー
if "%~1"=="" exit /b 1

rem 出力オプションは環境変数 JP1_OUTPUT_MODE から取得（必須）
if "%JP1_OUTPUT_MODE%"=="" exit /b 1

rem 引数を環境変数に設定（PowerShellに渡すため）
if not "%~2"=="" (
    rem 複数ジョブモード（2つの引数）
    set "JP1_UNIT_PATH="
    set "JP1_UNIT_PATH_1=%~1"
    set "JP1_UNIT_PATH_2=%~2"
) else (
    rem 単一ジョブモード（1つの引数）
    set "JP1_UNIT_PATH=%~1"
    set "JP1_UNIT_PATH_1="
    set "JP1_UNIT_PATH_2="
)

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
# JP1 REST API ジョブ実行ログ取得ツール（複数ジョブ対応版）
#
# 説明:
#   JP1/AJS3 Web Console REST APIを使用して、ジョブネットを即時実行し、
#   完了後にジョブの実行結果詳細を取得します。
#   同一ジョブネット内の複数ジョブを1回の実行で処理できます。
#   ※ JP1/AJS3 - Web Consoleが必要です
#
# 使い方:
#   単一ジョブモード:
#     JP1ジョブ実行情報取得.bat "/JobGroup/Jobnet/Job1"
#
#   複数ジョブモード（同一ジョブネット内）:
#     JP1ジョブ実行情報取得.bat "/JobGroup/Jobnet/Job1" "/JobGroup/Jobnet/Job2"
#
# 処理フロー:
#   STEP 1: DEFINITION で存在確認・ユニット種別確認（全ジョブ）
#   STEP 2: ルートジョブネット特定・同一確認
#   STEP 3: 親ジョブネットのコメント取得
#   STEP 4: 即時実行登録API でルートジョブネットを実行（1回のみ）
#   STEP 5: 状態監視（全ジョブの完了待ち）
#   STEP 6: 実行結果詳細取得（各ジョブ）
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
#   14: ルートジョブネット不一致（複数ジョブが同一ジョブネットではありません）
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
# ★ キー: ジョブパス（完全一致で検索）
# ★ 値: Excelファイル名、シート名、貼り付けセルをカンマ区切りで指定
# 例: "/グループ/ネット/ジョブ" = "ファイル名.xlsx,Sheet1,A1"
$jobExcelMapping = @{
    # === ジョブパスとExcelファイルのマッピング ===
    # 以下に「ジョブパス」=「Excelファイル名,シート名,セル」の形式で記載してください
    # 完全一致で検索されるため、呼び出し元で指定するパスと同じ形式で記載してください
    #
    # 例:
    # "/TIA/Jobnet/週単位ジョブ" = "TIA解析(自習当初)_週単位.xls,Sheet1,A1"
    # "/TIA/Jobnet/年単位ジョブ" = "TIA解析(自習当初)_年単位.xls,Sheet1,A1"
    #
    # ★ 以下を編集してください ★
    "/サンプル/Jobnet/週単位ジョブ" = "TIA解析(自習当初)_週単位.xls,Sheet1,A1"
    "/サンプル/Jobnet/年単位ジョブ" = "TIA解析(自習当初)_年単位.xls,Sheet1,A1"
    # "/グループ/ネット/ジョブ3" = "Excelファイル3.xls,Sheet1,A1"
    # "/グループ/ネット/ジョブ4" = "Excelファイル4.xls,Sheet1,A1"
    # "/グループ/ネット/ジョブ5" = "Excelファイル5.xls,Sheet1,A1"
    # "/グループ/ネット/ジョブ6" = "Excelファイル6.xls,Sheet1,A1"
}

# ジョブパスとテキストファイル名の紐づけ設定（クリップボード保存用）
# ★ キー: ジョブパス（完全一致で検索）
# ★ 値: 保存するテキストファイル名
# 例: "/グループ/ネット/ジョブ" = "runh_week.txt"
$jobTextFileMapping = @{
    # === ジョブパスとテキストファイルのマッピング ===
    # Excelに貼り付けた後、クリップボード内容を保存するファイル名を指定します
    # 完全一致で検索されるため、呼び出し元で指定するパスと同じ形式で記載してください
    #
    # ★ 以下を編集してください ★
    "/サンプル/Jobnet/週単位ジョブ" = "runh_week.txt"
    "/サンプル/Jobnet/年単位ジョブ" = "runh_year.txt"
    # "/グループ/ネット/ジョブ3" = "runh_file3.txt"
    # "/グループ/ネット/ジョブ4" = "runh_file4.txt"
    # "/グループ/ネット/ジョブ5" = "runh_file5.txt"
    # "/グループ/ネット/ジョブ6" = "runh_file6.txt"
}

# ==============================================================================
# ■ メイン処理（以下は通常編集不要）
# ==============================================================================

# ------------------------------------------------------------------------------
# 環境変数からユニットパスを取得
# ------------------------------------------------------------------------------
# 複数ジョブモードか単一ジョブモードかを判定
$unitPaths = @()

if ($env:JP1_UNIT_PATH_1) {
    # 複数ジョブモード
    $unitPaths += $env:JP1_UNIT_PATH_1
    if ($env:JP1_UNIT_PATH_2) {
        $unitPaths += $env:JP1_UNIT_PATH_2
    }
    Write-Console "[モード] 複数ジョブ: $($unitPaths.Count) 件"
} elseif ($env:JP1_UNIT_PATH) {
    # 単一ジョブモード
    $unitPaths += $env:JP1_UNIT_PATH
    Write-Console "[モード] 単一ジョブ"
} else {
    exit 1  # 引数エラー
}

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

# ==============================================================================
# 各ジョブの情報を格納する配列を初期化
# ==============================================================================
$jobInfoList = @()

foreach ($unitPath in $unitPaths) {
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

    $jobInfoList += @{
        UnitPath = $unitPath
        ParentPath = $parentPath
        JobName = $jobName
        GrandParentPath = $grandParentPath
        JobnetName = $jobnetName
        RootJobnetName = $null
        JobnetComment = ""
        JobStatus = $null
        JobExecId = $null
        JobStartTime = $null
        OutputFilePath = $null
    }
}

# ==============================================================================
# STEP 1: ユニット存在確認・種別確認（DEFINITION）- 全ジョブ
# ==============================================================================
Write-Console "[STEP 1] ユニット存在確認中..."

for ($i = 0; $i -lt $jobInfoList.Count; $i++) {
    $jobInfo = $jobInfoList[$i]
    $encodedParentPath = [System.Uri]::EscapeDataString($jobInfo.ParentPath)
    $encodedJobName = [System.Uri]::EscapeDataString($jobInfo.JobName)

    $defUrl = "${baseUrl}/objects/statuses?mode=search"
    $defUrl += "&manager=${managerHost}"
    $defUrl += "&serviceName=${schedulerService}"
    $defUrl += "&location=${encodedParentPath}"
    $defUrl += "&searchLowerUnits=NO"
    $defUrl += "&searchTarget=DEFINITION"
    $defUrl += "&unitName=${encodedJobName}"
    $defUrl += "&unitNameMatchMethods=EQ"

    try {
        $defResponse = Invoke-WebRequest -Uri $defUrl -Method GET -Headers $headers -TimeoutSec 30 -UseBasicParsing

        # UTF-8文字化け対策
        $defBytes = $defResponse.RawContentStream.ToArray()
        $defText = [System.Text.Encoding]::UTF8.GetString($defBytes)
        $defJson = $defText | ConvertFrom-Json

        # ユニット存在確認
        if (-not $defJson.statuses -or $defJson.statuses.Count -eq 0) {
            Write-Console "[エラー] ユニットが見つかりません: $($jobInfo.UnitPath)"
            exit 2  # ユニット未検出
        }

        # 最初のユニットの情報を取得
        $defUnit = $defJson.statuses[0]
        $unitTypeValue = $defUnit.definition.unitType
        $jobInfoList[$i].RootJobnetName = $defUnit.definition.rootJobnetName

        # ユニット種別確認（JOB系かどうか）
        if ($unitTypeValue -notmatch "JOB") {
            Write-Console "[エラー] ジョブではありません: $($jobInfo.UnitPath) (種別: $unitTypeValue)"
            exit 3  # ユニット種別エラー（ジョブではない）
        }

        Write-Console "  [$($i + 1)/$($jobInfoList.Count)] $($jobInfo.JobName) - OK"

    } catch {
        Write-Console "[エラー] API接続エラー: $($_.Exception.Message)"
        exit 9  # API接続エラー（存在確認）
    }
}

# ==============================================================================
# STEP 2: ルートジョブネット特定・同一確認
# ==============================================================================
Write-Console "[STEP 2] ルートジョブネット確認中..."

$rootJobnetName = $jobInfoList[0].RootJobnetName

if (-not $rootJobnetName) {
    exit 4  # ルートジョブネット特定エラー
}

# 複数ジョブの場合、全て同じルートジョブネットか確認
if ($jobInfoList.Count -gt 1) {
    foreach ($jobInfo in $jobInfoList) {
        if ($jobInfo.RootJobnetName -ne $rootJobnetName) {
            Write-Console "[エラー] ルートジョブネットが異なります"
            Write-Console "  ジョブ1: $rootJobnetName"
            Write-Console "  ジョブ2: $($jobInfo.RootJobnetName)"
            exit 14  # ルートジョブネット不一致
        }
    }
    Write-Console "  全ジョブが同一ジョブネット内です: $rootJobnetName"
} else {
    Write-Console "  ルートジョブネット: $rootJobnetName"
}

# ==============================================================================
# STEP 3: 親ジョブネットのコメント取得
# ==============================================================================
Write-Console "[STEP 3] 親ジョブネットのコメント取得中..."

# 最初のジョブの親ジョブネットからコメントを取得（同一ジョブネット内なので共通）
$jobInfo = $jobInfoList[0]
$encodedGrandParentPath = [System.Uri]::EscapeDataString($jobInfo.GrandParentPath)
$encodedJobnetName = [System.Uri]::EscapeDataString($jobInfo.JobnetName)

$jobnetUrl = "${baseUrl}/objects/statuses?mode=search"
$jobnetUrl += "&manager=${managerHost}"
$jobnetUrl += "&serviceName=${schedulerService}"
$jobnetUrl += "&location=${encodedGrandParentPath}"
$jobnetUrl += "&searchLowerUnits=NO"
$jobnetUrl += "&searchTarget=DEFINITION"
$jobnetUrl += "&unitName=${encodedJobnetName}"
$jobnetUrl += "&unitNameMatchMethods=EQ"

$jobnetComment = ""

try {
    $jobnetResponse = Invoke-WebRequest -Uri $jobnetUrl -Method GET -Headers $headers -TimeoutSec 30 -UseBasicParsing

    # UTF-8文字化け対策
    $jobnetBytes = $jobnetResponse.RawContentStream.ToArray()
    $jobnetText = [System.Text.Encoding]::UTF8.GetString($jobnetBytes)
    $jobnetJson = $jobnetText | ConvertFrom-Json

    # ジョブネットのコメントを取得
    if ($jobnetJson.statuses -and $jobnetJson.statuses.Count -gt 0) {
        $jobnetDef = $jobnetJson.statuses[0].definition
        if ($jobnetDef.unitComment) {
            $jobnetComment = $jobnetDef.unitComment
        }
    }
} catch {
    # コメント取得失敗は無視して続行（必須ではない）
    $jobnetComment = ""
}

# 全ジョブにコメントを設定
for ($i = 0; $i -lt $jobInfoList.Count; $i++) {
    $jobInfoList[$i].JobnetComment = $jobnetComment
}

# ==============================================================================
# STEP 4: 即時実行登録API（1回のみ）
# ==============================================================================
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

    Write-Console "  実行ID: $execIdFromRegister"

} catch {
    Write-Console "[エラー] 即時実行登録に失敗しました: $($_.Exception.Message)"
    exit 5  # 即時実行登録エラー
}

# ==============================================================================
# STEP 5: 状態監視（全ジョブの完了待ち）
# ==============================================================================
Write-Console "[STEP 5] ジョブ完了待機中..."

# 終了状態の定義
$finishedStatuses = @(
    "NORMAL", "NORMALFALSE", "WARNING", "ABNORMAL",
    "INTERRUPT", "KILL", "FAIL", "UNKNOWN",
    "MONITORCLOSE", "UNEXECMONITOR", "MONITORINTRPT", "MONITORNORMAL",
    "UNEXEC", "BYPASS", "EXECDEFFER", "INVALIDSEQ", "SHUTDOWN"
)

$startTime = Get-Date
$endTime = $startTime.AddSeconds($timeoutSeconds)

# 各ジョブの完了フラグ
$completedJobs = @{}
foreach ($jobInfo in $jobInfoList) {
    $completedJobs[$jobInfo.UnitPath] = $false
}

while ((Get-Date) -lt $endTime) {
    $allCompleted = $true

    for ($i = 0; $i -lt $jobInfoList.Count; $i++) {
        $jobInfo = $jobInfoList[$i]

        # 既に完了しているジョブはスキップ
        if ($completedJobs[$jobInfo.UnitPath]) {
            continue
        }

        $encodedParentPath = [System.Uri]::EscapeDataString($jobInfo.ParentPath)
        $encodedJobName = [System.Uri]::EscapeDataString($jobInfo.JobName)

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
                    $jobInfoList[$i].JobStatus = $unitStatus.status
                    $jobInfoList[$i].JobExecId = $unitStatus.execID
                    $jobInfoList[$i].JobStartTime = $unitStatus.startTime

                    # 終了状態かどうかをチェック
                    if ($finishedStatuses -contains $unitStatus.status) {
                        $completedJobs[$jobInfo.UnitPath] = $true
                        Write-Console "  [$($i + 1)/$($jobInfoList.Count)] $($jobInfo.JobName) - 完了 ($($unitStatus.status))"
                    }
                }
            }

        } catch {
            # API呼び出しエラーは無視して再試行
        }

        # 未完了のジョブがある
        if (-not $completedJobs[$jobInfo.UnitPath]) {
            $allCompleted = $false
        }
    }

    # 全ジョブが完了したらループを抜ける
    if ($allCompleted) {
        break
    }

    Start-Sleep -Seconds $pollIntervalSeconds
}

# タイムアウトチェック
if ((Get-Date) -ge $endTime) {
    foreach ($jobInfo in $jobInfoList) {
        if (-not $completedJobs[$jobInfo.UnitPath]) {
            Write-Console "[エラー] タイムアウト: $($jobInfo.JobName)"
        }
    }
    exit 6  # タイムアウト
}

# ==============================================================================
# STEP 6: 実行結果詳細取得API（各ジョブ）
# ==============================================================================
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

# 出力オプションを環境変数から取得
$outputMode = $env:JP1_OUTPUT_MODE
if (-not $outputMode) { $outputMode = "/NOTEPAD" }

# 出力ディレクトリを作成
$outputDir = Join-Path $scriptDir "..\02.Output"
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

# Excel処理用の共通変数（/EXCELモード時）
$dateFolderPath = $null
$templateCopied = $false

if ($outputMode.ToUpper() -eq "/EXCEL") {
    # --------------------------------------------------------------
    # Excel共通処理: 出力フォルダ準備（1回だけ実行）
    # --------------------------------------------------------------
    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Yellow
    Write-Host "  Excel貼り付け処理" -ForegroundColor Yellow
    Write-Host "================================================================" -ForegroundColor Yellow
    Write-Host ""

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
    # Excel共通処理: 雛形フォルダコピー（1回だけ実行）
    # --------------------------------------------------------------
    Write-Host "  [STEP 2] 雛形フォルダコピー" -ForegroundColor Cyan
    $templatePath = Join-Path $scriptDir $templateFolderName

    if (-not (Test-Path $templatePath)) {
        Write-Console "[エラー] 雛形フォルダが見つかりません: $templatePath"
        exit 13  # 雛形フォルダ未検出エラー
    }

    # 雛形フォルダの中身をyyyymmddフォルダに直接コピー
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
    $templateCopied = $true
}

# 各ジョブの実行結果を取得
for ($i = 0; $i -lt $jobInfoList.Count; $i++) {
    $jobInfo = $jobInfoList[$i]

    Write-Console "  [$($i + 1)/$($jobInfoList.Count)] $($jobInfo.JobName) の結果取得中..."

    # 開始日時をファイル名用フォーマットに変換（yyyyMMdd_HHmmss）
    $startTimeForFileName = ""
    if ($jobInfo.JobStartTime) {
        try {
            $dt = [DateTime]::Parse($jobInfo.JobStartTime)
            $startTimeForFileName = $dt.ToString("yyyyMMdd_HHmmss")
        } catch {
            $startTimeForFileName = ""
        }
    }

    $encodedPath = [System.Uri]::EscapeDataString($jobInfo.UnitPath)

    # 実行結果詳細取得APIのURL構築
    $detailUrl = "${baseUrl}/objects/statuses/${encodedPath}:$($jobInfo.JobExecId)/actions/execResultDetails/invoke"
    $detailUrl += "?manager=${managerHost}&serviceName=${schedulerService}"

    try {
        # APIリクエストを送信
        $resultResponse = Invoke-WebRequest -Uri $detailUrl -Method GET -Headers $headers -TimeoutSec 30 -UseBasicParsing

        # UTF-8文字化け対策
        $resultBytes = $resultResponse.RawContentStream.ToArray()
        $resultText = [System.Text.Encoding]::UTF8.GetString($resultBytes)
        $resultJson = $resultText | ConvertFrom-Json

        # all フラグのチェック（falseの場合は5MB超過で切り捨て）
        if ($resultJson.all -eq $false) {
            Write-Console "[警告] 5MB超過のため結果が切り捨てられました: $($jobInfo.JobName)"
        }

        # 実行結果の内容を取得
        $execResultContent = ""
        if ($resultJson.execResultDetails) {
            $execResultContent = $resultJson.execResultDetails
        }

        # 終了状態を取得（日本語変換済み）
        $endStatusDisplay = Get-StatusDisplayName -status $jobInfo.JobStatus

        # 出力ファイル名を生成
        $outputFileName = "【ジョブ実行結果】【${startTimeForFileName}実行分】【${endStatusDisplay}】$($jobInfo.JobnetName)_$($jobInfo.JobnetComment).txt"
        $outputFilePath = Join-Path $outputDir $outputFileName

        # 実行結果詳細をファイルに出力
        $execResultContent | Out-File -FilePath $outputFilePath -Encoding Default
        $jobInfoList[$i].OutputFilePath = $outputFilePath

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
                    for ($j = 0; $j -lt $lines.Count; $j++) {
                        if ($lines[$j] -match [regex]::Escape($scrollToText)) {
                            $scrollLineNum = $j + 1  # 1始まりの行番号
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
                # Excelに貼り付け
                Write-Host ""
                Write-Host "  [STEP 3-$($i + 1)] $($jobInfo.JobName) をExcelに貼り付け" -ForegroundColor Cyan

                # ジョブパスからExcel設定を取得
                $excelFileName = $null
                $excelSheetName = $null
                $excelPasteCell = $null

                # ジョブパスに一致するマッピングを検索（完全一致）
                foreach ($key in $jobExcelMapping.Keys) {
                    if ($jobInfo.UnitPath -eq $key) {
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
                    Write-Console "[エラー] ジョブパス '$($jobInfo.UnitPath)' に対応するExcel設定が見つかりません。"
                    continue
                }

                Write-Host "    ジョブパス    : $($jobInfo.UnitPath)"
                Write-Host "    Excelファイル : $excelFileName"
                Write-Host "    シート名      : $excelSheetName"
                Write-Host "    貼り付けセル  : $excelPasteCell"

                $excelPath = Join-Path $dateFolderPath $excelFileName

                if (-not (Test-Path $excelPath)) {
                    Write-Console "[エラー] Excelファイルが見つかりません: $excelPath"
                    continue
                }

                try {
                    # ログファイルの内容を読み込み
                    $logContent = Get-Content $outputFilePath -Encoding Default -Raw

                    # クリップボードにコピー
                    Set-Clipboard -Value $logContent

                    $excel = New-Object -ComObject Excel.Application
                    $excel.Visible = $true
                    $workbook = $excel.Workbooks.Open($excelPath)
                    $sheet = $workbook.Worksheets.Item($excelSheetName)

                    # 貼り付け先のセルを選択してクリップボードから貼り付け
                    $sheet.Range($excelPasteCell).Select()
                    $sheet.Paste()

                    $workbook.Save()
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) | Out-Null
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
                    Write-Host "    貼り付け完了: $excelFileName"

                    # テキストファイル保存
                    $textFileName = "runh_default.txt"
                    foreach ($key in $jobTextFileMapping.Keys) {
                        if ($jobInfo.UnitPath -eq $key) {
                            $textFileName = $jobTextFileMapping[$key]
                            break
                        }
                    }
                    $clipboardOutputFile = Join-Path $dateFolderPath $textFileName
                    $convertBatchFile = Join-Path $scriptDir "【削除禁止】ConvertNS932Result.bat"

                    Get-Clipboard | Out-File -FilePath $clipboardOutputFile -Encoding Default
                    Write-Host "    テキスト保存: $textFileName"

                    # 変換バッチファイルを実行
                    if (Test-Path $convertBatchFile) {
                        $env:OUTPUT_FOLDER = $dateFolderPath
                        & cmd /c "`"$convertBatchFile`""
                        $env:OUTPUT_FOLDER = $null
                    }

                } catch {
                    Write-Console "[エラー] Excel貼り付けに失敗しました: $($_.Exception.Message)"
                }
            }
            "/WINMERGE" {
                # WinMerge処理（既存のロジック）
                $winMergePath = "C:\Program Files\WinMerge\WinMergeU.exe"
                if (Test-Path $winMergePath) {
                    Start-Process $winMergePath -ArgumentList "`"$outputFilePath`""
                }
            }
        }

        Write-Console "    完了: $outputFilePath"

    } catch {
        Write-Console "[エラー] 詳細取得に失敗: $($jobInfo.JobName) - $($_.Exception.Message)"
    }
}

# ==============================================================================
# 完了サマリー
# ==============================================================================
Write-Host ""
Write-Host "================================================================" -ForegroundColor Green
Write-Host "  処理完了" -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  処理ジョブ数: $($jobInfoList.Count)"
if ($dateFolderPath) {
    Write-Host "  出力先: $dateFolderPath"
}
Write-Host ""

# 正常終了
exit 0
