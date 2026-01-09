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

rem 第2引数（比較用ジョブパス、オプション）を環境変数に設定
set "JP1_UNIT_PATH_2=%~2"

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
# JP1 REST API ジョブログ取得ツール
#
# 説明:
#   JP1/AJS3 Web Console REST APIを使用して、ジョブ/ジョブネットの
#   実行結果詳細を取得します（実行せずに既存の結果を取得）。
#   ※ JP1/AJS3 - Web Consoleが必要です
#
# 使い方:
#   ジョブのパスを指定して実行します:
#     JP1_REST_ジョブログ取得ツール.bat "/JobGroup/Jobnet/Job1"
#
#   2つのジョブを比較して新しい方を取得する場合:
#     JP1_REST_ジョブログ取得ツール.bat "/JobGroup/Jobnet/Job1" "/JobGroup/Jobnet/Job2"
#
#   ※ 親パス（ジョブネット）を検索対象とし、ジョブ名でフィルタします
#
# 処理フロー:
#   STEP 1: DEFINITION で存在確認・ユニット種別確認
#   STEP 2: DEFINITION_AND_STATUS で execID 取得
#   STEP 3: 実行結果詳細取得
#   ※ 2引数モード: 両方のジョブを取得し、START_TIMEを比較して新しい方を出力
#
# 終了コード（実行順）:
#   0: 正常終了
#   1: 引数エラー（ユニットパスが指定されていません）
#   2: ユニット未検出（STEP 1: 指定したユニットが存在しません）
#   3: ユニット種別エラー（STEP 1: 指定したユニットがジョブではありません）
#   4: 実行世代なし（STEP 2: 実行履歴が存在しません）
#   5: 5MB超過エラー（STEP 3: 実行結果が切り捨てられました）
#   6: 詳細取得エラー（STEP 3: 実行結果詳細の取得に失敗）
#   8: 比較モードで両方のジョブ取得に失敗
#   9: API接続エラー（各STEPでの接続失敗）
#  10: 比較モードで実行中のジョブが検出された
#  11: 実行中のジョブが検出された（待機タイムアウト）
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

# ------------------------------------------------------------------------------
# 実行中ジョブ待機設定
# ------------------------------------------------------------------------------
# ジョブが実行中の場合、終了するまで待機する最大秒数を指定します。
# 0を指定すると待機せずに即座にエラー終了します。
# デフォルト: 60秒（1分）
$maxWaitSeconds = 60

# 実行中ジョブのチェック間隔（秒）を指定します。
# デフォルト: 10秒
$checkIntervalSeconds = 10

# ==============================================================================
# ■ 検索条件設定セクション
# ==============================================================================
# このセクションでは、ユニット一覧取得APIの検索条件を設定します。

# ------------------------------------------------------------------------------
# (1) GenerationType - 世代指定
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
# (2) UnitStatus - ユニット状態フィルタ
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
# (3) DelayType - 遅延状態フィルタ
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
# (4) HoldPlan - 保留予定フィルタ
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
$unitPath2 = $env:JP1_UNIT_PATH_2

# 2引数モードかどうかを判定
$isCompareMode = $false
if ($unitPath2 -and $unitPath2.Trim() -ne "") {
    $isCompareMode = $true
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

# ==============================================================================
# ■ ユーティリティ関数（メイン処理の前に定義が必要）
# ==============================================================================

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

# ==============================================================================
# 2引数モード: 実行中チェック＆START_TIME比較処理
# ==============================================================================
# 2つのジョブパスが指定された場合:
# 1. まず両方のジョブが実行中かどうかをチェック
# 2. 実行中のジョブがあれば、終了を待機してそのジョブを最新として選択
# 3. どちらも実行中でなければ、START_TIMEを比較して新しい方を選択

$selectedPath = ""  # 比較モードで選択されたパス
$selectedTime = ""  # 比較モードで選択されたジョブの時間
$rejectedPath = ""  # 比較モードで選択されなかったパス
$rejectedTime = ""  # 比較モードで選択されなかったジョブの時間

if ($isCompareMode) {
    # ジョブが実行中かどうかをチェックする関数（execIDも取得）
    function Get-JobRunningStatus {
        param([string]$jobPath)

        $lastSlash = $jobPath.LastIndexOf("/")
        if ($lastSlash -le 0) { return $null }

        $parent = $jobPath.Substring(0, $lastSlash)
        $name = $jobPath.Substring($lastSlash + 1)
        if (-not $name) { return $null }

        $encParent = [System.Uri]::EscapeDataString($parent)
        $encName = [System.Uri]::EscapeDataString($name)

        $url = "${baseUrl}/objects/statuses?mode=search"
        $url += "&manager=${managerHost}"
        $url += "&serviceName=${schedulerService}"
        $url += "&location=${encParent}"
        $url += "&searchLowerUnits=NO"
        $url += "&searchTarget=DEFINITION_AND_STATUS"
        $url += "&unitName=${encName}"
        $url += "&unitNameMatchMethods=EQ"
        $url += "&generation=STATUS"
        $url += "&status=GRP_RUN"

        try {
            $resp = Invoke-WebRequest -Uri $url -Method GET -Headers $headers -TimeoutSec 30 -UseBasicParsing
            $bytes = $resp.RawContentStream.ToArray()
            $text = [System.Text.Encoding]::UTF8.GetString($bytes)
            $json = $text | ConvertFrom-Json

            if ($json.statuses -and $json.statuses.Count -gt 0) {
                $status = $json.statuses[0].unitStatus
                if ($status) {
                    return @{
                        IsRunning = $true
                        Status = $status.status
                        StartTime = $status.startTime
                        ExecID = $status.execID
                    }
                }
            }
        } catch {
            return $null
        }
        return @{ IsRunning = $false }
    }

    # 両方のジョブの実行中状態をチェック
    $runStatus1 = Get-JobRunningStatus -jobPath $unitPath
    $runStatus2 = Get-JobRunningStatus -jobPath $unitPath2

    $isRunning1 = $runStatus1 -and $runStatus1.IsRunning
    $isRunning2 = $runStatus2 -and $runStatus2.IsRunning

    # 実行中のジョブがある場合は待機
    if ($isRunning1 -or $isRunning2) {
        # 待機対象のジョブを決定（両方実行中の場合はジョブ1を優先）
        $waitTargetPath = if ($isRunning1) { $unitPath } else { $unitPath2 }
        $waitTargetStatus = if ($isRunning1) { $runStatus1 } else { $runStatus2 }
        $waitingExecId = $waitTargetStatus.ExecID
        $waitTargetStatusDisplay = Get-StatusDisplayName -status $waitTargetStatus.Status

        if ($isRunning1 -and $isRunning2) {
            [Console]::Error.WriteLine("COMPARE_INFO:両方のジョブが実行中です。$unitPath の終了を待機します")
        } else {
            [Console]::Error.WriteLine("COMPARE_INFO:実行中のジョブを検出しました - $waitTargetPath の終了を待機します")
        }

        # 待機ループ
        $waitedSeconds = 0
        $stillRunning = $true
        while ($stillRunning) {
            # 最大待機秒数を超えた場合はエラー終了
            if ($waitedSeconds -ge $maxWaitSeconds) {
                [Console]::WriteLine("RUNNING_ERROR:実行中のジョブが検出されました（待機タイムアウト）")
                [Console]::WriteLine("RUNNING_JOB:$waitTargetPath（ステータス: ${waitTargetStatusDisplay}, 開始日時: $($waitTargetStatus.StartTime)）")
                [Console]::WriteLine("WAIT_TIMEOUT:${maxWaitSeconds}秒待機しましたが、ジョブが終了しませんでした")
                exit 11  # 実行中のジョブが検出された（タイムアウト）
            }

            [Console]::Error.WriteLine("WAITING:実行中のジョブを検出しました。終了を待機しています...（${waitedSeconds}/${maxWaitSeconds}秒）")
            [Console]::Error.WriteLine("WAITING_JOB:$waitTargetPath（ステータス: ${waitTargetStatusDisplay}, 開始日時: $($waitTargetStatus.StartTime), execID: ${waitingExecId}）")

            Start-Sleep -Seconds $checkIntervalSeconds
            $waitedSeconds += $checkIntervalSeconds

            # 再度チェック
            $recheckStatus = Get-JobRunningStatus -jobPath $waitTargetPath
            if (-not $recheckStatus -or -not $recheckStatus.IsRunning) {
                $stillRunning = $false
                [Console]::Error.WriteLine("WAIT_COMPLETE:ジョブの終了を確認しました（${waitedSeconds}秒待機、execID: ${waitingExecId}）")
            } else {
                $waitTargetStatusDisplay = Get-StatusDisplayName -status $recheckStatus.Status
            }
        }

        # 待機完了後、実行中だったジョブを選択
        $originalUnitPath = $unitPath
        $unitPath = $waitTargetPath
        $selectedPath = $waitTargetPath
        $selectedTime = $waitTargetStatus.StartTime
        if ($waitTargetPath -eq $originalUnitPath) {
            $rejectedPath = $unitPath2
        } else {
            $rejectedPath = $originalUnitPath
        }
        $rejectedTime = "(実行中ジョブを優先)"

        [Console]::Error.WriteLine("INFO:待機していたジョブのexecID（${waitingExecId}）を使用してログを取得します")
    } else {
        # どちらも実行中でない場合はSTART_TIMEで比較
        # START_TIMEを取得する関数
        function Get-JobStartTime {
            param([string]$jobPath)

            $lastSlash = $jobPath.LastIndexOf("/")
            if ($lastSlash -le 0) { return $null }

            $parent = $jobPath.Substring(0, $lastSlash)
            $name = $jobPath.Substring($lastSlash + 1)
            if (-not $name) { return $null }

            $encParent = [System.Uri]::EscapeDataString($parent)
            $encName = [System.Uri]::EscapeDataString($name)

            $url = "${baseUrl}/objects/statuses?mode=search"
            $url += "&manager=${managerHost}"
            $url += "&serviceName=${schedulerService}"
            $url += "&location=${encParent}"
            $url += "&searchLowerUnits=NO"
            $url += "&searchTarget=DEFINITION_AND_STATUS"
            $url += "&unitName=${encName}"
            $url += "&unitNameMatchMethods=EQ"
            $url += "&generation=${generation}"

            try {
                $resp = Invoke-WebRequest -Uri $url -Method GET -Headers $headers -TimeoutSec 30 -UseBasicParsing
                $bytes = $resp.RawContentStream.ToArray()
                $text = [System.Text.Encoding]::UTF8.GetString($bytes)
                $json = $text | ConvertFrom-Json

                if ($json.statuses -and $json.statuses.Count -gt 0) {
                    $status = $json.statuses[0].unitStatus
                    if ($status -and $status.startTime) {
                        return $status.startTime
                    }
                }
            } catch {
                return $null
            }
            return $null
        }

        # 両方のジョブのSTART_TIMEを取得
        $startTime1 = Get-JobStartTime -jobPath $unitPath
        $startTime2 = Get-JobStartTime -jobPath $unitPath2

        # 両方失敗した場合
        if (-not $startTime1 -and -not $startTime2) {
            exit 8  # 比較モードで両方のジョブ取得に失敗
        }

        # 元のunitPathを保存（比較結果表示用）
        $originalUnitPath = $unitPath

        # 片方だけ失敗した場合
        if (-not $startTime1) {
            $unitPath = $unitPath2
            $selectedPath = $unitPath2
            $selectedTime = $startTime2
            $rejectedPath = $originalUnitPath
            $rejectedTime = "(取得失敗)"
        } elseif (-not $startTime2) {
            # $unitPath はそのまま
            $selectedPath = $unitPath
            $selectedTime = $startTime1
            $rejectedPath = $unitPath2
            $rejectedTime = "(取得失敗)"
        } else {
            # 両方成功した場合、日時を比較
            try {
                $dt1 = [DateTime]::Parse($startTime1)
                $dt2 = [DateTime]::Parse($startTime2)

                if ($dt2 -gt $dt1) {
                    $unitPath = $unitPath2
                    $selectedPath = $unitPath2
                    $selectedTime = $startTime2
                    $rejectedPath = $originalUnitPath
                    $rejectedTime = $startTime1
                } else {
                    $selectedPath = $unitPath
                    $selectedTime = $startTime1
                    $rejectedPath = $unitPath2
                    $rejectedTime = $startTime2
                }
            } catch {
                # パースエラーの場合は文字列比較
                if ($startTime2 -gt $startTime1) {
                    $unitPath = $unitPath2
                    $selectedPath = $unitPath2
                    $selectedTime = $startTime2
                    $rejectedPath = $originalUnitPath
                    $rejectedTime = $startTime1
                } else {
                    $selectedPath = $unitPath
                    $selectedTime = $startTime1
                    $rejectedPath = $unitPath2
                    $rejectedTime = $startTime2
                }
            }
        }
    }
}

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
# STEP 1.5: 親ジョブネットのコメント取得
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
# STEP 1.8: 実行中ジョブチェック（待機機能付き）
# ==============================================================================
# 同じジョブが現在実行中かどうかを確認します。
# 実行中の場合は、終了するまで待機します（最大待機秒数まで）。
# 最大待機秒数を超えても終了しない場合はエラー終了します。

$runningUrl = "${baseUrl}/objects/statuses?mode=search"
$runningUrl += "&manager=${managerHost}"
$runningUrl += "&serviceName=${schedulerService}"
$runningUrl += "&location=${encodedParentPath}"
$runningUrl += "&searchLowerUnits=NO"
$runningUrl += "&searchTarget=DEFINITION_AND_STATUS"
$runningUrl += "&unitName=${encodedJobName}"
$runningUrl += "&unitNameMatchMethods=EQ"
$runningUrl += "&generation=STATUS"
$runningUrl += "&status=GRP_RUN"

$waitedSeconds = 0
$isRunning = $true
$waitingExecId = $null  # 待機中のジョブのexecIDを保存

while ($isRunning) {
    try {
        $runningResponse = Invoke-WebRequest -Uri $runningUrl -Method GET -Headers $headers -TimeoutSec 30 -UseBasicParsing

        # UTF-8文字化け対策
        $runningBytes = $runningResponse.RawContentStream.ToArray()
        $runningText = [System.Text.Encoding]::UTF8.GetString($runningBytes)
        $runningJson = $runningText | ConvertFrom-Json

        # 実行中のジョブが存在するか確認
        if ($runningJson.statuses -and $runningJson.statuses.Count -gt 0) {
            $runningUnit = $runningJson.statuses[0]
            $runningStatus = $runningUnit.unitStatus.status
            $runningStartTime = $runningUnit.unitStatus.startTime
            $runningStatusDisplay = Get-StatusDisplayName -status $runningStatus

            # 最初の検出時にexecIDを保存（待機完了後にこのexecIDでログを取得）
            if (-not $waitingExecId) {
                $waitingExecId = $runningUnit.unitStatus.execID
            }

            # 最大待機秒数を超えた場合はエラー終了
            if ($waitedSeconds -ge $maxWaitSeconds) {
                [Console]::WriteLine("RUNNING_ERROR:実行中のジョブが検出されました（待機タイムアウト）")
                [Console]::WriteLine("RUNNING_JOB:$unitPath（ステータス: ${runningStatusDisplay}, 開始日時: ${runningStartTime}）")
                [Console]::WriteLine("WAIT_TIMEOUT:${maxWaitSeconds}秒待機しましたが、ジョブが終了しませんでした")
                exit 11  # 実行中のジョブが検出された（タイムアウト）
            }

            # 待機中メッセージを出力（標準エラー出力へ）
            [Console]::Error.WriteLine("WAITING:実行中のジョブを検出しました。終了を待機しています...（${waitedSeconds}/${maxWaitSeconds}秒）")
            [Console]::Error.WriteLine("WAITING_JOB:$unitPath（ステータス: ${runningStatusDisplay}, 開始日時: ${runningStartTime}, execID: ${waitingExecId}）")

            # 指定秒数待機
            Start-Sleep -Seconds $checkIntervalSeconds
            $waitedSeconds += $checkIntervalSeconds
        } else {
            # 実行中ではない → ループを抜ける
            $isRunning = $false

            # 待機していた場合は完了メッセージを出力
            if ($waitedSeconds -gt 0) {
                [Console]::Error.WriteLine("WAIT_COMPLETE:ジョブの終了を確認しました（${waitedSeconds}秒待機、execID: ${waitingExecId}）")
            }
        }
    } catch {
        # 実行中チェック失敗は無視して続行（必須ではない）
        $isRunning = $false
    }
}

# ==============================================================================
# STEP 2: 実行状態・execID取得（DEFINITION_AND_STATUS）
# ==============================================================================
# 存在確認・種別確認が成功したら、DEFINITION_AND_STATUS で execID を取得します。
# ※ 待機していた場合は、待機中に保存したexecIDを使用します。

$statusUrl = "${baseUrl}/objects/statuses?mode=search"
$statusUrl += "&manager=${managerHost}"
$statusUrl += "&serviceName=${schedulerService}"
$statusUrl += "&location=${encodedParentPath}"
$statusUrl += "&searchLowerUnits=NO"
$statusUrl += "&searchTarget=DEFINITION_AND_STATUS"
$statusUrl += "&unitName=${encodedJobName}"
$statusUrl += "&unitNameMatchMethods=EQ"

# 待機していた場合は、そのexecIDを使用（設定ファイルの世代指定を上書き）
if ($waitingExecId) {
    $statusUrl += "&generation=EXECID"
    $statusUrl += "&execID=${waitingExecId}"
    [Console]::Error.WriteLine("INFO:待機していたジョブのexecID（${waitingExecId}）を使用してログを取得します")
} else {
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
            $startTimeValue = if ($unitStatus) { $unitStatus.startTime } else { $null }

            # execIDがある場合のみリストに追加
            if ($execIdValue) {
                $execIdList += @{
                    Path = $unitFullName
                    ExecId = $execIdValue
                    Status = $statusValue
                    UnitType = $unitTypeValue
                    StartTime = $startTimeValue
                    EndStatus = $statusValue  # 終了状態を保持
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
        $targetStartTime = $item.StartTime
        $targetEndStatus = $item.EndStatus

        # 開始日時をファイル名用フォーマットに変換（yyyyMMdd_HHmmss）
        # 例: "2015-09-02T22:50:28+09:00" → "20150902_225028"
        $startTimeForFileName = ""
        if ($targetStartTime) {
            try {
                $dt = [DateTime]::Parse($targetStartTime)
                $startTimeForFileName = $dt.ToString("yyyyMMdd_HHmmss")
            } catch {
                $startTimeForFileName = ""
            }
        }

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

            # 実行結果の内容を取得
            $execResultContent = ""
            if ($resultJson.execResultDetails) {
                $execResultContent = $resultJson.execResultDetails
            }

            # 比較モードで選択されたパスと時間を出力（存在する場合）
            if ($selectedPath) {
                [Console]::WriteLine("SELECTED_PATH:$selectedPath")
                [Console]::WriteLine("SELECTED_TIME:$selectedTime")
                [Console]::WriteLine("REJECTED_PATH:$rejectedPath")
                [Console]::WriteLine("REJECTED_TIME:$rejectedTime")
            }

            # 開始日時を最初の行に出力（ファイル名用フォーマット）
            [Console]::WriteLine("START_TIME:$startTimeForFileName")

            # 終了状態を出力（ファイル名用、日本語変換済み）
            $endStatusDisplay = Get-StatusDisplayName -status $targetEndStatus
            [Console]::WriteLine("END_STATUS:$endStatusDisplay")

            # ジョブネット名を出力（ファイル名用）
            [Console]::WriteLine("JOBNET_NAME:$jobnetName")

            # ジョブネットコメントを出力（ファイル名用）
            [Console]::WriteLine("JOBNET_COMMENT:$jobnetComment")

            # 実行結果詳細を出力
            [Console]::WriteLine($execResultContent)
        } catch {
            exit 6  # 詳細取得エラー
        }
    }
}

# 正常終了
exit 0
