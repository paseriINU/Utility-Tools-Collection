<# :
@echo off
setlocal
chcp 932 >nul

rem ============================================================================
rem ■ バッチファイル部分（PowerShellを起動するためのラッパー）
rem ============================================================================
rem
rem 【このファイルの仕組み - ポリグロットパターン】
rem   このファイルは「ポリグロット」と呼ばれる特殊な形式で作成されています。
rem   1つのファイルがバッチファイル(.bat)としてもPowerShellスクリプト(.ps1)
rem   としても動作します。
rem
rem   - 最初の行「<# :」はバッチファイルでは無視され、PowerShellでは
rem     コメントブロックの開始として認識されます
rem   - 最後の「: 番号記号 大なり」はバッチファイルではラベル、PowerShellでは
rem     コメントブロックの終了として認識されます
rem   - これにより、ダブルクリックで直接実行できる.batファイルでありながら、
rem     PowerShellの高度な機能（REST API呼び出し等）を使用できます
rem
rem 【使い方】
rem   このファイルは直接実行せず、呼び出し元のバッチファイルから使用します。
rem
rem   呼び出し元での設定例:
rem     set "JP1_OUTPUT_MODE=/NOTEPAD"
rem     call JP1ジョブ情報取得.bat "/JobGroup/Jobnet/Job1"
rem
rem   2つのジョブを比較して新しい方を取得する場合:
rem     call JP1ジョブ情報取得.bat "/JobGroup/Jobnet/Job1" "/JobGroup/Jobnet/Job2"
rem
rem 【出力オプション】（環境変数 JP1_OUTPUT_MODE で指定、必須）
rem   /NOTEPAD  - 取得したログをメモ帳で開きます
rem   /EXCEL    - 取得したログをExcelの指定セルに貼り付けます
rem   /WINMERGE - WinMergeで比較します（未実装）
rem
rem 【必要な環境】
rem   - JP1/AJS3 - Web Console がインストールされていること
rem   - PowerShell 5.1 以降
rem   - ネットワーク経由でWeb Consoleサーバーに接続できること
rem
rem ============================================================================

rem ----------------------------------------------------------------------------
rem 引数チェック
rem ----------------------------------------------------------------------------
rem 第1引数（ジョブパス）が指定されていない場合はエラー終了します。
rem "%~1" は第1引数を取得します。~1 は引用符を除去した値を返します。
rem 空文字列と比較して、引数がない場合はエラーコード1で終了します。
if "%~1"=="" exit /b 1

rem ----------------------------------------------------------------------------
rem 環境変数への引数設定
rem ----------------------------------------------------------------------------
rem バッチファイルからPowerShellに値を渡すため、環境変数を使用します。
rem PowerShell内で $env:JP1_UNIT_PATH のようにアクセスできます。

rem 第1引数（ジョブパス）を環境変数に設定
rem 例: "/業務システム/日次バッチ/データ集計ジョブ"
set "JP1_UNIT_PATH=%~1"

rem 第2引数（比較用ジョブパス、オプション）を環境変数に設定
rem 2つのジョブを指定した場合、開始時刻を比較して新しい方を取得します
set "JP1_UNIT_PATH_2=%~2"

rem ----------------------------------------------------------------------------
rem 出力オプションチェック
rem ----------------------------------------------------------------------------
rem 出力オプション（JP1_OUTPUT_MODE）は呼び出し元で設定する必須項目です。
rem 設定されていない場合はエラーコード1で終了します。
if "%JP1_OUTPUT_MODE%"=="" exit /b 1

rem ----------------------------------------------------------------------------
rem UNCパス対応（ネットワークドライブ対応）
rem ----------------------------------------------------------------------------
rem このバッチファイルがネットワーク共有フォルダ（\\server\share\...）に
rem 置かれている場合でも正常に動作するようにします。
rem pushd は自動的に一時的なドライブ文字を割り当てます。
rem 例: \\server\share → Z: にマッピング
pushd "%~dp0"

rem ----------------------------------------------------------------------------
rem PowerShell実行
rem ----------------------------------------------------------------------------
rem このバッチファイル自体をPowerShellスクリプトとして実行します。
rem
rem オプションの説明:
rem   -NoProfile        : PowerShellプロファイルを読み込まない（起動高速化）
rem   -ExecutionPolicy Bypass : スクリプト実行ポリシーを一時的に回避
rem
rem コマンドの説明:
rem   $scriptDir=...    : バッチファイルのフォルダパスを変数に保存
rem   gc '%~f0'         : このファイル自体の内容を読み込む（gcはGet-Contentの略）
rem   -Encoding Default : Shift-JISエンコーディングで読み込む
rem   -join "`n"        : 各行を改行で連結して1つの文字列にする
rem   iex               : 読み込んだ内容をPowerShellコマンドとして実行（iexはInvoke-Expressionの略）
rem   try-finally       : エラーが発生しても必ずSet-Locationを実行
rem
powershell -NoProfile -ExecutionPolicy Bypass -Command "$scriptDir=('%~dp0' -replace '\\$',''); try { iex ((gc '%~f0' -Encoding Default) -join \"`n\") } finally { Set-Location C:\ }"

rem PowerShellの終了コードを保存
rem %ERRORLEVEL% は直前に実行したコマンドの終了コードを保持しています
set "EXITCODE=%ERRORLEVEL%"

rem 一時ドライブマッピングを解除
rem pushd で作成したマッピングを元に戻します
popd

rem PowerShellの終了コードをそのまま返す
rem 呼び出し元で終了コードを確認してエラー処理を行えます
exit /b %EXITCODE%
: #>

# ==============================================================================
# ■ JP1 REST API ジョブ情報取得ツール
# ==============================================================================
#
# 【このツールの目的】
#   JP1/AJS3で実行されたジョブの「実行結果詳細」（ログ）を取得するツールです。
#   JP1/AJS3 Web Console が提供する REST API を使用して、ジョブの標準出力や
#   標準エラー出力の内容を取得し、ファイルに保存します。
#
# 【REST APIとは？】
#   REST API は、Web経由でシステムの機能を呼び出す仕組みです。
#   このツールでは、HTTPリクエストを使ってJP1のサーバーと通信し、
#   ジョブの情報を取得しています。
#
# 【必要な環境】
#   - JP1/AJS3 - Web Console がインストール・起動されていること
#   - このツールからWeb Consoleサーバーにネットワーク接続できること
#   - JP1ユーザーアカウント（対象ジョブへの参照権限が必要）
#
# 【使い方】
#   このファイルは直接実行せず、呼び出し元のバッチファイルから使用します。
#
#   基本的な使い方（1つのジョブを取得）:
#     JP1ジョブ情報取得.bat "/JobGroup/Jobnet/Job1"
#
#   比較モード（2つのジョブを比較して新しい方を取得）:
#     JP1ジョブ情報取得.bat "/JobGroup/Jobnet/Job1" "/JobGroup/Jobnet/Job2"
#     ※ 両方のジョブの開始時刻を比較し、新しい方のログを取得します
#
# 【処理フロー】
#   このツールは以下の6つのステップで処理を行います:
#
#   STEP 1: ユニット存在確認
#           - 指定されたジョブがJP1上に存在するか確認します
#           - ジョブ（JOB系ユニット）かどうかを確認します
#
#   STEP 2: 親ジョブネットのコメント取得
#           - ジョブが属するジョブネットの説明文を取得します
#           - 出力ファイル名に使用します
#
#   STEP 3: 実行中ジョブチェック（待機機能付き）
#           - ジョブが現在実行中かどうかを確認します
#           - 実行中の場合は終了するまで待機します（設定で最大待機時間を指定可能）
#
#   STEP 4: execID（実行ID）取得
#           - ジョブの実行履歴から、ログ取得に必要な実行IDを取得します
#           - 実行IDは各実行を一意に識別する番号です
#
#   STEP 5: 実行結果詳細取得
#           - 実行IDを使って、ジョブの実行結果詳細（ログ）を取得します
#           - 取得したログをテキストファイルに保存します
#
#   STEP 6: 出力処理
#           - 設定に応じて、メモ帳で開く/Excelに貼り付けなどを行います
#
# 【終了コード一覧】
#   このツールは処理結果に応じて以下の終了コードを返します。
#   呼び出し元でこの値を確認してエラー処理を行えます。
#
#   コード | 意味                      | 発生タイミング
#   -------|---------------------------|------------------
#     0    | 正常終了                  | すべての処理が成功
#     1    | 引数エラー                | ジョブパスが指定されていない
#     2    | ユニット未検出            | STEP 1: 指定したジョブが存在しない
#     3    | ユニット種別エラー        | STEP 1: 指定したユニットがジョブではない
#     4    | 実行世代なし              | STEP 4: 実行履歴が存在しない
#     5    | 5MB超過エラー             | STEP 5: ログが大きすぎて切り捨てられた
#     6    | 詳細取得エラー            | STEP 5: ログ取得に失敗
#     8    | 比較モード失敗            | 両方のジョブ情報の取得に失敗
#     9    | API接続エラー             | Web Consoleへの接続に失敗
#    10    | Excel設定エラー           | Excel貼り付けの設定が不足
#    11    | 待機タイムアウト          | STEP 3: ジョブが終了しなかった
#    12    | Excelファイル未検出       | 指定したExcelファイルが見つからない
#
# 【公式ドキュメント】
#   JP1/AJS3 - Web Console REST API リファレンス:
#   https://itpfdoc.hitachi.co.jp/manuals/3021/30213b1920/AJSO0280.HTM
#
# ==============================================================================

# ------------------------------------------------------------------------------
# 出力エンコーディング設定
# ------------------------------------------------------------------------------
# 【この設定の意味】
# PowerShellが画面に出力する文字のエンコーディング（文字コード）を設定します。
# Shift-JIS（コードページ932）は、日本語Windowsの標準的な文字コードです。
# この設定により、コマンドプロンプトで日本語が文字化けせずに表示されます。
#
# 【補足】
# [Console]::OutputEncoding は .NET Framework のクラスで、
# コンソール出力のエンコーディングを制御します。
# GetEncoding(932) で Shift-JIS を指定しています。
[Console]::OutputEncoding = [System.Text.Encoding]::GetEncoding(932)

# タイトル表示
Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  JP1ジョブ情報取得ツール" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

# ==============================================================================
# ■ 接続設定セクション
# ==============================================================================
# 【このセクションについて】
# JP1/AJS3 Web Console への接続に必要な情報を設定します。
# 環境に合わせてこのセクションの値を変更してください。
#
# 【設定変更のポイント】
# 1. まず $webConsoleHost と $webConsolePort を環境に合わせて設定
# 2. 認証情報は Windows 資格情報マネージャーに登録するのがおすすめ
# 3. テスト環境では $useHttps = $false（HTTP接続）でOK
# ==============================================================================

# ------------------------------------------------------------------------------
# Web Consoleサーバー設定
# ------------------------------------------------------------------------------
# 【Web Consoleサーバーとは？】
# JP1/AJS3 - Web Console は、JP1のジョブ管理をWebブラウザで操作できる
# コンポーネントです。このツールはWeb ConsoleのREST APIを使用します。
#
# 【設定方法】
# Web Consoleサーバーのホスト名またはIPアドレスを指定します。
# - 同じPCにある場合: "localhost"
# - 別のサーバーの場合: サーバー名またはIPアドレス
#
# 設定例:
#   $webConsoleHost = "localhost"           # 同じPCの場合
#   $webConsoleHost = "192.168.1.100"       # IPアドレス指定
#   $webConsoleHost = "jp1server.example.com"  # ホスト名指定
$webConsoleHost = "localhost"

# 【ポート番号】
# Web Consoleが待ち受けているポート番号を指定します。
# インストール時のデフォルト値:
#   - HTTP接続の場合:  22252
#   - HTTPS接続の場合: 22253
# ※ インストール時に変更している場合は、その値を設定してください
$webConsolePort = "22252"

# 【HTTPS（暗号化通信）の使用】
# 通信を暗号化するかどうかを設定します。
#
# $false（HTTP接続）:
#   - 暗号化なし。通信内容が傍受される可能性があります
#   - 社内ネットワークなど安全な環境向け
#   - 設定が簡単（証明書不要）
#
# $true（HTTPS接続）:
#   - 暗号化あり。通信内容が保護されます
#   - インターネット経由やセキュリティ重視の環境向け
#   - SSL証明書の設定が必要
$useHttps = $false

# ------------------------------------------------------------------------------
# JP1/AJS3 Manager設定
# ------------------------------------------------------------------------------
# 【JP1/AJS3 Managerとは？】
# JP1/AJS3 Manager は、ジョブの定義や実行を管理するサーバーです。
# Web Console経由でManagerに接続し、ジョブ情報を取得します。
#
# 【設定方法】
# JP1/AJS3 Managerのホスト名またはIPアドレスを指定します。
# Web ConsoleとManagerが同じサーバーにある場合は "localhost" でOKです。
$managerHost = "localhost"

# 【スケジューラーサービス名】
# JP1/AJS3では複数のスケジューラーサービスを運用できます。
# デフォルトのサービス名は "AJSROOT1" です。
#
# 確認方法:
#   - JP1/AJS3 View でルートジョブグループ名を確認
#   - 通常は "AJSROOT1" から始まる名前
#
# 設定例:
#   $schedulerService = "AJSROOT1"        # デフォルト
#   $schedulerService = "AJSROOT2"        # 2つ目のサービス
#   $schedulerService = "PRODUCTION"      # カスタム名
$schedulerService = "AJSROOT1"

# ------------------------------------------------------------------------------
# 認証設定
# ------------------------------------------------------------------------------
# 【認証の仕組み】
# JP1/AJS3 Web Console REST API は、ユーザー名とパスワードで認証します。
# 認証情報は以下の優先順位で取得されます:
#   1. このファイルに直接記載された値
#   2. Windows 資格情報マネージャーに保存された値
#   3. 実行時に入力プロンプトで入力
#
# 【セキュリティ上の注意】
# パスワードをこのファイルに直接記載すると、ファイルが流出した際に
# パスワードも漏洩します。Windows 資格情報マネージャーの使用を推奨します。
#
# 【JP1ユーザー名】
# JP1/AJS3にログインするためのユーザー名です。
# このユーザーには、対象ジョブへの「参照権限」が必要です。
# ★ 空欄にすると、資格情報マネージャー → 入力プロンプトの順で取得します
$jp1User = ""

# 【JP1パスワード】
# JP1ユーザーのパスワードです。
# ★ セキュリティのため、空欄にしてWindows資格情報マネージャーの使用を推奨
$jp1Password = ""

# 【Windows資格情報マネージャーのターゲット名】
# 資格情報マネージャーに登録する際の「ターゲット名」を指定します。
#
# 【資格情報の登録方法】
# 方法1: コマンドで登録
#   cmdkey /generic:JP1_WebConsole /user:jp1admin /pass:yourpassword
#
# 方法2: GUIで登録
#   1. コントロールパネル → ユーザーアカウント → 資格情報マネージャー
#   2. 「Windows資格情報」→「汎用資格情報の追加」
#   3. インターネットまたはネットワークのアドレス: JP1_WebConsole
#   4. ユーザー名とパスワードを入力
$credentialTarget = "JP1_WebConsole"

# ------------------------------------------------------------------------------
# 実行中ジョブ待機設定
# ------------------------------------------------------------------------------
# 【この設定の目的】
# ジョブが現在実行中の場合、終了するまで待機してからログを取得できます。
# 実行中のジョブには実行結果詳細（ログ）がまだ存在しないためです。
#
# 【最大待機秒数】
# ジョブの終了を待つ最大時間を秒単位で指定します。
# - 0を指定すると、実行中の場合は待機せずにエラー終了します
# - この時間を超えても終了しない場合は、エラーコード11で終了します
#
# 設定例:
#   $maxWaitSeconds = 0     # 待機しない
#   $maxWaitSeconds = 60    # 最大1分待機（デフォルト）
#   $maxWaitSeconds = 300   # 最大5分待機
$maxWaitSeconds = 60

# 【チェック間隔】
# ジョブが終了したかどうかを確認する間隔を秒単位で指定します。
# 短くすると応答が早くなりますが、サーバーへの負荷が増えます。
#
# 設定例:
#   $checkIntervalSeconds = 5    # 5秒ごとにチェック
#   $checkIntervalSeconds = 10   # 10秒ごとにチェック（デフォルト）
#   $checkIntervalSeconds = 30   # 30秒ごとにチェック
$checkIntervalSeconds = 10

# ------------------------------------------------------------------------------
# 出力設定
# ------------------------------------------------------------------------------
# 【出力先フォルダ】
# 取得したログファイルを保存するフォルダを指定します。
#
# 指定方法:
#   - 相対パス: このスクリプトがあるフォルダからの相対位置
#   - フルパス: ドライブ名から始まる絶対パス
#
# ※ フォルダが存在しない場合は自動的に作成されます
#
# 設定例:
#   $outputFolder = "..\02_output"          # 相対パス（1つ上の階層の02_outputフォルダ）
#   $outputFolder = "C:\Logs\JP1"           # フルパス
#   $outputFolder = ".\Output"              # 同じフォルダ内のOutputフォルダ
$outputFolder = "..\02_output"

# 【出力ファイル名のプレフィックス】
# 出力ファイル名の先頭に付ける文字列を指定します。
#
# 最終的なファイル名の形式:
#   {プレフィックス}【{開始日時}実行分】【{終了状態}】{ジョブネット名}_{コメント}.txt
#
# 出力例:
#   【ジョブ実行結果】【20250111_093000実行分】【正常終了】日次バッチ_売上集計処理.txt
$outputFilePrefix = "【ジョブ実行結果】"

# ==============================================================================
# ■ 検索条件設定セクション
# ==============================================================================
# 【このセクションについて】
# ジョブの実行履歴を検索する際の条件を設定します。
# 通常はデフォルト値のままで問題ありませんが、特定の条件でログを
# 取得したい場合にこれらの設定を変更します。
#
# 【よく使う設定パターン】
# - 最新の実行結果を取得したい → $generation = "RESULT"（デフォルト）
# - エラーになったジョブだけ見たい → $statusFilter = "ABNORMAL"
# - 特定期間のログを取得したい → $generation = "PERIOD" + 期間指定
# ==============================================================================

# ------------------------------------------------------------------------------
# (1) GenerationType - 世代指定
# ------------------------------------------------------------------------------
# 【世代とは？】
# JP1/AJS3では、ジョブは同じ定義で何度も実行されます。
# 各実行を「世代」と呼び、それぞれに実行ID（execID）が割り当てられます。
# この設定で、どの世代のログを取得するかを指定します。
#
# 【指定可能な値】
#
# "NO" - 世代を検索条件にしない
#        すべての世代が検索対象になります
#
# "STATUS" - 最新状態の世代を取得
#            現在表示されている状態の世代を取得します
#            （JP1/AJS3 View で見える状態と同じ）
#
# "RESULT" - 最新結果の世代を取得（★推奨★）
#            直近で終了したジョブの実行結果を取得します
#            通常はこの設定がおすすめです
#
# "PERIOD" - 指定期間の世代を取得
#            下の $periodBegin と $periodEnd で指定した期間に
#            実行されたジョブを取得します
#
# "EXECID" - 特定の実行IDを取得
#            下の $execID で指定した特定の実行を取得します
#            過去の特定の実行を調べたい場合に使用
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
# ★ 値: Excelファイル名、貼り付けシート名、貼り付けセル、移動先シート名をカンマ区切りで指定
# 例: "/グループ/ネット/ジョブ" = "ファイル名.xlsx,Sheet2,A1,Sheet1"
#     → Sheet2のA1に貼り付け後、Sheet1のA1に移動して保存
$jobExcelMapping = @{
    # === ジョブパスとExcelファイルのマッピング ===
    # 以下に「ジョブパス」=「Excelファイル名,貼り付けシート名,貼り付けセル,移動先シート名」の形式で記載
    # 完全一致で検索されるため、呼び出し元で指定するパスと同じ形式で記載してください
    #
    # 例:
    # "/TIA/Jobnet/週単位ジョブ" = "TIA解析(自習当初)_週単位.xls,Sheet2,A1,Sheet1"
    # "/TIA/Jobnet/年単位ジョブ" = "TIA解析(自習当初)_年単位.xls,Sheet2,A1,Sheet1"
    #
    # ★ 以下を編集してください ★
    "/サンプル/Jobnet/週単位ジョブ" = "TIA解析(自習当初)_週単位.xls,Sheet2,A1,Sheet1"
    "/サンプル/Jobnet/年単位ジョブ" = "TIA解析(自習当初)_年単位.xls,Sheet2,A1,Sheet1"
    # "/グループ/ネット/ジョブ3" = "Excelファイル3.xls,Sheet2,A1,Sheet1"
    # "/グループ/ネット/ジョブ4" = "Excelファイル4.xls,Sheet2,A1,Sheet1"
    # "/グループ/ネット/ジョブ5" = "Excelファイル5.xls,Sheet2,A1,Sheet1"
    # "/グループ/ネット/ジョブ6" = "Excelファイル6.xls,Sheet2,A1,Sheet1"
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
# 【このセクションについて】
# ここから先は実際の処理を行う部分です。通常は編集する必要はありません。
# 上部の設定セクションを変更することで、動作をカスタマイズできます。
#
# 【処理の流れ】
# 1. 環境変数から引数（ジョブパス）を取得
# 2. 認証情報を取得（設定 → 資格情報マネージャー → 入力プロンプト）
# 3. REST APIリクエストの準備（ヘッダー、URL等）
# 4. STEP 1〜6 の処理を順番に実行
# ==============================================================================

# ------------------------------------------------------------------------------
# 環境変数からユニットパスを取得
# ------------------------------------------------------------------------------
# 【この処理の目的】
# バッチファイルから渡された引数（ジョブのパス）を取得します。
# バッチファイル部分で環境変数に設定した値を、ここで読み取ります。
#
# 【環境変数とは？】
# 環境変数は、プログラム間でデータを受け渡すための仕組みです。
# バッチファイル（set コマンド）で設定した値を、
# PowerShell（$env:変数名）で取得できます。
$unitPath = $env:JP1_UNIT_PATH
$unitPath2 = $env:JP1_UNIT_PATH_2

# ------------------------------------------------------------------------------
# 比較モードの判定
# ------------------------------------------------------------------------------
# 【比較モードとは？】
# 2つのジョブパスが指定された場合、「比較モード」として動作します。
# 比較モードでは、2つのジョブの開始時刻を比較し、
# より新しい（最近実行された）方のログを取得します。
#
# 【使用例】
# 同じ処理を行う2つのジョブがあり、どちらか新しい方のログが欲しい場合:
#   JP1ジョブ情報取得.bat "/日次バッチ/集計A" "/日次バッチ/集計B"
$isCompareMode = $false
if ($unitPath2 -and $unitPath2.Trim() -ne "") {
    $isCompareMode = $true
}

# 対象ジョブの表示
if ($isCompareMode) {
    Write-Host "  対象1: $unitPath"
    Write-Host "  対象2: $unitPath2"
    Write-Host ""
    Write-Host "  [比較モード] 新しい方のログを取得します" -ForegroundColor Yellow
} else {
    Write-Host "  対象: $unitPath"
}
Write-Host ""

# ------------------------------------------------------------------------------
# プロトコル設定
# ------------------------------------------------------------------------------
# 【この処理の目的】
# HTTPまたはHTTPSのどちらで通信するかを決定します。
# 設定セクションの $useHttps の値に基づいて、URLの先頭部分を切り替えます。
#
# 【HTTP と HTTPS の違い】
# - HTTP  : 暗号化なしの通信。設定が簡単。社内ネットワーク向け。
# - HTTPS : 暗号化あり（SSL/TLS）。通信が傍受されにくい。
$protocol = if ($useHttps) { "https" } else { "http" }

# ------------------------------------------------------------------------------
# Windows資格情報マネージャーからの認証情報取得
# ------------------------------------------------------------------------------
# 【この処理の目的】
# JP1ユーザーのログイン情報（ユーザー名とパスワード）を取得します。
#
# 【認証情報の取得順序】
# 1. 設定セクションに直接書かれた値を使用
# 2. 空欄の場合 → Windows資格情報マネージャーから取得
# 3. それでも取得できない場合 → ユーザーに入力を求める
#
# 【Windows資格情報マネージャーとは？】
# Windowsに組み込まれたパスワード管理機能です。
# パスワードを暗号化して保存できるため、スクリプトに直接書くより安全です。
# コントロールパネル → 資格情報マネージャー から確認・登録できます。

if (-not $jp1User -or -not $jp1Password) {
    # Windows API (CredRead) を使用して資格情報を取得
    # 【技術的な補足】
    # PowerShellから直接Windows APIを呼び出すため、
    # C#コードを動的にコンパイルして使用しています。
    # advapi32.dll はWindowsのセキュリティ関連APIを提供するDLLです。
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

# ------------------------------------------------------------------------------
# 入力プロンプトでの認証情報取得（フォールバック）
# ------------------------------------------------------------------------------
# 【この処理の目的】
# 設定ファイルにも資格情報マネージャーにも認証情報がない場合、
# ユーザーに直接入力を求めます。
#
# 【セキュリティ配慮】
# パスワード入力時は -AsSecureString オプションを使用して、
# 入力内容が画面に表示されないようにしています（**** と表示される）。
if (-not $jp1User) {
    $jp1User = Read-Host "JP1ユーザー名を入力してください"
}
if (-not $jp1Password) {
    $securePass = Read-Host "JP1パスワードを入力してください" -AsSecureString
    # SecureString を通常の文字列に変換（API呼び出しに必要）
    $jp1Password = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePass))
}

# ------------------------------------------------------------------------------
# 認証情報の作成（Base64エンコード）
# ------------------------------------------------------------------------------
# 【この処理の目的】
# REST API認証用のヘッダー値を作成します。
#
# 【Base64エンコードとは？】
# 文字列をアルファベット・数字・一部記号だけで表現する変換方式です。
# HTTP通信で特殊文字を含むデータを安全に送るために使われます。
# ※ 暗号化ではないので、傍受されると解読される可能性があります。
#    そのため、HTTPSの使用が推奨されます。
#
# 【処理の流れ】
# 1. "ユーザー名:パスワード" の形式で文字列を作成
# 2. UTF-8でバイト配列に変換
# 3. Base64でエンコード
$authString = "${jp1User}:${jp1Password}"
$authBytes = [System.Text.Encoding]::UTF8.GetBytes($authString)
$authBase64 = [System.Convert]::ToBase64String($authBytes)

# ------------------------------------------------------------------------------
# HTTPリクエストヘッダーの設定
# ------------------------------------------------------------------------------
# 【この処理の目的】
# REST APIリクエストに付加するヘッダー情報を設定します。
#
# 【各ヘッダーの意味】
# Accept-Language: "ja"
#   → APIからの応答を日本語で受け取る指定
#
# X-AJS-Authorization: (Base64エンコードされた認証情報)
#   → JP1 Web Console独自の認証ヘッダー
#   → これがないと「認証エラー」で拒否されます
$headers = @{
    "Accept-Language" = "ja"
    "X-AJS-Authorization" = $authBase64
}

# ------------------------------------------------------------------------------
# SSL証明書検証の設定（HTTPS使用時）
# ------------------------------------------------------------------------------
# 【この処理の目的】
# HTTPS通信時のSSL証明書検証ポリシーを設定します。
#
# 【SSL証明書とは？】
# サーバーの身元を証明するための電子証明書です。
# 正規の認証局（CA）から発行された証明書を使うと、
# 「このサーバーは本物です」ということが保証されます。
#
# 【自己署名証明書について】
# 開発・テスト環境では、正規の証明書を取得せずに
# 「自己署名証明書」を使うことがあります。
# 自己署名証明書は検証に失敗するため、このコードで検証をスキップします。
#
# 【注意】
# 本番環境では正規の証明書を使用し、このスキップ処理は無効にすることを推奨。
# 証明書検証をスキップすると、中間者攻撃のリスクがあります。
if ($useHttps) {
    # C#コードを動的にコンパイルして証明書検証ポリシーを定義
    Add-Type @"
        using System.Net;
        using System.Security.Cryptography.X509Certificates;
        public class TrustAllCertsPolicy : ICertificatePolicy {
            public bool CheckValidationResult(
                ServicePoint srvPoint, X509Certificate certificate,
                WebRequest request, int certificateProblem) {
                return true;  // すべての証明書を信頼（検証スキップ）
            }
        }
"@
    # 証明書検証ポリシーを適用
    [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
    # TLS 1.2を使用（古いSSL/TLSバージョンはセキュリティ上問題がある）
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
}

# ------------------------------------------------------------------------------
# ベースURLの構築
# ------------------------------------------------------------------------------
# 【この処理の目的】
# REST API呼び出しで使用する基本URLを作成します。
#
# 【URLの構成】
# {プロトコル}://{ホスト}:{ポート}/ajs/api/v1
#
# 例:
#   http://localhost:22252/ajs/api/v1
#   https://jp1server.example.com:22253/ajs/api/v1
#
# 【/ajs/api/v1 について】
# これはJP1/AJS3 Web ConsoleのREST APIのベースパスです。
# "v1" はAPIのバージョンを示しています。
$baseUrl = "${protocol}://${webConsoleHost}:${webConsolePort}/ajs/api/v1"

# ==============================================================================
# ■ ユーティリティ関数（メイン処理の前に定義が必要）
# ==============================================================================
# 【このセクションについて】
# メイン処理で使用する補助的な関数を定義しています。
# PowerShellでは、関数は呼び出される前に定義されている必要があるため、
# ここで先に定義しています。
#
# 【関数とは？（初心者向け）】
# 関数は「よく使う処理をまとめたもの」です。
# 同じ処理を何度も書く代わりに、関数名で呼び出すことができます。
# 例: Get-StatusDisplayName "NORMAL"  →  "正常終了" を返す
# ==============================================================================

# ------------------------------------------------------------------------------
# Write-Console関数 - コンソールへの直接出力
# ------------------------------------------------------------------------------
# 【この関数の目的】
# 画面（コンソール）にメッセージを表示します。
#
# 【なぜ Write-Host ではなく専用関数を使うのか？】
# このスクリプトが別のスクリプトから呼ばれた場合、
# 標準出力がファイルにリダイレクト（> output.txt）されることがあります。
# その場合でも、進捗メッセージは画面に表示したいため、
# 特殊なデバイス名「CON」に書き込んでリダイレクトを回避しています。
#
# 【CONデバイスとは？】
# Windowsでは「CON」は「コンソール（画面）」を表す特殊なファイル名です。
# ここに書き込むと、常に画面に表示されます。
function Write-Console {
    param([string]$Message)
    [Console]::WriteLine($Message)
}

# ------------------------------------------------------------------------------
# Get-StatusDisplayName関数 - ステータス値の日本語変換
# ------------------------------------------------------------------------------
# 【この関数の目的】
# JP1のステータスコード（英語）を日本語の表示名に変換します。
#
# 【使用例】
# Get-StatusDisplayName "NORMAL"    →  "正常終了"
# Get-StatusDisplayName "ABNORMAL"  →  "異常検出終了"
#
# 【ステータスコード一覧】
# JP1/AJS3では、ジョブの終了状態を英語のコードで管理しています。
# ユーザーに分かりやすく表示するため、日本語に変換しています。
function Get-StatusDisplayName {
    param([string]$status)
    switch ($status) {
        # --- 正常系の終了状態 ---
        "NORMAL"        { return "正常終了" }           # 正常に終了
        "NORMALFALSE"   { return "正常終了-偽" }        # 判定ジョブで条件不成立

        # --- 警告系の終了状態 ---
        "WARNING"       { return "警告検出終了" }       # 警告があったが終了

        # --- 異常系の終了状態 ---
        "ABNORMAL"      { return "異常検出終了" }       # エラーで終了
        "KILL"          { return "強制終了" }           # 強制的に停止された
        "INTERRUPT"     { return "中断" }               # ユーザーが中断した
        "FAIL"          { return "起動失敗" }           # 起動自体に失敗
        "UNKNOWN"       { return "終了状態不正" }       # 状態が不明
        "INVALIDSEQ"    { return "順序不正" }           # 実行順序に問題

        # --- 監視系の終了状態 ---
        "MONITORCLOSE"  { return "監視打ち切り終了" }   # 監視タイムアウト
        "UNEXECMONITOR" { return "監視未起動終了" }     # 監視対象が起動しなかった
        "MONITORINTRPT" { return "監視中断" }           # 監視が中断された
        "MONITORNORMAL" { return "監視正常終了" }       # 監視が正常に終了

        # --- 実行中の状態 ---
        "RUNNING"       { return "実行中" }             # 現在実行中
        "WACONT"        { return "警告検出実行中" }     # 警告があるが実行継続中
        "ABCONT"        { return "異常検出実行中" }     # 異常があるが実行継続中

        # --- その他 ---
        default         { return $status }               # 未定義のコードはそのまま返す
    }
}

# ==============================================================================
# 2引数モード: 実行中チェック＆START_TIME比較処理
# ==============================================================================
# 【このセクションについて】
# 2つのジョブパスが指定された場合（比較モード）の処理を行います。
#
# 【比較モードの処理フロー】
# 1. まず両方のジョブが実行中かどうかをチェック
# 2. 実行中のジョブがあれば、終了を待機してそのジョブを最新として選択
# 3. どちらも実行中でなければ、開始時刻（START_TIME）を比較
# 4. より新しい方のジョブを選択してログを取得
#
# 【使用例】
# 例えば、同じ処理を行う2つのジョブがあり、
# どちらか片方が実行されたときにそのログを見たい場合:
#   JP1ジョブ情報取得.bat "/日次バッチ/東京/売上集計" "/日次バッチ/大阪/売上集計"
#   → 直近で実行された方のログを自動で取得
# ==============================================================================

# --- 比較モード用の変数を初期化 ---
$selectedPath = ""  # 比較モードで選択されたジョブのパス
$selectedTime = ""  # 比較モードで選択されたジョブの開始時刻
$rejectedPath = ""  # 比較モードで選択されなかったジョブのパス
$rejectedTime = ""  # 比較モードで選択されなかったジョブの開始時刻

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
            Write-Console "[情報] 両方のジョブが実行中です。$unitPath の終了を待機します"
        } else {
            Write-Console "[情報] 実行中のジョブがあります - $waitTargetPath の終了を待機します"
        }

        # 待機ループ
        $waitedSeconds = 0
        $stillRunning = $true
        while ($stillRunning) {
            # 最大待機秒数を超えた場合はエラー終了
            if ($waitedSeconds -ge $maxWaitSeconds) {
                Write-Host ""
                Write-Host "[タイムアウト] 待機時間を超過しました" -ForegroundColor Red
                Write-Host "  ジョブ: $waitTargetPath"
                Write-Host "  ステータス: ${waitTargetStatusDisplay}"
                Write-Host "  開始日時: $($waitTargetStatus.StartTime)"
                Write-Host "  待機時間: ${maxWaitSeconds}秒"
                Write-Host ""
                exit 11  # 実行中のジョブが検出された（タイムアウト）
            }

            Write-Console "[待機中] ジョブ終了を待っています...（${waitedSeconds}/${maxWaitSeconds}秒）"
            Write-Console "         $waitTargetPath（${waitTargetStatusDisplay}）"

            Start-Sleep -Seconds $checkIntervalSeconds
            $waitedSeconds += $checkIntervalSeconds

            # 再度チェック
            $recheckStatus = Get-JobRunningStatus -jobPath $waitTargetPath
            if (-not $recheckStatus -or -not $recheckStatus.IsRunning) {
                $stillRunning = $false
                Write-Console "[完了] ジョブが終了しました（${waitedSeconds}秒待機）"
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
        $rejectedTime = "(実行中のジョブを優先)"

        Write-Console "[情報] 待機していたジョブのログを取得します"
    } else {
        # どちらも実行中でない場合はSTART_TIMEで比較
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
# 【この処理の目的】
# ユーザーが指定したジョブパスを分解して、API呼び出しに必要な情報を抽出します。
#
# 【JP1のパス構造】
# JP1/AJS3では、ジョブは階層構造で管理されています。
#
# 例: /業務システム/日次バッチ/データ集計/売上集計ジョブ
#      ↑           ↑         ↑          ↑
#      ルート       ジョブ     ジョブ      ジョブ
#      ジョブ       グループ   ネット      （実際のジョブ）
#      グループ
#
# 【分解後のデータ】
# - parentPath: "/業務システム/日次バッチ/データ集計" （ジョブの親）
# - jobName: "売上集計ジョブ" （ジョブ名）
# - grandParentPath: "/業務システム/日次バッチ" （ジョブネットの親）
# - jobnetName: "データ集計" （ジョブネット名）

# 最後のスラッシュの位置を見つけて、親パスとジョブ名に分割
$lastSlashIndex = $unitPath.LastIndexOf("/")
if ($lastSlashIndex -le 0) {
    exit 1  # パス形式エラー（スラッシュがない、またはルートのみ）
}
$parentPath = $unitPath.Substring(0, $lastSlashIndex)   # ジョブの親パス
$jobName = $unitPath.Substring($lastSlashIndex + 1)     # ジョブ名

if (-not $jobName) {
    exit 1  # ジョブ名が空
}

# 親ジョブネット名を取得（出力ファイル名に使用）
$grandParentSlashIndex = $parentPath.LastIndexOf("/")
$grandParentPath = if ($grandParentSlashIndex -gt 0) { $parentPath.Substring(0, $grandParentSlashIndex) } else { "/" }
$jobnetName = if ($grandParentSlashIndex -ge 0) { $parentPath.Substring($grandParentSlashIndex + 1) } else { $parentPath.TrimStart("/") }

# ------------------------------------------------------------------------------
# URLエンコード
# ------------------------------------------------------------------------------
# 【この処理の目的】
# パス名をURLで使用できる形式に変換します。
#
# 【URLエンコードとは？】
# URLでは、日本語やスペースなどの特殊文字をそのまま使えません。
# 例えば "/" は "%2F" に、日本語は "%E6%97%A5..." のように変換されます。
# これにより、どんな文字を含むパスでも安全にAPI呼び出しができます。
$encodedParentPath = [System.Uri]::EscapeDataString($parentPath)
$encodedJobName = [System.Uri]::EscapeDataString($jobName)

# ==============================================================================
# STEP 1: ユニット存在確認・種別確認（DEFINITION）
# ==============================================================================
# 【このステップの目的】
# ログ取得を行う前に、指定されたジョブが有効かどうかを確認します。
# 存在しないジョブや、ジョブ以外のユニット（ジョブネット等）を指定した場合、
# 早期にエラーとして処理を中断します。
#
# 【確認内容】
# 1. 指定したパスにユニットが存在するか
# 2. 指定したユニットがジョブ（JOB系）かどうか
#    - ジョブネットやジョブグループは対象外
#    - JOB, PJOB(PCジョブ), QJOB(キュージョブ)など
#
# 【searchTarget=DEFINITION について】
# DEFINITIONは「定義情報のみ」を検索するモードです。
# 実行状態に関係なく、ジョブが定義されているかどうかを確認できます。
# （実行履歴がなくても見つかる）
#
# 【エラー時の終了コード】
# - 終了コード 2: ユニットが見つからない
# - 終了コード 3: 指定されたユニットがジョブではない

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
# STEP 2: 親ジョブネットのコメント取得
# ==============================================================================
# 【このステップの目的】
# ジョブが属する親ジョブネットの「コメント」を取得します。
# 取得したコメントは、出力ファイル名の一部として使用されます。
#
# 【コメントとは？】
# JP1/AJS3では、ジョブネットに「コメント」（説明文）を設定できます。
# 例: ジョブネット名「DAILY_SALES」、コメント「日次売上集計処理」
# このコメントを取得してファイル名に含めることで、
# ファイル名だけでどの処理のログか分かるようになります。
#
# 【出力ファイル名の例】
# 【ジョブ実行結果】【20250111_093000実行分】【正常終了】DAILY_SALES_日次売上集計処理.txt
#                                            ↑           ↑
#                                            ジョブネット名  コメント
#
# 【エラー時の動作】
# コメント取得に失敗しても処理は続行します（必須情報ではないため）。
# その場合、コメント部分は空になります。

Write-Console "[STEP 2] 親ジョブネットのコメント取得中..."

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
# STEP 3: 実行中ジョブチェック（待機機能付き）
# ==============================================================================
# 【このステップの目的】
# 対象のジョブが現在実行中かどうかを確認し、
# 実行中であれば終了するまで待機します。
#
# 【なぜ待機が必要か？】
# 実行中のジョブには「実行結果詳細」（ログ）がまだ存在しません。
# ログはジョブが終了してから確定するためです。
# そのため、実行中のジョブを検出した場合は終了を待つ必要があります。
#
# 【待機の仕組み】
# 1. まずジョブが実行中（GRP_RUN状態）かどうかをチェック
# 2. 実行中なら、設定した間隔（$checkIntervalSeconds）で再チェック
# 3. 実行中でなくなるか、最大待機時間（$maxWaitSeconds）を超えるまで繰り返す
#
# 【待機中の表示例】
# WAITING:実行中のジョブを検出しました。終了を待機しています...（10/60秒）
# WAITING_JOB:/日次バッチ/売上集計（ステータス: 実行中, 開始日時: 2025-01-11T09:30:00）
#
# 【エラー時の終了コード】
# - 終了コード 11: 待機タイムアウト（最大待機時間を超えても終了しなかった）

Write-Console "[STEP 3] 実行中ジョブチェック中..."

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
                Write-Host ""
                Write-Host "[タイムアウト] 待機時間を超過しました" -ForegroundColor Red
                Write-Host "  ジョブ: $unitPath"
                Write-Host "  ステータス: ${runningStatusDisplay}"
                Write-Host "  開始日時: ${runningStartTime}"
                Write-Host "  待機時間: ${maxWaitSeconds}秒"
                Write-Host ""
                exit 11  # 実行中のジョブが検出された（タイムアウト）
            }

            # 待機中メッセージを出力（コンソールへ直接表示）
            Write-Console "[待機中] ジョブ終了を待っています...（${waitedSeconds}/${maxWaitSeconds}秒）"
            Write-Console "         $unitPath（${runningStatusDisplay}）"

            # 指定秒数待機
            Start-Sleep -Seconds $checkIntervalSeconds
            $waitedSeconds += $checkIntervalSeconds
        } else {
            # 実行中ではない → ループを抜ける
            $isRunning = $false

            # 待機していた場合は完了メッセージを出力
            if ($waitedSeconds -gt 0) {
                Write-Console "[完了] ジョブが終了しました（${waitedSeconds}秒待機）"
            }
        }
    } catch {
        # 実行中チェック失敗は無視して続行（必須ではない）
        $isRunning = $false
    }
}

# ==============================================================================
# STEP 4: 実行状態・execID取得（DEFINITION_AND_STATUS）
# ==============================================================================
# 【このステップの目的】
# ジョブの実行履歴から、ログ取得に必要な「実行ID（execID）」を取得します。
#
# 【execID（実行ID）とは？】
# JP1/AJS3では、ジョブを実行するたびに一意の実行IDが割り当てられます。
# 例: @A100, @B200, @10A300 など
# この実行IDを使って、特定の実行のログを取得します。
#
# 【searchTarget=DEFINITION_AND_STATUS について】
# DEFINITION_AND_STATUS は「定義情報と実行状態」の両方を検索するモードです。
# 実行履歴がある場合のみヒットします。
# （一度も実行されていないジョブは見つかりません）
#
# 【待機していた場合の動作】
# STEP 3でジョブ終了を待機していた場合、そのジョブのexecIDをすでに取得しています。
# その場合は、設定ファイルの世代指定（$generation）を無視して、
# 待機していたexecIDを直接使用します。
# これにより、待機完了直後に確実にそのジョブのログを取得できます。
#
# 【エラー時の終了コード】
# - 終了コード 4: 実行世代なし（一度も実行されていない、または条件に合う実行がない）

Write-Console "[STEP 4] execID取得中..."

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
    Write-Console "INFO:待機していたジョブのexecID（${waitingExecId}）を使用してログを取得します"
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
# STEP 5: 実行結果詳細取得API
# ==============================================================================
# 【このステップの目的】
# STEP 4で取得した実行ID（execID）を使って、
# ジョブの「実行結果詳細」（ログ）を取得します。
#
# 【実行結果詳細とは？】
# ジョブ実行時に出力された内容で、以下が含まれます：
# - 標準出力（コマンドの実行結果など）
# - 標準エラー出力（エラーメッセージなど）
# - JP1/AJS3が付加する実行情報
#
# 【API呼び出しの形式】
# /objects/statuses/{ユニットパス}:{execID}/actions/execResultDetails/invoke
#
# 例:
# /objects/statuses/%2F日次バッチ%2F売上集計:@A100/actions/execResultDetails/invoke
# (%2F は "/" のURLエンコード)
#
# 【ファイル出力】
# 取得したログは、設定セクションで指定したフォルダにテキストファイルとして保存されます。
# ファイル名形式: {プレフィックス}【{開始日時}実行分】【{終了状態}】{ジョブネット名}_{コメント}.txt
#
# 【5MB制限について】
# REST APIでは、実行結果詳細の取得サイズに5MBの上限があります。
# 5MBを超えるログは切り捨てられ、"all"フラグがfalseになります。
# その場合は終了コード5でエラー終了します。
#
# 【エラー時の終了コード】
# - 終了コード 5: 5MB超過エラー（ログが大きすぎて切り捨てられた）
# - 終了コード 6: 詳細取得エラー（API呼び出しに失敗）

Write-Console "[STEP 5] 実行結果詳細取得中..."

# 出力用の変数を初期化
$outputFilePath = ""
$startTimeForFileName = ""
$endStatusDisplay = ""

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

            # 終了状態を取得（日本語変換済み）
            $endStatusDisplay = Get-StatusDisplayName -status $targetEndStatus

            # NOTEPADモード時のみファイル出力
            if ($env:JP1_OUTPUT_MODE -eq $null -or $env:JP1_OUTPUT_MODE.ToUpper() -eq "/NOTEPAD") {
                # 出力ディレクトリを作成（設定セクションの$outputFolderを使用）
                if ([System.IO.Path]::IsPathRooted($outputFolder)) {
                    $outputDir = $outputFolder
                } else {
                    $outputDir = Join-Path $scriptDir $outputFolder
                }
                if (-not (Test-Path $outputDir)) {
                    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
                }

                # 出力ファイル名を生成（設定セクションの$outputFilePrefixを使用）
                $outputFileName = "${outputFilePrefix}【${startTimeForFileName}実行分】【${endStatusDisplay}】${jobnetName}_${jobnetComment}.txt"
                $outputFilePath = Join-Path $outputDir $outputFileName

                # 実行結果詳細をファイルに出力
                $execResultContent | Out-File -FilePath $outputFilePath -Encoding Default
            }

        } catch {
            exit 6  # 詳細取得エラー
        }
    }
}

# ------------------------------------------------------------------------------
# 比較結果の表示
# ------------------------------------------------------------------------------
if ($selectedPath) {
    Write-Host ""
    Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan
    Write-Host "  比較結果" -ForegroundColor Cyan
    Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  新しい方のジョブを選択しました"
    Write-Host ""
    Write-Host "    選択: $selectedPath" -ForegroundColor Green
    Write-Host "          開始日時: $selectedTime"
    Write-Host ""
    Write-Host "    除外: $rejectedPath" -ForegroundColor DarkGray
    # 取得失敗の場合は赤色で表示
    if ($rejectedTime -eq "(取得失敗)") {
        Write-Host "          開始日時: $rejectedTime" -ForegroundColor Red
    } else {
        Write-Host "          開始日時: $rejectedTime"
    }
    Write-Host ""
    Write-Host "  続行する場合は任意のキーを押してください..."
    Write-Host ""
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# ------------------------------------------------------------------------------
# 2日以上前のお知らせ
# ------------------------------------------------------------------------------
if ($startTimeForFileName) {
    try {
        $dt = [DateTime]::ParseExact($startTimeForFileName, "yyyyMMdd_HHmmss", $null)
        $daysDiff = ((Get-Date) - $dt).TotalDays
        if ($daysDiff -ge 2) {
            Write-Host ""
            Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan
            Write-Host "  ご確認ください" -ForegroundColor Cyan
            Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan
            Write-Host ""
            Write-Host "  ジョブ開始日時が2日以上前です"
            Write-Host ""
            Write-Host "    開始日時: $startTimeForFileName"
            Write-Host ""
            Write-Host "  意図した世代のログかご確認ください。"
            Write-Host "  続行する場合は任意のキーを押してください..."
            Write-Host ""
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        }
    } catch {
        # 日付パースエラーは無視
    }
}

# ==============================================================================
# STEP 6: 出力処理
# ==============================================================================
# 【このステップの目的】
# STEP 5で保存したログファイルを、ユーザーが指定した方法で表示・出力します。
#
# 【出力オプション】
# 呼び出し元のバッチファイルで環境変数 JP1_OUTPUT_MODE を設定することで、
# 以下の出力方法を選択できます：
#
# /NOTEPAD - メモ帳で開く（デフォルト）
#   - ログファイルをメモ帳で自動的に開きます
#   - JP1_SCROLL_TO_TEXT を設定すると、その文字列がある行にジャンプします
#
# /EXCEL - Excelに貼り付け
#   - 指定したExcelファイルの指定セルにログ内容を貼り付けます
#   - 必要な環境変数：
#     - EXCEL_FILE_NAME: Excelファイルのパス
#     - EXCEL_SHEET_NAME: シート名
#     - EXCEL_PASTE_CELL: 貼り付け先セル（例: "B5"）
#
# /WINMERGE - WinMergeで比較
#   - 2つのログファイルを比較表示します
#
# 【エラー時の終了コード】
# - 終了コード 10: Excel設定エラー（必要な環境変数が未設定）
# - 終了コード 11: Excel貼り付けエラー（COM操作に失敗）
# - 終了コード 12: Excelファイル未検出

Write-Console "[STEP 6] 出力処理中..."

# 出力オプションを環境変数から取得
$outputMode = $env:JP1_OUTPUT_MODE
if (-not $outputMode) { $outputMode = "/NOTEPAD" }

# 出力オプションに応じた後処理
switch ($outputMode.ToUpper()) {
    "/NOTEPAD" {
        # メモ帳で開く
        Start-Process notepad $outputFilePath

        # スクロール位置の設定を環境変数から取得
        $scrollToText = $env:JP1_SCROLL_TO_TEXT
        if ($scrollToText) {
            Write-Console "スクロール位置: $scrollToText"

            # 検索文字列を含む最初の行番号を特定
            $scrollLineNum = $null
            $lineIndex = 0
            $fileContent = Get-Content -Path $outputFilePath -Encoding Default
            foreach ($line in $fileContent) {
                $lineIndex++
                if ($line -match [regex]::Escape($scrollToText)) {
                    $scrollLineNum = $lineIndex
                    break
                }
            }

            if ($scrollLineNum) {
                Write-Console "ジャンプ先行番号: $scrollLineNum"

                # メモ帳がアクティブになるまで待機し、Ctrl+Gで行へ移動
                Start-Sleep -Milliseconds 600
                $wshell = New-Object -ComObject WScript.Shell
                $activated = $wshell.AppActivate("メモ帳")
                if (-not $activated) { $activated = $wshell.AppActivate("Notepad") }
                if ($activated) {
                    Start-Sleep -Milliseconds 100
                    $wshell.SendKeys("^g")  # Ctrl+G で「行へ移動」ダイアログを開く
                    Start-Sleep -Milliseconds 200
                    $wshell.SendKeys($scrollLineNum.ToString())
                    Start-Sleep -Milliseconds 100
                    $wshell.SendKeys("{ENTER}")
                }
            } else {
                Write-Console "[情報] 指定した文字列がファイル内に見つかりませんでした"
            }
        }
    }
    "/EXCEL" {
        # ======================================================================
        # Excel貼り付け処理（雛形フォルダコピー + ジョブ別Excel選択）
        # ======================================================================
        # 【処理概要】
        # 1. ジョブパスから対応するExcelファイルを特定
        # 2. output/yyyymmddフォルダを作成
        # 3. 雛形フォルダをコピー
        # 4. 対応するExcelファイルにログを貼り付け
        # ======================================================================

        # ------------------------------------------------------------------
        # STEP 1: ジョブパスからExcel設定を取得
        # ------------------------------------------------------------------
        $targetUnitPath = $env:JP1_UNIT_PATH
        $excelFileName = $null
        $excelSheetName = $null
        $excelPasteCell = $null
        $excelMoveToSheet = $null

        # ジョブパスに一致するマッピングを検索（完全一致）
        foreach ($key in $jobExcelMapping.Keys) {
            if ($targetUnitPath -eq $key) {
                $mappingValue = $jobExcelMapping[$key]
                $parts = $mappingValue -split ","
                if ($parts.Count -ge 3) {
                    $excelFileName = $parts[0].Trim()
                    $excelSheetName = $parts[1].Trim()
                    $excelPasteCell = $parts[2].Trim()
                    if ($parts.Count -ge 4) {
                        $excelMoveToSheet = $parts[3].Trim()
                    }
                }
                break
            }
        }

        # マッピングが見つからない場合はエラー
        if (-not $excelFileName) {
            Write-Host "[エラー] ジョブパス '$targetUnitPath' に対応するExcel設定が見つかりません。" -ForegroundColor Red
            Write-Host "[エラー] 設定セクションの `$jobExcelMapping を確認してください。" -ForegroundColor Red
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
        Write-Host "    ジョブパス    : $targetUnitPath"
        Write-Host "    Excelファイル : $excelFileName"
        Write-Host "    貼り付けシート: $excelSheetName"
        Write-Host "    貼り付けセル  : $excelPasteCell"
        if ($excelMoveToSheet) {
            Write-Host "    移動先シート  : $excelMoveToSheet"
        }
        Write-Host ""

        # ------------------------------------------------------------------
        # STEP 2: 02_output/yyyymmddフォルダを作成
        # ------------------------------------------------------------------
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

        # ------------------------------------------------------------------
        # STEP 3: 雛形フォルダの中身をコピー
        # ------------------------------------------------------------------
        Write-Host "  [STEP 2] 雛形フォルダコピー" -ForegroundColor Cyan
        $templatePath = Join-Path $scriptDir $templateFolderName

        if (-not (Test-Path $templatePath)) {
            Write-Host "[エラー] 雛形フォルダが見つかりません: $templatePath" -ForegroundColor Red
            exit 13  # 雛形フォルダ未検出エラー
        }

        # 雛形フォルダの中身をyyyymmddフォルダに直接コピー（存在しない場合のみ）
        # （雛形フォルダ自体はコピーせず、中のファイルのみ）
        $templateItems = Get-ChildItem -Path $templatePath
        $copiedCount = 0
        foreach ($item in $templateItems) {
            $destPath = Join-Path $dateFolderPath $item.Name
            if (-not (Test-Path $destPath)) {
                try {
                    Copy-Item -Path $item.FullName -Destination $destPath -Recurse -Force
                    $copiedCount++
                } catch {
                    Write-Host "    [警告] コピーに失敗: $($item.Name) - $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
        }
        if ($copiedCount -gt 0) {
            Write-Host "    コピー完了: $copiedCount 件のファイル/フォルダ"
        } else {
            Write-Host "    スキップ: 既にコピー済み"
        }
        Write-Host ""

        # ------------------------------------------------------------------
        # STEP 4: Excelファイルにログを貼り付け
        # ------------------------------------------------------------------
        Write-Host "  [STEP 3] Excelに貼り付け" -ForegroundColor Cyan
        $excelPath = Join-Path $dateFolderPath $excelFileName

        if (-not (Test-Path $excelPath)) {
            Write-Host "[エラー] Excelファイルが見つかりません: $excelPath" -ForegroundColor Red
            exit 12  # Excelファイル未検出エラー
        }

        try {
            # クリップボードにコピー（$execResultContentを直接使用）
            Set-Clipboard -Value $execResultContent

            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $true
            $workbook = $excel.Workbooks.Open($excelPath)
            $sheet = $workbook.Worksheets.Item($excelSheetName)

            # 貼り付け先のセルを選択してクリップボードから貼り付け
            $sheet.Range($excelPasteCell).Select()
            $sheet.Paste()

            # 貼り付け後、指定シートに移動してA1セルを選択
            if ($excelMoveToSheet) {
                $workbook.Worksheets.Item($excelMoveToSheet).Activate()
                $workbook.Worksheets.Item($excelMoveToSheet).Range("A1").Select()
            }

            $workbook.Save()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            Write-Host "    貼り付け完了: $excelFileName"
            Write-Host ""

            # ------------------------------------------------------------------
            # テキストファイル保存
            # ------------------------------------------------------------------
            Write-Host "  [STEP 4] テキストファイル保存" -ForegroundColor Cyan
            $textFileName = "runh_default.txt"
            foreach ($key in $jobTextFileMapping.Keys) {
                if ($unitPath -eq $key) {
                    $textFileName = $jobTextFileMapping[$key]
                    break
                }
            }
            $clipboardOutputFile = Join-Path $dateFolderPath $textFileName
            Get-Clipboard | Out-File -FilePath $clipboardOutputFile -Encoding Default
            Write-Host "    保存完了: $textFileName"
            Write-Host ""

            # ------------------------------------------------------------------
            # 変換バッチ実行
            # ------------------------------------------------------------------
            $convertBatchFile = Join-Path $scriptDir "【削除禁止】ConvertNS932Result.bat"
            if (Test-Path $convertBatchFile) {
                Write-Host "  [STEP 5] 変換バッチ実行" -ForegroundColor Cyan
                $env:OUTPUT_FOLDER = $dateFolderPath
                & cmd /c "`"$convertBatchFile`""
                $env:OUTPUT_FOLDER = $null
                Write-Host "    実行完了"
                Write-Host ""
            }

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
            Write-Host "[エラー] Excel貼り付けに失敗しました: $($_.Exception.Message)" -ForegroundColor Red
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
            Write-Host "[エラー] WinMergeが見つかりません: $winMergePath" -ForegroundColor Red
            Write-Host "        WinMergeをインストールするか、パスを確認してください。" -ForegroundColor Red
        }
    }
    default {
        # デフォルトはログファイル出力のみ
    }
}

# 完了表示
Write-Host ""
Write-Host "================================================================" -ForegroundColor Green
Write-Host "  取得完了" -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Green
Write-Host ""
Write-Host "ジョブネット名: $jobnetName"
Write-Host "コメント:       $jobnetComment"
Write-Host "ジョブ開始日時: $startTimeForFileName"
Write-Host "終了状態:       $endStatusDisplay"
Write-Host "出力ファイル:   $outputFilePath"
Write-Host ""

# 正常終了
exit 0
