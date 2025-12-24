Option Explicit

'==============================================================================
' JP1 ジョブ管理ツール - メインモジュール
'   - ジョブネット一覧取得（ajsprint経由）
'   - ジョブ実行処理
'
' 注意: 初期化処理は JP1_JobManager_Setup.bas にあります
'       定数はSetupモジュールで Public として定義されています
'==============================================================================

' デバッグモード（Trueにすると[DEBUG-XX]ログが出力されます）
Private Const DEBUG_MODE As Boolean = False

' 管理者権限状態を保持
Private g_AdminChecked As Boolean
Private g_IsAdmin As Boolean

' 現在の実行セッションのログファイルパス
Private g_LogFilePath As String

'==============================================================================
' ジョブ一覧取得
'==============================================================================
Public Sub GetJobList()
    On Error GoTo ErrorHandler

    Dim config As Object
    Set config = GetConfig()

    If config Is Nothing Then Exit Sub

    ' リモートモードの場合、管理者権限をチェック
    If Not EnsureAdminForRemoteMode(config) Then Exit Sub

    ' パスワード入力（リモートモードの場合のみリモートパスワードが必要）
    If config("ExecMode") <> "ローカル" Then
        If config("RemotePassword") = "" Then
            config("RemotePassword") = InputBox("リモートサーバのパスワードを入力してください:", "パスワード入力")
            If config("RemotePassword") = "" Then
                MsgBox "パスワードが入力されませんでした。", vbExclamation
                Exit Sub
            End If
        End If
    End If

    If config("JP1Password") = "" Then
        config("JP1Password") = InputBox("JP1パスワードを入力してください:", "パスワード入力")
        If config("JP1Password") = "" Then
            MsgBox "パスワードが入力されませんでした。", vbExclamation
            Exit Sub
        End If
    End If

    Application.ScreenUpdating = False
    Application.StatusBar = "ジョブ一覧を取得中..."

    ' 既存のオートフィルタを解除
    Dim wsJobList As Worksheet
    Set wsJobList = Worksheets(SHEET_JOBLIST)
    If wsJobList.AutoFilterMode Then
        wsJobList.AutoFilterMode = False
    End If

    ' PowerShellスクリプト生成・実行
    Dim psScript As String
    psScript = BuildGetJobListScript(config)

    Dim result As String
    result = ExecutePowerShell(psScript)

    ' 結果をパース（戻り値で成功/失敗を判定）
    ' コマンド実行時は -F オプションでスケジューラサービスを指定するため
    ' パスにはスケジューラサービス名を含めない（ルートパス + "/" + ユニット名）
    Dim parseSuccess As Boolean
    parseSuccess = ParseJobListResult(result, config("RootPath"))

    Application.StatusBar = False
    Application.ScreenUpdating = True

    ' エラーの場合は完了メッセージを表示しない
    If Not parseSuccess Then
        Exit Sub
    End If

    ' 種別「ジョブネット」でオートフィルタを適用
    Dim lastDataRow As Long
    lastDataRow = wsJobList.Cells(wsJobList.Rows.Count, COL_JOBNET_PATH).End(xlUp).row
    If lastDataRow >= ROW_JOBLIST_DATA_START Then
        ' ヘッダー行からデータ最終行までを範囲としてオートフィルタを設定
        wsJobList.Range(wsJobList.Cells(ROW_JOBLIST_HEADER, COL_SELECT), wsJobList.Cells(lastDataRow, COL_LAST_MESSAGE)).AutoFilter _
            Field:=COL_UNIT_TYPE - COL_SELECT + 1, Criteria1:="ジョブネット"
    End If

    MsgBox "ジョブ一覧の取得が完了しました。" & vbCrLf & _
           "ジョブ一覧シートを確認してください。", vbInformation

    Worksheets(SHEET_JOBLIST).Activate
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました。" & vbCrLf & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description & vbCrLf & _
           "発生場所: GetJobList", vbCritical, "VBAエラー"
End Sub

'==============================================================================
' グループ名取得（取得パスのドロップダウンリスト更新）
'==============================================================================
Public Sub GetGroupList()
    On Error GoTo ErrorHandler

    ' 確認メッセージ
    If MsgBox("JP1サーバから全てのグループ名を抽出します。" & vbCrLf & _
              "処理に時間がかかる場合がありますがよろしいですか？", _
              vbYesNo + vbQuestion, "グループ名取得") = vbNo Then
        Exit Sub
    End If

    Dim config As Object
    Set config = GetConfig()

    If config Is Nothing Then Exit Sub

    ' リモートモードの場合、管理者権限をチェック
    If Not EnsureAdminForRemoteMode(config) Then Exit Sub

    ' パスワード入力（リモートモードの場合のみリモートパスワードが必要）
    If config("ExecMode") <> "ローカル" Then
        If config("RemotePassword") = "" Then
            config("RemotePassword") = InputBox("リモートサーバのパスワードを入力してください:", "パスワード入力")
            If config("RemotePassword") = "" Then
                MsgBox "パスワードが入力されませんでした。", vbExclamation
                Exit Sub
            End If
        End If
    End If

    Application.StatusBar = "グループ名を取得中..."
    Application.ScreenUpdating = False

    ' PowerShellスクリプト生成・実行（ルート直下のみ取得、-Rなし）
    Dim psScript As String
    psScript = BuildGetGroupListScript(config)

    Dim result As String
    result = ExecutePowerShell(psScript)

    ' 結果をパースしてグループ名リストを作成
    Dim groupList As String
    groupList = ParseGroupListResult(result)

    If groupList = "" Then
        Application.StatusBar = False
        Application.ScreenUpdating = True
        MsgBox "グループが見つかりませんでした。" & vbCrLf & _
               "接続設定を確認してください。", vbExclamation, "グループ名取得"
        Exit Sub
    End If

    ' グループリストを設定シートのG列に書き込み（ドロップダウン用）
    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_SETTINGS)

    ' G列の既存データをクリア
    Dim lastGroupRow As Long
    lastGroupRow = ws.Cells(ws.Rows.Count, 7).End(xlUp).row
    If lastGroupRow >= 1 Then
        ws.Range(ws.Cells(1, 7), ws.Cells(lastGroupRow, 7)).ClearContents
    End If

    ' グループリストを配列に変換して書き込み
    Dim groupArray() As String
    groupArray = Split(groupList, ",")

    Dim i As Long
    Dim groupCount As Long
    groupCount = UBound(groupArray) + 1

    For i = 0 To UBound(groupArray)
        ws.Cells(i + 1, 7).Value = groupArray(i)
    Next i

    ' G列を非表示に
    ws.Columns("G").Hidden = True

    ' 取得パス欄（C15）にドロップダウンリストを設定（セル範囲参照）
    Dim listRange As String
    listRange = "=" & SHEET_SETTINGS & "!$G$1:$G$" & groupCount

    With ws.Cells(ROW_ROOT_PATH, COL_SETTING_VALUE).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=listRange
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    Application.StatusBar = False
    Application.ScreenUpdating = True

    MsgBox "グループ名を取得しました（" & groupCount & " 件）。" & vbCrLf & _
           "取得パス欄のドロップダウンから選択できます。", vbInformation, "グループ名取得"
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました。" & vbCrLf & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description & vbCrLf & _
           "発生場所: GetGroupList", vbCritical, "VBAエラー"
End Sub

Private Function BuildGetGroupListScript(config As Object) As String
    Dim script As String

    script = "$ErrorActionPreference = 'Stop'" & vbCrLf
    script = script & vbCrLf

    ' UTF-8エンコーディング設定（日本語パス対応）
    script = script & "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8" & vbCrLf
    script = script & "$OutputEncoding = [System.Text.Encoding]::UTF8" & vbCrLf
    script = script & "chcp 65001 | Out-Null" & vbCrLf
    script = script & vbCrLf

    ' ログ出力関数を定義（コンソール表示用）
    script = script & "function Write-Log {" & vbCrLf
    script = script & "  param([string]$Message)" & vbCrLf
    script = script & "  $timestamp = Get-Date -Format 'yyyy/MM/dd HH:mm:ss'" & vbCrLf
    script = script & "  Write-Host ""[$timestamp] $Message""" & vbCrLf
    script = script & "}" & vbCrLf
    script = script & vbCrLf

    ' ローカルモードとリモートモードで処理を分岐
    If config("ExecMode") = "ローカル" Then
        ' ローカル実行モード
        script = script & "try {" & vbCrLf
        script = script & "  Write-Log '--------------------------------------------------------------------------------'" & vbCrLf
        script = script & "  Write-Log 'グループ一覧取得開始'" & vbCrLf
        script = script & "  Write-Log '--------------------------------------------------------------------------------'" & vbCrLf
        script = script & vbCrLf
        script = script & "  $ajsprintPath = $null" & vbCrLf
        script = script & "  $searchPaths = @(" & vbCrLf
        script = script & "    'C:\Program Files\HITACHI\JP1AJS3\bin\ajsprint.exe'," & vbCrLf
        script = script & "    'C:\Program Files (x86)\HITACHI\JP1AJS3\bin\ajsprint.exe'," & vbCrLf
        script = script & "    'C:\Program Files\Hitachi\JP1AJS2\bin\ajsprint.exe'," & vbCrLf
        script = script & "    'C:\Program Files (x86)\Hitachi\JP1AJS2\bin\ajsprint.exe'" & vbCrLf
        script = script & "  )" & vbCrLf
        script = script & "  foreach ($path in $searchPaths) {" & vbCrLf
        script = script & "    if (Test-Path $path) { $ajsprintPath = $path; break }" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & "  if (-not $ajsprintPath) {" & vbCrLf
        script = script & "    Write-Log '[ERROR] JP1コマンド(ajsprint.exe)が見つかりません'" & vbCrLf
        script = script & "    Write-Output 'ERROR: JP1コマンド(ajsprint.exe)が見つかりません。'" & vbCrLf
        script = script & "    exit 1" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & "  Write-Log ""JP1コマンドパス: $ajsprintPath""" & vbCrLf
        script = script & vbCrLf
        script = script & "  # 全グループを再帰的に取得（-Rあり）" & vbCrLf
        script = script & "  Write-Log '[実行] ajsprint - グループ一覧取得'" & vbCrLf
        script = script & "  Write-Log ""コマンド: ajsprint.exe -F " & config("SchedulerService") & " /* -R""" & vbCrLf
        script = script & "  $result = & $ajsprintPath -F " & config("SchedulerService") & " '/*' -R 2>&1" & vbCrLf
        script = script & "  Write-Log '[成功] グループ一覧取得完了'" & vbCrLf
        script = script & "  $result | ForEach-Object { Write-Output $_ }" & vbCrLf
        script = script & "} catch {" & vbCrLf
        script = script & "  Write-Log ""[ERROR] $($_.Exception.Message)""" & vbCrLf
        script = script & "  Write-Output ""ERROR: $($_.Exception.Message)""" & vbCrLf
        script = script & "}" & vbCrLf
    Else
        ' リモート実行モード（WinRM使用）
        script = script & "Write-Log '--------------------------------------------------------------------------------'" & vbCrLf
        script = script & "Write-Log 'グループ一覧取得開始（リモート）'" & vbCrLf
        script = script & "Write-Log '--------------------------------------------------------------------------------'" & vbCrLf
        script = script & "Write-Log ""接続先: " & config("JP1Server") & """" & vbCrLf
        script = script & vbCrLf
        script = script & "$securePass = ConvertTo-SecureString '" & EscapePSString(config("RemotePassword")) & "' -AsPlainText -Force" & vbCrLf
        script = script & "$cred = New-Object System.Management.Automation.PSCredential('" & config("RemoteUser") & "', $securePass)" & vbCrLf
        script = script & vbCrLf

        ' WinRM設定の保存と自動設定
        script = script & "$originalTrustedHosts = $null" & vbCrLf
        script = script & "$winrmConfigChanged = $false" & vbCrLf
        script = script & "$winrmServiceWasStarted = $false" & vbCrLf
        script = script & vbCrLf

        script = script & "try {" & vbCrLf
        script = script & "  Write-Log '[実行] WinRMサービス確認'" & vbCrLf
        script = script & "  $winrmService = Get-Service -Name WinRM -ErrorAction SilentlyContinue" & vbCrLf
        script = script & "  if ($winrmService.Status -ne 'Running') {" & vbCrLf
        script = script & "    Start-Service -Name WinRM -ErrorAction Stop" & vbCrLf
        script = script & "    $winrmServiceWasStarted = $true" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & vbCrLf
        script = script & "  $originalTrustedHosts = (Get-Item WSMan:\localhost\Client\TrustedHosts -ErrorAction SilentlyContinue).Value" & vbCrLf
        script = script & vbCrLf
        script = script & "  if ($originalTrustedHosts -notmatch '" & config("JP1Server") & "') {" & vbCrLf
        script = script & "    if ($originalTrustedHosts) {" & vbCrLf
        script = script & "      Set-Item WSMan:\localhost\Client\TrustedHosts -Value ""$originalTrustedHosts," & config("JP1Server") & """ -Force" & vbCrLf
        script = script & "    } else {" & vbCrLf
        script = script & "      Set-Item WSMan:\localhost\Client\TrustedHosts -Value '" & config("JP1Server") & "' -Force" & vbCrLf
        script = script & "    }" & vbCrLf
        script = script & "    $winrmConfigChanged = $true" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & vbCrLf

        ' リモート実行
        script = script & "  Write-Log '[実行] リモートセッション作成'" & vbCrLf
        script = script & "  $session = New-PSSession -ComputerName '" & config("JP1Server") & "' -Credential $cred -ErrorAction Stop" & vbCrLf
        script = script & "  Write-Log '[成功] リモートセッション作成完了'" & vbCrLf
        script = script & vbCrLf
        script = script & "  Write-Log '[実行] ajsprint - グループ一覧取得'" & vbCrLf
        script = script & "  Write-Log ""コマンド: ajsprint.exe -F " & config("SchedulerService") & " /* -R""" & vbCrLf
        script = script & "  $result = Invoke-Command -Session $session -ScriptBlock {" & vbCrLf
        script = script & "    param($schedulerService)" & vbCrLf
        script = script & "    $ajsprintPath = $null" & vbCrLf
        script = script & "    $searchPaths = @('C:\Program Files\HITACHI\JP1AJS3\bin\ajsprint.exe','C:\Program Files (x86)\HITACHI\JP1AJS3\bin\ajsprint.exe','C:\Program Files\Hitachi\JP1AJS2\bin\ajsprint.exe','C:\Program Files (x86)\Hitachi\JP1AJS2\bin\ajsprint.exe')" & vbCrLf
        script = script & "    foreach ($p in $searchPaths) { if (Test-Path $p) { $ajsprintPath = $p; break } }" & vbCrLf
        script = script & "    if (-not $ajsprintPath) { Write-Output 'ERROR: ajsprint.exe not found'; return }" & vbCrLf
        script = script & "    # 全グループを再帰的に取得（-Rあり）" & vbCrLf
        script = script & "    $output = & $ajsprintPath '-F' $schedulerService '/*' '-R' 2>&1" & vbCrLf
        script = script & "    $output | Where-Object { $_ -notmatch '^KAVS\d+-I' }" & vbCrLf
        script = script & "  } -ArgumentList '" & config("SchedulerService") & "'" & vbCrLf
        script = script & "  Write-Log '[成功] グループ一覧取得完了'" & vbCrLf
        script = script & vbCrLf
        script = script & "  Remove-PSSession $session" & vbCrLf
        script = script & vbCrLf
        script = script & "  $result | ForEach-Object { Write-Output $_ }" & vbCrLf
        script = script & "} catch {" & vbCrLf
        script = script & "  Write-Log ""[ERROR] $($_.Exception.Message)""" & vbCrLf
        script = script & "  Write-Output ""ERROR: $($_.Exception.Message)""" & vbCrLf
        script = script & "} finally {" & vbCrLf
        script = script & "  if ($winrmConfigChanged) {" & vbCrLf
        script = script & "    if ($originalTrustedHosts) {" & vbCrLf
        script = script & "      Set-Item WSMan:\localhost\Client\TrustedHosts -Value $originalTrustedHosts -Force -ErrorAction SilentlyContinue" & vbCrLf
        script = script & "    } else {" & vbCrLf
        script = script & "      Clear-Item WSMan:\localhost\Client\TrustedHosts -Force -ErrorAction SilentlyContinue" & vbCrLf
        script = script & "    }" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & "  if ($winrmServiceWasStarted) {" & vbCrLf
        script = script & "    Stop-Service -Name WinRM -Force -ErrorAction SilentlyContinue" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & "}" & vbCrLf
    End If

    BuildGetGroupListScript = script
End Function

Private Function ParseGroupListResult(result As String) As String
    ' グループ名を抽出してドロップダウン用のリストを作成（ネスト対応）
    ' 戻り値: カンマ区切りのパスリスト（例: /,/グループA,/グループA/サブグループ）
    ' ※アスタリスクなしで保存（ジョブ一覧取得時に-Rオプションで再帰取得）

    ' エラーチェック
    If InStr(result, "ERROR:") > 0 Then
        ParseGroupListResult = ""
        Exit Function
    End If

    Dim lines() As String
    lines = Split(result, vbCrLf)

    Dim groupPaths As String
    groupPaths = "/"  ' デフォルトで全件取得オプションを追加（ルート）

    ' ネスト対応のためスタック構造を使用
    Const MAX_DEPTH As Long = 20
    Dim pathStack(1 To MAX_DEPTH) As String   ' パスのスタック
    Dim typeStack(1 To MAX_DEPTH) As String   ' ユニットタイプのスタック
    Dim stackDepth As Long
    stackDepth = 0

    Dim i As Long
    Dim pendingUnitName As String
    pendingUnitName = ""

    For i = LBound(lines) To UBound(lines)
        Dim lineStr As String
        lineStr = Trim(Replace(lines(i), vbTab, ""))

        ' 空行はスキップ
        If lineStr = "" Then GoTo NextGroupLine

        ' unit=行からユニット名を取得
        If InStr(lineStr, "unit=") > 0 Then
            ' unit=名前,,...; から名前を抽出
            Dim parts() As String
            Dim unitPart As String
            unitPart = Mid(lineStr, InStr(lineStr, "unit=") + 5)
            parts = Split(unitPart, ",")
            If UBound(parts) >= 0 Then
                pendingUnitName = Trim(parts(0))
                ' 末尾のセミコロンを除去
                If Right(pendingUnitName, 1) = ";" Then
                    pendingUnitName = Left(pendingUnitName, Len(pendingUnitName) - 1)
                End If
            End If
            GoTo NextGroupLine
        End If

        ' { でネストレベルを上げる
        If Left(lineStr, 1) = "{" Then
            If pendingUnitName <> "" Then
                stackDepth = stackDepth + 1
                If stackDepth <= MAX_DEPTH Then
                    If stackDepth = 1 Then
                        pathStack(stackDepth) = "/" & pendingUnitName
                    Else
                        pathStack(stackDepth) = pathStack(stackDepth - 1) & "/" & pendingUnitName
                    End If
                    typeStack(stackDepth) = ""  ' まだタイプ未確定
                End If
                pendingUnitName = ""
            End If
            GoTo NextGroupLine
        End If

        ' ty=行でユニットタイプを確認
        If InStr(lineStr, "ty=") > 0 Then
            If stackDepth > 0 And stackDepth <= MAX_DEPTH Then
                ' ty=g;（グループ）の場合、パスリストに追加
                If InStr(lineStr, "ty=g;") > 0 Then
                    typeStack(stackDepth) = "g"
                    groupPaths = groupPaths & "," & pathStack(stackDepth)
                End If
            End If
            GoTo NextGroupLine
        End If

        ' } でネストレベルを下げる
        If Left(lineStr, 1) = "}" Then
            If stackDepth > 0 Then
                stackDepth = stackDepth - 1
            End If
            GoTo NextGroupLine
        End If

NextGroupLine:
    Next i

    ParseGroupListResult = groupPaths
End Function

Private Function BuildGetJobListScript(config As Object) As String
    Dim script As String

    script = "$ErrorActionPreference = 'Stop'" & vbCrLf
    script = script & vbCrLf

    ' UTF-8エンコーディング設定（日本語パス対応）
    script = script & "# UTF-8エンコーディング設定" & vbCrLf
    script = script & "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8" & vbCrLf
    script = script & "$OutputEncoding = [System.Text.Encoding]::UTF8" & vbCrLf
    script = script & "chcp 65001 | Out-Null" & vbCrLf
    script = script & vbCrLf

    ' ログ出力関数を定義（コンソール表示用）
    script = script & "function Write-Log {" & vbCrLf
    script = script & "  param([string]$Message)" & vbCrLf
    script = script & "  $timestamp = Get-Date -Format 'yyyy/MM/dd HH:mm:ss'" & vbCrLf
    script = script & "  Write-Host ""[$timestamp] $Message""" & vbCrLf
    script = script & "}" & vbCrLf
    script = script & vbCrLf

    ' ローカルモードとリモートモードで処理を分岐
    If config("ExecMode") = "ローカル" Then
        ' ローカル実行モード（WinRM不使用）
        script = script & "try {" & vbCrLf
        script = script & "  Write-Log '--------------------------------------------------------------------------------'" & vbCrLf
        script = script & "  Write-Log 'ジョブ一覧取得開始'" & vbCrLf
        script = script & "  Write-Log ""対象パス: " & config("RootPath") & """" & vbCrLf
        script = script & "  Write-Log '--------------------------------------------------------------------------------'" & vbCrLf
        script = script & vbCrLf
        script = script & "  # JP1コマンドパスの検出" & vbCrLf
        script = script & "  $ajsprintPath = $null" & vbCrLf
        script = script & "  $searchPaths = @(" & vbCrLf
        script = script & "    'C:\Program Files\HITACHI\JP1AJS3\bin\ajsprint.exe'," & vbCrLf
        script = script & "    'C:\Program Files (x86)\HITACHI\JP1AJS3\bin\ajsprint.exe'," & vbCrLf
        script = script & "    'C:\Program Files\Hitachi\JP1AJS2\bin\ajsprint.exe'," & vbCrLf
        script = script & "    'C:\Program Files (x86)\Hitachi\JP1AJS2\bin\ajsprint.exe'" & vbCrLf
        script = script & "  )" & vbCrLf
        script = script & "  foreach ($path in $searchPaths) {" & vbCrLf
        script = script & "    if (Test-Path $path) { $ajsprintPath = $path; break }" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & "  if (-not $ajsprintPath) {" & vbCrLf
        script = script & "    Write-Log '[ERROR] JP1コマンド(ajsprint.exe)が見つかりません'" & vbCrLf
        script = script & "    Write-Output ""ERROR: JP1コマンド(ajsprint.exe)が見つかりません。JP1/AJS3 Managerがインストールされているか確認してください。""" & vbCrLf
        script = script & "    exit 1" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & "  Write-Log ""JP1コマンドパス: $ajsprintPath""" & vbCrLf
        script = script & vbCrLf
        script = script & "  # ローカルでajsprintを実行" & vbCrLf
        script = script & "  Write-Log '[実行] ajsprint - ジョブ一覧取得'" & vbCrLf
        script = script & "  Write-Log ""コマンド: ajsprint.exe -F " & config("SchedulerService") & " " & config("RootPath") & " -R""" & vbCrLf
        script = script & "  $result = & $ajsprintPath -F " & config("SchedulerService") & " '" & config("RootPath") & "' -R 2>&1" & vbCrLf
        script = script & "  Write-Log '[成功] ジョブ一覧取得完了'" & vbCrLf
        script = script & "  $result | ForEach-Object { Write-Output $_ }" & vbCrLf
        script = script & "} catch {" & vbCrLf
        script = script & "  Write-Log ""[ERROR] $($_.Exception.Message)""" & vbCrLf
        script = script & "  Write-Output ""ERROR: $($_.Exception.Message)""" & vbCrLf
        script = script & "}" & vbCrLf
    Else
        ' リモート実行モード（WinRM使用）
        script = script & "Write-Log '--------------------------------------------------------------------------------'" & vbCrLf
        script = script & "Write-Log 'ジョブ一覧取得開始（リモート）'" & vbCrLf
        script = script & "Write-Log ""対象パス: " & config("RootPath") & """" & vbCrLf
        script = script & "Write-Log ""接続先: " & config("JP1Server") & """" & vbCrLf
        script = script & "Write-Log '--------------------------------------------------------------------------------'" & vbCrLf
        script = script & vbCrLf
        ' 認証情報
        script = script & "$securePass = ConvertTo-SecureString '" & EscapePSString(config("RemotePassword")) & "' -AsPlainText -Force" & vbCrLf
        script = script & "$cred = New-Object System.Management.Automation.PSCredential('" & config("RemoteUser") & "', $securePass)" & vbCrLf
        script = script & vbCrLf

        ' WinRM設定の保存と自動設定
        script = script & "$originalTrustedHosts = $null" & vbCrLf
        script = script & "$winrmConfigChanged = $false" & vbCrLf
        script = script & "$winrmServiceWasStarted = $false" & vbCrLf
        script = script & vbCrLf

        script = script & "try {" & vbCrLf
        script = script & "  Write-Log '[実行] WinRMサービス確認'" & vbCrLf
        script = script & "  # WinRMサービスの起動確認（TrustedHosts取得前に起動が必要）" & vbCrLf
        script = script & "  $winrmService = Get-Service -Name WinRM -ErrorAction SilentlyContinue" & vbCrLf
        script = script & "  if ($winrmService.Status -ne 'Running') {" & vbCrLf
        script = script & "    Start-Service -Name WinRM -ErrorAction Stop" & vbCrLf
        script = script & "    $winrmServiceWasStarted = $true" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & vbCrLf
        script = script & "  # 現在のTrustedHostsを取得（WinRMサービス起動後に取得）" & vbCrLf
        script = script & "  $originalTrustedHosts = (Get-Item WSMan:\localhost\Client\TrustedHosts -ErrorAction SilentlyContinue).Value" & vbCrLf
        script = script & vbCrLf
        script = script & "  # TrustedHostsに接続先を追加（必要な場合のみ）" & vbCrLf
        script = script & "  if ($originalTrustedHosts -notmatch '" & config("JP1Server") & "') {" & vbCrLf
        script = script & "    if ($originalTrustedHosts) {" & vbCrLf
        script = script & "      Set-Item WSMan:\localhost\Client\TrustedHosts -Value ""$originalTrustedHosts," & config("JP1Server") & """ -Force" & vbCrLf
        script = script & "    } else {" & vbCrLf
        script = script & "      Set-Item WSMan:\localhost\Client\TrustedHosts -Value '" & config("JP1Server") & "' -Force" & vbCrLf
        script = script & "    }" & vbCrLf
        script = script & "    $winrmConfigChanged = $true" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & vbCrLf

        ' リモート実行
        script = script & "  Write-Log '[実行] リモートセッション作成'" & vbCrLf
        script = script & "  $session = New-PSSession -ComputerName '" & config("JP1Server") & "' -Credential $cred -ErrorAction Stop" & vbCrLf
        script = script & "  Write-Log '[成功] リモートセッション作成完了'" & vbCrLf
        script = script & vbCrLf
        script = script & "  Write-Log '[実行] ajsprint - ジョブ一覧取得'" & vbCrLf
        script = script & "  Write-Log ""コマンド: ajsprint.exe -F " & config("SchedulerService") & " " & config("RootPath") & " -R""" & vbCrLf
        script = script & "  $result = Invoke-Command -Session $session -ScriptBlock {" & vbCrLf
        script = script & "    param($schedulerService, $rootPath)" & vbCrLf
        script = script & "    if ([string]::IsNullOrWhiteSpace($rootPath)) { Write-Output 'ERROR: rootPath is empty'; return }" & vbCrLf
        script = script & "    $ajsprintPath = $null" & vbCrLf
        script = script & "    $searchPaths = @('C:\Program Files\HITACHI\JP1AJS3\bin\ajsprint.exe','C:\Program Files (x86)\HITACHI\JP1AJS3\bin\ajsprint.exe','C:\Program Files\Hitachi\JP1AJS2\bin\ajsprint.exe','C:\Program Files (x86)\Hitachi\JP1AJS2\bin\ajsprint.exe')" & vbCrLf
        script = script & "    foreach ($p in $searchPaths) { if (Test-Path $p) { $ajsprintPath = $p; break } }" & vbCrLf
        script = script & "    if (-not $ajsprintPath) { Write-Output 'ERROR: ajsprint.exe not found'; return }" & vbCrLf
        script = script & "    $output = & $ajsprintPath '-F' $schedulerService $rootPath '-R' 2>&1" & vbCrLf
        script = script & "    # KAVS情報メッセージ（-I）を除外、unit=行のみ出力" & vbCrLf
        script = script & "    $output | Where-Object { $_ -notmatch '^KAVS\d+-I' }" & vbCrLf
        script = script & "  } -ArgumentList '" & config("SchedulerService") & "', '" & config("RootPath") & "'" & vbCrLf
        script = script & "  Write-Log '[成功] ジョブ一覧取得完了'" & vbCrLf
        script = script & vbCrLf
        script = script & "  Remove-PSSession $session" & vbCrLf
        script = script & vbCrLf
        script = script & "  $result | ForEach-Object { Write-Output $_ }" & vbCrLf
        script = script & "} catch {" & vbCrLf
        script = script & "  Write-Log ""[ERROR] $($_.Exception.Message)""" & vbCrLf
        script = script & "  Write-Output ""ERROR: $($_.Exception.Message)""" & vbCrLf
        script = script & "} finally {" & vbCrLf
        script = script & "  # WinRM設定の復元" & vbCrLf
        script = script & "  if ($winrmConfigChanged) {" & vbCrLf
        script = script & "    if ($originalTrustedHosts) {" & vbCrLf
        script = script & "      Set-Item WSMan:\localhost\Client\TrustedHosts -Value $originalTrustedHosts -Force -ErrorAction SilentlyContinue" & vbCrLf
        script = script & "    } else {" & vbCrLf
        script = script & "      Clear-Item WSMan:\localhost\Client\TrustedHosts -Force -ErrorAction SilentlyContinue" & vbCrLf
        script = script & "    }" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & "  if ($winrmServiceWasStarted) {" & vbCrLf
        script = script & "    Stop-Service -Name WinRM -Force -ErrorAction SilentlyContinue" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & "}" & vbCrLf
    End If

    BuildGetJobListScript = script
End Function

Private Function ParseJobListResult(result As String, rootPath As String) As Boolean
    ' 戻り値: True=成功, False=エラー
    ' JP1 ajsprint出力形式（ネスト対応）:
    '   unit=ユニット名,,admin,グループ;    ← 2番目のフィールドは空
    '   {
    '       ty=n;
    '       cm="コメント";
    '       unit=子ユニット名,,admin,グループ;  ← ネストされたユニット
    '       {
    '           ty=n;
    '           ...
    '       }
    '   }
    ' フルパス = ルートパス + "/" + ユニット名
    ' ※コマンド実行時は -F オプションでスケジューラサービスを指定するため
    '   パスにはスケジューラサービス名を含めない
    ParseJobListResult = False

    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_JOBLIST)

    ' 既存データをクリア
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_JOBNET_PATH).End(xlUp).Row
    If lastRow >= ROW_JOBLIST_DATA_START Then
        ws.Range(ws.Cells(ROW_JOBLIST_DATA_START, COL_SELECT), ws.Cells(lastRow, COL_LAST_MESSAGE)).Clear
    End If

    ' 結果をパース
    Dim lines() As String
    lines = Split(result, vbCrLf)

    Dim row As Long
    row = ROW_JOBLIST_DATA_START

    Dim i As Long

    ' ルートパスの末尾のスラッシュを正規化
    ' ただし、"/" のみの場合は空文字列にならないよう除外
    Dim basePath As String
    basePath = rootPath
    If Len(basePath) > 1 And Right(basePath, 1) = "/" Then
        basePath = Left(basePath, Len(basePath) - 1)
    End If

    ' ネスト対応のためスタック構造を使用
    ' 配列でスタックをシミュレート（最大ネスト深度10）
    Const MAX_DEPTH As Long = 10
    Dim unitStack(1 To MAX_DEPTH) As String   ' unit=...ヘッダーのスタック
    Dim blockStack(1 To MAX_DEPTH) As String  ' ブロック内容のスタック
    Dim pathStack(1 To MAX_DEPTH) As String   ' フルパスのスタック
    Dim rowStack(1 To MAX_DEPTH) As Long      ' 書き込み行番号のスタック（親を先に確保）
    Dim stackDepth As Long                     ' 現在のスタック深度

    stackDepth = 0

    For i = LBound(lines) To UBound(lines)
        Dim line As String
        ' Trim()はスペースのみ除去するため、TAB文字も明示的に除去
        line = Trim(Replace(lines(i), vbTab, ""))

        ' 空行はスキップ
        If line = "" Then GoTo NextLine

        ' エラーチェック
        If InStr(line, "ERROR:") > 0 Then
            MsgBox "エラーが発生しました:" & vbCrLf & line, vbExclamation
            Exit Function
        End If

        ' unit= で始まる行（ヘッダー）- ブロック内外問わず検出
        If InStr(line, "unit=") > 0 Then
            ' 次の行が{かどうかを先読み
            Dim nextIdx As Long
            nextIdx = i + 1
            If nextIdx <= UBound(lines) Then
                Dim nextLine As String
                nextLine = Trim(Replace(lines(nextIdx), vbTab, ""))
                If Left(nextLine, 1) = "{" Then
                    ' 新しいユニット定義開始 - スタックにプッシュ
                    stackDepth = stackDepth + 1
                    If stackDepth <= MAX_DEPTH Then
                        unitStack(stackDepth) = line
                        blockStack(stackDepth) = ""

                        ' ユニット名を取得（unit=の最初のフィールド）
                        Dim unitName As String
                        unitName = ExtractUnitName(line)

                        ' フルパスを構築
                        If stackDepth = 1 Then
                            ' ルートレベル: ajsprintで指定したパスのユニット自体が
                            ' 最初に出力されるため、basePathをそのまま使用
                            pathStack(stackDepth) = basePath
                        Else
                            ' ネストレベル: 親のパス + "/" + ユニット名
                            ' ただし、親パスが"/"の場合は"/"を追加しない（"//"を防ぐ）
                            If pathStack(stackDepth - 1) = "/" Then
                                pathStack(stackDepth) = "/" & unitName
                            Else
                                pathStack(stackDepth) = pathStack(stackDepth - 1) & "/" & unitName
                            End If
                        End If

                        ' 行番号を確保（親が先に行番号を取得するため、親が上に表示される）
                        rowStack(stackDepth) = row
                        row = row + 1
                    End If
                End If
            End If
            GoTo NextLine
        End If

        ' ブロック開始 {
        If Left(line, 1) = "{" Then
            ' {の後に内容がある場合
            If Len(line) > 1 And stackDepth > 0 Then
                blockStack(stackDepth) = blockStack(stackDepth) & " " & Mid(line, 2)
            End If
            GoTo NextLine
        End If

        ' ブロック終了 }
        If Right(line, 1) = "}" Or line = "}" Then
            ' }の前に内容がある場合
            If Len(line) > 1 And stackDepth > 0 Then
                blockStack(stackDepth) = blockStack(stackDepth) & " " & Left(line, Len(line) - 1)
            End If

            ' スタックからポップして処理
            If stackDepth > 0 Then
                Dim currentHeader As String
                Dim currentBlock As String
                Dim currentFullPath As String
                Dim currentRow As Long
                currentHeader = unitStack(stackDepth)
                currentBlock = blockStack(stackDepth)
                currentFullPath = pathStack(stackDepth)
                currentRow = rowStack(stackDepth)  ' 事前に確保した行番号を使用

                ' ユニットタイプを抽出（ty=xxx; から xxx を取得）
                Dim unitType As String
                Dim unitTypeDisplay As String
                unitType = ExtractUnitType(currentBlock)
                unitTypeDisplay = GetUnitTypeDisplayName(unitType)

                ' ty=が存在し、グループ以外の場合に一覧に追加
                ' グループ(g, mg)は実行できないため除外
                If unitType <> "" And currentFullPath <> "" And unitType <> "g" And unitType <> "mg" Then
                    ws.Cells(currentRow, COL_SELECT).Value = ChrW(&H2610)  ' ☐（空のチェックボックス）
                    ws.Cells(currentRow, COL_ORDER).Value = ""
                    ' 種別を設定
                    ws.Cells(currentRow, COL_UNIT_TYPE).Value = unitTypeDisplay
                    ws.Cells(currentRow, COL_UNIT_TYPE).HorizontalAlignment = xlCenter
                    ' フルパス（ルートからのパス）を設定
                    ws.Cells(currentRow, COL_JOBNET_PATH).Value = currentFullPath
                    ' ユニット名を設定（unit=の最初のフィールド）
                    ws.Cells(currentRow, COL_JOBNET_NAME).Value = ExtractUnitName(currentHeader)
                    ws.Cells(currentRow, COL_COMMENT).Value = ExtractCommentFromBlock(currentBlock)
                    ' スクリプトファイル名 (sc=)
                    ws.Cells(currentRow, COL_SCRIPT).Value = ExtractAttributeFromBlock(currentBlock, "sc")
                    ' パラメーター (prm=)
                    ws.Cells(currentRow, COL_PARAMETER).Value = ExtractAttributeFromBlock(currentBlock, "prm")
                    ' ワークパス (wkp=)
                    ws.Cells(currentRow, COL_WORK_PATH).Value = ExtractAttributeFromBlock(currentBlock, "wkp")

                    ' 保留状態を解析
                    Dim isHold As Boolean
                    isHold = (InStr(currentBlock, "hd=h") > 0) Or (InStr(currentBlock, "hd=H") > 0)

                    If isHold Then
                        ws.Cells(currentRow, COL_HOLD).Value = "保留中"
                        ws.Cells(currentRow, COL_HOLD).HorizontalAlignment = xlCenter
                        ws.Cells(currentRow, COL_HOLD).Interior.Color = RGB(255, 235, 156)  ' 保留列のみ黄色
                        ws.Cells(currentRow, COL_HOLD).Font.Bold = True
                        ws.Cells(currentRow, COL_HOLD).Font.Color = RGB(156, 87, 0)
                    Else
                        ws.Cells(currentRow, COL_HOLD).Value = ""
                    End If

                    ' 選択列・順序列の書式
                    With ws.Cells(currentRow, COL_SELECT)
                        .HorizontalAlignment = xlCenter
                    End With
                    With ws.Cells(currentRow, COL_ORDER)
                        .HorizontalAlignment = xlCenter
                    End With

                    ' 罫線
                    ws.Range(ws.Cells(currentRow, COL_SELECT), ws.Cells(currentRow, COL_LAST_MESSAGE)).Borders.LineStyle = xlContinuous
                End If

                ' スタックをクリアしてポップ
                unitStack(stackDepth) = ""
                blockStack(stackDepth) = ""
                pathStack(stackDepth) = ""
                rowStack(stackDepth) = 0
                stackDepth = stackDepth - 1
            End If
            GoTo NextLine
        End If

        ' ブロック内のコンテンツを収集（ただし、次のunit=行の前まで）
        If stackDepth > 0 Then
            ' 次の行がunit=かチェック
            Dim isNextUnit As Boolean
            isNextUnit = False
            If i + 1 <= UBound(lines) Then
                Dim checkLine As String
                checkLine = Trim(Replace(lines(i + 1), vbTab, ""))
                If InStr(checkLine, "unit=") > 0 Then
                    isNextUnit = True
                End If
            End If

            ' unit=の直前でなければブロック内容に追加
            If Not isNextUnit Then
                blockStack(stackDepth) = blockStack(stackDepth) & " " & line
            End If
        End If

NextLine:
    Next i

    ' グループ除外により空になった行を削除（下から上に削除）
    Dim deleteRow As Long
    For deleteRow = row - 1 To ROW_JOBLIST_DATA_START Step -1
        If ws.Cells(deleteRow, COL_JOBNET_PATH).Value = "" Then
            ws.Rows(deleteRow).Delete
        End If
    Next deleteRow

    ' データがない場合（空行削除後に再チェック）
    Dim actualLastRow As Long
    actualLastRow = ws.Cells(ws.Rows.Count, COL_JOBNET_PATH).End(xlUp).row
    If actualLastRow < ROW_JOBLIST_DATA_START Then
        MsgBox "実行可能なユニットが見つかりませんでした。" & vbCrLf & _
               "（グループは除外されます）", vbExclamation
        Exit Function
    End If

    ' 成功
    ParseJobListResult = True
End Function

Private Function ExtractJobNameFromHeader(header As String) As String
    ' unit=パス,名前,admin,グループ; から名前を抽出
    Dim parts() As String
    Dim unitPart As String
    Dim startPos As Long

    startPos = InStr(header, "unit=")
    If startPos > 0 Then
        unitPart = Mid(header, startPos + 5)
        ' セミコロンを除去
        If Right(unitPart, 1) = ";" Then
            unitPart = Left(unitPart, Len(unitPart) - 1)
        End If
        ' カンマで分割
        parts = Split(unitPart, ",")
        If UBound(parts) >= 1 Then
            ExtractJobNameFromHeader = parts(1)
        End If
    End If
End Function

Private Function ExtractCommentFromBlock(blockContent As String) As String
    ' cm="コメント"; からコメントを抽出
    Dim startPos As Long
    Dim endPos As Long

    startPos = InStr(blockContent, "cm=""")
    If startPos > 0 Then
        startPos = startPos + 4
        endPos = InStr(startPos, blockContent, """")
        If endPos > startPos Then
            ExtractCommentFromBlock = Mid(blockContent, startPos, endPos - startPos)
        End If
    End If
End Function

Private Function ExtractAttributeFromBlock(blockContent As String, attrName As String) As String
    ' 指定された属性名の値を抽出
    ' 形式1: attr="value"; (ダブルクォート囲み)
    ' 形式2: attr=value; (クォートなし)
    Dim startPos As Long
    Dim endPos As Long
    Dim searchStr As String

    ExtractAttributeFromBlock = ""

    ' ダブルクォート形式を先にチェック: attr="value"
    searchStr = attrName & "="""
    startPos = InStr(blockContent, searchStr)
    If startPos > 0 Then
        startPos = startPos + Len(searchStr)
        endPos = InStr(startPos, blockContent, """")
        If endPos > startPos Then
            ExtractAttributeFromBlock = Mid(blockContent, startPos, endPos - startPos)
            Exit Function
        End If
    End If

    ' クォートなし形式: attr=value;
    searchStr = attrName & "="
    startPos = InStr(blockContent, searchStr)
    If startPos > 0 Then
        startPos = startPos + Len(searchStr)
        ' セミコロンまたはスペースまでを取得
        endPos = InStr(startPos, blockContent, ";")
        Dim endPosSpace As Long
        endPosSpace = InStr(startPos, blockContent, " ")

        If endPos > startPos Then
            If endPosSpace > startPos And endPosSpace < endPos Then
                endPos = endPosSpace
            End If
            ExtractAttributeFromBlock = Mid(blockContent, startPos, endPos - startPos)
        End If
    End If
End Function

Private Function ExtractUnitPath(line As String) As String
    ' unit=/path/to/jobnet から /path/to/jobnet を抽出
    ' 注: JP1のajsprintでは最初のフィールドがユニット名
    Dim startPos As Long
    Dim endPos As Long

    startPos = InStr(line, "unit=")
    If startPos > 0 Then
        startPos = startPos + 5
        endPos = InStr(startPos, line, ",")
        If endPos > startPos Then
            ExtractUnitPath = Mid(line, startPos, endPos - startPos)
        End If
    End If
End Function

Private Function ExtractUnitName(line As String) As String
    ' unit=ユニット名,,admin,group; からユニット名を抽出
    ' 最初のフィールド（カンマまで）を返す
    ExtractUnitName = ExtractUnitPath(line)
End Function

Private Function ExtractUnitType(blockContent As String) As String
    ' ty=xxx; から xxx を抽出
    ' 例: ty=n; → n, ty=pj; → pj, ty=jdj; → jdj
    Dim startPos As Long
    Dim endPos As Long
    Dim tyValue As String

    ExtractUnitType = ""

    startPos = InStr(blockContent, "ty=")
    If startPos > 0 Then
        startPos = startPos + 3
        ' セミコロンまたはスペースまでを取得
        endPos = InStr(startPos, blockContent, ";")
        Dim endPosSpace As Long
        endPosSpace = InStr(startPos, blockContent, " ")

        If endPos > startPos Then
            If endPosSpace > startPos And endPosSpace < endPos Then
                endPos = endPosSpace
            End If
            ExtractUnitType = Mid(blockContent, startPos, endPos - startPos)
        End If
    End If
End Function

Private Function GetUnitTypeDisplayName(unitType As String) As String
    ' ユニットタイプコードを日本語表示名に変換
    ' JP1/AJS3の全ユニット種別に対応
    Select Case LCase(unitType)
        ' グループ系
        Case "g"
            GetUnitTypeDisplayName = "グループ"
        Case "mg"
            GetUnitTypeDisplayName = "マネージャーグループ"

        ' ジョブネット系
        Case "n"
            GetUnitTypeDisplayName = "ジョブネット"
        Case "rn"
            GetUnitTypeDisplayName = "リカバリーネット"
        Case "rm"
            GetUnitTypeDisplayName = "リモートネット"
        Case "mn"
            GetUnitTypeDisplayName = "マネージャーネット"

        ' 標準ジョブ系
        Case "j"
            GetUnitTypeDisplayName = "ジョブ"
        Case "rj"
            GetUnitTypeDisplayName = "リカバリージョブ"
        Case "pj"
            GetUnitTypeDisplayName = "判定ジョブ"
        Case "rp"
            GetUnitTypeDisplayName = "リカバリー判定"
        Case "qj"
            GetUnitTypeDisplayName = "キュージョブ"
        Case "rq"
            GetUnitTypeDisplayName = "リカバリーキュー"

        ' 判定変数系
        Case "jdj"
            GetUnitTypeDisplayName = "判定変数参照"
        Case "rjdj"
            GetUnitTypeDisplayName = "リカバリー判定変数"
        Case "orj"
            GetUnitTypeDisplayName = "OR分岐"
        Case "rorj"
            GetUnitTypeDisplayName = "リカバリーOR分岐"

        ' イベント監視系
        Case "evwj"
            GetUnitTypeDisplayName = "イベント監視"
        Case "revwj"
            GetUnitTypeDisplayName = "リカバリーイベント"
        Case "flwj"
            GetUnitTypeDisplayName = "ファイル監視"
        Case "rflwj"
            GetUnitTypeDisplayName = "リカバリーファイル監視"
        Case "mlwj"
            GetUnitTypeDisplayName = "メール受信監視"
        Case "rmlwj"
            GetUnitTypeDisplayName = "リカバリーメール受信"
        Case "mqwj"
            GetUnitTypeDisplayName = "MQ受信監視"
        Case "rmqwj"
            GetUnitTypeDisplayName = "リカバリーMQ受信"
        Case "mswj"
            GetUnitTypeDisplayName = "MSMQメッセージ受信監視"
        Case "rmswj"
            GetUnitTypeDisplayName = "リカバリーMSMQ受信"
        Case "lfwj"
            GetUnitTypeDisplayName = "ログファイル監視"
        Case "rlfwj"
            GetUnitTypeDisplayName = "リカバリーログ監視"
        Case "ntwj"
            GetUnitTypeDisplayName = "Windows NT イベントログ監視"
        Case "rntwj"
            GetUnitTypeDisplayName = "リカバリー NT イベントログ"
        Case "tmwj"
            GetUnitTypeDisplayName = "実行間隔制御"
        Case "rtmwj"
            GetUnitTypeDisplayName = "リカバリー実行間隔"

        ' 送信系
        Case "evsj"
            GetUnitTypeDisplayName = "イベント送信"
        Case "revsj"
            GetUnitTypeDisplayName = "リカバリーイベント送信"
        Case "mlsj"
            GetUnitTypeDisplayName = "メール送信"
        Case "rmlsj"
            GetUnitTypeDisplayName = "リカバリーメール送信"
        Case "mqsj"
            GetUnitTypeDisplayName = "MQ送信"
        Case "rmqsj"
            GetUnitTypeDisplayName = "リカバリーMQ送信"
        Case "mssj"
            GetUnitTypeDisplayName = "MSMQメッセージ送信"
        Case "rmssj"
            GetUnitTypeDisplayName = "リカバリーMSMQ送信"
        Case "cmsj"
            GetUnitTypeDisplayName = "JP1イベント送信"
        Case "rcmsj"
            GetUnitTypeDisplayName = "リカバリーJP1送信"

        ' PowerShell系
        Case "pwlj"
            GetUnitTypeDisplayName = "ローカルPowerShell"
        Case "rpwlj"
            GetUnitTypeDisplayName = "リカバリーローカルPS"
        Case "pwrj"
            GetUnitTypeDisplayName = "リモートPowerShell"
        Case "rpwrj"
            GetUnitTypeDisplayName = "リカバリーリモートPS"

        ' カスタム系
        Case "cj"
            GetUnitTypeDisplayName = "カスタムジョブ"
        Case "rcj"
            GetUnitTypeDisplayName = "リカバリーカスタム"
        Case "cpj"
            GetUnitTypeDisplayName = "カスタムPCジョブ"
        Case "rcpj"
            GetUnitTypeDisplayName = "リカバリーカスタムPC"

        ' 外部連携系
        Case "fxj"
            GetUnitTypeDisplayName = "ファイル転送"
        Case "rfxj"
            GetUnitTypeDisplayName = "リカバリーファイル転送"
        Case "htpj"
            GetUnitTypeDisplayName = "HTTP接続"
        Case "rhtpj"
            GetUnitTypeDisplayName = "リカバリーHTTP"

        ' その他
        Case "nc"
            GetUnitTypeDisplayName = "ネットコネクタ"
        Case "hln"
            GetUnitTypeDisplayName = "リンク"
        Case "rc"
            GetUnitTypeDisplayName = "リリースコネクタ"
        Case "rr"
            GetUnitTypeDisplayName = "ルートジョブネット起動条件"

        ' 未知のタイプはそのまま表示
        Case Else
            If unitType <> "" Then
                GetUnitTypeDisplayName = unitType
            Else
                GetUnitTypeDisplayName = ""
            End If
    End Select
End Function

Private Function ExtractJobName(line As String) As String
    ' unit=/path/to/jobnet,ジョブ名,ty=n から ジョブ名 を抽出
    ' ajsprintの出力形式: unit=/path,name,ty=type,cm="comment";
    Dim startPos As Long
    Dim endPos As Long
    Dim fields() As String
    Dim unitPart As String

    ' unit= の後ろを取得
    startPos = InStr(line, "unit=")
    If startPos > 0 Then
        unitPart = Mid(line, startPos + 5)
        ' セミコロンまでを取得
        endPos = InStr(unitPart, ";")
        If endPos > 0 Then
            unitPart = Left(unitPart, endPos - 1)
        End If

        ' カンマで分割
        fields = Split(unitPart, ",")

        ' 2番目のフィールドがジョブ名（ty=で始まらない場合）
        If UBound(fields) >= 1 Then
            If InStr(fields(1), "ty=") = 0 And InStr(fields(1), "cm=") = 0 Then
                ExtractJobName = Trim(fields(1))
                Exit Function
            End If
        End If

        ' 2番目がty=の場合はパスの最後の部分を使用
        If UBound(fields) >= 0 Then
            ExtractJobName = GetLastPathComponent(fields(0))
        End If
    End If
End Function

Private Function ExtractComment(line As String) As String
    ' cm="comment" からコメントを抽出
    Dim startPos As Long
    Dim endPos As Long

    startPos = InStr(line, "cm=""")
    If startPos > 0 Then
        startPos = startPos + 4
        endPos = InStr(startPos, line, """")
        If endPos > startPos Then
            ExtractComment = Mid(line, startPos, endPos - startPos)
        End If
    End If
End Function

Private Function ExtractHoldStatus(line As String) As Boolean
    ' hd=y（保留）を検出
    ' JP1のajsprint出力で hd=y はホールド(保留)を示す
    ExtractHoldStatus = (InStr(line, ",hd=y") > 0 Or InStr(line, " hd=y") > 0)
End Function

Private Function GetLastPathComponent(path As String) As String
    Dim parts() As String
    parts = Split(path, "/")
    If UBound(parts) >= 0 Then
        GetLastPathComponent = parts(UBound(parts))
    End If
End Function

'==============================================================================
' 選択ジョブ実行
'==============================================================================
Public Sub ExecuteCheckedJobs()
    On Error GoTo ErrorHandler

    Dim config As Object
    Set config = GetConfig()

    If config Is Nothing Then Exit Sub

    ' リモートモードの場合、管理者権限をチェック
    If Not EnsureAdminForRemoteMode(config) Then Exit Sub

    ' パスワード入力（リモートモードの場合のみリモートパスワードが必要）
    If config("ExecMode") <> "ローカル" Then
        If config("RemotePassword") = "" Then
            config("RemotePassword") = InputBox("リモートサーバのパスワードを入力してください:", "パスワード入力")
            If config("RemotePassword") = "" Then
                MsgBox "パスワードが入力されませんでした。", vbExclamation
                Exit Sub
            End If
        End If
    End If

    If config("JP1Password") = "" Then
        config("JP1Password") = InputBox("JP1パスワードを入力してください:", "パスワード入力")
        If config("JP1Password") = "" Then
            MsgBox "パスワードが入力されませんでした。", vbExclamation
            Exit Sub
        End If
    End If

    ' 順序が指定されたジョブを取得
    Dim jobs As Collection
    Set jobs = GetOrderedJobs()

    If jobs.Count = 0 Then
        MsgBox "実行するジョブが選択されていません。" & vbCrLf & _
               "ジョブ一覧シートの「順序」列に数字（1, 2, 3...）を入力してください。", vbExclamation
        Exit Sub
    End If

    ' 順序のバリデーション
    Dim validationError As String
    validationError = ValidateJobOrder(jobs)
    If validationError <> "" Then
        MsgBox validationError, vbExclamation, "順序指定エラー"
        Exit Sub
    End If

    ' 保留中のジョブ数をカウント
    Dim holdCount As Long
    holdCount = 0
    Dim j As Variant
    For Each j In jobs
        If j("IsHold") Then holdCount = holdCount + 1
    Next j

    ' 確認
    Dim msg As String
    msg = "以下の " & jobs.Count & " 件のジョブを実行します：" & vbCrLf & vbCrLf
    Dim cnt As Long
    cnt = 0
    For Each j In jobs
        cnt = cnt + 1
        If cnt <= 5 Then
            Dim holdMark As String
            If j("IsHold") Then
                holdMark = " [保留中]"
            Else
                holdMark = ""
            End If
            msg = msg & cnt & ". " & j("Path") & holdMark & vbCrLf
        ElseIf cnt = 6 Then
            msg = msg & "..." & vbCrLf
        End If
    Next j

    If holdCount > 0 Then
        msg = msg & vbCrLf & "※ 保留中のジョブが " & holdCount & " 件あります。自動で保留解除してから実行します。" & vbCrLf
    End If
    msg = msg & vbCrLf & "実行しますか？"

    If MsgBox(msg, vbYesNo + vbQuestion, "実行確認") = vbNo Then Exit Sub

    ' ログファイルの初期化
    g_LogFilePath = CreateLogFile()

    ' 実行
    Application.ScreenUpdating = False

    Dim wsLog As Worksheet
    Set wsLog = Worksheets(SHEET_LOG)
    Dim logRow As Long
    logRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1
    If logRow < 5 Then logRow = 5

    Dim success As Boolean
    success = True

    For Each j In jobs
        Application.StatusBar = "実行中: " & j("Path")

        Dim execResult As Object
        Set execResult = ExecuteSingleJob(config, j("Path"), j("IsHold"), g_LogFilePath)

        ' 結果をログに記録
        wsLog.Cells(logRow, 1).Value = Now
        wsLog.Cells(logRow, 2).Value = j("Path")
        wsLog.Cells(logRow, 3).Value = execResult("Status")
        wsLog.Cells(logRow, 4).Value = execResult("StartTime")
        wsLog.Cells(logRow, 5).Value = execResult("EndTime")

        ' F列にログパスをハイパーリンク付きで設定
        If execResult("LogPath") <> "" Then
            wsLog.Cells(logRow, 6).Value = execResult("LogPath")
            On Error Resume Next
            wsLog.Hyperlinks.Add Anchor:=wsLog.Cells(logRow, 6), _
                                 Address:=execResult("LogPath"), _
                                 TextToDisplay:=execResult("LogPath")
            On Error GoTo 0
        End If

        ' 色付け（ジョブ一覧シートと同じ配色）
        If execResult("Status") = "正常終了" Then
            wsLog.Cells(logRow, 3).Interior.Color = RGB(198, 239, 206)  ' 緑（正常）
        ElseIf execResult("Status") = "起動成功" Then
            wsLog.Cells(logRow, 3).Interior.Color = RGB(255, 235, 156)  ' 黄（起動のみ）
        ElseIf execResult("Status") = "警告検出終了" Or execResult("Status") = "警告終了" Then
            wsLog.Cells(logRow, 3).Interior.Color = RGB(255, 192, 0)    ' オレンジ（警告）
        Else
            wsLog.Cells(logRow, 3).Interior.Color = RGB(255, 199, 206)  ' 赤（異常）
        End If

        wsLog.Range(wsLog.Cells(logRow, 1), wsLog.Cells(logRow, 6)).Borders.LineStyle = xlContinuous

        ' ジョブ一覧シートも更新
        UpdateJobListStatus j("Row"), execResult

        logRow = logRow + 1

        ' エラー・警告時は停止
        If execResult("Status") <> "正常終了" And execResult("Status") <> "起動成功" Then
            success = False
            ' 警告検出終了と異常終了で異なるメッセージを表示
            If execResult("Status") = "警告検出終了" Or execResult("Status") = "警告終了" Then
                MsgBox "ジョブ「" & j("Path") & "」で警告が検出されました。" & vbCrLf & _
                       "処理を中断します。" & vbCrLf & vbCrLf & _
                       "詳細: " & execResult("Message") & vbCrLf & vbCrLf & _
                       "実行ログ: " & g_LogFilePath, vbExclamation, "警告検出"
            Else
                MsgBox "ジョブ「" & j("Path") & "」が失敗しました。" & vbCrLf & _
                       "処理を中断します。" & vbCrLf & vbCrLf & _
                       "詳細: " & execResult("Message") & vbCrLf & vbCrLf & _
                       "実行ログ: " & g_LogFilePath, vbCritical, "異常終了"
            End If
            Exit For
        End If
    Next j

    Application.StatusBar = False
    Application.ScreenUpdating = True

    If success Then
        MsgBox "すべてのジョブが正常に完了しました。" & vbCrLf & vbCrLf & _
               "実行ログ: " & g_LogFilePath, vbInformation
    End If

    Worksheets(SHEET_LOG).Activate
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました。" & vbCrLf & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description & vbCrLf & _
           "発生場所: ExecuteCheckedJobs", vbCritical, "VBAエラー"
End Sub

Private Function GetOrderedJobs() As Collection
    Dim jobs As New Collection
    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_JOBLIST)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_JOBNET_PATH).End(xlUp).Row

    ' 順序が入力されている行を収集
    Dim orderedRows As New Collection
    Dim row As Long
    For row = ROW_JOBLIST_DATA_START To lastRow
        Dim orderValue As Variant
        orderValue = ws.Cells(row, COL_ORDER).Value

        ' 順序列に数字が入っている場合のみ対象
        If IsNumeric(orderValue) And orderValue <> "" Then
            Dim job As Object
            Set job = CreateObject("Scripting.Dictionary")
            job("Row") = row
            job("Path") = ws.Cells(row, COL_JOBNET_PATH).Value
            job("Order") = CLng(orderValue)
            ' 保留状態を取得
            job("IsHold") = (ws.Cells(row, COL_HOLD).Value = "保留中")

            orderedRows.Add job
        End If
    Next row

    ' 実行順でソート（単純なバブルソート）
    If orderedRows.Count = 0 Then
        Set GetOrderedJobs = jobs
        Exit Function
    End If

    Dim arr() As Variant
    ReDim arr(1 To orderedRows.Count)
    Dim i As Long
    Dim k As Long  ' ソート用ループ変数
    For i = 1 To orderedRows.Count
        Set arr(i) = orderedRows(i)
    Next i

    Dim temp As Object
    For i = 1 To UBound(arr) - 1
        For k = i + 1 To UBound(arr)
            If arr(i)("Order") > arr(k)("Order") Then
                Set temp = arr(i)
                Set arr(i) = arr(k)
                Set arr(k) = temp
            End If
        Next k
    Next i

    For i = 1 To UBound(arr)
        jobs.Add arr(i)
    Next i

    Set GetOrderedJobs = jobs
End Function

Private Function ValidateJobOrder(jobs As Collection) As String
    ' 順序指定のバリデーション
    ' 戻り値: エラーメッセージ（正常な場合は空文字）

    If jobs.Count = 0 Then
        ValidateJobOrder = ""
        Exit Function
    End If

    ' 順序番号を配列に収集
    Dim orders() As Long
    ReDim orders(1 To jobs.Count)
    Dim i As Long
    i = 0
    Dim j As Variant
    For Each j In jobs
        i = i + 1
        orders(i) = j("Order")
    Next j

    ' 重複チェック
    Dim k As Long
    For i = 1 To UBound(orders) - 1
        For k = i + 1 To UBound(orders)
            If orders(i) = orders(k) Then
                ValidateJobOrder = "順序番号 " & orders(i) & " が重複しています。" & vbCrLf & _
                                   "各ジョブには異なる順序番号を指定してください。"
                Exit Function
            End If
        Next k
    Next i

    ' 連続性チェック（1から始まって連続しているか）
    ' まずソート
    Dim temp As Long
    For i = 1 To UBound(orders) - 1
        For k = i + 1 To UBound(orders)
            If orders(i) > orders(k) Then
                temp = orders(i)
                orders(i) = orders(k)
                orders(k) = temp
            End If
        Next k
    Next i

    ' 1から始まっているか
    If orders(1) <> 1 Then
        ValidateJobOrder = "順序番号は 1 から開始してください。" & vbCrLf & _
                           "現在の最小値: " & orders(1)
        Exit Function
    End If

    ' 連続しているか
    For i = 2 To UBound(orders)
        If orders(i) <> orders(i - 1) + 1 Then
            ValidateJobOrder = "順序番号が連続していません。" & vbCrLf & _
                               orders(i - 1) & " の次は " & (orders(i - 1) + 1) & " を指定してください。" & vbCrLf & _
                               "（現在: " & orders(i) & "）"
            Exit Function
        End If
    Next i

    ValidateJobOrder = ""
End Function

Private Function ExecuteSingleJob(ByVal config As Object, ByVal jobnetPath As String, ByVal isHold As Boolean, ByVal logFilePath As String) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result("Status") = ""
    result("StartTime") = ""
    result("EndTime") = ""
    result("Message") = ""

    Dim waitCompletion As Boolean
    waitCompletion = (config("WaitCompletion") = "はい")

    Dim psScript As String
    psScript = BuildExecuteJobScript(config, jobnetPath, waitCompletion, isHold, logFilePath)

    Dim output As String
    output = ExecutePowerShell(psScript)

    ' 結果をパース
    Dim lines() As String
    lines = Split(output, vbCrLf)

    Dim line As String
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        line = lines(i)
        If InStr(line, "RESULT_STATUS:") > 0 Then
            result("Status") = Trim(Replace(line, "RESULT_STATUS:", ""))
        ElseIf InStr(line, "RESULT_START:") > 0 Then
            result("StartTime") = Trim(Replace(line, "RESULT_START:", ""))
        ElseIf InStr(line, "RESULT_END:") > 0 Then
            result("EndTime") = Trim(Replace(line, "RESULT_END:", ""))
        ElseIf InStr(line, "RESULT_MESSAGE:") > 0 Then
            result("Message") = Trim(Replace(line, "RESULT_MESSAGE:", ""))
        ElseIf InStr(line, "RESULT_LOGPATH:") > 0 Then
            result("LogPath") = Trim(Replace(line, "RESULT_LOGPATH:", ""))
        ElseIf InStr(line, "RESULT_DETAIL:") > 0 Then
            result("Detail") = Trim(Replace(line, "RESULT_DETAIL:", ""))
        ElseIf InStr(line, "ERROR:") > 0 Then
            result("Status") = "エラー"
            result("Message") = line
        End If
    Next i

    If result("Status") = "" Then
        result("Status") = "不明"
        result("Message") = output
    End If

    Set ExecuteSingleJob = result
End Function

Private Function BuildExecuteJobScript(ByVal config As Object, ByVal jobnetPath As String, ByVal waitCompletion As Boolean, ByVal isHold As Boolean, ByVal logFilePath As String) As String
    Dim script As String
    Dim isRemote As Boolean
    isRemote = (config("ExecMode") <> "ローカル")

    script = "$ErrorActionPreference = 'Stop'" & vbCrLf
    script = script & vbCrLf

    ' UTF-8エンコーディング設定
    script = script & "# UTF-8エンコーディング設定" & vbCrLf
    script = script & "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8" & vbCrLf
    script = script & "$OutputEncoding = [System.Text.Encoding]::UTF8" & vbCrLf
    script = script & "chcp 65001 | Out-Null" & vbCrLf
    script = script & vbCrLf

    ' ログ出力関数を定義
    script = script & "# デバッグモードフラグ" & vbCrLf
    script = script & "$debugMode = $" & IIf(DEBUG_MODE, "true", "false") & vbCrLf
    script = script & vbCrLf
    script = script & "# ログ出力関数" & vbCrLf
    script = script & "$logFile = '" & Replace(logFilePath, "'", "''") & "'" & vbCrLf
    script = script & "function Write-Log {" & vbCrLf
    script = script & "  param([string]$Message)" & vbCrLf
    script = script & "  if ($Message -match '^\[DEBUG-' -and -not $debugMode) { return }" & vbCrLf
    script = script & "  $timestamp = Get-Date -Format 'yyyy/MM/dd HH:mm:ss'" & vbCrLf
    script = script & "  $logLine = ""[$timestamp] $Message""" & vbCrLf
    script = script & "  Write-Host $logLine" & vbCrLf
    script = script & "  Add-Content -Path $logFile -Value $logLine -Encoding UTF8" & vbCrLf
    script = script & "}" & vbCrLf
    script = script & vbCrLf

    ' 実行モードフラグ
    script = script & "# 実行モード設定" & vbCrLf
    script = script & "$isRemote = $" & IIf(isRemote, "true", "false") & vbCrLf
    script = script & "$session = $null" & vbCrLf
    script = script & vbCrLf

    ' JP1コマンド実行関数（ローカル/リモート共通）
    script = script & "# JP1コマンド実行関数（ローカル/リモート共通）" & vbCrLf
    script = script & "function Invoke-JP1Command {" & vbCrLf
    script = script & "  param([string]$CommandName, [string[]]$Arguments)" & vbCrLf
    script = script & "  $searchPaths = @('C:\Program Files\HITACHI\JP1AJS3\bin','C:\Program Files (x86)\HITACHI\JP1AJS3\bin','C:\Program Files\Hitachi\JP1AJS2\bin','C:\Program Files (x86)\Hitachi\JP1AJS2\bin')" & vbCrLf
    script = script & "  if ($isRemote) {" & vbCrLf
    script = script & "    Invoke-Command -Session $session -ScriptBlock {" & vbCrLf
    script = script & "      param($cmdName, $cmdArgs, $paths)" & vbCrLf
    script = script & "      $cmdPath = $null" & vbCrLf
    script = script & "      foreach ($p in $paths) { if (Test-Path ""$p\$cmdName"") { $cmdPath = ""$p\$cmdName""; break } }" & vbCrLf
    script = script & "      if (-not $cmdPath) { return @{ ExitCode = 1; Output = ""$cmdName not found"" } }" & vbCrLf
    script = script & "      $output = & $cmdPath $cmdArgs 2>&1" & vbCrLf
    script = script & "      @{ ExitCode = $LASTEXITCODE; Output = $output }" & vbCrLf
    script = script & "    } -ArgumentList $CommandName, $Arguments, $searchPaths" & vbCrLf
    script = script & "  } else {" & vbCrLf
    script = script & "    $cmdPath = $null" & vbCrLf
    script = script & "    foreach ($p in $searchPaths) { if (Test-Path ""$p\$CommandName"") { $cmdPath = ""$p\$CommandName""; break } }" & vbCrLf
    script = script & "    if (-not $cmdPath) { return @{ ExitCode = 1; Output = ""$CommandName not found"" } }" & vbCrLf
    script = script & "    $output = & $cmdPath $Arguments 2>&1" & vbCrLf
    script = script & "    @{ ExitCode = $LASTEXITCODE; Output = $output }" & vbCrLf
    script = script & "  }" & vbCrLf
    script = script & "}" & vbCrLf
    script = script & vbCrLf

    ' ファイル読み取り関数（ローカル/リモート共通）
    script = script & "# ファイル読み取り関数（ローカル/リモート共通）" & vbCrLf
    script = script & "function Read-FileContent {" & vbCrLf
    script = script & "  param([string]$FilePath)" & vbCrLf
    script = script & "  if ($isRemote) {" & vbCrLf
    script = script & "    Invoke-Command -Session $session -ScriptBlock {" & vbCrLf
    script = script & "      param($path)" & vbCrLf
    script = script & "      if (Test-Path $path) { Get-Content $path -Encoding Default -ErrorAction SilentlyContinue }" & vbCrLf
    script = script & "    } -ArgumentList $FilePath" & vbCrLf
    script = script & "  } else {" & vbCrLf
    script = script & "    if (Test-Path $FilePath) { Get-Content $FilePath -Encoding Default -ErrorAction SilentlyContinue }" & vbCrLf
    script = script & "  }" & vbCrLf
    script = script & "}" & vbCrLf
    script = script & vbCrLf

    ' リモートモードの場合: WinRM設定変数
    If isRemote Then
        script = script & "# WinRM設定変数" & vbCrLf
        script = script & "$originalTrustedHosts = $null" & vbCrLf
        script = script & "$winrmConfigChanged = $false" & vbCrLf
        script = script & "$winrmServiceWasStarted = $false" & vbCrLf
        script = script & "$securePass = ConvertTo-SecureString '" & EscapePSString(config("RemotePassword")) & "' -AsPlainText -Force" & vbCrLf
        script = script & "$cred = New-Object System.Management.Automation.PSCredential('" & config("RemoteUser") & "', $securePass)" & vbCrLf
        script = script & vbCrLf
    End If

    script = script & "try {" & vbCrLf
    script = script & "  Write-Log '--------------------------------------------------------------------------------'" & vbCrLf
    script = script & "  Write-Log 'ジョブネット: " & jobnetPath & "'" & vbCrLf

    If isRemote Then
        script = script & "  Write-Log '接続先: " & config("JP1Server") & " (リモートモード)'" & vbCrLf
    End If

    script = script & "  Write-Log '--------------------------------------------------------------------------------'" & vbCrLf
    script = script & vbCrLf

    ' リモートモードの場合: WinRMセットアップ
    If isRemote Then
        script = script & "  # WinRMサービス起動確認" & vbCrLf
        script = script & "  $winrmService = Get-Service -Name WinRM -ErrorAction SilentlyContinue" & vbCrLf
        script = script & "  if ($winrmService.Status -ne 'Running') {" & vbCrLf
        script = script & "    Write-Log '[準備] WinRMサービスを起動'" & vbCrLf
        script = script & "    Start-Service -Name WinRM -ErrorAction Stop" & vbCrLf
        script = script & "    $winrmServiceWasStarted = $true" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & "  $originalTrustedHosts = (Get-Item WSMan:\localhost\Client\TrustedHosts -ErrorAction SilentlyContinue).Value" & vbCrLf
        script = script & "  if ($originalTrustedHosts -notmatch '" & config("JP1Server") & "') {" & vbCrLf
        script = script & "    Write-Log '[準備] TrustedHostsに接続先を追加'" & vbCrLf
        script = script & "    if ($originalTrustedHosts) {" & vbCrLf
        script = script & "      Set-Item WSMan:\localhost\Client\TrustedHosts -Value ""$originalTrustedHosts," & config("JP1Server") & """ -Force" & vbCrLf
        script = script & "    } else {" & vbCrLf
        script = script & "      Set-Item WSMan:\localhost\Client\TrustedHosts -Value '" & config("JP1Server") & "' -Force" & vbCrLf
        script = script & "    }" & vbCrLf
        script = script & "    $winrmConfigChanged = $true" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & "  Write-Log '[接続] リモートセッション作成中...'" & vbCrLf
        script = script & "  $session = New-PSSession -ComputerName '" & config("JP1Server") & "' -Credential $cred -ErrorAction Stop" & vbCrLf
        script = script & "  Write-Log '[接続] セッション確立完了'" & vbCrLf
        script = script & vbCrLf
    End If

    ' 保留解除処理（共通）
    If isHold Then
        script = script & "  # 保留解除" & vbCrLf
        script = script & "  Write-Log '[実行] ajsplan -r - 保留解除'" & vbCrLf
        script = script & "  $releaseResult = Invoke-JP1Command 'ajsplan.exe' @('-F', '" & config("SchedulerService") & "', '-r', '" & jobnetPath & "')" & vbCrLf
        script = script & "  Write-Log ""結果: $($releaseResult.Output -join ' ')""" & vbCrLf
        script = script & "  if ($releaseResult.ExitCode -ne 0) {" & vbCrLf
        script = script & "    Write-Log '[ERROR] 保留解除失敗'" & vbCrLf
        script = script & "    Write-Output ""RESULT_STATUS:保留解除失敗""" & vbCrLf
        script = script & "    Write-Output ""RESULT_MESSAGE:$($releaseResult.Output -join ' ')""" & vbCrLf
        script = script & "    Write-Output ""RESULT_LOGPATH:$logFile""" & vbCrLf
        script = script & "    exit" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & "  Write-Log '[成功] 保留解除完了'" & vbCrLf
        script = script & vbCrLf
    End If

    ' ajsentry実行前に現在の最新実行IDを取得（比較用）
    script = script & "  # ajsentry実行前の実行IDを取得（比較用）" & vbCrLf
    script = script & "  $beforeIdResult = Invoke-JP1Command 'ajsshow.exe' @('-F', '" & config("SchedulerService") & "', '-g', '1', '-i', '%##', '" & jobnetPath & "')" & vbCrLf
    script = script & "  $beforeExecId = ($beforeIdResult.Output -join '').Trim()" & vbCrLf
    script = script & vbCrLf

    ' ajsentry実行（共通）-n: 即時実行, -w: 完了待ち
    script = script & "  # ajsentry実行（即時実行・完了待ち）" & vbCrLf
    script = script & "  Write-Log '[実行] ajsentry - ジョブ起動（-wオプションで完了待ち）'" & vbCrLf
    script = script & "  $entryResult = Invoke-JP1Command 'ajsentry.exe' @('-F', '" & config("SchedulerService") & "', '-n', '-w', '" & jobnetPath & "')" & vbCrLf
    script = script & "  $entryOutput = if ($entryResult.Output) { $entryResult.Output -join ' ' } else { '' }" & vbCrLf
    script = script & "  $entryExitCode = $entryResult.ExitCode" & vbCrLf
    script = script & "  if ($entryOutput) { Write-Log ""結果: $entryOutput"" } else { Write-Log ""結果: 正常終了 (ExitCode=$entryExitCode)"" }" & vbCrLf
    script = script & vbCrLf
    script = script & "  # ajsentryの実行結果をチェック" & vbCrLf
    script = script & "  if ($entryExitCode -ne 0 -or $entryOutput -match 'KAVS\d+-E') {" & vbCrLf
    script = script & "    $errMsg = if ($entryOutput) { $entryOutput } else { ""ExitCode=$entryExitCode"" }" & vbCrLf
    script = script & "    Write-Log ""[ERROR] ajsentryエラー: $errMsg""" & vbCrLf
    script = script & "    Write-Output ""RESULT_STATUS:実行エラー""" & vbCrLf
    script = script & "    Write-Output ""RESULT_MESSAGE:$errMsg""" & vbCrLf
    script = script & "    Write-Output ""RESULT_LOGPATH:$logFile""" & vbCrLf
    script = script & "    exit" & vbCrLf
    script = script & "  }" & vbCrLf
    script = script & vbCrLf
    script = script & "  # 実行IDと実行登録番号を取得（ajsentry後の最新世代）" & vbCrLf
    script = script & "  $execIdResult = Invoke-JP1Command 'ajsshow.exe' @('-F', '" & config("SchedulerService") & "', '-g', '1', '-i', '%## %ll', '" & jobnetPath & "')" & vbCrLf
    script = script & "  $execIdOutput = ($execIdResult.Output -join '').Trim()" & vbCrLf
    script = script & "  # 実行IDが空または不正な場合のチェック" & vbCrLf
    script = script & "  if (-not $execIdOutput -or $execIdOutput -match 'KAVS\d+-E') {" & vbCrLf
    script = script & "    Write-Log '[ERROR] 実行ID/実行登録番号の取得に失敗しました'" & vbCrLf
    script = script & "    Write-Output ""RESULT_STATUS:実行ID取得失敗""" & vbCrLf
    script = script & "    Write-Output ""RESULT_MESSAGE:$($execIdResult.Output -join ' ')""" & vbCrLf
    script = script & "    Write-Output ""RESULT_LOGPATH:$logFile""" & vbCrLf
    script = script & "    exit" & vbCrLf
    script = script & "  }" & vbCrLf
    script = script & "  $execIdParts = $execIdOutput -split '\s+'" & vbCrLf
    script = script & "  $execId = $execIdParts[0]" & vbCrLf
    script = script & "  $execRegNum = $execIdParts[1]" & vbCrLf
    script = script & "  Write-Log ""実行ID: $execId / 実行登録番号: $execRegNum""" & vbCrLf
    script = script & vbCrLf
    script = script & "  # 実行IDが変わったことを確認（今回の実行であることを保証）" & vbCrLf
    script = script & "  if ($execId -eq $beforeExecId) {" & vbCrLf
    script = script & "    Write-Log '[ERROR] 実行IDが変化していません。ジョブが実行されませんでした。'" & vbCrLf
    script = script & "    Write-Output ""RESULT_STATUS:実行ID未変化""" & vbCrLf
    script = script & "    Write-Output ""RESULT_MESSAGE:実行IDが変化していません（前回: $beforeExecId）""" & vbCrLf
    script = script & "    Write-Output ""RESULT_LOGPATH:$logFile""" & vbCrLf
    script = script & "    exit" & vbCrLf
    script = script & "  }" & vbCrLf
    script = script & vbCrLf

    If waitCompletion Then
        ' ajsentry -w終了後、ajsshowで1回だけ結果を取得
        ' （ajsentryの戻り値はコマンド実行成否であり、ジョブネット結果ではない）
        script = script & "  # ajsentry終了後、ajsshowで1回だけ結果を取得" & vbCrLf
        script = script & "  $statusResult = Invoke-JP1Command 'ajsshow.exe' @('-F', '" & config("SchedulerService") & "', '-B', $execRegNum, '-i', '%CC', '" & jobnetPath & "')" & vbCrLf
        script = script & "  $jobStatus = ($statusResult.Output -join ' ').Trim()" & vbCrLf
        script = script & "  Write-Log ""ジョブネット状態: $jobStatus""" & vbCrLf
        script = script & vbCrLf
        script = script & "  # ajsshowコマンド自体のエラーチェック" & vbCrLf
        script = script & "  if ($statusResult.ExitCode -ne 0 -or $jobStatus -match 'KAVS\d+-E') {" & vbCrLf
        script = script & "    Write-Log ""[ERROR] ajsshowコマンドエラー: $jobStatus""" & vbCrLf
        script = script & "    Write-Output ""RESULT_STATUS:コマンドエラー""" & vbCrLf
        script = script & "    Write-Output ""RESULT_MESSAGE:$jobStatus""" & vbCrLf
        script = script & "    Write-Output ""RESULT_LOGPATH:$logFile""" & vbCrLf
        script = script & "    exit" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & vbCrLf
        script = script & "  # 状態判定" & vbCrLf
        script = script & "  if ($jobStatus -match '正常終了') {" & vbCrLf
        script = script & "    Write-Log '[完了] 正常終了'" & vbCrLf
        script = script & "    Write-Output ""RESULT_STATUS:正常終了""" & vbCrLf
        script = script & "    Write-Output ""RESULT_LOGPATH:$logFile""" & vbCrLf
        script = script & "  } elseif ($jobStatus -match '警告検出終了|警告終了') {" & vbCrLf
        script = script & "    Write-Log '[完了] 警告検出終了'" & vbCrLf
        script = script & "    Write-Output ""RESULT_STATUS:警告検出終了""" & vbCrLf
        script = script & "  } elseif ($jobStatus -match '異常終了|異常検出終了|強制終了|中断') {" & vbCrLf
        script = script & "    Write-Log '[完了] 異常終了'" & vbCrLf
        script = script & "    Write-Output ""RESULT_STATUS:異常終了""" & vbCrLf
        script = script & "  } else {" & vbCrLf
        script = script & "    Write-Log ""[完了] 状態: $jobStatus""" & vbCrLf
        script = script & "    Write-Output ""RESULT_STATUS:$jobStatus""" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & vbCrLf

        ' 詳細情報取得（ajsshowで実行結果を確認）
        script = script & "  # 詳細情報取得" & vbCrLf
        script = script & "  $detailStatusResult = Invoke-JP1Command 'ajsshow.exe' @('-F', '" & config("SchedulerService") & "', '-B', $execRegNum, '-i', '%JJ %CC %SS %EE', '" & jobnetPath & "')" & vbCrLf
        script = script & "  $lastStatusStr = $detailStatusResult.Output -join ' '" & vbCrLf
        script = script & "  Write-Log ""詳細ステータス: $lastStatusStr""" & vbCrLf
        script = script & "  # 詳細取得エラーチェック" & vbCrLf
        script = script & "  if ($detailStatusResult.ExitCode -ne 0 -or $lastStatusStr -match 'KAVS\d+-E') {" & vbCrLf
        script = script & "    Write-Log ""[ERROR] 詳細情報取得エラー: $lastStatusStr""" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & vbCrLf

        ' 時間抽出（共通）
        script = script & "  # 時間抽出" & vbCrLf
        script = script & "  $timePattern = '\d{4}/\d{2}/\d{2}\s+\d{2}:\d{2}'" & vbCrLf
        script = script & "  $allTimes = [regex]::Matches($lastStatusStr, $timePattern)" & vbCrLf
        script = script & "  $startTimeStr = ''" & vbCrLf
        script = script & "  $endTimeStr = ''" & vbCrLf
        script = script & "  if ($allTimes.Count -ge 1) {" & vbCrLf
        script = script & "    $startTimeStr = $allTimes[0].Value" & vbCrLf
        script = script & "    Write-Output ""RESULT_START:$startTimeStr""" & vbCrLf
        script = script & "    Write-Log ""開始時刻: $startTimeStr""" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & "  if ($allTimes.Count -ge 2) {" & vbCrLf
        script = script & "    $endTimeStr = $allTimes[1].Value" & vbCrLf
        script = script & "    Write-Output ""RESULT_END:$endTimeStr""" & vbCrLf
        script = script & "    Write-Log ""終了時刻: $endTimeStr""" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & "  # 実行時間計算" & vbCrLf
        script = script & "  if ($startTimeStr -and $endTimeStr) {" & vbCrLf
        script = script & "    try {" & vbCrLf
        script = script & "      $startDt = [datetime]::ParseExact($startTimeStr, 'yyyy/MM/dd HH:mm', $null)" & vbCrLf
        script = script & "      $endDt = [datetime]::ParseExact($endTimeStr, 'yyyy/MM/dd HH:mm', $null)" & vbCrLf
        script = script & "      $duration = $endDt - $startDt" & vbCrLf
        script = script & "      $durationStr = '{0:D2}:{1:D2}:{2:D2}' -f [int]$duration.TotalHours, $duration.Minutes, $duration.Seconds" & vbCrLf
        script = script & "      Write-Log ""実行時間: $durationStr""" & vbCrLf
        script = script & "    } catch { }" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & "  $cleanMsg = $lastStatusStr -replace 'KAVS\d+-[IEW][^\r\n]*', '' -replace '\s+', ' '" & vbCrLf
        script = script & "  Write-Output ""RESULT_MESSAGE:$cleanMsg""" & vbCrLf
        script = script & vbCrLf

        ' ジョブネット内のジョブ一覧を表示（詳細ログ）
        script = script & "  # ジョブネット内のジョブ状態一覧を取得" & vbCrLf
        script = script & "  Write-Log ''" & vbCrLf
        script = script & "  Write-Log '【ジョブ実行結果一覧】'" & vbCrLf
        script = script & "  $jobListResult = Invoke-JP1Command 'ajsshow.exe' @('-F', '" & config("SchedulerService") & "', '-B', $execRegNum, '-R', '-f', '%JJ %TT %CC %RR', '" & jobnetPath & "')" & vbCrLf
        script = script & "  Write-Log ('  ' + '-' * 78)" & vbCrLf
        script = script & "  Write-Log ('  {0,-40} {1,-10} {2,-12} {3}' -f 'ジョブ名', 'タイプ', '状態', '戻り値')" & vbCrLf
        script = script & "  Write-Log ('  ' + '-' * 78)" & vbCrLf
        script = script & "  foreach ($jobLine in $jobListResult.Output) {" & vbCrLf
        script = script & "    if ($jobLine -match '^(/[^\s]+)\s+(\S+)\s+(\S+)\s*(.*)$') {" & vbCrLf
        script = script & "      $jName = $matches[1]" & vbCrLf
        script = script & "      $jType = $matches[2]" & vbCrLf
        script = script & "      $jStatus = $matches[3]" & vbCrLf
        script = script & "      $jReturn = $matches[4].Trim()" & vbCrLf
        script = script & "      # ジョブ名が長い場合は省略" & vbCrLf
        script = script & "      if ($jName.Length -gt 40) { $jName = '...' + $jName.Substring($jName.Length - 37) }" & vbCrLf
        script = script & "      $statusMark = switch -Regex ($jStatus) {" & vbCrLf
        script = script & "        '正常終了' { '[OK]' }" & vbCrLf
        script = script & "        '異常終了|異常検出' { '[NG]' }" & vbCrLf
        script = script & "        '警告' { '[!]' }" & vbCrLf
        script = script & "        '未実行|未起動' { '[-]' }" & vbCrLf
        script = script & "        default { '' }" & vbCrLf
        script = script & "      }" & vbCrLf
        script = script & "      Write-Log ('  {0,-40} {1,-10} {2} {3,-10} {4}' -f $jName, $jType, $statusMark, $jStatus, $jReturn)" & vbCrLf
        script = script & "    }" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & "  Write-Log ('  ' + '-' * 78)" & vbCrLf
        script = script & "  Write-Log ''" & vbCrLf
        script = script & vbCrLf

        ' エラー詳細取得（異常終了の場合）
        script = script & "  # エラー詳細取得" & vbCrLf
        script = script & "  if ($jobStatus -match '警告検出終了|警告終了|異常終了|異常検出終了|強制終了|中断') {" & vbCrLf
        script = script & "    Write-Log '[詳細取得] 異常終了したジョブを検索中...'" & vbCrLf
        script = script & "    $failedJobsResult = Invoke-JP1Command 'ajsshow.exe' @('-F', '" & config("SchedulerService") & "', '-R', '-f', '%J %T %C %R', '" & jobnetPath & "')" & vbCrLf
        script = script & "    $failedJobsStr = $failedJobsResult.Output -join ""`n""" & vbCrLf
        script = script & "    Write-Log ""[DEBUG] ajsshow -R -f 結果: $failedJobsStr""" & vbCrLf
        script = script & vbCrLf
        script = script & "    $failedJobPath = ''" & vbCrLf
        script = script & "    $nonZeroReturnJobPath = ''" & vbCrLf
        script = script & "    foreach ($line in $failedJobsResult.Output) {" & vbCrLf
        script = script & "      if ($line -match '^(/[^\s]+)\s+(\w*job|\w*jb)\s+(異常終了|警告終了|警告検出終了|Abnormal|Warning|ended abnormally|ended with warning)') {" & vbCrLf
        script = script & "        $failedJobPath = $matches[1]" & vbCrLf
        script = script & "        Write-Log ""[DEBUG] 異常終了ジョブ検出: $failedJobPath""" & vbCrLf
        script = script & "        break" & vbCrLf
        script = script & "      }" & vbCrLf
        script = script & "      if (-not $nonZeroReturnJobPath -and $line -match '^(/[^\s]+)\s+(\w*job|\w*jb)\s+\S+\s+([1-9]\d*|-\d+)') {" & vbCrLf
        script = script & "        $nonZeroReturnJobPath = $matches[1]" & vbCrLf
        script = script & "      }" & vbCrLf
        script = script & "    }" & vbCrLf
        script = script & "    if (-not $failedJobPath -and $nonZeroReturnJobPath) { $failedJobPath = $nonZeroReturnJobPath }" & vbCrLf
        script = script & vbCrLf
        script = script & "    if ($failedJobPath) {" & vbCrLf
        script = script & "      Write-Log ""[DEBUG] failedJobPath: $failedJobPath""" & vbCrLf
        script = script & "      $detailResult = Invoke-JP1Command 'ajsshow.exe' @('-F', '" & config("SchedulerService") & "', '-B', $execRegNum, '-i', '%## %ll %rr', $failedJobPath)" & vbCrLf
        script = script & "      $detailStr = $detailResult.Output -join ""`n""" & vbCrLf
        script = script & "      Write-Log ""[DEBUG] 詳細結果: $detailStr""" & vbCrLf
        script = script & vbCrLf
        script = script & "      $stderrFile = ''" & vbCrLf
        script = script & "      if ($detailStr -match '[A-Za-z]:[^\r\n]+\.err') { $stderrFile = $matches[0] }" & vbCrLf
        script = script & "      if ($stderrFile) {" & vbCrLf
        script = script & "        Write-Output ""RESULT_LOGPATH:$logFile""" & vbCrLf
        script = script & "        Write-Log ""[DEBUG] 標準エラーファイル: $stderrFile""" & vbCrLf
        script = script & "        $logContent = Read-FileContent $stderrFile" & vbCrLf
        script = script & "        if ($logContent) {" & vbCrLf
        script = script & "          Write-Log '[詳細] 標準エラーログ:'" & vbCrLf
        script = script & "          foreach ($line in $logContent) { Write-Log ""  $line"" }" & vbCrLf
        script = script & "        } else {" & vbCrLf
        script = script & "          Write-Log '標準エラーログを取得できませんでした'" & vbCrLf
        script = script & "        }" & vbCrLf
        script = script & "      }" & vbCrLf
        script = script & "    } else {" & vbCrLf
        script = script & "      Write-Log '異常終了したジョブが見つかりませんでした'" & vbCrLf
        script = script & "    }" & vbCrLf
        script = script & "  }" & vbCrLf
    Else
        script = script & "  Write-Log '[完了] 起動成功（完了待ちなし）'" & vbCrLf
        script = script & "  Write-Output ""RESULT_STATUS:起動成功""" & vbCrLf
        script = script & "  Write-Output ""RESULT_LOGPATH:$logFile""" & vbCrLf
        script = script & "  $msgOutput = if ($entryOutput) { $entryOutput } else { ""起動完了 (ExitCode=$entryExitCode)"" }" & vbCrLf
        script = script & "  Write-Output ""RESULT_MESSAGE:$msgOutput""" & vbCrLf
    End If

    ' リモートモードの場合: セッション終了
    If isRemote Then
        script = script & "  Write-Log '[クリーンアップ] セッション終了'" & vbCrLf
        script = script & "  Remove-PSSession $session" & vbCrLf
    End If

    script = script & "} catch {" & vbCrLf
    script = script & "  Write-Log ""[EXCEPTION] $($_.Exception.Message)""" & vbCrLf
    script = script & "  Write-Output ""ERROR: $($_.Exception.Message)""" & vbCrLf
    script = script & "}" & vbCrLf

    ' リモートモードの場合: WinRM設定復元
    If isRemote Then
        script = script & "finally {" & vbCrLf
        script = script & "  Write-Log '[クリーンアップ] WinRM設定を復元中...'" & vbCrLf
        script = script & "  if ($winrmConfigChanged) {" & vbCrLf
        script = script & "    if ($originalTrustedHosts) {" & vbCrLf
        script = script & "      Set-Item WSMan:\localhost\Client\TrustedHosts -Value $originalTrustedHosts -Force -ErrorAction SilentlyContinue" & vbCrLf
        script = script & "    } else {" & vbCrLf
        script = script & "      Clear-Item WSMan:\localhost\Client\TrustedHosts -Force -ErrorAction SilentlyContinue" & vbCrLf
        script = script & "    }" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & "  if ($winrmServiceWasStarted) {" & vbCrLf
        script = script & "    Stop-Service -Name WinRM -Force -ErrorAction SilentlyContinue" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & "  Write-Log '[完了] 処理終了'" & vbCrLf
        script = script & "}" & vbCrLf
    End If

    BuildExecuteJobScript = script
End Function

Private Sub UpdateJobListStatus(ByVal row As Long, ByVal result As Object)
    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_JOBLIST)

    ws.Cells(row, COL_LAST_STATUS).Value = result("Status")
    ws.Cells(row, COL_LAST_EXEC_TIME).Value = result("StartTime")
    ws.Cells(row, COL_LAST_END_TIME).Value = result("EndTime")

    ' ログパスを記録（N列にハイパーリンク設定）
    If result("LogPath") <> "" Then
        Dim logPath As String
        logPath = result("LogPath")
        ws.Cells(row, COL_LAST_MESSAGE).Value = logPath
        ' ファイルパスの場合はハイパーリンクを設定
        On Error Resume Next
        ws.Hyperlinks.Add Anchor:=ws.Cells(row, COL_LAST_MESSAGE), _
                          Address:=logPath, _
                          TextToDisplay:=logPath
        On Error GoTo 0
    End If

    ' 保留解除された場合（成功時）、保留列をクリアしてハイライトを解除
    If result("Status") = "正常終了" Or result("Status") = "起動成功" Then
        If ws.Cells(row, COL_HOLD).Value = "保留中" Then
            ws.Cells(row, COL_HOLD).Value = ""
            ws.Cells(row, COL_HOLD).Font.Bold = False
            ws.Cells(row, COL_HOLD).Font.Color = RGB(0, 0, 0)
            ' 行のハイライトを解除
            ws.Range(ws.Cells(row, COL_SELECT), ws.Cells(row, COL_LAST_MESSAGE)).Interior.ColorIndex = xlNone
        End If
    End If

    ' 色付け
    If result("Status") = "正常終了" Then
        ws.Cells(row, COL_LAST_STATUS).Interior.Color = RGB(198, 239, 206)  ' 緑（正常）
    ElseIf result("Status") = "起動成功" Then
        ws.Cells(row, COL_LAST_STATUS).Interior.Color = RGB(255, 235, 156)  ' 黄（起動のみ）
    ElseIf result("Status") = "警告検出終了" Then
        ws.Cells(row, COL_LAST_STATUS).Interior.Color = RGB(255, 192, 0)    ' オレンジ（警告）
    Else
        ws.Cells(row, COL_LAST_STATUS).Interior.Color = RGB(255, 199, 206)  ' 赤（異常）
    End If
End Sub

'==============================================================================
' 一覧クリア（実行結果のみクリア、ジョブ定義は保持）
'==============================================================================
Public Sub ClearJobList()
    If MsgBox("実行結果をクリアしますか？" & vbCrLf & _
              "（ジョブ定義情報は保持されます）", vbYesNo + vbQuestion) = vbNo Then Exit Sub

    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_JOBLIST)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_JOBNET_PATH).End(xlUp).Row

    If lastRow >= ROW_JOBLIST_DATA_START Then
        ' 選択列（☑/☐）と順序列をクリア
        ws.Range(ws.Cells(ROW_JOBLIST_DATA_START, COL_SELECT), ws.Cells(lastRow, COL_SELECT)).ClearContents
        ws.Range(ws.Cells(ROW_JOBLIST_DATA_START, COL_ORDER), ws.Cells(lastRow, COL_ORDER)).ClearContents

        ' 実行結果・ログパス列をクリア
        ws.Range(ws.Cells(ROW_JOBLIST_DATA_START, COL_LAST_STATUS), ws.Cells(lastRow, COL_LAST_MESSAGE)).ClearContents

        ' ハイパーリンクも削除
        On Error Resume Next
        ws.Range(ws.Cells(ROW_JOBLIST_DATA_START, COL_LAST_MESSAGE), ws.Cells(lastRow, COL_LAST_MESSAGE)).Hyperlinks.Delete
        On Error GoTo 0

        ' 背景色を初期状態に戻す
        Dim row As Long
        For row = ROW_JOBLIST_DATA_START To lastRow
            ' まず行全体の背景色をクリア
            ws.Range(ws.Cells(row, COL_SELECT), ws.Cells(row, COL_LAST_MESSAGE)).Interior.ColorIndex = xlNone
            ' 保留中の場合は保留列のみ黄色を適用
            If ws.Cells(row, COL_HOLD).Value = "保留中" Then
                ws.Cells(row, COL_HOLD).Interior.Color = RGB(255, 235, 156)
                ws.Cells(row, COL_HOLD).Font.Bold = True
                ws.Cells(row, COL_HOLD).Font.Color = RGB(156, 87, 0)
            End If
        Next row

        ' オートフィルタを再適用（種別「ジョブネット」）
        If ws.AutoFilterMode Then
            ws.AutoFilterMode = False
        End If
        ws.Range(ws.Cells(ROW_JOBLIST_HEADER, COL_SELECT), ws.Cells(lastRow, COL_LAST_MESSAGE)).AutoFilter _
            Field:=COL_UNIT_TYPE - COL_SELECT + 1, Criteria1:="ジョブネット"
    End If

    MsgBox "実行結果をクリアしました。", vbInformation
End Sub

'==============================================================================
' ジョブ一覧のダブルクリック時処理（シートモジュールから呼び出される）
' 選択列（COL_SELECT）をダブルクリックすると☑/☐を切り替え
'==============================================================================
Public Sub OnJobListDoubleClick(row As Long, col As Long, ByRef Cancel As Boolean)
    ' 選択列以外は通常の編集を許可
    If col <> COL_SELECT Then Exit Sub

    ' ヘッダー行以前は無視
    If row < ROW_JOBLIST_DATA_START Then Exit Sub

    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_JOBLIST)

    ' ジョブパスがない行は無視
    If ws.Cells(row, COL_JOBNET_PATH).Value = "" Then Exit Sub

    ' セル編集をキャンセル（重要：これがないとセル内編集モードになる）
    Cancel = True

    ' チェック状態を切り替え
    ToggleCheckMark row
End Sub

'==============================================================================
' チェックマーク（☑/☐）を切り替え
'==============================================================================
Private Sub ToggleCheckMark(row As Long)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_JOBLIST)

    Application.EnableEvents = False

    Dim currentValue As String
    currentValue = CStr(ws.Cells(row, COL_SELECT).Value)

    If currentValue = ChrW(&H2611) Then  ' ☑ → ☐
        ' チェックを外す
        ws.Cells(row, COL_SELECT).Value = ChrW(&H2610)  ' ☐
        ' 順序をクリアして再採番
        ws.Cells(row, COL_ORDER).Value = ""
        RenumberJobOrder

        ' 背景色を元に戻す
        ' まず行全体の背景色をクリア
        ws.Range(ws.Cells(row, COL_SELECT), ws.Cells(row, COL_LAST_MESSAGE)).Interior.ColorIndex = xlNone
        ' 保留中の場合は保留列のみ黄色を適用
        If ws.Cells(row, COL_HOLD).Value = "保留中" Then
            ws.Cells(row, COL_HOLD).Interior.Color = RGB(255, 235, 156)
        End If
    Else
        ' チェックを入れる（☐ または空白）
        ws.Cells(row, COL_SELECT).Value = ChrW(&H2611)  ' ☑
        ' 順序を自動採番
        Dim maxOrder As Long
        maxOrder = GetMaxOrderNumber()
        ws.Cells(row, COL_ORDER).Value = maxOrder + 1

        ' 行全体に水色の背景色を設定
        ws.Range(ws.Cells(row, COL_SELECT), ws.Cells(row, COL_LAST_MESSAGE)).Interior.Color = RGB(221, 235, 247)
        ' 保留中の場合は保留列のみ黄色を優先適用
        If ws.Cells(row, COL_HOLD).Value = "保留中" Then
            ws.Cells(row, COL_HOLD).Interior.Color = RGB(255, 235, 156)
        End If
    End If

    ' セルの書式を中央揃えに
    ws.Cells(row, COL_SELECT).HorizontalAlignment = xlCenter

    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    Application.EnableEvents = True
End Sub


'==============================================================================
' 現在の最大順序番号を取得
'==============================================================================
Private Function GetMaxOrderNumber() As Long
    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_JOBLIST)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_JOBNET_PATH).End(xlUp).Row

    Dim maxOrder As Long
    maxOrder = 0

    Dim row As Long
    For row = ROW_JOBLIST_DATA_START To lastRow
        Dim orderValue As Variant
        orderValue = ws.Cells(row, COL_ORDER).Value
        If IsNumeric(orderValue) And orderValue <> "" Then
            If CLng(orderValue) > maxOrder Then
                maxOrder = CLng(orderValue)
            End If
        End If
    Next row

    GetMaxOrderNumber = maxOrder
End Function

'==============================================================================
' 順序番号を再採番（チェックが外された時に呼び出し）
'==============================================================================
Private Sub RenumberJobOrder()
    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_JOBLIST)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_JOBNET_PATH).End(xlUp).Row

    ' 順序が入っている行を収集（順序番号順）
    Dim orderedRows As New Collection
    Dim row As Long
    For row = ROW_JOBLIST_DATA_START To lastRow
        Dim orderValue As Variant
        orderValue = ws.Cells(row, COL_ORDER).Value
        If IsNumeric(orderValue) And orderValue <> "" Then
            Dim item As Object
            Set item = CreateObject("Scripting.Dictionary")
            item("Row") = row
            item("Order") = CLng(orderValue)
            orderedRows.Add item
        End If
    Next row

    ' 順序番号が無い場合は終了
    If orderedRows.Count = 0 Then Exit Sub

    ' 順序でソート
    Dim arr() As Variant
    ReDim arr(1 To orderedRows.Count)
    Dim i As Long
    For i = 1 To orderedRows.Count
        Set arr(i) = orderedRows(i)
    Next i

    Dim j As Long
    Dim temp As Object
    For i = 1 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i)("Order") > arr(j)("Order") Then
                Set temp = arr(i)
                Set arr(i) = arr(j)
                Set arr(j) = temp
            End If
        Next j
    Next i

    ' 1から連番で再採番
    For i = 1 To UBound(arr)
        ws.Cells(arr(i)("Row"), COL_ORDER).Value = i
    Next i
End Sub

'==============================================================================
' 実行ログ履歴クリア
'==============================================================================
Public Sub ClearLogHistory()
    If MsgBox("実行ログの履歴をすべて削除しますか？", vbYesNo + vbQuestion) = vbNo Then Exit Sub

    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_LOG)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' データ行がある場合のみ削除（5行目以降がデータ）
    If lastRow >= 5 Then
        ' データ行を削除
        ws.Range(ws.Cells(5, 1), ws.Cells(lastRow, 6)).ClearContents

        ' ハイパーリンクも削除（F列）
        On Error Resume Next
        ws.Range(ws.Cells(5, 6), ws.Cells(lastRow, 6)).Hyperlinks.Delete
        On Error GoTo 0

        ' 背景色もクリア
        ws.Range(ws.Cells(5, 1), ws.Cells(lastRow, 6)).Interior.ColorIndex = xlNone

        ' 罫線もクリア
        ws.Range(ws.Cells(5, 1), ws.Cells(lastRow, 6)).Borders.LineStyle = xlNone
    End If

    MsgBox "実行ログの履歴を削除しました。", vbInformation
End Sub

'==============================================================================
' ユーティリティ
'==============================================================================
Private Function GetConfig() As Object
    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_SETTINGS)

    Dim config As Object
    Set config = CreateObject("Scripting.Dictionary")

    ' 実行モード（ローカル/リモート）
    config("ExecMode") = CStr(ws.Cells(ROW_EXEC_MODE, COL_SETTING_VALUE).Value)

    config("JP1Server") = CStr(ws.Cells(ROW_JP1_SERVER, COL_SETTING_VALUE).Value)
    config("RemoteUser") = CStr(ws.Cells(ROW_REMOTE_USER, COL_SETTING_VALUE).Value)
    config("RemotePassword") = CStr(ws.Cells(ROW_REMOTE_PASSWORD, COL_SETTING_VALUE).Value)
    config("JP1User") = CStr(ws.Cells(ROW_JP1_USER, COL_SETTING_VALUE).Value)
    config("JP1Password") = CStr(ws.Cells(ROW_JP1_PASSWORD, COL_SETTING_VALUE).Value)
    config("SchedulerService") = CStr(ws.Cells(ROW_SCHEDULER_SERVICE, COL_SETTING_VALUE).Value)
    config("RootPath") = CStr(ws.Cells(ROW_ROOT_PATH, COL_SETTING_VALUE).Value)
    config("WaitCompletion") = CStr(ws.Cells(ROW_WAIT_COMPLETION, COL_SETTING_VALUE).Value)
    config("Timeout") = CLng(ws.Cells(ROW_TIMEOUT, COL_SETTING_VALUE).Value)
    config("PollingInterval") = CLng(ws.Cells(ROW_POLLING_INTERVAL, COL_SETTING_VALUE).Value)

    ' 必須項目チェック（ローカルモードとリモートモードで異なる）
    If config("ExecMode") = "ローカル" Then
        ' ローカルモード: JP1ユーザーのみ必須
        If config("JP1User") = "" Then
            MsgBox "JP1ユーザーを入力してください。", vbExclamation
            Set GetConfig = Nothing
            Exit Function
        End If
    Else
        ' リモートモード: 接続情報が必須
        If config("JP1Server") = "" Or config("RemoteUser") = "" Or config("JP1User") = "" Then
            MsgBox "接続設定が不完全です。設定シートで設定を入力してください。", vbExclamation
            Set GetConfig = Nothing
            Exit Function
        End If
    End If

    Set GetConfig = config
End Function

Private Function ExecutePowerShell(script As String) As String
    ' 一時ファイルにスクリプトを保存（UTF-8 BOMなしで保存）
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim tempFolder As String
    tempFolder = fso.GetSpecialFolder(2) ' Temp folder

    Dim timestamp As String
    timestamp = Format(Now, "yyyymmddhhnnss") & "_" & Int(Rnd * 10000)

    Dim scriptPath As String
    scriptPath = tempFolder & "\jp1_temp_" & timestamp & ".ps1"

    Dim outputPath As String
    outputPath = tempFolder & "\jp1_output_" & timestamp & ".txt"

    ' スクリプトをラップして結果をファイルに出力
    Dim wrappedScript As String
    wrappedScript = script & vbCrLf
    wrappedScript = wrappedScript & "# 出力完了マーカー" & vbCrLf

    ' ADODB.Streamを使用してUTF-8（BOM付き）で保存
    ' PowerShellはBOM付きUTF-8を自動認識するため、日本語パスが正しく処理される
    Dim utfStream As Object
    Set utfStream = CreateObject("ADODB.Stream")
    utfStream.Type = 2 ' adTypeText
    utfStream.Charset = "UTF-8"
    utfStream.Open
    utfStream.WriteText wrappedScript
    utfStream.SaveToFile scriptPath, 2 ' adSaveCreateOverWrite（BOM付きで保存）
    utfStream.Close
    Set utfStream = Nothing

    ' PowerShell実行（リアルタイム表示・結果をファイルに出力）
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")

    Dim cmd As String
    ' PowerShellウィンドウを直接表示して実行
    ' Start-Transcriptでログを取りながらリアルタイム表示
    cmd = "powershell -NoProfile -ExecutionPolicy Bypass -Command ""& {" & _
          "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; " & _
          "Start-Transcript -Path '" & outputPath & "' -Force | Out-Null; " & _
          "try { & '" & scriptPath & "' } finally { Stop-Transcript | Out-Null }" & _
          "}"""

    ' 1 = vbNormalFocus（通常表示）、True で完了まで待機
    shell.Run cmd, 1, True

    ' 結果ファイルを読み込む
    Dim output As String
    output = ""

    If fso.FileExists(outputPath) Then
        ' UTF-8で読み込み
        Set utfStream = CreateObject("ADODB.Stream")
        utfStream.Type = 2 ' adTypeText
        utfStream.Charset = "UTF-8"
        utfStream.Open
        utfStream.LoadFromFile outputPath

        If Not utfStream.EOS Then
            output = utfStream.ReadText
        End If

        utfStream.Close
        Set utfStream = Nothing

        ' 出力ファイル削除
        On Error Resume Next
        fso.DeleteFile outputPath
        On Error GoTo 0
    End If

    ' スクリプトファイル削除
    On Error Resume Next
    fso.DeleteFile scriptPath
    On Error GoTo 0

    ExecutePowerShell = output
End Function

Private Function EscapePSString(str As String) As String
    ' PowerShell文字列内のシングルクォートをエスケープ
    EscapePSString = Replace(str, "'", "''")
End Function

'==============================================================================
' 管理者権限チェック
'==============================================================================
Private Function IsRunningAsAdmin() As Boolean
    ' キャッシュを利用
    If g_AdminChecked Then
        IsRunningAsAdmin = g_IsAdmin
        Exit Function
    End If

    ' PowerShellで管理者権限をチェック
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")

    Dim cmd As String
    cmd = "powershell -NoProfile -Command ""$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent()); if ($principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { exit 0 } else { exit 1 }"""

    Dim exitCode As Long
    exitCode = shell.Run(cmd, 0, True)

    g_IsAdmin = (exitCode = 0)
    g_AdminChecked = True

    IsRunningAsAdmin = g_IsAdmin
End Function

Private Function EnsureAdminForRemoteMode(config As Object) As Boolean
    ' ローカルモードなら管理者不要
    If config("ExecMode") = "ローカル" Then
        EnsureAdminForRemoteMode = True
        Exit Function
    End If

    ' 既に管理者なら問題なし
    If IsRunningAsAdmin() Then
        EnsureAdminForRemoteMode = True
        Exit Function
    End If

    ' 管理者でない場合、ユーザーに選択させる
    Dim response As VbMsgBoxResult
    response = MsgBox( _
        "リモート実行モードでは、WinRM設定の変更に管理者権限が必要です。" & vbCrLf & vbCrLf & _
        "現在、管理者権限で実行されていません。" & vbCrLf & vbCrLf & _
        "[はい] 管理者としてExcelを再起動して実行" & vbCrLf & _
        "[いいえ] このまま続行（WinRMが既に設定済みの場合）" & vbCrLf & _
        "[キャンセル] 処理を中止", _
        vbYesNoCancel + vbExclamation, "管理者権限が必要")

    Select Case response
        Case vbYes
            ' 管理者権限でExcelを再起動
            RestartAsAdmin
            EnsureAdminForRemoteMode = False

        Case vbNo
            ' そのまま続行
            EnsureAdminForRemoteMode = True

        Case vbCancel
            ' 処理を中止
            EnsureAdminForRemoteMode = False
    End Select
End Function

Private Sub RestartAsAdmin()
    ' 現在のブックを保存
    If ThisWorkbook.Saved = False Then
        Dim saveResponse As VbMsgBoxResult
        saveResponse = MsgBox("ブックを保存しますか？", vbYesNoCancel + vbQuestion, "保存確認")
        If saveResponse = vbYes Then
            ThisWorkbook.Save
        ElseIf saveResponse = vbCancel Then
            Exit Sub
        End If
    End If

    ' 管理者権限でExcelを再起動
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")

    Dim excelPath As String
    excelPath = Application.Path & "\EXCEL.EXE"

    Dim workbookPath As String
    workbookPath = ThisWorkbook.FullName

    ' PowerShellでStart-Process -Verb RunAsを実行
    Dim cmd As String
    cmd = "powershell -NoProfile -Command ""Start-Process -FilePath '" & Replace(excelPath, "'", "''") & "' -ArgumentList '""" & Replace(workbookPath, "'", "''") & """' -Verb RunAs"""

    shell.Run cmd, 0, False

    ' このブックのみを閉じる（他のExcelブックは維持）
    ThisWorkbook.Close SaveChanges:=False
End Sub

'==============================================================================
' ログファイル関連
'==============================================================================
Private Function CreateLogFile() As String
    ' ログファイルのパスを生成して初期ヘッダーを書き込む
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' ログフォルダ（Excelブックと同じフォルダにLogsサブフォルダ）
    Dim logFolder As String
    logFolder = ThisWorkbook.Path & "\Logs"

    ' Logsフォルダが存在しない場合は作成
    If Not fso.FolderExists(logFolder) Then
        fso.CreateFolder logFolder
    End If

    ' ログファイル名（JP1_実行ログ_yyyyMMdd_HHmmss.txt）
    Dim logFileName As String
    logFileName = "JP1_実行ログ_" & Format(Now, "yyyyMMdd_HHmmss") & ".txt"

    Dim logFilePath As String
    logFilePath = logFolder & "\" & logFileName

    ' ADODB.Streamを使用してUTF-8（BOMなし）でヘッダーを書き込む
    Dim logContent As String
    logContent = "================================================================================" & vbCrLf
    logContent = logContent & "JP1 ジョブ管理ツール - 実行ログ" & vbCrLf
    logContent = logContent & "================================================================================" & vbCrLf
    logContent = logContent & "開始日時: " & Format(Now, "yyyy/mm/dd HH:mm:ss") & vbCrLf
    logContent = logContent & "実行モード: " & Worksheets(SHEET_SETTINGS).Cells(ROW_EXEC_MODE, COL_SETTING_VALUE).Value & vbCrLf
    logContent = logContent & "================================================================================" & vbCrLf
    logContent = logContent & "" & vbCrLf

    Dim utfStream As Object
    Set utfStream = CreateObject("ADODB.Stream")
    utfStream.Type = 2 ' adTypeText
    utfStream.Charset = "UTF-8"
    utfStream.Open
    utfStream.WriteText logContent

    ' BOMをスキップしてバイナリで保存
    utfStream.Position = 0
    utfStream.Type = 1 ' adTypeBinary
    utfStream.Position = 3 ' BOM（3バイト）をスキップ

    Dim binStream As Object
    Set binStream = CreateObject("ADODB.Stream")
    binStream.Type = 1 ' adTypeBinary
    binStream.Open
    utfStream.CopyTo binStream
    binStream.SaveToFile logFilePath, 2 ' adSaveCreateOverWrite

    binStream.Close
    utfStream.Close
    Set binStream = Nothing
    Set utfStream = Nothing

    CreateLogFile = logFilePath
End Function

Private Function GetLogFilePath() As String
    ' 現在のログファイルパスを返す
    GetLogFilePath = g_LogFilePath
End Function
