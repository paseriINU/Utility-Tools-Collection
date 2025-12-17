Option Explicit

'==============================================================================
' JP1 ジョブ管理ツール（バッチ版） - メインモジュール
'   - 外部バッチファイル（JP1_ジョブ実行.bat）を呼び出す方式
'   - 接続設定・パスワードはバッチファイルの設定セクションに記載
'   - 注意: 定数はSetupモジュールで Public として定義されています
'==============================================================================

' 管理者権限状態を保持
Private g_AdminChecked As Boolean
Private g_IsAdmin As Boolean

' 現在の実行セッションのログファイルパス
Private g_LogFilePath As String

'==============================================================================
' ジョブ一覧取得
'==============================================================================
Public Sub GetJobList()
    Dim config As Object
    Set config = GetConfig()

    If config Is Nothing Then Exit Sub

    ' リモートモードの場合、管理者権限を確認
    If config("ExecMode") = "リモート" Then
        If Not CheckAndRequestAdminPrivileges() Then Exit Sub
    End If

    Application.StatusBar = "ジョブ一覧を取得中..."
    Application.ScreenUpdating = False

    Dim psScript As String
    psScript = BuildGetJobListScript(config)

    Dim result As String
    result = ExecutePowerShell(psScript)

    ' 結果をパース
    ParseJobListResult result

    Application.StatusBar = False
    Application.ScreenUpdating = True

    Worksheets(SHEET_JOBLIST).Activate

    MsgBox "ジョブ一覧を取得しました。", vbInformation
End Sub

'==============================================================================
' ジョブ一覧クリア
'==============================================================================
Public Sub ClearJobList()
    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_JOBLIST)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_JOBNET_PATH).End(xlUp).Row

    If lastRow >= ROW_JOBLIST_DATA_START Then
        ws.Range(ws.Cells(ROW_JOBLIST_DATA_START, COL_ORDER), ws.Cells(lastRow, COL_LAST_MESSAGE)).Clear
    End If

    MsgBox "ジョブ一覧をクリアしました。", vbInformation
End Sub

'==============================================================================
' 設定取得（メインシートから）
'==============================================================================
Private Function GetConfig() As Object
    Dim config As Object
    Set config = CreateObject("Scripting.Dictionary")

    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_MAIN)

    config("ExecMode") = ws.Cells(ROW_EXEC_MODE, COL_SETTING_VALUE).Value
    config("JP1Server") = ws.Cells(ROW_JP1_SERVER, COL_SETTING_VALUE).Value
    config("RemoteUser") = ws.Cells(ROW_REMOTE_USER, COL_SETTING_VALUE).Value
    config("RemotePassword") = ws.Cells(ROW_REMOTE_PASSWORD, COL_SETTING_VALUE).Value
    config("JP1User") = ws.Cells(ROW_JP1_USER, COL_SETTING_VALUE).Value
    config("JP1Password") = ws.Cells(ROW_JP1_PASSWORD, COL_SETTING_VALUE).Value
    config("RootPath") = ws.Cells(ROW_ROOT_PATH, COL_SETTING_VALUE).Value
    config("WaitCompletion") = ws.Cells(ROW_WAIT_COMPLETION + 4, COL_SETTING_VALUE).Value
    config("Timeout") = ws.Cells(ROW_TIMEOUT + 4, COL_SETTING_VALUE).Value
    config("PollingInterval") = ws.Cells(ROW_POLLING_INTERVAL + 4, COL_SETTING_VALUE).Value

    ' パスワード入力（空の場合）
    If config("JP1Password") = "" Then
        config("JP1Password") = InputBox("JP1パスワードを入力してください:", "パスワード入力")
        If config("JP1Password") = "" Then
            MsgBox "JP1パスワードが入力されていません。", vbExclamation
            Set GetConfig = Nothing
            Exit Function
        End If
    End If

    If config("ExecMode") = "リモート" Then
        If config("RemotePassword") = "" Then
            config("RemotePassword") = InputBox("リモートパスワードを入力してください:", "パスワード入力")
            If config("RemotePassword") = "" Then
                MsgBox "リモートパスワードが入力されていません。", vbExclamation
                Set GetConfig = Nothing
                Exit Function
            End If
        End If

        ' 必須項目チェック
        If config("JP1Server") = "" Or config("RemoteUser") = "" Then
            MsgBox "接続設定が不完全です。JP1サーバとリモートユーザーを入力してください。", vbExclamation
            Set GetConfig = Nothing
            Exit Function
        End If
    End If

    Set GetConfig = config
End Function

'==============================================================================
' ジョブ一覧取得スクリプト構築
'==============================================================================
Private Function BuildGetJobListScript(config As Object) As String
    Dim script As String

    script = "$ErrorActionPreference = 'Stop'" & vbCrLf
    script = script & "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8" & vbCrLf
    script = script & vbCrLf

    If config("ExecMode") = "ローカル" Then
        ' ローカル実行
        script = script & "$jp1BinPath = 'C:\Program Files\HITACHI\JP1AJS3\bin'" & vbCrLf
        script = script & "if (-not (Test-Path ""$jp1BinPath\ajsprint.exe"")) {" & vbCrLf
        script = script & "  $jp1BinPath = 'C:\Program Files\Hitachi\JP1AJS2\bin'" & vbCrLf
        script = script & "}" & vbCrLf
        script = script & "if (-not (Test-Path ""$jp1BinPath\ajsprint.exe"")) {" & vbCrLf
        script = script & "  Write-Error 'JP1コマンドが見つかりません'" & vbCrLf
        script = script & "  exit 1" & vbCrLf
        script = script & "}" & vbCrLf
        script = script & vbCrLf
        script = script & "$output = & ""$jp1BinPath\ajsprint.exe"" -h localhost -u '" & config("JP1User") & "' -p '" & EscapePSString(config("JP1Password")) & "' -a -R '" & config("RootPath") & "' 2>&1" & vbCrLf
        script = script & "Write-Output $output" & vbCrLf
    Else
        ' リモート実行
        script = script & "$securePass = ConvertTo-SecureString '" & EscapePSString(config("RemotePassword")) & "' -AsPlainText -Force" & vbCrLf
        script = script & "$cred = New-Object System.Management.Automation.PSCredential('" & config("RemoteUser") & "', $securePass)" & vbCrLf
        script = script & vbCrLf

        ' WinRM設定
        script = script & "$originalTrustedHosts = (Get-Item WSMan:\localhost\Client\TrustedHosts -ErrorAction SilentlyContinue).Value" & vbCrLf
        script = script & "$winrmConfigChanged = $false" & vbCrLf
        script = script & "$winrmServiceWasStarted = $false" & vbCrLf
        script = script & vbCrLf

        script = script & "try {" & vbCrLf
        script = script & "  $winrmService = Get-Service -Name WinRM -ErrorAction SilentlyContinue" & vbCrLf
        script = script & "  if ($winrmService.Status -ne 'Running') {" & vbCrLf
        script = script & "    Start-Service -Name WinRM -ErrorAction Stop" & vbCrLf
        script = script & "    $winrmServiceWasStarted = $true" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & vbCrLf
        script = script & "  if ($originalTrustedHosts -notmatch '" & config("JP1Server") & "') {" & vbCrLf
        script = script & "    if ($originalTrustedHosts) {" & vbCrLf
        script = script & "      Set-Item WSMan:\localhost\Client\TrustedHosts -Value ""$originalTrustedHosts," & config("JP1Server") & """ -Force -Confirm:`$false" & vbCrLf
        script = script & "    } else {" & vbCrLf
        script = script & "      Set-Item WSMan:\localhost\Client\TrustedHosts -Value '" & config("JP1Server") & "' -Force -Confirm:`$false" & vbCrLf
        script = script & "    }" & vbCrLf
        script = script & "    $winrmConfigChanged = $true" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & vbCrLf
        script = script & "  $session = New-PSSession -ComputerName '" & config("JP1Server") & "' -Credential $cred -ErrorAction Stop" & vbCrLf
        script = script & vbCrLf
        script = script & "  $result = Invoke-Command -Session $session -ScriptBlock {" & vbCrLf
        script = script & "    param($jp1User, $jp1Pass, $rootPath)" & vbCrLf
        script = script & "    $jp1BinPath = 'C:\Program Files\HITACHI\JP1AJS3\bin'" & vbCrLf
        script = script & "    if (-not (Test-Path ""$jp1BinPath\ajsprint.exe"")) { $jp1BinPath = 'C:\Program Files\Hitachi\JP1AJS2\bin' }" & vbCrLf
        script = script & "    & ""$jp1BinPath\ajsprint.exe"" -h localhost -u $jp1User -p $jp1Pass -a -R $rootPath 2>&1" & vbCrLf
        script = script & "  } -ArgumentList '" & config("JP1User") & "', '" & EscapePSString(config("JP1Password")) & "', '" & config("RootPath") & "'" & vbCrLf
        script = script & vbCrLf
        script = script & "  Write-Output $result" & vbCrLf
        script = script & "  Remove-PSSession $session" & vbCrLf
        script = script & "} finally {" & vbCrLf
        script = script & "  if ($winrmConfigChanged) {" & vbCrLf
        script = script & "    if ($originalTrustedHosts) {" & vbCrLf
        script = script & "      Set-Item WSMan:\localhost\Client\TrustedHosts -Value $originalTrustedHosts -Force -Confirm:`$false -ErrorAction SilentlyContinue" & vbCrLf
        script = script & "    } else {" & vbCrLf
        script = script & "      Clear-Item WSMan:\localhost\Client\TrustedHosts -Force -Confirm:`$false -ErrorAction SilentlyContinue" & vbCrLf
        script = script & "    }" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & "  if ($winrmServiceWasStarted) {" & vbCrLf
        script = script & "    Stop-Service -Name WinRM -Force -ErrorAction SilentlyContinue" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & "}" & vbCrLf
    End If

    BuildGetJobListScript = script
End Function

'==============================================================================
' ジョブ一覧パース
'==============================================================================
Private Sub ParseJobListResult(result As String)
    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_JOBLIST)

    ' 既存データをクリア
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_JOBNET_PATH).End(xlUp).Row
    If lastRow >= ROW_JOBLIST_DATA_START Then
        ws.Range(ws.Cells(ROW_JOBLIST_DATA_START, COL_ORDER), ws.Cells(lastRow, COL_LAST_MESSAGE)).Clear
    End If

    ' 結果をパース
    Dim lines() As String
    lines = Split(result, vbCrLf)

    Dim row As Long
    row = ROW_JOBLIST_DATA_START

    Dim line As Variant
    For Each line In lines
        If Len(line) > 0 Then
            ' un= で始まる行がジョブ定義
            Dim unitMatch As String
            unitMatch = ExtractUnitPath(CStr(line))

            If unitMatch <> "" And InStr(line, ",ty=n") > 0 Then
                ' ジョブネット（ty=n）のみ追加
                ws.Cells(row, COL_ORDER).Value = ""
                ws.Cells(row, COL_JOBNET_PATH).Value = unitMatch
                ws.Cells(row, COL_JOBNET_NAME).Value = ExtractJobName(CStr(line))
                ws.Cells(row, COL_COMMENT).Value = ExtractComment(CStr(line))

                ' 保留状態を解析
                Dim isHold As Boolean
                isHold = ExtractHoldStatus(CStr(line))

                If isHold Then
                    ws.Cells(row, COL_HOLD).Value = "保留中"
                    ws.Cells(row, COL_HOLD).HorizontalAlignment = xlCenter

                    ' 保留中のジョブは行全体をハイライト（オレンジ系）
                    ws.Range(ws.Cells(row, COL_ORDER), ws.Cells(row, COL_LAST_MESSAGE)).Interior.Color = RGB(255, 235, 156)
                    ws.Cells(row, COL_HOLD).Font.Bold = True
                    ws.Cells(row, COL_HOLD).Font.Color = RGB(156, 87, 0)
                Else
                    ws.Cells(row, COL_HOLD).Value = ""
                End If

                ' 順序列の書式
                With ws.Cells(row, COL_ORDER)
                    .HorizontalAlignment = xlCenter
                End With

                ' 罫線
                ws.Range(ws.Cells(row, COL_ORDER), ws.Cells(row, COL_LAST_MESSAGE)).Borders.LineStyle = xlContinuous

                row = row + 1
            End If
        End If
    Next line
End Sub

'==============================================================================
' ユニットパス抽出
'==============================================================================
Private Function ExtractUnitPath(line As String) As String
    Dim startPos As Long, endPos As Long
    startPos = InStr(line, "un=""")
    If startPos > 0 Then
        startPos = startPos + 4
        endPos = InStr(startPos, line, """")
        If endPos > startPos Then
            ExtractUnitPath = Mid(line, startPos, endPos - startPos)
        End If
    End If
End Function

Private Function ExtractJobName(line As String) As String
    Dim unitPath As String
    unitPath = ExtractUnitPath(line)
    If unitPath <> "" Then
        ExtractJobName = GetLastPathComponent(unitPath)
    End If
End Function

Private Function ExtractComment(line As String) As String
    Dim startPos As Long, endPos As Long
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
    ' バッチファイルの存在確認
    Dim batPath As String
    batPath = ThisWorkbook.path & "\JP1_ジョブ実行.bat"
    If Dir(batPath) = "" Then
        MsgBox "バッチファイルが見つかりません。" & vbCrLf & vbCrLf & _
               "以下のファイルをExcelブックと同じフォルダに配置してください：" & vbCrLf & _
               "JP1_ジョブ実行.bat" & vbCrLf & vbCrLf & _
               "※接続設定はバッチファイルの設定セクションを編集してください。", vbCritical
        Exit Sub
    End If

    ' 順序が入力されているジョブを取得
    Dim jobs As Collection
    Set jobs = GetOrderedJobs()

    If jobs.Count = 0 Then
        MsgBox "実行するジョブが選択されていません。" & vbCrLf & _
               "ジョブ一覧シートの「順序」列に数字（1, 2, 3...）を入力してください。", vbExclamation
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

    msg = msg & vbCrLf & "※ 接続設定はバッチファイル（JP1_ジョブ実行.bat）を確認してください。" & vbCrLf
    msg = msg & vbCrLf & "実行しますか？"

    If MsgBox(msg, vbYesNo + vbQuestion, "実行確認") = vbNo Then Exit Sub

    ' ログファイルの初期化
    g_LogFilePath = CreateLogFile()

    ' 実行
    Application.ScreenUpdating = False

    Dim wsLog As Worksheet
    Set wsLog = Worksheets(SHEET_LOG)
    Dim logRow As Long
    logRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).row + 1
    If logRow < 4 Then logRow = 4

    Dim success As Boolean
    success = True

    For Each j In jobs
        Application.StatusBar = "実行中: " & j("Path")

        Dim execResult As Object
        Set execResult = ExecuteSingleJob(j("Path"), j("IsHold"), g_LogFilePath)

        ' 結果をログに記録
        wsLog.Cells(logRow, 1).Value = Now
        wsLog.Cells(logRow, 2).Value = j("Path")
        wsLog.Cells(logRow, 3).Value = execResult("Status")
        wsLog.Cells(logRow, 4).Value = execResult("StartTime")
        wsLog.Cells(logRow, 5).Value = execResult("EndTime")
        wsLog.Cells(logRow, 6).Value = execResult("Message")

        ' ジョブ一覧シートも更新
        UpdateJobListStatus j("Row"), execResult

        logRow = logRow + 1

        ' エラー時は停止
        If execResult("Status") <> "正常終了" And execResult("Status") <> "起動成功" Then
            success = False
            MsgBox "ジョブ「" & j("Path") & "」が失敗しました。" & vbCrLf & _
                   "処理を中断します。" & vbCrLf & vbCrLf & _
                   "詳細: " & execResult("Message") & vbCrLf & vbCrLf & _
                   "実行ログ: " & g_LogFilePath, vbCritical
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
End Sub

Private Function GetOrderedJobs() As Collection
    Dim jobs As New Collection
    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_JOBLIST)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_JOBNET_PATH).End(xlUp).row

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
            job("IsHold") = (ws.Cells(row, COL_HOLD).Value = "保留中")

            orderedRows.Add job
        End If
    Next row

    ' 実行順でソート（単純なバブルソート）
    If orderedRows.Count = 0 Then
        Set GetOrderedJobs = jobs
        Exit Function
    End If

    Dim arr() As Object
    ReDim arr(1 To orderedRows.Count)
    Dim i As Long
    For i = 1 To orderedRows.Count
        Set arr(i) = orderedRows(i)
    Next i

    Dim swapped As Boolean
    Dim temp As Object
    Do
        swapped = False
        For i = 1 To UBound(arr) - 1
            If arr(i)("Order") > arr(i + 1)("Order") Then
                Set temp = arr(i)
                Set arr(i) = arr(i + 1)
                Set arr(i + 1) = temp
                swapped = True
            End If
        Next i
    Loop While swapped

    For i = 1 To UBound(arr)
        jobs.Add arr(i)
    Next i

    Set GetOrderedJobs = jobs
End Function

'==============================================================================
' 単一ジョブ実行（バッチファイル呼び出し版）
'==============================================================================
Private Function ExecuteSingleJob(jobnetPath As String, isHold As Boolean, logFilePath As String) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result("Status") = ""
    result("StartTime") = ""
    result("EndTime") = ""
    result("Message") = ""

    ' バッチファイルのパス
    Dim batPath As String
    batPath = ThisWorkbook.path & "\JP1_ジョブ実行.bat"

    ' 引数を構築（設定はバッチファイル内で定義）
    Dim args As String
    args = """" & jobnetPath & """"
    If isHold Then
        args = args & " -IsHold true"
    Else
        args = args & " -IsHold false"
    End If
    args = args & " -LogFile """ & logFilePath & """"

    ' バッチファイルを実行して出力を取得
    Dim output As String
    output = ExecuteBatchFile(batPath, args)

    ' 結果をパース
    Dim lines() As String
    lines = Split(output, vbCrLf)

    Dim line As Variant
    For Each line In lines
        If InStr(line, "RESULT_STATUS:") > 0 Then
            result("Status") = Mid(line, InStr(line, "RESULT_STATUS:") + 14)
        ElseIf InStr(line, "RESULT_MESSAGE:") > 0 Then
            result("Message") = Mid(line, InStr(line, "RESULT_MESSAGE:") + 15)
        ElseIf InStr(line, "RESULT_START:") > 0 Then
            result("StartTime") = Mid(line, InStr(line, "RESULT_START:") + 13)
        ElseIf InStr(line, "RESULT_END:") > 0 Then
            result("EndTime") = Mid(line, InStr(line, "RESULT_END:") + 11)
        End If
    Next line

    If result("Status") = "" Then
        result("Status") = "不明"
        result("Message") = output
    End If

    Set ExecuteSingleJob = result
End Function

'==============================================================================
' バッチファイル実行
'==============================================================================
Private Function ExecuteBatchFile(batPath As String, args As String) As String
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")

    Dim tempFile As String
    tempFile = Environ("TEMP") & "\jp1_batch_output_" & Format(Now, "yyyyMMddHHmmss") & ".txt"

    ' バッチファイルを実行して出力をファイルにリダイレクト
    Dim cmd As String
    cmd = """" & batPath & """ " & args & " > """ & tempFile & """ 2>&1"

    Dim exitCode As Long
    exitCode = shell.Run("cmd /c " & cmd, 0, True)

    ' 出力を読み込み
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim result As String
    result = ""

    If fso.FileExists(tempFile) Then
        Dim ts As Object
        Set ts = fso.OpenTextFile(tempFile, 1, False, -1) ' Unicode
        If Not ts.AtEndOfStream Then
            result = ts.ReadAll
        End If
        ts.Close
        fso.DeleteFile tempFile
    End If

    ExecuteBatchFile = result
End Function

'==============================================================================
' ジョブ一覧ステータス更新
'==============================================================================
Private Sub UpdateJobListStatus(row As Long, result As Object)
    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_JOBLIST)

    ws.Cells(row, COL_LAST_STATUS).Value = result("Status")
    ws.Cells(row, COL_LAST_EXEC_TIME).Value = result("StartTime")
    ws.Cells(row, COL_LAST_END_TIME).Value = result("EndTime")

    If result("Message") <> "" Then
        ws.Cells(row, COL_LAST_MESSAGE).Value = result("Message")
    End If

    ' 保留解除された場合（成功時）、保留列をクリアしてハイライトを解除
    If result("Status") = "正常終了" Or result("Status") = "起動成功" Then
        If ws.Cells(row, COL_HOLD).Value = "保留中" Then
            ws.Cells(row, COL_HOLD).Value = ""
            ws.Cells(row, COL_HOLD).Font.Bold = False
            ws.Cells(row, COL_HOLD).Font.Color = RGB(0, 0, 0)
            ws.Range(ws.Cells(row, COL_ORDER), ws.Cells(row, COL_LAST_MESSAGE)).Interior.ColorIndex = xlNone
        End If
    End If

    ' 色付け
    If result("Status") = "正常終了" Then
        ws.Cells(row, COL_LAST_STATUS).Interior.Color = RGB(198, 239, 206)
    ElseIf result("Status") = "起動成功" Then
        ws.Cells(row, COL_LAST_STATUS).Interior.Color = RGB(255, 235, 156)
    Else
        ws.Cells(row, COL_LAST_STATUS).Interior.Color = RGB(255, 199, 206)
    End If
End Sub

'==============================================================================
' PowerShell実行
'==============================================================================
Private Function ExecutePowerShell(script As String) As String
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")

    Dim tempScript As String, tempOutput As String
    tempScript = Environ("TEMP") & "\jp1_script_" & Format(Now, "yyyyMMddHHmmss") & ".ps1"
    tempOutput = Environ("TEMP") & "\jp1_output_" & Format(Now, "yyyyMMddHHmmss") & ".txt"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ts As Object
    Set ts = fso.CreateTextFile(tempScript, True, True)
    ts.Write script
    ts.Close

    Dim cmd As String
    cmd = "powershell -NoProfile -ExecutionPolicy Bypass -File """ & tempScript & """ > """ & tempOutput & """ 2>&1"
    shell.Run cmd, 0, True

    Dim result As String
    result = ""
    If fso.FileExists(tempOutput) Then
        Set ts = fso.OpenTextFile(tempOutput, 1, False, -1)
        If Not ts.AtEndOfStream Then
            result = ts.ReadAll
        End If
        ts.Close
    End If

    On Error Resume Next
    fso.DeleteFile tempScript
    fso.DeleteFile tempOutput
    On Error GoTo 0

    ExecutePowerShell = result
End Function

Private Function EscapePSString(s As String) As String
    EscapePSString = Replace(s, "'", "''")
End Function

'==============================================================================
' 管理者権限チェック
'==============================================================================
Private Function CheckAndRequestAdminPrivileges() As Boolean
    If g_AdminChecked Then
        CheckAndRequestAdminPrivileges = g_IsAdmin
        Exit Function
    End If

    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")

    Dim isAdmin As Boolean
    isAdmin = False

    On Error Resume Next
    Dim result As Long
    result = shell.Run("net session", 0, True)
    If result = 0 Then
        isAdmin = True
    End If
    On Error GoTo 0

    g_AdminChecked = True
    g_IsAdmin = isAdmin

    If Not isAdmin Then
        Dim response As VbMsgBoxResult
        response = MsgBox("リモートモードでは管理者権限が必要です。" & vbCrLf & vbCrLf & _
                          "管理者権限でExcelを再起動しますか？" & vbCrLf & vbCrLf & _
                          "「はい」→ 管理者として再起動" & vbCrLf & _
                          "「いいえ」→ このまま続行", _
                          vbYesNo + vbQuestion, "管理者権限が必要")

        If response = vbYes Then
            RestartAsAdmin
            CheckAndRequestAdminPrivileges = False
            Exit Function
        Else
            g_IsAdmin = True
        End If
    End If

    CheckAndRequestAdminPrivileges = True
End Function

Private Sub RestartAsAdmin()
    If Not ThisWorkbook.Saved Then
        ThisWorkbook.Save
    End If

    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")

    Dim excelPath As String
    excelPath = Application.path & "\EXCEL.EXE"

    Dim workbookPath As String
    workbookPath = ThisWorkbook.FullName

    Dim cmd As String
    cmd = "powershell -NoProfile -Command ""Start-Process -FilePath '" & Replace(excelPath, "'", "''") & "' -ArgumentList '""" & Replace(workbookPath, "'", "''") & """' -Verb RunAs"""

    shell.Run cmd, 0, False
    Application.Quit
End Sub

'==============================================================================
' ログファイル関連
'==============================================================================
Private Function CreateLogFile() As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim logFolder As String
    logFolder = ThisWorkbook.path & "\Logs"

    If Not fso.FolderExists(logFolder) Then
        fso.CreateFolder logFolder
    End If

    Dim logFileName As String
    logFileName = "JP1_実行ログ_" & Format(Now, "yyyyMMdd_HHmmss") & ".txt"

    Dim logFilePath As String
    logFilePath = logFolder & "\" & logFileName

    Dim ts As Object
    Set ts = fso.CreateTextFile(logFilePath, True, True)
    ts.WriteLine "================================================================================"
    ts.WriteLine "JP1 ジョブ管理ツール（バッチ版） - 実行ログ"
    ts.WriteLine "================================================================================"
    ts.WriteLine "開始日時: " & Format(Now, "yyyy/mm/dd HH:mm:ss")
    ts.WriteLine "================================================================================"
    ts.WriteLine ""
    ts.Close

    CreateLogFile = logFilePath
End Function

Private Function GetLogFilePath() As String
    GetLogFilePath = g_LogFilePath
End Function
