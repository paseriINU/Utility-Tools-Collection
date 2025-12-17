Option Explicit

'==============================================================================
' JP1 ジョブ管理ツール - メインモジュール
'   - ジョブネット一覧取得（ajsprint経由）
'   - ジョブ実行処理
'
' 注意: 初期化処理は JP1_ジョブ管理ツール_Setup.bas にあります
'       定数はSetupモジュールで Public として定義されています
'==============================================================================

'==============================================================================
' ジョブ一覧取得
'==============================================================================
Public Sub GetJobList()
    Dim config As Object
    Set config = GetConfig()

    If config Is Nothing Then Exit Sub

    ' パスワード入力
    If config("RemotePassword") = "" Then
        config("RemotePassword") = InputBox("リモートサーバのパスワードを入力してください:", "パスワード入力")
        If config("RemotePassword") = "" Then
            MsgBox "パスワードが入力されませんでした。", vbExclamation
            Exit Sub
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

    ' PowerShellスクリプト生成・実行
    Dim psScript As String
    psScript = BuildGetJobListScript(config)

    Dim result As String
    result = ExecutePowerShell(psScript)

    ' 結果をパース
    ParseJobListResult result

    Application.StatusBar = False
    Application.ScreenUpdating = True

    MsgBox "ジョブ一覧の取得が完了しました。" & vbCrLf & _
           "ジョブ一覧シートを確認してください。", vbInformation

    Worksheets(SHEET_JOBLIST).Activate
End Sub

Private Function BuildGetJobListScript(config As Object) As String
    Dim script As String

    script = "$ErrorActionPreference = 'Stop'" & vbCrLf
    script = script & "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8" & vbCrLf
    script = script & vbCrLf

    ' 認証情報
    script = script & "$securePass = ConvertTo-SecureString '" & EscapePSString(config("RemotePassword")) & "' -AsPlainText -Force" & vbCrLf
    script = script & "$cred = New-Object System.Management.Automation.PSCredential('" & config("RemoteUser") & "', $securePass)" & vbCrLf
    script = script & vbCrLf

    ' WinRM設定
    script = script & "try {" & vbCrLf
    script = script & "  $originalTH = (Get-Item WSMan:\localhost\Client\TrustedHosts -ErrorAction SilentlyContinue).Value" & vbCrLf
    script = script & "  if ($originalTH -notmatch '" & config("JP1Server") & "') {" & vbCrLf
    script = script & "    if ($originalTH) { Set-Item WSMan:\localhost\Client\TrustedHosts -Value ""$originalTH," & config("JP1Server") & """ -Force -Confirm:`$false }" & vbCrLf
    script = script & "    else { Set-Item WSMan:\localhost\Client\TrustedHosts -Value '" & config("JP1Server") & "' -Force -Confirm:`$false }" & vbCrLf
    script = script & "  }" & vbCrLf
    script = script & vbCrLf

    ' リモート実行
    script = script & "  $session = New-PSSession -ComputerName '" & config("JP1Server") & "' -Credential $cred -ErrorAction Stop" & vbCrLf
    script = script & vbCrLf
    script = script & "  $result = Invoke-Command -Session $session -ScriptBlock {" & vbCrLf
    script = script & "    param($jp1User, $jp1Pass, $rootPath)" & vbCrLf
    script = script & "    $ajsprintPath = 'C:\Program Files\HITACHI\JP1AJS3\bin\ajsprint.exe'" & vbCrLf
    script = script & "    if (-not (Test-Path $ajsprintPath)) {" & vbCrLf
    script = script & "      $ajsprintPath = 'C:\Program Files\Hitachi\JP1AJS2\bin\ajsprint.exe'" & vbCrLf
    script = script & "    }" & vbCrLf
    script = script & "    & $ajsprintPath -h localhost -u $jp1User -p $jp1Pass -F $rootPath -R 2>&1" & vbCrLf
    script = script & "  } -ArgumentList '" & config("JP1User") & "', '" & EscapePSString(config("JP1Password")) & "', '" & config("RootPath") & "'" & vbCrLf
    script = script & vbCrLf
    script = script & "  Remove-PSSession $session" & vbCrLf
    script = script & vbCrLf
    script = script & "  $result | ForEach-Object { Write-Output $_ }" & vbCrLf
    script = script & "} catch {" & vbCrLf
    script = script & "  Write-Output ""ERROR: $($_.Exception.Message)""" & vbCrLf
    script = script & "} finally {" & vbCrLf
    script = script & "  if ($originalTH -ne $null) {" & vbCrLf
    script = script & "    Set-Item WSMan:\localhost\Client\TrustedHosts -Value $originalTH -Force -Confirm:`$false -ErrorAction SilentlyContinue" & vbCrLf
    script = script & "  }" & vbCrLf
    script = script & "}" & vbCrLf

    BuildGetJobListScript = script
End Function

Private Sub ParseJobListResult(result As String)
    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_JOBLIST)

    ' 既存データをクリア
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_JOBNET_PATH).End(xlUp).Row
    If lastRow >= ROW_JOBLIST_DATA_START Then
        ws.Range(ws.Cells(ROW_JOBLIST_DATA_START, COL_ORDER), ws.Cells(lastRow, COL_LAST_MESSAGE)).ClearContents
    End If

    ' 結果をパース
    Dim lines() As String
    lines = Split(result, vbCrLf)

    Dim row As Long
    row = ROW_JOBLIST_DATA_START

    Dim i As Long
    Dim currentJobnet As String
    Dim jobnetName As String
    Dim jobnetComment As String

    For i = LBound(lines) To UBound(lines)
        Dim line As String
        line = Trim(lines(i))

        ' エラーチェック
        If InStr(line, "ERROR:") > 0 Then
            MsgBox "エラーが発生しました:" & vbCrLf & line, vbExclamation
            Exit Sub
        End If

        ' ジョブネット定義の行を検出（unit=で始まる行）
        If InStr(line, "unit=") > 0 Then
            ' unit=/path/to/jobnet,name,ty=n; 形式
            Dim unitMatch As String
            unitMatch = ExtractUnitPath(line)

            If unitMatch <> "" And InStr(line, ",ty=n") > 0 Then
                ' ジョブネット（ty=n）のみ追加
                ws.Cells(row, COL_ORDER).Value = ""
                ws.Cells(row, COL_JOBNET_PATH).Value = unitMatch
                ws.Cells(row, COL_JOBNET_NAME).Value = ExtractJobName(line)
                ws.Cells(row, COL_COMMENT).Value = ExtractComment(line)

                ' 順序列の書式
                With ws.Cells(row, COL_ORDER)
                    .HorizontalAlignment = xlCenter
                End With

                ' 罫線
                ws.Range(ws.Cells(row, COL_ORDER), ws.Cells(row, COL_LAST_MESSAGE)).Borders.LineStyle = xlContinuous

                row = row + 1
            End If
        End If
    Next i

    ' データがない場合
    If row = ROW_JOBLIST_DATA_START Then
        MsgBox "ジョブネットが見つかりませんでした。" & vbCrLf & _
               "取得パスを確認してください。", vbExclamation
    End If
End Sub

Private Function ExtractUnitPath(line As String) As String
    ' unit=/path/to/jobnet から /path/to/jobnet を抽出
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
    Dim config As Object
    Set config = GetConfig()

    If config Is Nothing Then Exit Sub

    ' パスワード入力
    If config("RemotePassword") = "" Then
        config("RemotePassword") = InputBox("リモートサーバのパスワードを入力してください:", "パスワード入力")
        If config("RemotePassword") = "" Then
            MsgBox "パスワードが入力されませんでした。", vbExclamation
            Exit Sub
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

    ' 確認
    Dim msg As String
    msg = "以下の " & jobs.Count & " 件のジョブを実行します：" & vbCrLf & vbCrLf
    Dim j As Variant
    Dim cnt As Long
    cnt = 0
    For Each j In jobs
        cnt = cnt + 1
        If cnt <= 5 Then
            msg = msg & cnt & ". " & j("Path") & vbCrLf
        ElseIf cnt = 6 Then
            msg = msg & "..." & vbCrLf
        End If
    Next j
    msg = msg & vbCrLf & "実行しますか？"

    If MsgBox(msg, vbYesNo + vbQuestion, "実行確認") = vbNo Then Exit Sub

    ' 実行
    Application.ScreenUpdating = False

    Dim wsLog As Worksheet
    Set wsLog = Worksheets(SHEET_LOG)
    Dim logRow As Long
    logRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1
    If logRow < 4 Then logRow = 4

    Dim success As Boolean
    success = True

    For Each j In jobs
        Application.StatusBar = "実行中: " & j("Path")

        Dim execResult As Object
        Set execResult = ExecuteSingleJob(config, j("Path"))

        ' 結果をログに記録
        wsLog.Cells(logRow, 1).Value = Now
        wsLog.Cells(logRow, 2).Value = j("Path")
        wsLog.Cells(logRow, 3).Value = execResult("Status")
        wsLog.Cells(logRow, 4).Value = execResult("StartTime")
        wsLog.Cells(logRow, 5).Value = execResult("EndTime")
        wsLog.Cells(logRow, 6).Value = execResult("Message")

        ' 色付け
        If execResult("Status") = "正常終了" Then
            wsLog.Cells(logRow, 3).Interior.Color = RGB(198, 239, 206)
        Else
            wsLog.Cells(logRow, 3).Interior.Color = RGB(255, 199, 206)
        End If

        wsLog.Range(wsLog.Cells(logRow, 1), wsLog.Cells(logRow, 6)).Borders.LineStyle = xlContinuous

        ' ジョブ一覧シートも更新
        UpdateJobListStatus j("Row"), execResult

        logRow = logRow + 1

        ' エラー時は停止
        If execResult("Status") <> "正常終了" And execResult("Status") <> "起動成功" Then
            success = False
            MsgBox "ジョブ「" & j("Path") & "」が失敗しました。" & vbCrLf & _
                   "処理を中断します。" & vbCrLf & vbCrLf & _
                   "詳細: " & execResult("Message"), vbCritical
            Exit For
        End If
    Next j

    Application.StatusBar = False
    Application.ScreenUpdating = True

    If success Then
        MsgBox "すべてのジョブが正常に完了しました。", vbInformation
    End If

    Worksheets(SHEET_LOG).Activate
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
    For i = 1 To orderedRows.Count
        Set arr(i) = orderedRows(i)
    Next i

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

    For i = 1 To UBound(arr)
        jobs.Add arr(i)
    Next i

    Set GetOrderedJobs = jobs
End Function

Private Function ExecuteSingleJob(config As Object, jobnetPath As String) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result("Status") = ""
    result("StartTime") = ""
    result("EndTime") = ""
    result("Message") = ""

    Dim waitCompletion As Boolean
    waitCompletion = (config("WaitCompletion") = "はい")

    Dim psScript As String
    psScript = BuildExecuteJobScript(config, jobnetPath, waitCompletion)

    Dim output As String
    output = ExecutePowerShell(psScript)

    ' 結果をパース
    Dim lines() As String
    lines = Split(output, vbCrLf)

    Dim line As String
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

Private Function BuildExecuteJobScript(config As Object, jobnetPath As String, waitCompletion As Boolean) As String
    Dim script As String

    script = "$ErrorActionPreference = 'Stop'" & vbCrLf
    script = script & "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8" & vbCrLf
    script = script & vbCrLf

    ' 認証情報
    script = script & "$securePass = ConvertTo-SecureString '" & EscapePSString(config("RemotePassword")) & "' -AsPlainText -Force" & vbCrLf
    script = script & "$cred = New-Object System.Management.Automation.PSCredential('" & config("RemoteUser") & "', $securePass)" & vbCrLf
    script = script & vbCrLf

    ' WinRM設定
    script = script & "try {" & vbCrLf
    script = script & "  $originalTH = (Get-Item WSMan:\localhost\Client\TrustedHosts -ErrorAction SilentlyContinue).Value" & vbCrLf
    script = script & "  if ($originalTH -notmatch '" & config("JP1Server") & "') {" & vbCrLf
    script = script & "    if ($originalTH) { Set-Item WSMan:\localhost\Client\TrustedHosts -Value ""$originalTH," & config("JP1Server") & """ -Force -Confirm:`$false }" & vbCrLf
    script = script & "    else { Set-Item WSMan:\localhost\Client\TrustedHosts -Value '" & config("JP1Server") & "' -Force -Confirm:`$false }" & vbCrLf
    script = script & "  }" & vbCrLf
    script = script & vbCrLf

    ' リモート実行
    script = script & "  $session = New-PSSession -ComputerName '" & config("JP1Server") & "' -Credential $cred -ErrorAction Stop" & vbCrLf
    script = script & vbCrLf

    ' ajsentry実行
    script = script & "  $entryResult = Invoke-Command -Session $session -ScriptBlock {" & vbCrLf
    script = script & "    param($jp1User, $jp1Pass, $jobnetPath)" & vbCrLf
    script = script & "    $ajsentryPath = 'C:\Program Files\HITACHI\JP1AJS3\bin\ajsentry.exe'" & vbCrLf
    script = script & "    if (-not (Test-Path $ajsentryPath)) { $ajsentryPath = 'C:\Program Files\Hitachi\JP1AJS2\bin\ajsentry.exe' }" & vbCrLf
    script = script & "    $output = & $ajsentryPath -h localhost -u $jp1User -p $jp1Pass -F $jobnetPath 2>&1" & vbCrLf
    script = script & "    @{ ExitCode = $LASTEXITCODE; Output = ($output -join ' ') }" & vbCrLf
    script = script & "  } -ArgumentList '" & config("JP1User") & "', '" & EscapePSString(config("JP1Password")) & "', '" & jobnetPath & "'" & vbCrLf
    script = script & vbCrLf

    script = script & "  if ($entryResult.ExitCode -ne 0) {" & vbCrLf
    script = script & "    Write-Output ""RESULT_STATUS:起動失敗""" & vbCrLf
    script = script & "    Write-Output ""RESULT_MESSAGE:$($entryResult.Output)""" & vbCrLf
    script = script & "    Remove-PSSession $session" & vbCrLf
    script = script & "    exit" & vbCrLf
    script = script & "  }" & vbCrLf
    script = script & vbCrLf

    If waitCompletion Then
        ' 完了待ち
        script = script & "  $timeout = " & config("Timeout") & vbCrLf
        script = script & "  $interval = " & config("PollingInterval") & vbCrLf
        script = script & "  $startTime = Get-Date" & vbCrLf
        script = script & "  $isRunning = $true" & vbCrLf
        script = script & vbCrLf
        script = script & "  while ($isRunning) {" & vbCrLf
        script = script & "    if ($timeout -gt 0 -and ((Get-Date) - $startTime).TotalSeconds -ge $timeout) {" & vbCrLf
        script = script & "      Write-Output ""RESULT_STATUS:タイムアウト""" & vbCrLf
        script = script & "      break" & vbCrLf
        script = script & "    }" & vbCrLf
        script = script & vbCrLf
        script = script & "    $statusResult = Invoke-Command -Session $session -ScriptBlock {" & vbCrLf
        script = script & "      param($jp1User, $jp1Pass, $jobnetPath)" & vbCrLf
        script = script & "      $ajsstatusPath = 'C:\Program Files\HITACHI\JP1AJS3\bin\ajsstatus.exe'" & vbCrLf
        script = script & "      if (-not (Test-Path $ajsstatusPath)) { $ajsstatusPath = 'C:\Program Files\Hitachi\JP1AJS2\bin\ajsstatus.exe' }" & vbCrLf
        script = script & "      & $ajsstatusPath -h localhost -u $jp1User -p $jp1Pass -F $jobnetPath 2>&1" & vbCrLf
        script = script & "    } -ArgumentList '" & config("JP1User") & "', '" & EscapePSString(config("JP1Password")) & "', '" & jobnetPath & "'" & vbCrLf
        script = script & vbCrLf
        script = script & "    $statusStr = ($statusResult -join ' ').ToLower()" & vbCrLf
        script = script & "    if ($statusStr -match 'ended abnormally|abnormal end|abend|killed|failed') {" & vbCrLf
        script = script & "      Write-Output ""RESULT_STATUS:異常終了""" & vbCrLf
        script = script & "      $isRunning = $false" & vbCrLf
        script = script & "    } elseif ($statusStr -match 'end normally|ended normally|normal end|completed') {" & vbCrLf
        script = script & "      Write-Output ""RESULT_STATUS:正常終了""" & vbCrLf
        script = script & "      $isRunning = $false" & vbCrLf
        script = script & "    } else {" & vbCrLf
        script = script & "      Start-Sleep -Seconds $interval" & vbCrLf
        script = script & "    }" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & vbCrLf

        ' 詳細取得
        script = script & "  $showResult = Invoke-Command -Session $session -ScriptBlock {" & vbCrLf
        script = script & "    param($jp1User, $jp1Pass, $jobnetPath)" & vbCrLf
        script = script & "    $ajsshowPath = 'C:\Program Files\HITACHI\JP1AJS3\bin\ajsshow.exe'" & vbCrLf
        script = script & "    if (-not (Test-Path $ajsshowPath)) { $ajsshowPath = 'C:\Program Files\Hitachi\JP1AJS2\bin\ajsshow.exe' }" & vbCrLf
        script = script & "    if (Test-Path $ajsshowPath) { & $ajsshowPath -h localhost -u $jp1User -p $jp1Pass -F $jobnetPath -E 2>&1 }" & vbCrLf
        script = script & "  } -ArgumentList '" & config("JP1User") & "', '" & EscapePSString(config("JP1Password")) & "', '" & jobnetPath & "'" & vbCrLf
        script = script & "  Write-Output ""RESULT_MESSAGE:$($showResult -join ' ')""" & vbCrLf
    Else
        script = script & "  Write-Output ""RESULT_STATUS:起動成功""" & vbCrLf
        script = script & "  Write-Output ""RESULT_MESSAGE:$($entryResult.Output)""" & vbCrLf
    End If

    script = script & "  Remove-PSSession $session" & vbCrLf
    script = script & "} catch {" & vbCrLf
    script = script & "  Write-Output ""ERROR: $($_.Exception.Message)""" & vbCrLf
    script = script & "} finally {" & vbCrLf
    script = script & "  if ($originalTH -ne $null) {" & vbCrLf
    script = script & "    Set-Item WSMan:\localhost\Client\TrustedHosts -Value $originalTH -Force -Confirm:`$false -ErrorAction SilentlyContinue" & vbCrLf
    script = script & "  }" & vbCrLf
    script = script & "}" & vbCrLf

    BuildExecuteJobScript = script
End Function

Private Sub UpdateJobListStatus(row As Long, result As Object)
    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_JOBLIST)

    ws.Cells(row, COL_LAST_STATUS).Value = result("Status")
    ws.Cells(row, COL_LAST_EXEC_TIME).Value = result("StartTime")
    ws.Cells(row, COL_LAST_END_TIME).Value = result("EndTime")

    ' 詳細メッセージを記録
    If result("Message") <> "" Then
        ws.Cells(row, COL_LAST_MESSAGE).Value = result("Message")
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
' 一覧クリア
'==============================================================================
Public Sub ClearJobList()
    If MsgBox("ジョブ一覧をクリアしますか？ (y/n)", vbYesNo + vbQuestion) = vbNo Then Exit Sub

    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_JOBLIST)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_JOBNET_PATH).End(xlUp).Row

    If lastRow >= ROW_JOBLIST_DATA_START Then
        ws.Range(ws.Cells(ROW_JOBLIST_DATA_START, COL_ORDER), ws.Cells(lastRow, COL_LAST_MESSAGE)).Clear
    End If

    MsgBox "クリアしました。", vbInformation
End Sub

'==============================================================================
' ユーティリティ
'==============================================================================
Private Function GetConfig() As Object
    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_MAIN)

    Dim config As Object
    Set config = CreateObject("Scripting.Dictionary")

    config("JP1Server") = CStr(ws.Cells(ROW_JP1_SERVER, COL_SETTING_VALUE).Value)
    config("RemoteUser") = CStr(ws.Cells(ROW_REMOTE_USER, COL_SETTING_VALUE).Value)
    config("RemotePassword") = CStr(ws.Cells(ROW_REMOTE_PASSWORD, COL_SETTING_VALUE).Value)
    config("JP1User") = CStr(ws.Cells(ROW_JP1_USER, COL_SETTING_VALUE).Value)
    config("JP1Password") = CStr(ws.Cells(ROW_JP1_PASSWORD, COL_SETTING_VALUE).Value)
    config("RootPath") = CStr(ws.Cells(ROW_ROOT_PATH, COL_SETTING_VALUE).Value)
    config("WaitCompletion") = CStr(ws.Cells(ROW_WAIT_COMPLETION + 4, COL_SETTING_VALUE).Value)
    config("Timeout") = CLng(ws.Cells(ROW_TIMEOUT + 4, COL_SETTING_VALUE).Value)
    config("PollingInterval") = CLng(ws.Cells(ROW_POLLING_INTERVAL + 4, COL_SETTING_VALUE).Value)

    ' 必須項目チェック
    If config("JP1Server") = "" Or config("RemoteUser") = "" Or config("JP1User") = "" Then
        MsgBox "接続設定が不完全です。メインシートで設定を入力してください。", vbExclamation
        Set GetConfig = Nothing
        Exit Function
    End If

    Set GetConfig = config
End Function

Private Function ExecutePowerShell(script As String) As String
    ' 一時ファイルにスクリプトを保存
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim tempFolder As String
    tempFolder = fso.GetSpecialFolder(2) ' Temp folder

    Dim scriptPath As String
    scriptPath = tempFolder & "\jp1_temp_" & Format(Now, "yyyymmddhhnnss") & ".ps1"

    Dim ts As Object
    Set ts = fso.CreateTextFile(scriptPath, True, True) ' Unicode
    ts.Write script
    ts.Close

    ' PowerShell実行
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")

    Dim cmd As String
    cmd = "powershell -NoProfile -ExecutionPolicy Bypass -File """ & scriptPath & """"

    Dim exec As Object
    Set exec = shell.exec(cmd)

    ' 結果を取得
    Dim output As String
    output = ""

    Do While exec.Status = 0
        DoEvents
    Loop

    output = exec.StdOut.ReadAll

    ' 一時ファイル削除
    On Error Resume Next
    fso.DeleteFile scriptPath
    On Error GoTo 0

    ExecutePowerShell = output
End Function

Private Function EscapePSString(str As String) As String
    ' PowerShell文字列内のシングルクォートをエスケープ
    EscapePSString = Replace(str, "'", "''")
End Function
