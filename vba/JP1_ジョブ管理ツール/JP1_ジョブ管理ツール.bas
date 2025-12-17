Attribute VB_Name = "JP1_JobManager"
Option Explicit

'==============================================================================
' JP1 ジョブ管理ツール
'   - ジョブネット一覧取得（ajsprint経由）
'   - Excelでチェック・実行順指定
'   - チェックしたジョブを順次実行
'==============================================================================

' シート名定数
Private Const SHEET_MAIN As String = "メイン"
Private Const SHEET_JOBLIST As String = "ジョブ一覧"
Private Const SHEET_LOG As String = "実行ログ"

' 設定セル位置（メインシート）
Private Const ROW_JP1_SERVER As Long = 5
Private Const ROW_REMOTE_USER As Long = 6
Private Const ROW_REMOTE_PASSWORD As Long = 7
Private Const ROW_JP1_USER As Long = 8
Private Const ROW_JP1_PASSWORD As Long = 9
Private Const ROW_ROOT_PATH As Long = 10
Private Const ROW_WAIT_COMPLETION As Long = 11
Private Const ROW_TIMEOUT As Long = 12
Private Const ROW_POLLING_INTERVAL As Long = 13
Private Const COL_SETTING_VALUE As Long = 3

' ジョブ一覧シートの列位置
Private Const COL_CHECK As Long = 1
Private Const COL_ORDER As Long = 2
Private Const COL_JOBNET_PATH As Long = 3
Private Const COL_JOBNET_NAME As Long = 4
Private Const COL_COMMENT As Long = 5
Private Const COL_LAST_STATUS As Long = 6
Private Const COL_LAST_EXEC_TIME As Long = 7
Private Const COL_LAST_END_TIME As Long = 8
Private Const COL_LAST_RETURN_CODE As Long = 9
Private Const ROW_JOBLIST_HEADER As Long = 3
Private Const ROW_JOBLIST_DATA_START As Long = 4

'==============================================================================
' 初期化
'==============================================================================
Public Sub InitializeJP1Manager()
    Application.ScreenUpdating = False

    ' シート作成
    CreateSheet SHEET_MAIN
    CreateSheet SHEET_JOBLIST
    CreateSheet SHEET_LOG

    ' メインシートのフォーマット
    FormatMainSheet

    ' ジョブ一覧シートのフォーマット
    FormatJobListSheet

    ' ログシートのフォーマット
    FormatLogSheet

    ' メインシートをアクティブに
    Worksheets(SHEET_MAIN).Activate

    Application.ScreenUpdating = True

    MsgBox "初期化が完了しました。" & vbCrLf & vbCrLf & _
           "1. メインシートで接続設定を入力してください" & vbCrLf & _
           "2. 「ジョブ一覧取得」ボタンでジョブを取得" & vbCrLf & _
           "3. ジョブ一覧シートでチェック・実行順を設定" & vbCrLf & _
           "4. 「選択ジョブ実行」ボタンで実行", _
           vbInformation, "JP1 ジョブ管理ツール"
End Sub

Private Sub CreateSheet(sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
        ws.Name = sheetName
    End If
End Sub

'==============================================================================
' メインシートのフォーマット
'==============================================================================
Private Sub FormatMainSheet()
    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_MAIN)

    ws.Cells.Clear

    ' タイトル
    With ws.Range("A1:F1")
        .Merge
        .Value = "JP1 ジョブ管理ツール"
        .Font.Size = 16
        .Font.Bold = True
        .Interior.Color = RGB(0, 112, 192)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .RowHeight = 30
    End With

    ' 説明
    ws.Range("A2").Value = "JP1サーバに接続してジョブネット一覧を取得し、選択したジョブを実行します。"

    ' 設定セクション
    ws.Range("A4").Value = "■ 接続設定"
    ws.Range("A4").Font.Bold = True

    ws.Cells(ROW_JP1_SERVER, 1).Value = "JP1サーバ"
    ws.Cells(ROW_JP1_SERVER, COL_SETTING_VALUE).Value = "192.168.1.100"

    ws.Cells(ROW_REMOTE_USER, 1).Value = "リモートユーザー"
    ws.Cells(ROW_REMOTE_USER, COL_SETTING_VALUE).Value = "Administrator"

    ws.Cells(ROW_REMOTE_PASSWORD, 1).Value = "リモートパスワード"
    ws.Cells(ROW_REMOTE_PASSWORD, COL_SETTING_VALUE).Value = ""
    ws.Cells(ROW_REMOTE_PASSWORD, 4).Value = "※空の場合は実行時に入力"
    ws.Cells(ROW_REMOTE_PASSWORD, 4).Font.Color = RGB(128, 128, 128)

    ws.Cells(ROW_JP1_USER, 1).Value = "JP1ユーザー"
    ws.Cells(ROW_JP1_USER, COL_SETTING_VALUE).Value = "jp1admin"

    ws.Cells(ROW_JP1_PASSWORD, 1).Value = "JP1パスワード"
    ws.Cells(ROW_JP1_PASSWORD, COL_SETTING_VALUE).Value = ""
    ws.Cells(ROW_JP1_PASSWORD, 4).Value = "※空の場合は実行時に入力"
    ws.Cells(ROW_JP1_PASSWORD, 4).Font.Color = RGB(128, 128, 128)

    ws.Cells(ROW_ROOT_PATH, 1).Value = "取得パス"
    ws.Cells(ROW_ROOT_PATH, COL_SETTING_VALUE).Value = "/"
    ws.Cells(ROW_ROOT_PATH, 4).Value = "※ジョブネット取得の起点パス（/で全件）"
    ws.Cells(ROW_ROOT_PATH, 4).Font.Color = RGB(128, 128, 128)

    ' 実行設定セクション
    ws.Range("A14").Value = "■ 実行設定"
    ws.Range("A14").Font.Bold = True

    ws.Cells(ROW_WAIT_COMPLETION + 4, 1).Value = "完了待ち"
    ws.Cells(ROW_WAIT_COMPLETION + 4, COL_SETTING_VALUE).Value = "はい"
    AddDropdown ws, ws.Cells(ROW_WAIT_COMPLETION + 4, COL_SETTING_VALUE), "はい,いいえ"

    ws.Cells(ROW_TIMEOUT + 4, 1).Value = "タイムアウト（秒）"
    ws.Cells(ROW_TIMEOUT + 4, COL_SETTING_VALUE).Value = 3600

    ws.Cells(ROW_POLLING_INTERVAL + 4, 1).Value = "状態確認間隔（秒）"
    ws.Cells(ROW_POLLING_INTERVAL + 4, COL_SETTING_VALUE).Value = 10

    ' ボタン追加
    AddButton ws, 200, 50, 150, 30, "GetJobList", "ジョブ一覧取得"
    AddButton ws, 200, 90, 150, 30, "ExecuteCheckedJobs", "選択ジョブ実行"
    AddButton ws, 200, 130, 150, 30, "ClearJobList", "一覧クリア"

    ' 列幅調整
    ws.Columns("A").ColumnWidth = 20
    ws.Columns("B").ColumnWidth = 5
    ws.Columns("C").ColumnWidth = 30
    ws.Columns("D").ColumnWidth = 40

    ' 入力セルの書式
    With ws.Range(ws.Cells(ROW_JP1_SERVER, COL_SETTING_VALUE), ws.Cells(ROW_POLLING_INTERVAL + 4, COL_SETTING_VALUE))
        .Interior.Color = RGB(255, 255, 204)
        .Borders.LineStyle = xlContinuous
    End With
End Sub

Private Sub AddDropdown(ws As Worksheet, cell As Range, options As String)
    With cell.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=options
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
End Sub

Private Sub AddButton(ws As Worksheet, left As Double, top As Double, width As Double, height As Double, macroName As String, caption As String)
    Dim btn As Button
    Set btn = ws.Buttons.Add(left, top, width, height)
    btn.OnAction = macroName
    btn.caption = caption
    btn.Font.Size = 10
End Sub

'==============================================================================
' ジョブ一覧シートのフォーマット
'==============================================================================
Private Sub FormatJobListSheet()
    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_JOBLIST)

    ws.Cells.Clear

    ' タイトル
    With ws.Range("A1:I1")
        .Merge
        .Value = "ジョブネット一覧"
        .Font.Size = 14
        .Font.Bold = True
        .Interior.Color = RGB(0, 176, 80)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .RowHeight = 25
    End With

    ' 説明
    ws.Range("A2").Value = "実行するジョブにチェックを入れ、実行順を数字で指定してください。"

    ' ヘッダー
    ws.Cells(ROW_JOBLIST_HEADER, COL_CHECK).Value = "実行"
    ws.Cells(ROW_JOBLIST_HEADER, COL_ORDER).Value = "順序"
    ws.Cells(ROW_JOBLIST_HEADER, COL_JOBNET_PATH).Value = "ジョブネットパス"
    ws.Cells(ROW_JOBLIST_HEADER, COL_JOBNET_NAME).Value = "ジョブネット名"
    ws.Cells(ROW_JOBLIST_HEADER, COL_COMMENT).Value = "コメント"
    ws.Cells(ROW_JOBLIST_HEADER, COL_LAST_STATUS).Value = "最終実行結果"
    ws.Cells(ROW_JOBLIST_HEADER, COL_LAST_EXEC_TIME).Value = "開始時刻"
    ws.Cells(ROW_JOBLIST_HEADER, COL_LAST_END_TIME).Value = "終了時刻"
    ws.Cells(ROW_JOBLIST_HEADER, COL_LAST_RETURN_CODE).Value = "戻り値"

    With ws.Range(ws.Cells(ROW_JOBLIST_HEADER, COL_CHECK), ws.Cells(ROW_JOBLIST_HEADER, COL_LAST_RETURN_CODE))
        .Font.Bold = True
        .Interior.Color = RGB(79, 129, 189)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With

    ' 列幅調整
    ws.Columns(COL_CHECK).ColumnWidth = 6
    ws.Columns(COL_ORDER).ColumnWidth = 6
    ws.Columns(COL_JOBNET_PATH).ColumnWidth = 50
    ws.Columns(COL_JOBNET_NAME).ColumnWidth = 25
    ws.Columns(COL_COMMENT).ColumnWidth = 30
    ws.Columns(COL_LAST_STATUS).ColumnWidth = 15
    ws.Columns(COL_LAST_EXEC_TIME).ColumnWidth = 18
    ws.Columns(COL_LAST_END_TIME).ColumnWidth = 18
    ws.Columns(COL_LAST_RETURN_CODE).ColumnWidth = 8

    ' フィルター設定
    ws.Range(ws.Cells(ROW_JOBLIST_HEADER, COL_CHECK), ws.Cells(ROW_JOBLIST_HEADER, COL_LAST_RETURN_CODE)).AutoFilter
End Sub

'==============================================================================
' ログシートのフォーマット
'==============================================================================
Private Sub FormatLogSheet()
    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_LOG)

    ws.Cells.Clear

    ' タイトル
    With ws.Range("A1:F1")
        .Merge
        .Value = "実行ログ"
        .Font.Size = 14
        .Font.Bold = True
        .Interior.Color = RGB(192, 80, 77)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .RowHeight = 25
    End With

    ' ヘッダー
    ws.Cells(3, 1).Value = "実行日時"
    ws.Cells(3, 2).Value = "ジョブネットパス"
    ws.Cells(3, 3).Value = "結果"
    ws.Cells(3, 4).Value = "開始時刻"
    ws.Cells(3, 5).Value = "終了時刻"
    ws.Cells(3, 6).Value = "詳細メッセージ"

    With ws.Range("A3:F3")
        .Font.Bold = True
        .Interior.Color = RGB(192, 80, 77)
        .Font.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
    End With

    ' 列幅調整
    ws.Columns("A").ColumnWidth = 18
    ws.Columns("B").ColumnWidth = 50
    ws.Columns("C").ColumnWidth = 12
    ws.Columns("D").ColumnWidth = 18
    ws.Columns("E").ColumnWidth = 18
    ws.Columns("F").ColumnWidth = 60
End Sub

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
        ws.Range(ws.Cells(ROW_JOBLIST_DATA_START, COL_CHECK), ws.Cells(lastRow, COL_LAST_RETURN_CODE)).ClearContents
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
                ws.Cells(row, COL_CHECK).Value = ""
                ws.Cells(row, COL_ORDER).Value = ""
                ws.Cells(row, COL_JOBNET_PATH).Value = unitMatch
                ws.Cells(row, COL_JOBNET_NAME).Value = ExtractJobName(line)
                ws.Cells(row, COL_COMMENT).Value = ExtractComment(line)

                ' チェックボックス用の書式
                With ws.Cells(row, COL_CHECK)
                    .HorizontalAlignment = xlCenter
                End With
                With ws.Cells(row, COL_ORDER)
                    .HorizontalAlignment = xlCenter
                End With

                ' 罫線
                ws.Range(ws.Cells(row, COL_CHECK), ws.Cells(row, COL_LAST_RETURN_CODE)).Borders.LineStyle = xlContinuous

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

    ' チェックされたジョブを取得
    Dim jobs As Collection
    Set jobs = GetCheckedJobs()

    If jobs.Count = 0 Then
        MsgBox "実行するジョブが選択されていません。" & vbCrLf & _
               "ジョブ一覧シートで「実行」列にチェック（任意の文字）を入れてください。", vbExclamation
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

Private Function GetCheckedJobs() As Collection
    Dim jobs As New Collection
    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_JOBLIST)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_JOBNET_PATH).End(xlUp).Row

    ' チェックされた行を収集
    Dim checkedRows As New Collection
    Dim row As Long
    For row = ROW_JOBLIST_DATA_START To lastRow
        If ws.Cells(row, COL_CHECK).Value <> "" Then
            Dim job As Object
            Set job = CreateObject("Scripting.Dictionary")
            job("Row") = row
            job("Path") = ws.Cells(row, COL_JOBNET_PATH).Value
            job("Order") = ws.Cells(row, COL_ORDER).Value

            If job("Order") = "" Then job("Order") = 9999

            checkedRows.Add job
        End If
    Next row

    ' 実行順でソート（単純なバブルソート）
    Dim arr() As Variant
    ReDim arr(1 To checkedRows.Count)
    Dim i As Long
    For i = 1 To checkedRows.Count
        Set arr(i) = checkedRows(i)
    Next i

    Dim temp As Object
    For i = 1 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If CLng(arr(i)("Order")) > CLng(arr(j)("Order")) Then
                Set temp = arr(i)
                Set arr(i) = arr(j)
                Set arr(j) = temp
            End If
        Next j
    Next i

    For i = 1 To UBound(arr)
        jobs.Add arr(i)
    Next i

    Set GetCheckedJobs = jobs
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
    If MsgBox("ジョブ一覧をクリアしますか？", vbYesNo + vbQuestion) = vbNo Then Exit Sub

    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_JOBLIST)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_JOBNET_PATH).End(xlUp).Row

    If lastRow >= ROW_JOBLIST_DATA_START Then
        ws.Range(ws.Cells(ROW_JOBLIST_DATA_START, COL_CHECK), ws.Cells(lastRow, COL_LAST_RETURN_CODE)).Clear
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
