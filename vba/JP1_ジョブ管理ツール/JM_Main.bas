Attribute VB_Name = "JM_Main"
Option Explicit

'==============================================================================
' JP1 ジョブ管理ツール - メインモジュール
' エントリーポイント、UI操作処理を提供
' ※定数・設定はJM_Config、パース処理はJM_Parser、実行処理はJM_Executorを参照
'==============================================================================

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

    ' 結果をパース
    Dim parseSuccess As Boolean
    parseSuccess = ParseJobListResult(result, config("RootPath"))

    Application.StatusBar = False
    Application.ScreenUpdating = True

    If Not parseSuccess Then
        Exit Sub
    End If

    ' 種別「ジョブネット」でオートフィルタを適用
    Dim lastDataRow As Long
    lastDataRow = wsJobList.Cells(wsJobList.Rows.Count, COL_JOBNET_PATH).End(xlUp).row
    If lastDataRow >= ROW_JOBLIST_DATA_START Then
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

    ' PowerShellスクリプト生成・実行
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

    ' 取得パス欄にドロップダウンリストを設定
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
    logRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).row + 1
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
            On Error GoTo ErrorHandler
        End If

        ' 色付け
        If execResult("Status") = "正常終了" Then
            wsLog.Cells(logRow, 3).Interior.Color = RGB(198, 239, 206)
        ElseIf execResult("Status") = "起動成功" Then
            wsLog.Cells(logRow, 3).Interior.Color = RGB(255, 235, 156)
        ElseIf execResult("Status") = "警告検出終了" Or execResult("Status") = "警告終了" Then
            wsLog.Cells(logRow, 3).Interior.Color = RGB(255, 192, 0)
        Else
            wsLog.Cells(logRow, 3).Interior.Color = RGB(255, 199, 206)
        End If

        wsLog.Range(wsLog.Cells(logRow, 1), wsLog.Cells(logRow, 6)).Borders.LineStyle = xlContinuous

        ' ジョブ一覧シートも更新
        UpdateJobListStatus j("Row"), execResult

        logRow = logRow + 1

        ' エラー・警告時は停止
        If execResult("Status") <> "正常終了" And execResult("Status") <> "起動成功" Then
            success = False
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

'==============================================================================
' 一覧クリア（実行結果のみクリア、ジョブ定義は保持）
'==============================================================================
Public Sub ClearJobList()
    If MsgBox("実行結果をクリアしますか？" & vbCrLf & _
              "（ジョブ定義情報は保持されます）", vbYesNo + vbQuestion) = vbNo Then Exit Sub

    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_JOBLIST)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_JOBNET_PATH).End(xlUp).row

    If lastRow >= ROW_JOBLIST_DATA_START Then
        ' 選択列と順序列をクリア
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
            ws.Range(ws.Cells(row, COL_SELECT), ws.Cells(row, COL_LAST_MESSAGE)).Interior.ColorIndex = xlNone
            If ws.Cells(row, COL_HOLD).Value = "保留中" Then
                ws.Cells(row, COL_HOLD).Interior.Color = RGB(255, 235, 156)
                ws.Cells(row, COL_HOLD).Font.Bold = True
                ws.Cells(row, COL_HOLD).Font.Color = RGB(156, 87, 0)
            End If
        Next row

        ' オートフィルタを再適用
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

    ' セル編集をキャンセル
    Cancel = True

    ' チェック状態を切り替え
    ToggleCheckMark row
End Sub

'==============================================================================
' 実行ログ履歴クリア
'==============================================================================
Public Sub ClearLogHistory()
    If MsgBox("実行ログの履歴をすべて削除しますか？", vbYesNo + vbQuestion) = vbNo Then Exit Sub

    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_LOG)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    If lastRow >= 5 Then
        ws.Range(ws.Cells(5, 1), ws.Cells(lastRow, 6)).ClearContents

        On Error Resume Next
        ws.Range(ws.Cells(5, 6), ws.Cells(lastRow, 6)).Hyperlinks.Delete
        On Error GoTo 0

        ws.Range(ws.Cells(5, 1), ws.Cells(lastRow, 6)).Interior.ColorIndex = xlNone
        ws.Range(ws.Cells(5, 1), ws.Cells(lastRow, 6)).Borders.LineStyle = xlNone
    End If

    MsgBox "実行ログの履歴を削除しました。", vbInformation
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

    If currentValue = ChrW(&H2611) Then
        ' チェックを外す
        ws.Cells(row, COL_SELECT).Value = ChrW(&H2610)
        ws.Cells(row, COL_ORDER).Value = ""
        RenumberJobOrder

        ' 背景色を元に戻す
        ws.Range(ws.Cells(row, COL_SELECT), ws.Cells(row, COL_LAST_MESSAGE)).Interior.ColorIndex = xlNone
        If ws.Cells(row, COL_HOLD).Value = "保留中" Then
            ws.Cells(row, COL_HOLD).Interior.Color = RGB(255, 235, 156)
        End If
    Else
        ' チェックを入れる
        ws.Cells(row, COL_SELECT).Value = ChrW(&H2611)
        Dim maxOrder As Long
        maxOrder = GetMaxOrderNumber()
        ws.Cells(row, COL_ORDER).Value = maxOrder + 1

        ' 行全体に水色の背景色を設定
        ws.Range(ws.Cells(row, COL_SELECT), ws.Cells(row, COL_LAST_MESSAGE)).Interior.Color = RGB(221, 235, 247)
        If ws.Cells(row, COL_HOLD).Value = "保留中" Then
            ws.Cells(row, COL_HOLD).Interior.Color = RGB(255, 235, 156)
        End If
    End If

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
    lastRow = ws.Cells(ws.Rows.Count, COL_JOBNET_PATH).End(xlUp).row

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
' 順序番号を再採番
'==============================================================================
Private Sub RenumberJobOrder()
    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_JOBLIST)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_JOBNET_PATH).End(xlUp).row

    ' 順序が入っている行を収集
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
' 順序が指定されたジョブを取得
'==============================================================================
Private Function GetOrderedJobs() As Collection
    Dim jobs As New Collection
    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_JOBLIST)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_JOBNET_PATH).End(xlUp).row

    ' 順序が入力されている行を収集
    Dim orderedRows As New Collection
    Dim row As Long
    For row = ROW_JOBLIST_DATA_START To lastRow
        Dim orderValue As Variant
        orderValue = ws.Cells(row, COL_ORDER).Value

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

    If orderedRows.Count = 0 Then
        Set GetOrderedJobs = jobs
        Exit Function
    End If

    ' 実行順でソート
    Dim arr() As Variant
    ReDim arr(1 To orderedRows.Count)
    Dim i As Long
    Dim k As Long
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

'==============================================================================
' 順序指定のバリデーション
'==============================================================================
Private Function ValidateJobOrder(jobs As Collection) As String
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

    ' 連続性チェック（まずソート）
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
