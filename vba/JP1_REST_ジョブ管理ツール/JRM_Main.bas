Attribute VB_Name = "JRM_Main"
Option Explicit

'==============================================================================
' JP1 REST ジョブ管理ツール - メインモジュール
' エントリーポイント、ツリー操作、ジョブ実行機能を提供
'==============================================================================

'==============================================================================
' ツリー取得（メインエントリポイント）
'==============================================================================
Public Sub RefreshTree()
    On Error GoTo ErrorHandler

    Dim config As Object
    Set config = GetConfig()

    If config Is Nothing Then Exit Sub

    ' パスワード入力
    If config("JP1Password") = "" Then
        config("JP1Password") = InputBox("JP1パスワードを入力してください:", "パスワード入力")
        If config("JP1Password") = "" Then
            MsgBox "パスワードが入力されませんでした。", vbExclamation
            Exit Sub
        End If
    End If

    Application.ScreenUpdating = False
    Application.StatusBar = "ツリーを取得中..."

    ' ツリー表示シートをクリア
    Dim wsTree As Worksheet
    Set wsTree = Worksheets(SHEET_TREE)

    Dim lastRow As Long
    lastRow = wsTree.Cells(wsTree.Rows.Count, COL_UNIT_NAME).End(xlUp).Row
    If lastRow >= ROW_TREE_DATA_START Then
        wsTree.Range(wsTree.Cells(ROW_TREE_DATA_START, 1), wsTree.Cells(lastRow, COL_SELECT)).ClearContents
    End If

    ' REST API でルート配下のユニットを取得
    Dim units As Collection
    Set units = GetUnitList(config, config("RootPath"))

    If units Is Nothing Then
        Application.StatusBar = False
        Application.ScreenUpdating = True
        MsgBox "ユニットの取得に失敗しました。" & vbCrLf & _
               "接続設定を確認してください。", vbExclamation
        Exit Sub
    End If

    ' ツリー表示シートに書き込み
    Dim row As Long
    row = ROW_TREE_DATA_START

    Dim unit As Object
    For Each unit In units
        WriteUnitToSheet wsTree, row, unit, 0
        row = row + 1
    Next unit

    Application.StatusBar = False
    Application.ScreenUpdating = True

    MsgBox "ツリーの取得が完了しました（" & units.Count & " 件）。" & vbCrLf & _
           "[>]をダブルクリックで展開できます。", vbInformation

    wsTree.Activate
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました。" & vbCrLf & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, vbCritical, "VBAエラー"
End Sub

'==============================================================================
' ツリー展開
'==============================================================================
Public Sub ExpandUnit(row As Long)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_TREE)

    Dim unitPath As String
    unitPath = ws.Cells(row, COL_UNIT_PATH).Value

    If unitPath = "" Then Exit Sub

    ' 既に展開済みかチェック
    Dim expandIcon As String
    expandIcon = ws.Cells(row, COL_EXPAND).Value

    If expandIcon = ICON_EXPANDED Then
        ' 既に展開済み → 折りたたみ
        CollapseUnit row
        Exit Sub
    End If

    If expandIcon <> ICON_COLLAPSED Then
        ' 展開可能なユニットではない
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.StatusBar = "展開中..."

    Dim config As Object
    Set config = GetConfig()

    If config Is Nothing Then
        Application.StatusBar = False
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' パスワード入力
    If config("JP1Password") = "" Then
        config("JP1Password") = InputBox("JP1パスワードを入力してください:", "パスワード入力")
        If config("JP1Password") = "" Then
            Application.StatusBar = False
            Application.ScreenUpdating = True
            Exit Sub
        End If
    End If

    ' 子ユニット取得
    Dim children As Collection
    Set children = GetUnitList(config, unitPath)

    If children Is Nothing Or children.Count = 0 Then
        ' 子がいない場合はアイコンを消す
        ws.Cells(row, COL_EXPAND).Value = ""
        Application.StatusBar = False
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' 子ユニットを挿入
    Dim indentLevel As Long
    indentLevel = GetIndentLevel(unitPath) + 1

    Dim insertRow As Long
    insertRow = row + 1

    ' 必要な行数を挿入
    ws.Rows(insertRow & ":" & insertRow + children.Count - 1).Insert Shift:=xlDown

    Dim child As Object
    For Each child In children
        WriteUnitToSheet ws, insertRow, child, indentLevel
        insertRow = insertRow + 1
    Next child

    ' 展開アイコンを更新
    ws.Cells(row, COL_EXPAND).Value = ICON_EXPANDED

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "展開中にエラーが発生しました。" & vbCrLf & _
           Err.Description, vbExclamation
End Sub

'==============================================================================
' ツリー折りたたみ
'==============================================================================
Public Sub CollapseUnit(row As Long)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_TREE)

    Dim expandIcon As String
    expandIcon = ws.Cells(row, COL_EXPAND).Value

    If expandIcon <> ICON_EXPANDED Then Exit Sub

    Application.ScreenUpdating = False

    ' 親のインデントレベルを取得
    Dim parentPath As String
    parentPath = ws.Cells(row, COL_UNIT_PATH).Value

    Dim parentLevel As Long
    parentLevel = GetIndentLevel(parentPath)

    ' 子行を削除（親のレベルより深い行をすべて削除）
    Dim deleteStartRow As Long
    Dim deleteEndRow As Long
    deleteStartRow = row + 1
    deleteEndRow = deleteStartRow - 1

    Dim checkRow As Long
    checkRow = row + 1

    Do While checkRow <= ws.Cells(ws.Rows.Count, COL_UNIT_PATH).End(xlUp).Row
        Dim checkPath As String
        checkPath = ws.Cells(checkRow, COL_UNIT_PATH).Value

        If checkPath = "" Then Exit Do

        Dim checkLevel As Long
        checkLevel = GetIndentLevel(checkPath)

        If checkLevel <= parentLevel Then
            ' 同レベル以上に達したら終了
            Exit Do
        End If

        deleteEndRow = checkRow
        checkRow = checkRow + 1
    Loop

    ' 削除実行
    If deleteEndRow >= deleteStartRow Then
        ws.Rows(deleteStartRow & ":" & deleteEndRow).Delete Shift:=xlUp
    End If

    ' 折りたたみアイコンに変更
    ws.Cells(row, COL_EXPAND).Value = ICON_COLLAPSED

    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
End Sub

'==============================================================================
' 全展開
'==============================================================================
Public Sub ExpandAll()
    MsgBox "全展開機能は現在実装中です。" & vbCrLf & _
           "個別に[>]をダブルクリックして展開してください。", vbInformation
End Sub

'==============================================================================
' 全折りたたみ
'==============================================================================
Public Sub CollapseAll()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_TREE)

    Application.ScreenUpdating = False

    ' ルートレベル以外の行を削除
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_UNIT_PATH).End(xlUp).Row

    If lastRow < ROW_TREE_DATA_START Then
        Application.ScreenUpdating = True
        Exit Sub
    End If

    Dim row As Long
    For row = lastRow To ROW_TREE_DATA_START Step -1
        Dim unitPath As String
        unitPath = ws.Cells(row, COL_UNIT_PATH).Value

        If unitPath <> "" Then
            Dim level As Long
            level = GetIndentLevel(unitPath)

            If level > 0 Then
                ws.Rows(row).Delete Shift:=xlUp
            Else
                ' ルートレベルは折りたたみアイコンに戻す
                Dim unitType As String
                unitType = ws.Cells(row, COL_UNIT_TYPE).Value
                If unitType = TYPE_GROUP Or unitType = TYPE_ROOTNET Or unitType = TYPE_NET Then
                    ws.Cells(row, COL_EXPAND).Value = ICON_COLLAPSED
                End If
            End If
        End If
    Next row

    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
End Sub

'==============================================================================
' ダブルクリックイベントハンドラ
'==============================================================================
Public Sub OnTreeDoubleClick(row As Long, col As Long, ByRef Cancel As Boolean)
    ' データ行以外は無視
    If row < ROW_TREE_DATA_START Then Exit Sub

    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_TREE)

    If col = COL_EXPAND Then
        ' 展開/折りたたみ
        Cancel = True
        Dim expandIcon As String
        expandIcon = ws.Cells(row, COL_EXPAND).Value

        If expandIcon = ICON_COLLAPSED Then
            ExpandUnit row
        ElseIf expandIcon = ICON_EXPANDED Then
            CollapseUnit row
        End If

    ElseIf col = COL_SELECT Then
        ' 選択チェック切り替え
        Cancel = True
        Dim currentValue As String
        currentValue = ws.Cells(row, COL_SELECT).Value

        If currentValue = CHECK_ON Then
            ws.Cells(row, COL_SELECT).Value = CHECK_OFF
        Else
            ws.Cells(row, COL_SELECT).Value = CHECK_ON
        End If
    End If
End Sub

'==============================================================================
' 選択ジョブネット実行
'==============================================================================
Public Sub ExecuteSelectedJobnet()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_TREE)

    ' 選択されたジョブネットを収集
    Dim selectedUnits As Collection
    Set selectedUnits = New Collection

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_UNIT_PATH).End(xlUp).Row

    If lastRow < ROW_TREE_DATA_START Then
        MsgBox "ツリーが空です。先に「ツリー取得」を実行してください。", vbExclamation
        Exit Sub
    End If

    Dim row As Long
    For row = ROW_TREE_DATA_START To lastRow
        If ws.Cells(row, COL_SELECT).Value = CHECK_ON Then
            Dim unitType As String
            unitType = ws.Cells(row, COL_UNIT_TYPE).Value

            ' ジョブネット（ROOTNET/NET）のみ実行可能
            If unitType = TYPE_ROOTNET Or unitType = TYPE_NET Then
                Dim unitInfo As Object
                Set unitInfo = CreateObject("Scripting.Dictionary")
                unitInfo("Path") = ws.Cells(row, COL_UNIT_PATH).Value
                unitInfo("Name") = ws.Cells(row, COL_UNIT_NAME).Value
                unitInfo("Row") = row
                selectedUnits.Add unitInfo
            End If
        End If
    Next row

    If selectedUnits.Count = 0 Then
        MsgBox "実行するジョブネットが選択されていません。" & vbCrLf & _
               "ジョブネットの「選択」列をチェックしてください。" & vbCrLf & vbCrLf & _
               "※ジョブネット（" & TYPE_ROOTNET & "/" & TYPE_NET & "）のみ実行可能です。", vbExclamation
        Exit Sub
    End If

    ' 確認
    Dim msg As String
    msg = "以下の " & selectedUnits.Count & " 件のジョブネットを即時実行しますか？" & vbCrLf & vbCrLf

    Dim unitObj As Object
    Dim cnt As Long
    cnt = 0
    For Each unitObj In selectedUnits
        cnt = cnt + 1
        If cnt <= 5 Then
            msg = msg & "  " & cnt & ". " & unitObj("Path") & vbCrLf
        ElseIf cnt = 6 Then
            msg = msg & "  ... 他 " & (selectedUnits.Count - 5) & " 件" & vbCrLf
        End If
    Next unitObj

    If MsgBox(msg, vbYesNo + vbQuestion, "実行確認") = vbNo Then
        Exit Sub
    End If

    Dim config As Object
    Set config = GetConfig()

    If config Is Nothing Then Exit Sub

    ' パスワード入力
    If config("JP1Password") = "" Then
        config("JP1Password") = InputBox("JP1パスワードを入力してください:", "パスワード入力")
        If config("JP1Password") = "" Then
            Exit Sub
        End If
    End If

    Application.ScreenUpdating = False

    ' 各ジョブネットを実行
    Dim successCount As Long
    Dim failCount As Long
    successCount = 0
    failCount = 0

    For Each unitObj In selectedUnits
        Application.StatusBar = "実行中: " & unitObj("Path")

        Dim execResult As Object
        Set execResult = ExecuteImmediateExec(config, unitObj("Path"))

        If execResult("Success") Then
            successCount = successCount + 1

            ' execIDを取得してシートに反映
            ws.Cells(unitObj("Row"), COL_EXEC_ID).Value = execResult("ExecID")

            ' ログ記録
            WriteLogEntry unitObj("Path"), "即時実行", "成功", execResult("ExecID"), "", ""

            ' 完了待ち設定の場合
            If config("WaitCompletion") = "はい" Then
                Application.StatusBar = "完了待機中: " & unitObj("Path")

                Dim pollResult As Object
                Set pollResult = PollExecutionStatus(config, unitObj("Path"), execResult("ExecID"))

                If pollResult("Success") Then
                    ' 状態を更新
                    ws.Cells(unitObj("Row"), COL_STATUS).Value = pollResult("Status")
                    ws.Cells(unitObj("Row"), COL_START_TIME).Value = pollResult("StartTime")
                    ws.Cells(unitObj("Row"), COL_END_TIME).Value = pollResult("EndTime")

                    ' ログ取得
                    Dim logResult As Object
                    Set logResult = GetExecResultDetails(config, unitObj("Path"), execResult("ExecID"))

                    If logResult("Success") Then
                        WriteLogEntry unitObj("Path"), "ログ取得", "成功", execResult("ExecID"), _
                                      pollResult("StartTime"), pollResult("EndTime")
                    End If
                End If
            End If
        Else
            failCount = failCount + 1
            WriteLogEntry unitObj("Path"), "即時実行", "失敗: " & execResult("ErrorMessage"), "", "", ""
        End If

        ' 選択解除
        ws.Cells(unitObj("Row"), COL_SELECT).Value = CHECK_OFF
    Next unitObj

    Application.StatusBar = False
    Application.ScreenUpdating = True

    MsgBox "実行が完了しました。" & vbCrLf & vbCrLf & _
           "成功: " & successCount & " 件" & vbCrLf & _
           "失敗: " & failCount & " 件", vbInformation

    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "実行中にエラーが発生しました。" & vbCrLf & _
           Err.Description, vbCritical
End Sub

'==============================================================================
' ログ取得（手動）
'==============================================================================
Public Sub GetExecutionLog()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_TREE)

    ' 選択されたユニットでexecIDがあるものを対象
    Dim targetUnits As Collection
    Set targetUnits = New Collection

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_UNIT_PATH).End(xlUp).Row

    If lastRow < ROW_TREE_DATA_START Then
        MsgBox "ツリーが空です。", vbExclamation
        Exit Sub
    End If

    Dim row As Long
    For row = ROW_TREE_DATA_START To lastRow
        If ws.Cells(row, COL_SELECT).Value = CHECK_ON Then
            Dim execID As String
            execID = ws.Cells(row, COL_EXEC_ID).Value

            If execID <> "" Then
                Dim unitInfo As Object
                Set unitInfo = CreateObject("Scripting.Dictionary")
                unitInfo("Path") = ws.Cells(row, COL_UNIT_PATH).Value
                unitInfo("ExecID") = execID
                unitInfo("Row") = row
                targetUnits.Add unitInfo
            End If
        End If
    Next row

    If targetUnits.Count = 0 Then
        MsgBox "ログ取得対象が選択されていません。" & vbCrLf & _
               "execIDがあるユニットを選択してください。", vbExclamation
        Exit Sub
    End If

    Dim config As Object
    Set config = GetConfig()

    If config Is Nothing Then Exit Sub

    ' パスワード入力
    If config("JP1Password") = "" Then
        config("JP1Password") = InputBox("JP1パスワードを入力してください:", "パスワード入力")
        If config("JP1Password") = "" Then
            Exit Sub
        End If
    End If

    Application.ScreenUpdating = False

    Dim successCount As Long
    successCount = 0

    Dim unitObj As Object
    For Each unitObj In targetUnits
        Application.StatusBar = "ログ取得中: " & unitObj("Path")

        Dim logResult As Object
        Set logResult = GetExecResultDetails(config, unitObj("Path"), unitObj("ExecID"))

        If logResult("Success") Then
            successCount = successCount + 1
            WriteLogEntry unitObj("Path"), "ログ取得", "成功", unitObj("ExecID"), "", ""

            ' ログ内容を表示（5MB制限があるので注意）
            If Len(logResult("Details")) > 0 Then
                ' ログをテキストファイルに保存して開く
                SaveAndOpenLog unitObj("Path"), unitObj("ExecID"), logResult("Details")
            End If
        Else
            WriteLogEntry unitObj("Path"), "ログ取得", "失敗", unitObj("ExecID"), "", ""
        End If

        ' 選択解除
        ws.Cells(unitObj("Row"), COL_SELECT).Value = CHECK_OFF
    Next unitObj

    Application.StatusBar = False
    Application.ScreenUpdating = True

    MsgBox "ログ取得が完了しました（" & successCount & " 件）。", vbInformation
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "ログ取得中にエラーが発生しました。" & vbCrLf & _
           Err.Description, vbCritical
End Sub

'==============================================================================
' ログ履歴クリア
'==============================================================================
Public Sub ClearLogHistory()
    If MsgBox("実行ログ履歴をクリアしますか？", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_LOG)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If lastRow >= ROW_LOG_DATA_START Then
        ws.Range(ws.Cells(ROW_LOG_DATA_START, 1), ws.Cells(lastRow, 7)).ClearContents
    End If

    MsgBox "ログ履歴をクリアしました。", vbInformation
End Sub

'==============================================================================
' ユニット行をシートに書き込み
'==============================================================================
Private Sub WriteUnitToSheet(ws As Worksheet, row As Long, unit As Object, indentLevel As Long)
    Dim displayName As String
    displayName = String(indentLevel * 2, " ") & unit("Name")

    Dim typeDisplay As String
    typeDisplay = GetTypeDisplayName(unit("Type"))

    ' 展開可能なユニットには▶を表示
    Dim expandIcon As String
    If typeDisplay = TYPE_GROUP Or typeDisplay = TYPE_ROOTNET Or typeDisplay = TYPE_NET Then
        expandIcon = ICON_COLLAPSED
    Else
        expandIcon = ""
    End If

    ws.Cells(row, COL_EXPAND).Value = expandIcon
    ws.Cells(row, COL_UNIT_NAME).Value = displayName
    ws.Cells(row, COL_UNIT_PATH).Value = unit("Path")
    ws.Cells(row, COL_UNIT_TYPE).Value = typeDisplay
    ws.Cells(row, COL_STATUS).Value = unit("Status")
    ws.Cells(row, COL_EXEC_ID).Value = unit("ExecID")
    ws.Cells(row, COL_START_TIME).Value = unit("StartTime")
    ws.Cells(row, COL_END_TIME).Value = unit("EndTime")
    ws.Cells(row, COL_SELECT).Value = CHECK_OFF
End Sub

'==============================================================================
' ログ記録
'==============================================================================
Private Sub WriteLogEntry(unitPath As String, operation As String, result As String, _
                          execID As String, startTime As String, endTime As String)
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_LOG)

    Dim newRow As Long
    newRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    If newRow < ROW_LOG_DATA_START Then newRow = ROW_LOG_DATA_START

    ws.Cells(newRow, 1).Value = Format(Now, "yyyy/mm/dd hh:nn:ss")
    ws.Cells(newRow, 2).Value = unitPath
    ws.Cells(newRow, 3).Value = operation
    ws.Cells(newRow, 4).Value = result
    ws.Cells(newRow, 5).Value = execID
    ws.Cells(newRow, 6).Value = startTime
    ws.Cells(newRow, 7).Value = endTime

    On Error GoTo 0
End Sub

'==============================================================================
' ログをファイルに保存して開く
'==============================================================================
Private Sub SaveAndOpenLog(unitPath As String, execID As String, logContent As String)
    On Error GoTo ErrorHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' ファイル名を作成（パスからファイル名を生成）
    Dim safeName As String
    safeName = Replace(Replace(unitPath, "/", "_"), "\", "_")
    If Left(safeName, 1) = "_" Then safeName = Mid(safeName, 2)

    Dim timestamp As String
    timestamp = Format(Now, "yyyymmdd_hhnnss")

    Dim logFileName As String
    logFileName = fso.GetSpecialFolder(2) & "\JP1_Log_" & safeName & "_" & timestamp & ".txt"

    ' UTF-8で保存
    Dim utfStream As Object
    Set utfStream = CreateObject("ADODB.Stream")
    utfStream.Type = 2
    utfStream.Charset = "UTF-8"
    utfStream.Open
    utfStream.WriteText "=== JP1 実行ログ ===" & vbCrLf
    utfStream.WriteText "ユニットパス: " & unitPath & vbCrLf
    utfStream.WriteText "execID: " & execID & vbCrLf
    utfStream.WriteText "取得日時: " & Format(Now, "yyyy/mm/dd hh:nn:ss") & vbCrLf
    utfStream.WriteText "==================" & vbCrLf & vbCrLf
    utfStream.WriteText logContent
    utfStream.SaveToFile logFileName, 2
    utfStream.Close
    Set utfStream = Nothing

    ' メモ帳で開く
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    shell.Run "notepad.exe """ & logFileName & """", 1, False

    Exit Sub

ErrorHandler:
    MsgBox "ログファイルの保存に失敗しました。" & vbCrLf & _
           Err.Description, vbExclamation
End Sub

