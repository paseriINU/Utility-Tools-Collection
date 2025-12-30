Option Explicit

'==============================================================================
' JP1 REST ジョブ管理ツール - メインモジュール
'   - REST API呼び出し処理
'   - ツリー表示・展開/折りたたみ
'   - ジョブ実行・ログ取得
'
' 注意: 初期化処理は JP1_REST_Manager_Setup.bas にあります
'       定数はSetupモジュールで Public として定義されています
'==============================================================================

' デバッグモード（設定シートから読み込み）
Private m_DebugMode As Boolean

' ユニット種別の日本語表示名
Private Const TYPE_GROUP As String = "グループ"
Private Const TYPE_ROOTNET As String = "ルートジョブネット"
Private Const TYPE_NET As String = "ネストジョブネット"
Private Const TYPE_JOB As String = "ジョブ"
Private Const TYPE_UNKNOWN As String = "その他"

' 展開アイコン
Private Const ICON_COLLAPSED As String = ">"  ' 折りたたみ状態
Private Const ICON_EXPANDED As String = "v"   ' 展開状態

' チェックボックス文字
Private Const CHECK_ON As String = "v"
Private Const CHECK_OFF As String = ""

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
' REST API: ユニット一覧取得
'==============================================================================
Private Function GetUnitList(config As Object, location As String) As Collection
    On Error GoTo ErrorHandler

    Dim psScript As String
    psScript = BuildStatusesAPIScript(config, location, "NO")

    Dim result As String
    result = ExecutePowerShell(psScript)

    ' 結果をパース
    Set GetUnitList = ParseStatusesResponse(result)
    Exit Function

ErrorHandler:
    Set GetUnitList = Nothing
End Function

'==============================================================================
' REST API: 即時実行登録
'==============================================================================
Private Function ExecuteImmediateExec(config As Object, unitPath As String) As Object
    On Error GoTo ErrorHandler

    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result("Success") = False

    Dim psScript As String
    psScript = BuildImmediateExecAPIScript(config, unitPath)

    Dim output As String
    output = ExecutePowerShell(psScript)

    ' 結果をパース
    If InStr(output, "EXEC_ID:") > 0 Then
        Dim execID As String
        execID = ExtractValue(output, "EXEC_ID:")
        result("Success") = True
        result("ExecID") = Trim(execID)
    Else
        result("Success") = False
        result("ErrorMessage") = "execIDの取得に失敗しました"

        If InStr(output, "API_ERROR:") > 0 Then
            result("ErrorMessage") = ExtractValue(output, "ERROR_MESSAGE:")
        End If
    End If

    Set ExecuteImmediateExec = result
    Exit Function

ErrorHandler:
    Set result = CreateObject("Scripting.Dictionary")
    result("Success") = False
    result("ErrorMessage") = Err.Description
    Set ExecuteImmediateExec = result
End Function

'==============================================================================
' REST API: 実行状態ポーリング
'==============================================================================
Private Function PollExecutionStatus(config As Object, unitPath As String, execID As String) As Object
    On Error GoTo ErrorHandler

    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result("Success") = False

    Dim pollingInterval As Long
    pollingInterval = CLng(config("PollingInterval"))
    If pollingInterval < 1 Then pollingInterval = 5

    Dim timeout As Long
    timeout = CLng(config("Timeout"))

    Dim startTime As Date
    startTime = Now

    Do
        ' API呼び出し
        Dim psScript As String
        psScript = BuildStatusesAPIScriptWithExecID(config, unitPath, execID)

        Dim output As String
        output = ExecutePowerShell(psScript)

        ' 状態を確認
        Dim status As String
        status = ExtractValue(output, "STATUS:")

        result("Status") = status
        result("StartTime") = ExtractValue(output, "START_TIME:")
        result("EndTime") = ExtractValue(output, "END_TIME:")

        ' 終了状態かチェック
        If IsTerminalStatus(status) Then
            result("Success") = True
            Exit Do
        End If

        ' タイムアウトチェック
        If timeout > 0 Then
            If DateDiff("s", startTime, Now) > timeout Then
                result("Success") = False
                result("ErrorMessage") = "タイムアウト"
                Exit Do
            End If
        End If

        ' 待機
        Application.Wait Now + TimeSerial(0, 0, pollingInterval)
        DoEvents
    Loop

    Set PollExecutionStatus = result
    Exit Function

ErrorHandler:
    Set result = CreateObject("Scripting.Dictionary")
    result("Success") = False
    result("ErrorMessage") = Err.Description
    Set PollExecutionStatus = result
End Function

'==============================================================================
' REST API: 実行結果詳細取得
'==============================================================================
Private Function GetExecResultDetails(config As Object, unitPath As String, execID As String) As Object
    On Error GoTo ErrorHandler

    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result("Success") = False

    Dim psScript As String
    psScript = BuildExecResultDetailsAPIScript(config, unitPath, execID)

    Dim output As String
    output = ExecutePowerShell(psScript)

    If InStr(output, "RESULT_DETAILS_START") > 0 Then
        Dim startPos As Long
        Dim endPos As Long
        startPos = InStr(output, "RESULT_DETAILS_START") + Len("RESULT_DETAILS_START") + 2
        endPos = InStr(output, "RESULT_DETAILS_END")

        If endPos > startPos Then
            result("Details") = Mid(output, startPos, endPos - startPos)
            result("Success") = True
        End If
    Else
        result("Success") = False
        result("ErrorMessage") = "ログの取得に失敗しました"
    End If

    Set GetExecResultDetails = result
    Exit Function

ErrorHandler:
    Set result = CreateObject("Scripting.Dictionary")
    result("Success") = False
    result("ErrorMessage") = Err.Description
    Set GetExecResultDetails = result
End Function

'==============================================================================
' PowerShellスクリプト生成: statuses API
'==============================================================================
Private Function BuildStatusesAPIScript(config As Object, location As String, searchLowerUnits As String) As String
    Dim script As String

    script = BuildAPIHeader(config)

    ' URLエンコード
    script = script & "$location = '" & EscapePSString(location) & "'" & vbCrLf
    script = script & "$encodedLocation = [System.Uri]::EscapeDataString($location)" & vbCrLf
    script = script & vbCrLf

    ' URL構築
    script = script & "$url = ""${baseUrl}/objects/statuses?mode=search""" & vbCrLf
    script = script & "$url += ""&manager=${managerHost}""" & vbCrLf
    script = script & "$url += ""&serviceName=${schedulerService}""" & vbCrLf
    script = script & "$url += ""&location=${encodedLocation}""" & vbCrLf
    script = script & "$url += ""&searchLowerUnits=" & searchLowerUnits & """" & vbCrLf
    script = script & "$url += ""&searchTarget=DEFINITION_AND_STATUS""" & vbCrLf
    script = script & vbCrLf

    ' API呼び出し
    script = script & "try {" & vbCrLf
    script = script & "  $response = Invoke-WebRequest -Uri $url -Method GET -Headers $headers -TimeoutSec 30 -UseBasicParsing" & vbCrLf
    script = script & "  $responseBytes = $response.RawContentStream.ToArray()" & vbCrLf
    script = script & "  $responseText = [System.Text.Encoding]::UTF8.GetString($responseBytes)" & vbCrLf
    script = script & "  $json = $responseText | ConvertFrom-Json" & vbCrLf
    script = script & vbCrLf
    script = script & "  Write-Output 'JSON_START'" & vbCrLf
    script = script & "  Write-Output ($json | ConvertTo-Json -Depth 10 -Compress)" & vbCrLf
    script = script & "  Write-Output 'JSON_END'" & vbCrLf
    script = script & "} catch {" & vbCrLf
    script = script & "  Write-Output ""API_ERROR:$($_.Exception.Response.StatusCode.Value__)""" & vbCrLf
    script = script & "  Write-Output ""ERROR_MESSAGE:$($_.Exception.Message)""" & vbCrLf
    script = script & "}" & vbCrLf

    BuildStatusesAPIScript = script
End Function

'==============================================================================
' PowerShellスクリプト生成: statuses API (execID指定)
'==============================================================================
Private Function BuildStatusesAPIScriptWithExecID(config As Object, unitPath As String, execID As String) As String
    Dim script As String

    script = BuildAPIHeader(config)

    ' パスを分解
    script = script & "$unitPath = '" & EscapePSString(unitPath) & "'" & vbCrLf
    script = script & "$lastSlash = $unitPath.LastIndexOf('/')" & vbCrLf
    script = script & "$parentPath = $unitPath.Substring(0, $lastSlash)" & vbCrLf
    script = script & "if (-not $parentPath) { $parentPath = '/' }" & vbCrLf
    script = script & "$unitName = $unitPath.Substring($lastSlash + 1)" & vbCrLf
    script = script & "$encodedParent = [System.Uri]::EscapeDataString($parentPath)" & vbCrLf
    script = script & "$encodedName = [System.Uri]::EscapeDataString($unitName)" & vbCrLf
    script = script & vbCrLf

    ' URL構築
    script = script & "$url = ""${baseUrl}/objects/statuses?mode=search""" & vbCrLf
    script = script & "$url += ""&manager=${managerHost}""" & vbCrLf
    script = script & "$url += ""&serviceName=${schedulerService}""" & vbCrLf
    script = script & "$url += ""&location=${encodedParent}""" & vbCrLf
    script = script & "$url += ""&searchLowerUnits=NO""" & vbCrLf
    script = script & "$url += ""&searchTarget=DEFINITION_AND_STATUS""" & vbCrLf
    script = script & "$url += ""&unitName=${encodedName}""" & vbCrLf
    script = script & "$url += ""&unitNameMatchMethods=EQ""" & vbCrLf
    script = script & "$url += ""&generation=EXECID""" & vbCrLf
    script = script & "$url += ""&execID=" & execID & """" & vbCrLf
    script = script & vbCrLf

    ' API呼び出し
    script = script & "try {" & vbCrLf
    script = script & "  $response = Invoke-WebRequest -Uri $url -Method GET -Headers $headers -TimeoutSec 30 -UseBasicParsing" & vbCrLf
    script = script & "  $responseBytes = $response.RawContentStream.ToArray()" & vbCrLf
    script = script & "  $responseText = [System.Text.Encoding]::UTF8.GetString($responseBytes)" & vbCrLf
    script = script & "  $json = $responseText | ConvertFrom-Json" & vbCrLf
    script = script & vbCrLf
    script = script & "  if ($json.statuses -and $json.statuses.Count -gt 0) {" & vbCrLf
    script = script & "    $unit = $json.statuses[0]" & vbCrLf
    script = script & "    $status = if ($unit.unitStatus) { $unit.unitStatus.status } else { 'N/A' }" & vbCrLf
    script = script & "    $startTime = if ($unit.unitStatus) { $unit.unitStatus.startTime } else { '' }" & vbCrLf
    script = script & "    $endTime = if ($unit.unitStatus) { $unit.unitStatus.endTime } else { '' }" & vbCrLf
    script = script & "    Write-Output ""STATUS:$status""" & vbCrLf
    script = script & "    Write-Output ""START_TIME:$startTime""" & vbCrLf
    script = script & "    Write-Output ""END_TIME:$endTime""" & vbCrLf
    script = script & "  } else {" & vbCrLf
    script = script & "    Write-Output 'STATUS:NOT_FOUND'" & vbCrLf
    script = script & "  }" & vbCrLf
    script = script & "} catch {" & vbCrLf
    script = script & "  Write-Output ""API_ERROR:$($_.Exception.Response.StatusCode.Value__)""" & vbCrLf
    script = script & "  Write-Output ""ERROR_MESSAGE:$($_.Exception.Message)""" & vbCrLf
    script = script & "}" & vbCrLf

    BuildStatusesAPIScriptWithExecID = script
End Function

'==============================================================================
' PowerShellスクリプト生成: 即時実行登録 API
'==============================================================================
Private Function BuildImmediateExecAPIScript(config As Object, unitPath As String) As String
    Dim script As String

    script = BuildAPIHeader(config)

    ' URLエンコード
    script = script & "$unitPath = '" & EscapePSString(unitPath) & "'" & vbCrLf
    script = script & "$encodedPath = [System.Uri]::EscapeDataString($unitPath)" & vbCrLf
    script = script & vbCrLf

    ' URL構築
    script = script & "$url = ""${baseUrl}/objects/definitions/${encodedPath}/actions/registerImmediateExec/invoke""" & vbCrLf
    script = script & "$url += ""?manager=${managerHost}&serviceName=${schedulerService}""" & vbCrLf
    script = script & vbCrLf

    ' API呼び出し（POST）
    script = script & "try {" & vbCrLf
    script = script & "  $response = Invoke-WebRequest -Uri $url -Method POST -Headers $headers -TimeoutSec 30 -UseBasicParsing" & vbCrLf
    script = script & "  $responseBytes = $response.RawContentStream.ToArray()" & vbCrLf
    script = script & "  $responseText = [System.Text.Encoding]::UTF8.GetString($responseBytes)" & vbCrLf
    script = script & "  $json = $responseText | ConvertFrom-Json" & vbCrLf
    script = script & "  Write-Output ""EXEC_ID:$($json.execID)""" & vbCrLf
    script = script & "} catch {" & vbCrLf
    script = script & "  Write-Output ""API_ERROR:$($_.Exception.Response.StatusCode.Value__)""" & vbCrLf
    script = script & "  Write-Output ""ERROR_MESSAGE:$($_.Exception.Message)""" & vbCrLf
    script = script & "}" & vbCrLf

    BuildImmediateExecAPIScript = script
End Function

'==============================================================================
' PowerShellスクリプト生成: 実行結果詳細取得 API
'==============================================================================
Private Function BuildExecResultDetailsAPIScript(config As Object, unitPath As String, execID As String) As String
    Dim script As String

    script = BuildAPIHeader(config)

    ' URLエンコード
    script = script & "$unitPath = '" & EscapePSString(unitPath) & "'" & vbCrLf
    script = script & "$execID = '" & execID & "'" & vbCrLf
    script = script & "$encodedPath = [System.Uri]::EscapeDataString($unitPath)" & vbCrLf
    script = script & vbCrLf

    ' URL構築
    script = script & "$url = ""${baseUrl}/objects/statuses/${encodedPath}:${execID}/actions/execResultDetails/invoke""" & vbCrLf
    script = script & "$url += ""?manager=${managerHost}&serviceName=${schedulerService}""" & vbCrLf
    script = script & vbCrLf

    ' API呼び出し
    script = script & "try {" & vbCrLf
    script = script & "  $response = Invoke-WebRequest -Uri $url -Method GET -Headers $headers -TimeoutSec 60 -UseBasicParsing" & vbCrLf
    script = script & "  $responseBytes = $response.RawContentStream.ToArray()" & vbCrLf
    script = script & "  $responseText = [System.Text.Encoding]::UTF8.GetString($responseBytes)" & vbCrLf
    script = script & "  $json = $responseText | ConvertFrom-Json" & vbCrLf
    script = script & "  Write-Output 'RESULT_DETAILS_START'" & vbCrLf
    script = script & "  Write-Output $json.execResultDetails" & vbCrLf
    script = script & "  Write-Output 'RESULT_DETAILS_END'" & vbCrLf
    script = script & "  Write-Output ""ALL:$($json.all)""" & vbCrLf
    script = script & "} catch {" & vbCrLf
    script = script & "  Write-Output ""API_ERROR:$($_.Exception.Response.StatusCode.Value__)""" & vbCrLf
    script = script & "  Write-Output ""ERROR_MESSAGE:$($_.Exception.Message)""" & vbCrLf
    script = script & "}" & vbCrLf

    BuildExecResultDetailsAPIScript = script
End Function

'==============================================================================
' PowerShellスクリプト共通ヘッダー
'==============================================================================
Private Function BuildAPIHeader(config As Object) As String
    Dim script As String

    ' UTF-8エンコーディング
    script = "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8" & vbCrLf
    script = script & "$OutputEncoding = [System.Text.Encoding]::UTF8" & vbCrLf
    script = script & vbCrLf

    ' 接続設定
    Dim protocol As String
    If config("UseHttps") = "はい" Then
        protocol = "https"
    Else
        protocol = "http"
    End If

    script = script & "$protocol = '" & protocol & "'" & vbCrLf
    script = script & "$webConsoleHost = '" & config("WebConsoleHost") & "'" & vbCrLf
    script = script & "$webConsolePort = '" & config("WebConsolePort") & "'" & vbCrLf
    script = script & "$managerHost = '" & config("ManagerHost") & "'" & vbCrLf
    script = script & "$schedulerService = '" & config("SchedulerService") & "'" & vbCrLf
    script = script & "$baseUrl = ""${protocol}://${webConsoleHost}:${webConsolePort}/ajs/api/v1""" & vbCrLf
    script = script & vbCrLf

    ' 認証ヘッダー
    script = script & "$authString = '" & config("JP1User") & ":" & config("JP1Password") & "'" & vbCrLf
    script = script & "$authBytes = [System.Text.Encoding]::UTF8.GetBytes($authString)" & vbCrLf
    script = script & "$authBase64 = [System.Convert]::ToBase64String($authBytes)" & vbCrLf
    script = script & "$headers = @{ 'Accept-Language' = 'ja'; 'X-AJS-Authorization' = $authBase64 }" & vbCrLf
    script = script & vbCrLf

    ' HTTPS設定
    If config("UseHttps") = "はい" Then
        script = script & "Add-Type @""" & vbCrLf
        script = script & "using System.Net;" & vbCrLf
        script = script & "using System.Security.Cryptography.X509Certificates;" & vbCrLf
        script = script & "public class TrustAllCertsPolicy : ICertificatePolicy {" & vbCrLf
        script = script & "    public bool CheckValidationResult(ServicePoint sp, X509Certificate cert, WebRequest req, int problem) { return true; }" & vbCrLf
        script = script & "}" & vbCrLf
        script = script & """@" & vbCrLf
        script = script & "[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy" & vbCrLf
        script = script & "[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12" & vbCrLf
        script = script & vbCrLf
    End If

    BuildAPIHeader = script
End Function

'==============================================================================
' PowerShell実行
'==============================================================================
Private Function ExecutePowerShell(script As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim tempFolder As String
    tempFolder = fso.GetSpecialFolder(2) ' Temp folder

    Dim timestamp As String
    timestamp = Format(Now, "yyyymmddhhnnss") & "_" & Int(Rnd * 10000)

    Dim scriptPath As String
    scriptPath = tempFolder & "\jp1rest_" & timestamp & ".ps1"

    Dim outputPath As String
    outputPath = tempFolder & "\jp1rest_output_" & timestamp & ".txt"

    ' ADODB.Streamを使用してUTF-8（BOM付き）で保存
    Dim utfStream As Object
    Set utfStream = CreateObject("ADODB.Stream")
    utfStream.Type = 2 ' adTypeText
    utfStream.Charset = "UTF-8"
    utfStream.Open
    utfStream.WriteText script
    utfStream.SaveToFile scriptPath, 2
    utfStream.Close
    Set utfStream = Nothing

    ' PowerShell実行
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")

    Dim cmd As String
    cmd = "powershell -NoProfile -ExecutionPolicy Bypass -Command ""& {" & _
          "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; " & _
          "& '" & scriptPath & "'" & _
          "}"" > """ & outputPath & """ 2>&1"

    ' デバッグモード: 1 = 表示, 0 = 非表示
    Dim windowStyle As Long
    If m_DebugMode Then
        windowStyle = 1  ' 表示
    Else
        windowStyle = 0  ' 非表示
    End If

    shell.Run cmd, windowStyle, True

    ' 結果ファイルを読み込む
    Dim output As String
    output = ""

    If fso.FileExists(outputPath) Then
        Set utfStream = CreateObject("ADODB.Stream")
        utfStream.Type = 2
        utfStream.Charset = "UTF-8"
        utfStream.Open
        utfStream.LoadFromFile outputPath

        If Not utfStream.EOS Then
            output = utfStream.ReadText
        End If

        utfStream.Close
        Set utfStream = Nothing

        ' デバッグモードでない場合は出力ファイル削除
        If Not m_DebugMode Then
            On Error Resume Next
            fso.DeleteFile outputPath
            On Error GoTo 0
        Else
            ' デバッグモード: ファイルパスをデバッグ出力
            Debug.Print "API Output File: " & outputPath
        End If
    End If

    ' デバッグモードでない場合はスクリプトファイル削除
    If Not m_DebugMode Then
        On Error Resume Next
        fso.DeleteFile scriptPath
        On Error GoTo 0
    Else
        ' デバッグモード: ファイルパスをデバッグ出力
        Debug.Print "Script File: " & scriptPath
        Debug.Print "API Response Length: " & Len(output)
        If Len(output) > 0 Then
            Debug.Print "API Response (first 500 chars): " & Left(output, 500)
        End If
    End If

    ExecutePowerShell = output
End Function

'==============================================================================
' statuses APIレスポンスパース
'==============================================================================
Private Function ParseStatusesResponse(response As String) As Collection
    On Error GoTo ErrorHandler

    Set ParseStatusesResponse = New Collection

    ' JSON部分を抽出
    If InStr(response, "JSON_START") = 0 Then
        Exit Function
    End If

    Dim startPos As Long
    Dim endPos As Long
    startPos = InStr(response, "JSON_START") + Len("JSON_START") + 2
    endPos = InStr(response, "JSON_END")

    If endPos <= startPos Then
        Exit Function
    End If

    Dim jsonStr As String
    jsonStr = Trim(Mid(response, startPos, endPos - startPos))

    ' 簡易JSONパース（statusesの各要素を抽出）
    ' 注: VBAには標準のJSONパーサーがないため、簡易的にパースします

    Dim units As Collection
    Set units = New Collection

    ' statuses配列を探す
    Dim statusesStart As Long
    statusesStart = InStr(jsonStr, """statuses"":")

    If statusesStart = 0 Then
        Set ParseStatusesResponse = units
        Exit Function
    End If

    ' 各ユニットを抽出（簡易パース）
    Dim pos As Long
    pos = statusesStart

    Do
        ' "definition"を探す
        Dim defStart As Long
        defStart = InStr(pos, jsonStr, """definition"":")

        If defStart = 0 Then Exit Do

        ' unitNameを抽出
        Dim unitNameStart As Long
        unitNameStart = InStr(defStart, jsonStr, """unitName"":""")

        If unitNameStart = 0 Then Exit Do

        unitNameStart = unitNameStart + Len("""unitName"":""")
        Dim unitNameEnd As Long
        unitNameEnd = InStr(unitNameStart, jsonStr, """")

        Dim unitName As String
        unitName = Mid(jsonStr, unitNameStart, unitNameEnd - unitNameStart)

        ' simpleUnitNameを抽出
        Dim simpleNameStart As Long
        simpleNameStart = InStr(defStart, jsonStr, """simpleUnitName"":""")

        Dim simpleName As String
        If simpleNameStart > 0 And simpleNameStart < defStart + 500 Then
            simpleNameStart = simpleNameStart + Len("""simpleUnitName"":""")
            Dim simpleNameEnd As Long
            simpleNameEnd = InStr(simpleNameStart, jsonStr, """")
            simpleName = Mid(jsonStr, simpleNameStart, simpleNameEnd - simpleNameStart)
        Else
            simpleName = unitName
        End If

        ' unitTypeを抽出
        Dim unitTypeStart As Long
        unitTypeStart = InStr(defStart, jsonStr, """unitType"":""")

        Dim unitType As String
        If unitTypeStart > 0 And unitTypeStart < defStart + 500 Then
            unitTypeStart = unitTypeStart + Len("""unitType"":""")
            Dim unitTypeEnd As Long
            unitTypeEnd = InStr(unitTypeStart, jsonStr, """")
            unitType = Mid(jsonStr, unitTypeStart, unitTypeEnd - unitTypeStart)
        Else
            unitType = "UNKNOWN"
        End If

        ' unitStatusを探す
        Dim statusStart As Long
        statusStart = InStr(defStart, jsonStr, """unitStatus"":")

        Dim execID As String
        Dim status As String
        Dim startTime As String
        Dim endTime As String

        execID = ""
        status = ""
        startTime = ""
        endTime = ""

        If statusStart > 0 And statusStart < defStart + 2000 Then
            ' execIDを抽出
            Dim execIDStart As Long
            execIDStart = InStr(statusStart, jsonStr, """execID"":""")
            If execIDStart > 0 And execIDStart < statusStart + 500 Then
                execIDStart = execIDStart + Len("""execID"":""")
                Dim execIDEnd As Long
                execIDEnd = InStr(execIDStart, jsonStr, """")
                execID = Mid(jsonStr, execIDStart, execIDEnd - execIDStart)
            End If

            ' statusを抽出
            Dim statusValStart As Long
            statusValStart = InStr(statusStart, jsonStr, """status"":""")
            If statusValStart > 0 And statusValStart < statusStart + 500 Then
                statusValStart = statusValStart + Len("""status"":""")
                Dim statusValEnd As Long
                statusValEnd = InStr(statusValStart, jsonStr, """")
                status = Mid(jsonStr, statusValStart, statusValEnd - statusValStart)
            End If

            ' startTimeを抽出
            Dim startTimeStart As Long
            startTimeStart = InStr(statusStart, jsonStr, """startTime"":""")
            If startTimeStart > 0 And startTimeStart < statusStart + 800 Then
                startTimeStart = startTimeStart + Len("""startTime"":""")
                Dim startTimeEnd As Long
                startTimeEnd = InStr(startTimeStart, jsonStr, """")
                startTime = Mid(jsonStr, startTimeStart, startTimeEnd - startTimeStart)
            End If

            ' endTimeを抽出
            Dim endTimeStart As Long
            endTimeStart = InStr(statusStart, jsonStr, """endTime"":""")
            If endTimeStart > 0 And endTimeStart < statusStart + 1000 Then
                endTimeStart = endTimeStart + Len("""endTime"":""")
                Dim endTimeEnd As Long
                endTimeEnd = InStr(endTimeStart, jsonStr, """")
                endTime = Mid(jsonStr, endTimeStart, endTimeEnd - endTimeStart)
            End If
        End If

        ' ユニット情報を作成
        Dim unitInfo As Object
        Set unitInfo = CreateObject("Scripting.Dictionary")
        unitInfo("Path") = unitName
        unitInfo("Name") = simpleName
        unitInfo("Type") = unitType
        unitInfo("ExecID") = execID
        unitInfo("Status") = status
        unitInfo("StartTime") = startTime
        unitInfo("EndTime") = endTime

        units.Add unitInfo

        pos = unitNameEnd + 1
    Loop

    Set ParseStatusesResponse = units
    Exit Function

ErrorHandler:
    Set ParseStatusesResponse = New Collection
End Function

'==============================================================================
' ユーティリティ関数
'==============================================================================

' 設定取得
Private Function GetConfig() As Object
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_SETTINGS)

    Dim config As Object
    Set config = CreateObject("Scripting.Dictionary")

    config("WebConsoleHost") = Trim(CStr(ws.Cells(ROW_WEB_CONSOLE_HOST, COL_SETTING_VALUE).Value))
    config("WebConsolePort") = Trim(CStr(ws.Cells(ROW_WEB_CONSOLE_PORT, COL_SETTING_VALUE).Value))
    config("UseHttps") = Trim(CStr(ws.Cells(ROW_USE_HTTPS, COL_SETTING_VALUE).Value))
    config("ManagerHost") = Trim(CStr(ws.Cells(ROW_MANAGER_HOST, COL_SETTING_VALUE).Value))
    config("SchedulerService") = Trim(CStr(ws.Cells(ROW_SCHEDULER_SERVICE, COL_SETTING_VALUE).Value))
    config("JP1User") = Trim(CStr(ws.Cells(ROW_JP1_USER, COL_SETTING_VALUE).Value))
    config("JP1Password") = Trim(CStr(ws.Cells(ROW_JP1_PASSWORD, COL_SETTING_VALUE).Value))
    config("RootPath") = Trim(CStr(ws.Cells(ROW_ROOT_PATH, COL_SETTING_VALUE).Value))
    config("WaitCompletion") = Trim(CStr(ws.Cells(ROW_WAIT_COMPLETION, COL_SETTING_VALUE).Value))
    config("PollingInterval") = ws.Cells(ROW_POLLING_INTERVAL, COL_SETTING_VALUE).Value
    config("Timeout") = ws.Cells(ROW_TIMEOUT, COL_SETTING_VALUE).Value
    config("DebugMode") = Trim(CStr(ws.Cells(ROW_DEBUG_MODE, COL_SETTING_VALUE).Value))

    ' デバッグモード設定をモジュール変数に保存
    m_DebugMode = (config("DebugMode") = "はい")

    ' 必須項目チェック
    If config("WebConsoleHost") = "" Or config("SchedulerService") = "" Or config("JP1User") = "" Then
        MsgBox "接続設定が不完全です。" & vbCrLf & _
               "設定シートで必須項目を入力してください。", vbExclamation
        Set GetConfig = Nothing
        Exit Function
    End If

    Set GetConfig = config
    Exit Function

ErrorHandler:
    MsgBox "設定の取得に失敗しました。" & vbCrLf & Err.Description, vbCritical
    Set GetConfig = Nothing
End Function

' インデントレベル取得
Private Function GetIndentLevel(unitPath As String) As Long
    If unitPath = "" Or unitPath = "/" Then
        GetIndentLevel = 0
        Exit Function
    End If

    ' パスの "/" の数 - 1 がインデントレベル
    Dim slashCount As Long
    slashCount = Len(unitPath) - Len(Replace(unitPath, "/", ""))

    GetIndentLevel = slashCount - 1
    If GetIndentLevel < 0 Then GetIndentLevel = 0
End Function

' ユニットタイプの日本語表示名を取得
Private Function GetTypeDisplayName(unitType As String) As String
    Select Case UCase(unitType)
        Case "GROUP"
            GetTypeDisplayName = TYPE_GROUP
        Case "ROOTNET"
            GetTypeDisplayName = TYPE_ROOTNET
        Case "NET"
            GetTypeDisplayName = TYPE_NET
        Case Else
            If InStr(UCase(unitType), "JOB") > 0 Then
                GetTypeDisplayName = TYPE_JOB
            Else
                GetTypeDisplayName = TYPE_UNKNOWN
            End If
    End Select
End Function

' ユニット行をシートに書き込み
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

' PowerShell文字列エスケープ
Private Function EscapePSString(s As String) As String
    EscapePSString = Replace(Replace(s, "'", "''"), "`", "``")
End Function

' 値抽出
Private Function ExtractValue(text As String, prefix As String) As String
    Dim startPos As Long
    startPos = InStr(text, prefix)

    If startPos = 0 Then
        ExtractValue = ""
        Exit Function
    End If

    startPos = startPos + Len(prefix)

    Dim endPos As Long
    endPos = InStr(startPos, text, vbCrLf)

    If endPos = 0 Then
        endPos = InStr(startPos, text, vbLf)
    End If

    If endPos = 0 Then
        endPos = Len(text) + 1
    End If

    ExtractValue = Trim(Mid(text, startPos, endPos - startPos))
End Function

' 終了状態かどうかをチェック
Private Function IsTerminalStatus(status As String) As Boolean
    Select Case UCase(status)
        Case "NORMAL", "WARNING", "ABNORMAL", "KILLED", "BYPASS", "NOTRUN", "END", _
             "NEST_END_NORMAL", "NEST_END_WARNING", "NEST_END_ABNORMAL"
            IsTerminalStatus = True
        Case Else
            IsTerminalStatus = False
    End Select
End Function

' ログ記録
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

' ログをファイルに保存して開く
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
