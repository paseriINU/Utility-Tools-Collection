Attribute VB_Name = "JM_Setup"
Option Explicit

'==============================================================================
' JP1 ジョブ管理ツール - セットアップモジュール
' シート作成・フォーマット処理を提供
' ※このモジュールは初期化後に削除しても動作します（定数はJM_Configに定義）
'==============================================================================

'==============================================================================
' 初期化（メインエントリポイント）
'==============================================================================
Public Sub InitializeJP1Manager()
    Application.ScreenUpdating = False

    ' シート作成
    CreateSheet SHEET_SETTINGS
    CreateSheet SHEET_JOBLIST
    CreateSheet SHEET_LOG

    ' 設定シートのフォーマット
    FormatSettingsSheet

    ' ジョブ一覧シートのフォーマット
    FormatJobListSheet

    ' ログシートのフォーマット
    FormatLogSheet

    ' 設定シートをアクティブに
    Worksheets(SHEET_SETTINGS).Activate

    Application.ScreenUpdating = True

    MsgBox "初期化が完了しました。" & vbCrLf & vbCrLf & _
           "1. 設定シートで接続設定を入力してください" & vbCrLf & _
           "2. 「ジョブ一覧取得」ボタンでジョブを取得" & vbCrLf & _
           "3. ジョブ一覧シートで順序を設定" & vbCrLf & _
           "4. 「選択ジョブ実行」ボタンで実行", _
           vbInformation, "JP1 ジョブ管理ツール"
End Sub

'==============================================================================
' シート作成
'==============================================================================
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
' 設定シートのフォーマット
'==============================================================================
Private Sub FormatSettingsSheet()
    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_SETTINGS)

    ws.Cells.Clear

    ' タイトル
    With ws.Range("A1:F1")
        .Merge
        .Value = "JP1 ジョブ管理ツール - 接続設定"
        .Font.Size = 16
        .Font.Bold = True
        .Interior.Color = RGB(0, 112, 192)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .RowHeight = 30
    End With

    ' 説明
    ws.Range("A2").Value = "JP1サーバに接続してジョブネット一覧を取得し、選択したジョブを実行します。"

    ' ボタン追加
    AddButton ws, 20, 55, 130, 32, "GetGroupList", "グループ名取得", RGB(91, 155, 213)
    AddButton ws, 160, 55, 130, 32, "GetJobList", "ジョブ一覧取得", RGB(0, 112, 192)

    ' 設定セクション
    ws.Range("A6").Value = "■ 接続設定"
    ws.Range("A6").Font.Bold = True

    ws.Cells(ROW_EXEC_MODE, 1).Value = "実行モード"
    ws.Cells(ROW_EXEC_MODE, COL_SETTING_VALUE).Value = "リモート"
    AddDropdown ws, ws.Cells(ROW_EXEC_MODE, COL_SETTING_VALUE), "ローカル,リモート"
    ws.Cells(ROW_EXEC_MODE, 4).Value = "※ローカル: このPCのJP1を使用、リモート: WinRM経由で接続"
    ws.Cells(ROW_EXEC_MODE, 4).Font.Color = RGB(128, 128, 128)

    ws.Range("A8").Value = "【リモート接続設定】（ローカルモード時は不要）"
    ws.Range("A8").Font.Color = RGB(128, 128, 128)

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

    ws.Cells(ROW_SCHEDULER_SERVICE, 1).Value = "スケジューラーサービス"
    ws.Cells(ROW_SCHEDULER_SERVICE, COL_SETTING_VALUE).Value = "AJSROOT1"
    ws.Cells(ROW_SCHEDULER_SERVICE, 4).Value = "※JP1/AJS3のスケジューラーサービス名（例: AJSROOT1）"
    ws.Cells(ROW_SCHEDULER_SERVICE, 4).Font.Color = RGB(128, 128, 128)

    ws.Cells(ROW_ROOT_PATH, 1).Value = "取得パス"
    ws.Cells(ROW_ROOT_PATH, COL_SETTING_VALUE).Value = "/"
    ws.Cells(ROW_ROOT_PATH, 4).Value = "※「グループ名取得」でリスト更新（例: / または /グループ名）"
    ws.Cells(ROW_ROOT_PATH, 4).Font.Color = RGB(128, 128, 128)

    ' 実行設定セクション
    ws.Range("A16").Value = "■ 実行設定"
    ws.Range("A16").Font.Bold = True

    ws.Cells(ROW_WAIT_COMPLETION, 1).Value = "完了待ち"
    ws.Cells(ROW_WAIT_COMPLETION, COL_SETTING_VALUE).Value = "はい"
    AddDropdown ws, ws.Cells(ROW_WAIT_COMPLETION, COL_SETTING_VALUE), "はい,いいえ"

    ws.Cells(ROW_TIMEOUT, 1).Value = "タイムアウト（秒）"
    ws.Cells(ROW_TIMEOUT, COL_SETTING_VALUE).Value = 0
    ws.Cells(ROW_TIMEOUT, 4).Value = "※0=無制限"
    ws.Cells(ROW_TIMEOUT, 4).Font.Color = RGB(128, 128, 128)

    ws.Cells(ROW_POLLING_INTERVAL, 1).Value = "状態確認間隔（秒）"
    ws.Cells(ROW_POLLING_INTERVAL, COL_SETTING_VALUE).Value = 10

    ' 使い方セクション
    ws.Range("A21").Value = "■ 使い方"
    ws.Range("A21").Font.Bold = True

    ws.Range("A22").Value = "1. 上記の接続設定・実行設定を入力します"
    ws.Range("A23").Value = "2. 「グループ名取得」で取得パスのリストを更新できます（任意）"
    ws.Range("A24").Value = "3. 「ジョブ一覧取得」ボタンをクリックしてジョブネット一覧を取得します"
    ws.Range("A25").Value = "4. ジョブ一覧シートで実行するジョブの「順序」列に数字（1, 2, 3...）を入力します"
    ws.Range("A26").Value = "5. 「選択ジョブ実行」ボタンをクリックしてジョブを順番に実行します"
    ws.Range("A27").Value = "6. 実行結果は実行ログシートに記録されます"

    ws.Range("A29").Value = "■ 動作説明"
    ws.Range("A29").Font.Bold = True

    ws.Range("A30").Value = "・ジョブが保留中の場合、実行時に自動で保留解除されます"
    ws.Range("A31").Value = "・完了待ち「はい」の場合、ジョブ終了まで待機して結果を取得します"
    ws.Range("A32").Value = "・異常終了または警告終了した場合、後続のジョブは実行されません"
    ws.Range("A33").Value = "・実行ログにはジョブごとの開始・終了時刻、結果、ログパスが記録されます"
    ws.Range("A34").Value = "・警告・異常終了時はJP1サーバ上の標準エラーログも取得されます"

    ' 列幅調整
    ws.Columns("A").ColumnWidth = 20
    ws.Columns("B").ColumnWidth = 5
    ws.Columns("C").ColumnWidth = 30
    ws.Columns("D").ColumnWidth = 40

    ' 入力セルの書式
    Dim settingCells As Variant
    settingCells = Array(ROW_EXEC_MODE, ROW_JP1_SERVER, ROW_REMOTE_USER, ROW_REMOTE_PASSWORD, _
                         ROW_JP1_USER, ROW_JP1_PASSWORD, ROW_SCHEDULER_SERVICE, ROW_ROOT_PATH, _
                         ROW_WAIT_COMPLETION, ROW_TIMEOUT, ROW_POLLING_INTERVAL)
    Dim r As Variant
    For Each r In settingCells
        With ws.Cells(CLng(r), COL_SETTING_VALUE)
            .Interior.Color = RGB(255, 255, 204)
            .Borders.LineStyle = xlContinuous
        End With
    Next r
End Sub

'==============================================================================
' ジョブ一覧シートのフォーマット
'==============================================================================
Private Sub FormatJobListSheet()
    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_JOBLIST)

    ws.Cells.Clear

    ' タイトル
    With ws.Range("A1:N1")
        .Merge
        .Value = "ジョブネット一覧"
        .Font.Size = 14
        .Font.Bold = True
        .Interior.Color = RGB(0, 176, 80)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .RowHeight = 25
    End With

    ' ボタン追加
    AddButton ws, 20, 30, 130, 28, "ExecuteCheckedJobs", "選択ジョブ実行", RGB(0, 176, 80)
    AddButton ws, 160, 30, 130, 28, "ClearJobList", "一覧クリア", RGB(192, 80, 77)

    ws.Rows(2).RowHeight = 35

    ' 説明
    ws.Range("A3").Value = "「選択」列をダブルクリックすると" & ChrW(&H2611) & "/" & ChrW(&H2610) & "が切り替わり、「順序」列に自動採番されます。順序は手動でも変更可能です（1から連番で入力）。保留中のジョブは実行時に自動で保留解除されます。"

    ' シートモジュールにイベントを追加
    AddWorksheetDoubleClickEvent ws

    ' ヘッダー
    ws.Cells(ROW_JOBLIST_HEADER, COL_SELECT).Value = "選択"
    ws.Cells(ROW_JOBLIST_HEADER, COL_ORDER).Value = "順序"
    ws.Cells(ROW_JOBLIST_HEADER, COL_UNIT_TYPE).Value = "種別"
    ws.Cells(ROW_JOBLIST_HEADER, COL_JOBNET_PATH).Value = "ユニットパス"
    ws.Cells(ROW_JOBLIST_HEADER, COL_JOBNET_NAME).Value = "ユニット名"
    ws.Cells(ROW_JOBLIST_HEADER, COL_COMMENT).Value = "コメント"
    ws.Cells(ROW_JOBLIST_HEADER, COL_SCRIPT).Value = "スクリプト"
    ws.Cells(ROW_JOBLIST_HEADER, COL_PARAMETER).Value = "パラメーター"
    ws.Cells(ROW_JOBLIST_HEADER, COL_WORK_PATH).Value = "ワークパス"
    ws.Cells(ROW_JOBLIST_HEADER, COL_HOLD).Value = "保留"
    ws.Cells(ROW_JOBLIST_HEADER, COL_LAST_STATUS).Value = "最終実行結果"
    ws.Cells(ROW_JOBLIST_HEADER, COL_LAST_EXEC_TIME).Value = "開始時刻"
    ws.Cells(ROW_JOBLIST_HEADER, COL_LAST_END_TIME).Value = "終了時刻"
    ws.Cells(ROW_JOBLIST_HEADER, COL_LAST_MESSAGE).Value = "ログパス"

    With ws.Range(ws.Cells(ROW_JOBLIST_HEADER, COL_SELECT), ws.Cells(ROW_JOBLIST_HEADER, COL_LAST_MESSAGE))
        .Font.Bold = True
        .Interior.Color = RGB(79, 129, 189)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With

    ' 列幅調整
    ws.Columns(COL_SELECT).ColumnWidth = 6
    ws.Columns(COL_ORDER).ColumnWidth = 6
    ws.Columns(COL_UNIT_TYPE).ColumnWidth = 12
    ws.Columns(COL_JOBNET_PATH).ColumnWidth = 50
    ws.Columns(COL_JOBNET_NAME).ColumnWidth = 25
    ws.Columns(COL_COMMENT).ColumnWidth = 80
    ws.Columns(COL_SCRIPT).ColumnWidth = 40
    ws.Columns(COL_PARAMETER).ColumnWidth = 30
    ws.Columns(COL_WORK_PATH).ColumnWidth = 30
    ws.Columns(COL_HOLD).ColumnWidth = 8
    ws.Columns(COL_LAST_STATUS).ColumnWidth = 15
    ws.Columns(COL_LAST_EXEC_TIME).ColumnWidth = 18
    ws.Columns(COL_LAST_END_TIME).ColumnWidth = 18
    ws.Columns(COL_LAST_MESSAGE).ColumnWidth = 60

    ' G〜I列をグループ化
    ws.Columns("G:I").Group
    ws.Outline.ShowLevels ColumnLevels:=1

    ' フィルター設定
    ws.Range(ws.Cells(ROW_JOBLIST_HEADER, COL_SELECT), ws.Cells(ROW_JOBLIST_HEADER, COL_LAST_MESSAGE)).AutoFilter

    ' ウィンドウ枠の固定
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Cells(ROW_JOBLIST_DATA_START, 1).Select
    ActiveWindow.FreezePanes = True
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

    ' ボタン追加
    AddButton ws, 20, 30, 100, 28, "ClearLogHistory", "履歴クリア", RGB(192, 80, 77)
    ws.Rows(2).RowHeight = 35

    ' 説明
    ws.Range("A3").Value = "ジョブ実行の履歴ログです。"

    ' ヘッダー
    ws.Cells(4, 1).Value = "実行日時"
    ws.Cells(4, 2).Value = "ジョブネットパス"
    ws.Cells(4, 3).Value = "結果"
    ws.Cells(4, 4).Value = "開始時刻"
    ws.Cells(4, 5).Value = "終了時刻"
    ws.Cells(4, 6).Value = "ログパス"

    With ws.Range("A4:F4")
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

    ' ウィンドウ枠の固定
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Cells(5, 1).Select
    ActiveWindow.FreezePanes = True
End Sub

'==============================================================================
' ドロップダウン追加
'==============================================================================
Private Sub AddDropdown(ws As Worksheet, cell As Range, options As String)
    With cell.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=options
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
End Sub

'==============================================================================
' ボタン追加
'==============================================================================
Private Sub AddButton(ws As Worksheet, left As Double, top As Double, width As Double, height As Double, macroName As String, caption As String, Optional fillColor As Long = -1)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, left, top, width, height)

    With shp
        .Name = "btn_" & macroName
        .OnAction = macroName

        If fillColor = -1 Then
            .Fill.ForeColor.RGB = RGB(0, 112, 192)
        Else
            .Fill.ForeColor.RGB = fillColor
        End If

        .Line.ForeColor.RGB = RGB(0, 80, 150)
        .Line.Weight = 1

        .TextFrame2.TextRange.Text = caption
        .TextFrame2.TextRange.Font.Size = 10
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.MarginLeft = 0
        .TextFrame2.MarginRight = 0

        .Placement = xlFreeFloating
    End With
End Sub

'==============================================================================
' シートモジュールにWorksheet_BeforeDoubleClickイベントを追加
'==============================================================================
Private Sub AddWorksheetDoubleClickEvent(ws As Worksheet)
    On Error Resume Next

    ' VBAプロジェクトへのアクセスを確認
    Dim vbProj As Object
    Set vbProj = ThisWorkbook.VBProject

    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If

    ' シートのコードモジュールを取得
    Dim sheetModule As Object
    Set sheetModule = vbProj.VBComponents(ws.CodeName).CodeModule

    ' 既にイベントコードが追加されているか確認
    Dim existingCode As String
    existingCode = sheetModule.Lines(1, sheetModule.CountOfLines)

    If InStr(existingCode, "Private Sub Worksheet_BeforeDoubleClick") > 0 Then
        Exit Sub
    End If

    ' イベントコードを追加
    Dim eventCode As String
    eventCode = vbCrLf & _
        "Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)" & vbCrLf & _
        "    ' 選択列のダブルクリックでチェック切り替え" & vbCrLf & _
        "    On Error Resume Next" & vbCrLf & _
        "    OnJobListDoubleClick Target.Row, Target.Column, Cancel" & vbCrLf & _
        "    On Error GoTo 0" & vbCrLf & _
        "End Sub" & vbCrLf

    sheetModule.InsertLines sheetModule.CountOfLines + 1, eventCode

    On Error GoTo 0
End Sub
