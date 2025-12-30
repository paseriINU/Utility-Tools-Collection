Option Explicit

'==============================================================================
' JP1 REST ジョブ管理ツール - 初期化モジュール
'   - シート作成・フォーマット処理
'   - 初回セットアップ時のみ実行
'   - WinRM不使用、REST API専用
'==============================================================================

' シート名定数（Publicで共有）
Public Const SHEET_SETTINGS As String = "設定"
Public Const SHEET_TREE As String = "ツリー表示"
Public Const SHEET_LOG As String = "実行ログ"

' 設定セル位置（設定シート）
Public Const ROW_WEB_CONSOLE_HOST As Long = 7
Public Const ROW_WEB_CONSOLE_PORT As Long = 8
Public Const ROW_USE_HTTPS As Long = 9
Public Const ROW_MANAGER_HOST As Long = 10
Public Const ROW_SCHEDULER_SERVICE As Long = 11
Public Const ROW_JP1_USER As Long = 12
Public Const ROW_JP1_PASSWORD As Long = 13
Public Const ROW_ROOT_PATH As Long = 14
Public Const ROW_WAIT_COMPLETION As Long = 16
Public Const ROW_POLLING_INTERVAL As Long = 17
Public Const ROW_TIMEOUT As Long = 18
Public Const ROW_DEBUG_MODE As Long = 20
Public Const COL_SETTING_VALUE As Long = 3

' ツリー表示シートの列位置
Public Const COL_EXPAND As Long = 1          ' 展開/折りたたみ（>[v]）
Public Const COL_UNIT_NAME As Long = 2       ' インデント付きユニット名
Public Const COL_UNIT_PATH As Long = 3       ' ユニットパス（フルパス）
Public Const COL_UNIT_TYPE As Long = 4       ' ユニット種別
Public Const COL_STATUS As Long = 5          ' 状態
Public Const COL_LAST_RESULT As Long = 6     ' 最終実行結果
Public Const COL_EXEC_ID As Long = 7         ' execID
Public Const COL_START_TIME As Long = 8      ' 開始時刻
Public Const COL_END_TIME As Long = 9        ' 終了時刻
Public Const COL_SELECT As Long = 10         ' 選択チェック
Public Const ROW_TREE_HEADER As Long = 4
Public Const ROW_TREE_DATA_START As Long = 5

' ログシートの行位置
Public Const ROW_LOG_HEADER As Long = 4
Public Const ROW_LOG_DATA_START As Long = 5

'==============================================================================
' 初期化（メインエントリポイント）
'==============================================================================
Public Sub InitializeJP1RESTManager()
    Application.ScreenUpdating = False

    ' シート作成
    CreateSheet SHEET_SETTINGS
    CreateSheet SHEET_TREE
    CreateSheet SHEET_LOG

    ' 設定シートのフォーマット
    FormatSettingsSheet

    ' ツリー表示シートのフォーマット
    FormatTreeSheet

    ' ログシートのフォーマット
    FormatLogSheet

    ' 設定シートをアクティブに
    Worksheets(SHEET_SETTINGS).Activate

    Application.ScreenUpdating = True

    MsgBox "初期化が完了しました。" & vbCrLf & vbCrLf & _
           "1. 設定シートで接続設定を入力してください" & vbCrLf & _
           "2. 「ツリー取得」ボタンでジョブ階層を取得" & vbCrLf & _
           "3. ツリー表示でユニットをダブルクリックで展開" & vbCrLf & _
           "4. 「選択実行」ボタンでジョブネットを実行", _
           vbInformation, "JP1 REST ジョブ管理ツール"
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
        .Value = "JP1 REST ジョブ管理ツール - 接続設定"
        .Font.Size = 16
        .Font.Bold = True
        .Interior.Color = RGB(0, 112, 192)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .RowHeight = 30
    End With

    ' 説明
    ws.Range("A2").Value = "JP1/AJS3 Web Console REST APIを使用してジョブ階層を取得し、選択したジョブネットを実行します。（WinRM不使用）"

    ' ボタン追加
    AddButton ws, 20, 55, 130, 32, "RefreshTree", "ツリー取得", RGB(0, 112, 192)
    AddButton ws, 160, 55, 130, 32, "ExecuteSelectedJobnet", "選択実行", RGB(0, 176, 80)

    ' 設定セクション
    ws.Range("A6").Value = "■ Web Console接続設定"
    ws.Range("A6").Font.Bold = True

    ws.Cells(ROW_WEB_CONSOLE_HOST, 1).Value = "Web Consoleサーバ"
    ws.Cells(ROW_WEB_CONSOLE_HOST, COL_SETTING_VALUE).Value = "localhost"
    ws.Cells(ROW_WEB_CONSOLE_HOST, 4).Value = "※JP1/AJS3 Web Consoleのホスト名またはIPアドレス"
    ws.Cells(ROW_WEB_CONSOLE_HOST, 4).Font.Color = RGB(128, 128, 128)

    ws.Cells(ROW_WEB_CONSOLE_PORT, 1).Value = "Web Consoleポート"
    ws.Cells(ROW_WEB_CONSOLE_PORT, COL_SETTING_VALUE).Value = "22252"
    ws.Cells(ROW_WEB_CONSOLE_PORT, 4).Value = "※HTTP:22252 / HTTPS:22253"
    ws.Cells(ROW_WEB_CONSOLE_PORT, 4).Font.Color = RGB(128, 128, 128)

    ws.Cells(ROW_USE_HTTPS, 1).Value = "HTTPS使用"
    ws.Cells(ROW_USE_HTTPS, COL_SETTING_VALUE).Value = "いいえ"
    AddDropdown ws, ws.Cells(ROW_USE_HTTPS, COL_SETTING_VALUE), "いいえ,はい"
    ws.Cells(ROW_USE_HTTPS, 4).Value = "※自己署名証明書も使用可"
    ws.Cells(ROW_USE_HTTPS, 4).Font.Color = RGB(128, 128, 128)

    ws.Cells(ROW_MANAGER_HOST, 1).Value = "Managerホスト"
    ws.Cells(ROW_MANAGER_HOST, COL_SETTING_VALUE).Value = "localhost"
    ws.Cells(ROW_MANAGER_HOST, 4).Value = "※JP1/AJS3 Managerのホスト名"
    ws.Cells(ROW_MANAGER_HOST, 4).Font.Color = RGB(128, 128, 128)

    ws.Cells(ROW_SCHEDULER_SERVICE, 1).Value = "スケジューラーサービス"
    ws.Cells(ROW_SCHEDULER_SERVICE, COL_SETTING_VALUE).Value = "AJSROOT1"
    ws.Cells(ROW_SCHEDULER_SERVICE, 4).Value = "※JP1/AJS3のスケジューラーサービス名"
    ws.Cells(ROW_SCHEDULER_SERVICE, 4).Font.Color = RGB(128, 128, 128)

    ws.Cells(ROW_JP1_USER, 1).Value = "JP1ユーザー"
    ws.Cells(ROW_JP1_USER, COL_SETTING_VALUE).Value = "jp1admin"

    ws.Cells(ROW_JP1_PASSWORD, 1).Value = "JP1パスワード"
    ws.Cells(ROW_JP1_PASSWORD, COL_SETTING_VALUE).Value = ""
    ws.Cells(ROW_JP1_PASSWORD, 4).Value = "※空の場合は実行時に入力"
    ws.Cells(ROW_JP1_PASSWORD, 4).Font.Color = RGB(128, 128, 128)

    ws.Cells(ROW_ROOT_PATH, 1).Value = "ルートパス"
    ws.Cells(ROW_ROOT_PATH, COL_SETTING_VALUE).Value = "/"
    ws.Cells(ROW_ROOT_PATH, 4).Value = "※ツリー取得の起点パス（例: / または /グループ名）"
    ws.Cells(ROW_ROOT_PATH, 4).Font.Color = RGB(128, 128, 128)

    ' 実行設定セクション
    ws.Range("A15").Value = "■ 実行設定"
    ws.Range("A15").Font.Bold = True

    ws.Cells(ROW_WAIT_COMPLETION, 1).Value = "完了待ち"
    ws.Cells(ROW_WAIT_COMPLETION, COL_SETTING_VALUE).Value = "はい"
    AddDropdown ws, ws.Cells(ROW_WAIT_COMPLETION, COL_SETTING_VALUE), "はい,いいえ"
    ws.Cells(ROW_WAIT_COMPLETION, 4).Value = "※実行完了まで待機してログを自動取得"
    ws.Cells(ROW_WAIT_COMPLETION, 4).Font.Color = RGB(128, 128, 128)

    ws.Cells(ROW_POLLING_INTERVAL, 1).Value = "状態確認間隔（秒）"
    ws.Cells(ROW_POLLING_INTERVAL, COL_SETTING_VALUE).Value = 5

    ws.Cells(ROW_TIMEOUT, 1).Value = "タイムアウト（秒）"
    ws.Cells(ROW_TIMEOUT, COL_SETTING_VALUE).Value = 300
    ws.Cells(ROW_TIMEOUT, 4).Value = "※0=無制限"
    ws.Cells(ROW_TIMEOUT, 4).Font.Color = RGB(128, 128, 128)

    ' デバッグセクション
    ws.Range("A19").Value = "■ デバッグ設定"
    ws.Range("A19").Font.Bold = True

    ws.Cells(ROW_DEBUG_MODE, 1).Value = "デバッグモード"
    ws.Cells(ROW_DEBUG_MODE, COL_SETTING_VALUE).Value = "いいえ"
    AddDropdown ws, ws.Cells(ROW_DEBUG_MODE, COL_SETTING_VALUE), "いいえ,はい"
    ws.Cells(ROW_DEBUG_MODE, 4).Value = "※はい=PowerShellウィンドウ表示・API応答ログ出力"
    ws.Cells(ROW_DEBUG_MODE, 4).Font.Color = RGB(128, 128, 128)

    ' 使い方セクション
    ws.Range("A22").Value = "■ 使い方"
    ws.Range("A22").Font.Bold = True

    ws.Range("A23").Value = "1. 上記の接続設定を入力します"
    ws.Range("A24").Value = "2. 「ツリー取得」ボタンでジョブ階層を取得します"
    ws.Range("A25").Value = "3. ツリー表示シートで[>]をダブルクリックして展開します"
    ws.Range("A26").Value = "4. 実行したいジョブネットの「選択」列をチェックします"
    ws.Range("A27").Value = "5. 「選択実行」ボタンでジョブネットを即時実行します"
    ws.Range("A28").Value = "6. 完了待ち「はい」の場合、自動でログを取得します"

    ws.Range("A30").Value = "■ 注意事項"
    ws.Range("A30").Font.Bold = True

    ws.Range("A31").Value = "・このツールはREST APIのみを使用し、WinRMは使用しません"
    ws.Range("A32").Value = "・実行可能なのはジョブネット（ROOTNET/NET）のみです"
    ws.Range("A33").Value = "・ログ取得は最大5MBまでです（超過分は取得できません）"

    ' 列幅調整
    ws.Columns("A").ColumnWidth = 22
    ws.Columns("B").ColumnWidth = 5
    ws.Columns("C").ColumnWidth = 25
    ws.Columns("D").ColumnWidth = 50

    ' 入力セルの書式（黄色背景）
    Dim settingCells As Variant
    settingCells = Array(ROW_WEB_CONSOLE_HOST, ROW_WEB_CONSOLE_PORT, ROW_USE_HTTPS, _
                         ROW_MANAGER_HOST, ROW_SCHEDULER_SERVICE, ROW_JP1_USER, _
                         ROW_JP1_PASSWORD, ROW_ROOT_PATH, ROW_WAIT_COMPLETION, _
                         ROW_POLLING_INTERVAL, ROW_TIMEOUT, ROW_DEBUG_MODE)
    Dim r As Variant
    For Each r In settingCells
        With ws.Cells(CLng(r), COL_SETTING_VALUE)
            .Interior.Color = RGB(255, 255, 204)
            .Borders.LineStyle = xlContinuous
        End With
    Next r
End Sub

'==============================================================================
' ツリー表示シートのフォーマット
'==============================================================================
Private Sub FormatTreeSheet()
    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_TREE)

    ws.Cells.Clear

    ' タイトル
    With ws.Range("A1:J1")
        .Merge
        .Value = "JP1 ジョブ階層ツリー"
        .Font.Size = 14
        .Font.Bold = True
        .Interior.Color = RGB(0, 176, 80)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .RowHeight = 25
    End With

    ' ボタン追加
    AddButton ws, 20, 30, 100, 28, "RefreshTree", "更新", RGB(0, 112, 192)
    AddButton ws, 130, 30, 100, 28, "ExecuteSelectedJobnet", "選択実行", RGB(0, 176, 80)
    AddButton ws, 240, 30, 100, 28, "GetExecutionLog", "ログ取得", RGB(91, 155, 213)
    AddButton ws, 350, 30, 100, 28, "ExpandAll", "全展開", RGB(128, 128, 128)
    AddButton ws, 460, 30, 100, 28, "CollapseAll", "全折りたたみ", RGB(128, 128, 128)

    ws.Rows(2).RowHeight = 35

    ' 説明
    ws.Range("A3").Value = "[>]をダブルクリックで展開、[v]で折りたたみ。ジョブネットを選択して「選択実行」で即時実行。"

    ' ダブルクリックイベントを追加
    AddWorksheetDoubleClickEvent ws

    ' ヘッダー
    ws.Cells(ROW_TREE_HEADER, COL_EXPAND).Value = ""
    ws.Cells(ROW_TREE_HEADER, COL_UNIT_NAME).Value = "ユニット名"
    ws.Cells(ROW_TREE_HEADER, COL_UNIT_PATH).Value = "ユニットパス"
    ws.Cells(ROW_TREE_HEADER, COL_UNIT_TYPE).Value = "種別"
    ws.Cells(ROW_TREE_HEADER, COL_STATUS).Value = "状態"
    ws.Cells(ROW_TREE_HEADER, COL_LAST_RESULT).Value = "最終結果"
    ws.Cells(ROW_TREE_HEADER, COL_EXEC_ID).Value = "execID"
    ws.Cells(ROW_TREE_HEADER, COL_START_TIME).Value = "開始時刻"
    ws.Cells(ROW_TREE_HEADER, COL_END_TIME).Value = "終了時刻"
    ws.Cells(ROW_TREE_HEADER, COL_SELECT).Value = "選択"

    With ws.Range(ws.Cells(ROW_TREE_HEADER, COL_EXPAND), ws.Cells(ROW_TREE_HEADER, COL_SELECT))
        .Font.Bold = True
        .Interior.Color = RGB(79, 129, 189)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With

    ' 列幅調整
    ws.Columns(COL_EXPAND).ColumnWidth = 4
    ws.Columns(COL_UNIT_NAME).ColumnWidth = 45
    ws.Columns(COL_UNIT_PATH).ColumnWidth = 50
    ws.Columns(COL_UNIT_TYPE).ColumnWidth = 12
    ws.Columns(COL_STATUS).ColumnWidth = 12
    ws.Columns(COL_LAST_RESULT).ColumnWidth = 12
    ws.Columns(COL_EXEC_ID).ColumnWidth = 12
    ws.Columns(COL_START_TIME).ColumnWidth = 18
    ws.Columns(COL_END_TIME).ColumnWidth = 18
    ws.Columns(COL_SELECT).ColumnWidth = 6

    ' ユニットパス列を非表示（操作用に保持）
    ws.Columns(COL_UNIT_PATH).Hidden = True

    ' ウィンドウ枠の固定
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Cells(ROW_TREE_DATA_START, 1).Select
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
    With ws.Range("A1:G1")
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
    AddButton ws, 130, 30, 100, 28, "GetExecutionLog", "ログ取得", RGB(91, 155, 213)
    ws.Rows(2).RowHeight = 35

    ' 説明
    ws.Range("A3").Value = "ジョブ実行の履歴ログです。「ログ取得」で選択ジョブのログを手動取得できます。"

    ' ヘッダー
    ws.Cells(ROW_LOG_HEADER, 1).Value = "実行日時"
    ws.Cells(ROW_LOG_HEADER, 2).Value = "ユニットパス"
    ws.Cells(ROW_LOG_HEADER, 3).Value = "操作"
    ws.Cells(ROW_LOG_HEADER, 4).Value = "結果"
    ws.Cells(ROW_LOG_HEADER, 5).Value = "execID"
    ws.Cells(ROW_LOG_HEADER, 6).Value = "開始時刻"
    ws.Cells(ROW_LOG_HEADER, 7).Value = "終了時刻"

    With ws.Range("A4:G4")
        .Font.Bold = True
        .Interior.Color = RGB(192, 80, 77)
        .Font.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
    End With

    ' 列幅調整
    ws.Columns("A").ColumnWidth = 18
    ws.Columns("B").ColumnWidth = 50
    ws.Columns("C").ColumnWidth = 10
    ws.Columns("D").ColumnWidth = 12
    ws.Columns("E").ColumnWidth = 12
    ws.Columns("F").ColumnWidth = 18
    ws.Columns("G").ColumnWidth = 18

    ' ウィンドウ枠の固定
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Cells(ROW_LOG_DATA_START, 1).Select
    ActiveWindow.FreezePanes = True
End Sub

'==============================================================================
' ユーティリティ（初期化用）
'==============================================================================
Private Sub AddDropdown(ws As Worksheet, cell As Range, options As String)
    With cell.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=options
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
End Sub

Private Sub AddButton(ws As Worksheet, left As Double, top As Double, width As Double, height As Double, macroName As String, caption As String, Optional fillColor As Long = -1)
    ' 図形ボタンを追加（固定サイズ・色付き）
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, left, top, width, height)

    With shp
        .Name = "btn_" & macroName
        .OnAction = macroName

        ' 塗りつぶし色
        If fillColor = -1 Then
            .Fill.ForeColor.RGB = RGB(0, 112, 192)
        Else
            .Fill.ForeColor.RGB = fillColor
        End If

        ' 枠線
        .Line.ForeColor.RGB = RGB(0, 80, 150)
        .Line.Weight = 1

        ' テキスト設定
        .TextFrame2.TextRange.Text = caption
        .TextFrame2.TextRange.Font.Size = 10
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.MarginLeft = 0
        .TextFrame2.MarginRight = 0

        ' セルに依存しない（固定位置・固定サイズ）
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
        ' VBAプロジェクトへのアクセスが許可されていない場合は何もしない
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
        ' 既に追加済み
        Exit Sub
    End If

    ' Worksheet_BeforeDoubleClickイベントコードを追加
    Dim eventCode As String
    eventCode = vbCrLf & _
        "Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)" & vbCrLf & _
        "    ' ツリー展開/折りたたみ処理" & vbCrLf & _
        "    On Error Resume Next" & vbCrLf & _
        "    OnTreeDoubleClick Target.Row, Target.Column, Cancel" & vbCrLf & _
        "    On Error GoTo 0" & vbCrLf & _
        "End Sub" & vbCrLf

    sheetModule.InsertLines sheetModule.CountOfLines + 1, eventCode

    On Error GoTo 0
End Sub
