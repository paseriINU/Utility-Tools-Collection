Option Explicit

'==============================================================================
' JP1 ジョブ管理ツール - 初期化モジュール
'   - シート作成・フォーマット処理
'   - 初回セットアップ時のみ実行
'==============================================================================

' シート名定数（Publicで共有）
Public Const SHEET_MAIN As String = "メイン"
Public Const SHEET_JOBLIST As String = "ジョブ一覧"
Public Const SHEET_LOG As String = "実行ログ"

' 設定セル位置（メインシート）- Publicで共有
Public Const ROW_EXEC_MODE As Long = 5
Public Const ROW_JP1_SERVER As Long = 7
Public Const ROW_REMOTE_USER As Long = 8
Public Const ROW_REMOTE_PASSWORD As Long = 9
Public Const ROW_JP1_USER As Long = 10
Public Const ROW_JP1_PASSWORD As Long = 11
Public Const ROW_ROOT_PATH As Long = 12
Public Const ROW_WAIT_COMPLETION As Long = 13
Public Const ROW_TIMEOUT As Long = 14
Public Const ROW_POLLING_INTERVAL As Long = 15
Public Const COL_SETTING_VALUE As Long = 3

' ジョブ一覧シートの列位置 - Publicで共有
Public Const COL_ORDER As Long = 1
Public Const COL_JOBNET_PATH As Long = 2
Public Const COL_JOBNET_NAME As Long = 3
Public Const COL_COMMENT As Long = 4
Public Const COL_LAST_STATUS As Long = 5
Public Const COL_LAST_EXEC_TIME As Long = 6
Public Const COL_LAST_END_TIME As Long = 7
Public Const COL_LAST_RETURN_CODE As Long = 8
Public Const COL_LAST_MESSAGE As Long = 9
Public Const ROW_JOBLIST_HEADER As Long = 3
Public Const ROW_JOBLIST_DATA_START As Long = 4

'==============================================================================
' 初期化（メインエントリポイント）
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

    ws.Cells(ROW_EXEC_MODE, 1).Value = "実行モード"
    ws.Cells(ROW_EXEC_MODE, COL_SETTING_VALUE).Value = "リモート"
    AddDropdown ws, ws.Cells(ROW_EXEC_MODE, COL_SETTING_VALUE), "ローカル,リモート"
    ws.Cells(ROW_EXEC_MODE, 4).Value = "※ローカル: このPCのJP1を使用、リモート: WinRM経由で接続"
    ws.Cells(ROW_EXEC_MODE, 4).Font.Color = RGB(128, 128, 128)

    ws.Range("A6").Value = "【リモート接続設定】（ローカルモード時は不要）"
    ws.Range("A6").Font.Color = RGB(128, 128, 128)

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
    ws.Range("A16").Value = "■ 実行設定"
    ws.Range("A16").Font.Bold = True

    ws.Cells(ROW_WAIT_COMPLETION + 4, 1).Value = "完了待ち"
    ws.Cells(ROW_WAIT_COMPLETION + 4, COL_SETTING_VALUE).Value = "はい"
    AddDropdown ws, ws.Cells(ROW_WAIT_COMPLETION + 4, COL_SETTING_VALUE), "はい,いいえ"

    ws.Cells(ROW_TIMEOUT + 4, 1).Value = "タイムアウト（秒）"
    ws.Cells(ROW_TIMEOUT + 4, COL_SETTING_VALUE).Value = 0
    ws.Cells(ROW_TIMEOUT + 4, 4).Value = "※0=無制限"
    ws.Cells(ROW_TIMEOUT + 4, 4).Font.Color = RGB(128, 128, 128)

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
    With ws.Range(ws.Cells(ROW_EXEC_MODE, COL_SETTING_VALUE), ws.Cells(ROW_POLLING_INTERVAL + 4, COL_SETTING_VALUE))
        .Interior.Color = RGB(255, 255, 204)
        .Borders.LineStyle = xlContinuous
    End With
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
    ws.Range("A2").Value = "実行するジョブの「順序」列に数字（1, 2, 3...）を入力してください。順序が入っているジョブを1番から順に実行します。"

    ' ヘッダー
    ws.Cells(ROW_JOBLIST_HEADER, COL_ORDER).Value = "順序"
    ws.Cells(ROW_JOBLIST_HEADER, COL_JOBNET_PATH).Value = "ジョブネットパス"
    ws.Cells(ROW_JOBLIST_HEADER, COL_JOBNET_NAME).Value = "ジョブネット名"
    ws.Cells(ROW_JOBLIST_HEADER, COL_COMMENT).Value = "コメント"
    ws.Cells(ROW_JOBLIST_HEADER, COL_LAST_STATUS).Value = "最終実行結果"
    ws.Cells(ROW_JOBLIST_HEADER, COL_LAST_EXEC_TIME).Value = "開始時刻"
    ws.Cells(ROW_JOBLIST_HEADER, COL_LAST_END_TIME).Value = "終了時刻"
    ws.Cells(ROW_JOBLIST_HEADER, COL_LAST_RETURN_CODE).Value = "戻り値"
    ws.Cells(ROW_JOBLIST_HEADER, COL_LAST_MESSAGE).Value = "詳細メッセージ"

    With ws.Range(ws.Cells(ROW_JOBLIST_HEADER, COL_ORDER), ws.Cells(ROW_JOBLIST_HEADER, COL_LAST_MESSAGE))
        .Font.Bold = True
        .Interior.Color = RGB(79, 129, 189)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With

    ' 列幅調整
    ws.Columns(COL_ORDER).ColumnWidth = 6
    ws.Columns(COL_JOBNET_PATH).ColumnWidth = 50
    ws.Columns(COL_JOBNET_NAME).ColumnWidth = 25
    ws.Columns(COL_COMMENT).ColumnWidth = 30
    ws.Columns(COL_LAST_STATUS).ColumnWidth = 15
    ws.Columns(COL_LAST_EXEC_TIME).ColumnWidth = 18
    ws.Columns(COL_LAST_END_TIME).ColumnWidth = 18
    ws.Columns(COL_LAST_RETURN_CODE).ColumnWidth = 8
    ws.Columns(COL_LAST_MESSAGE).ColumnWidth = 50

    ' フィルター設定
    ws.Range(ws.Cells(ROW_JOBLIST_HEADER, COL_ORDER), ws.Cells(ROW_JOBLIST_HEADER, COL_LAST_MESSAGE)).AutoFilter
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

Private Sub AddButton(ws As Worksheet, left As Double, top As Double, width As Double, height As Double, macroName As String, caption As String)
    Dim btn As Button
    Set btn = ws.Buttons.Add(left, top, width, height)
    btn.OnAction = macroName
    btn.caption = caption
    btn.Font.Size = 10
End Sub
