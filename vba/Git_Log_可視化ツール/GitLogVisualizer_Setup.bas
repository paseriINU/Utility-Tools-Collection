'==============================================================================
' Git Log 可視化ツール - セットアップモジュール
' モジュール名: GitLogVisualizer_Setup
'==============================================================================
' 概要:
'   GitLogVisualizerの初期化とシートフォーマット機能を提供します。
'
' 含まれる機能:
'   - メインシート初期化
'   - シートフォーマット設定
'
' 作成日: 2025-12-17
'==============================================================================

Option Explicit

'==============================================================================
' 定数（初期化・フォーマット用）
'==============================================================================
' シート名
Public Const SHEET_MAIN As String = "メイン"
Public Const SHEET_DASHBOARD As String = "ダッシュボード"
Public Const SHEET_HISTORY As String = "コミット履歴"
Public Const SHEET_BRANCH_GRAPH As String = "ブランチグラフ"

' メインシートのセル位置
Public Const CELL_REPO_PATH As String = "D8"
Public Const CELL_COMMIT_COUNT As String = "D10"

' 注意: GIT_COMMAND定数は業務ロジックで使用するため、メインモジュールに配置されています

'==============================================================================
' メインシート初期化
'==============================================================================
Public Sub InitializeGitLog可視化ツール()
    Dim ws As Worksheet

    On Error Resume Next
    Application.DisplayAlerts = False

    ' 既存のメインシートがあれば削除
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = SHEET_MAIN Then
            ws.Delete
            Exit For
        End If
    Next ws

    Application.DisplayAlerts = True
    On Error GoTo 0

    ' 新しいシートを作成
    Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
    ws.Name = SHEET_MAIN

    ' シートを初期化
    FormatMainSheet ws

    MsgBox "メインシートを初期化しました。" & vbCrLf & vbCrLf & _
           "リポジトリパスと取得件数を設定して、" & vbCrLf & _
           "「実行」ボタンをクリックしてください。", vbInformation, "初期化完了"
End Sub

'==============================================================================
' メインシートのフォーマット
'==============================================================================
Private Sub FormatMainSheet(ByRef ws As Worksheet)
    Dim shp As Shape

    Application.ScreenUpdating = False

    With ws
        ' 全体の背景色を白に
        .Cells.Interior.Color = RGB(255, 255, 255)

        ' =================================================================
        ' タイトルエリア (行1-3)
        ' =================================================================
        .Range("B2:G2").Merge
        .Range("B2").Value = "Git Log 可視化ツール"
        With .Range("B2")
            .Font.Name = "Meiryo UI"
            .Font.Size = 20
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        .Range("B2:G3").Interior.Color = RGB(68, 114, 196)
        .Rows(2).RowHeight = 40
        .Rows(3).RowHeight = 5

        ' =================================================================
        ' 説明エリア (行5)
        ' =================================================================
        .Range("B5:G5").Merge
        .Range("B5").Value = "Gitリポジトリのコミット履歴を取得して視覚化します。"
        With .Range("B5")
            .Font.Name = "Meiryo UI"
            .Font.Size = 11
            .Font.Color = RGB(64, 64, 64)
        End With

        ' =================================================================
        ' 設定セクション (行7-12)
        ' =================================================================
        .Range("B7:G7").Merge
        .Range("B7").Value = "設定"
        With .Range("B7")
            .Font.Name = "Meiryo UI"
            .Font.Size = 14
            .Font.Bold = True
            .Font.Color = RGB(68, 114, 196)
        End With
        With .Range("B7:G7").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(68, 114, 196)
            .Weight = xlMedium
        End With

        ' リポジトリパス
        .Range("B8").Value = "リポジトリパス:"
        With .Range("B8")
            .Font.Name = "Meiryo UI"
            .Font.Size = 11
            .Font.Bold = True
        End With

        .Range("D8:G8").Merge
        .Range("D8").Value = "C:\Users\%USERNAME%\source\Git\project"
        With .Range("D8:G8")
            .Interior.Color = RGB(255, 255, 230)
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
        End With
        With .Range("D8:G8").Borders
            .LineStyle = xlContinuous
            .Color = RGB(200, 200, 200)
        End With

        ' 環境変数の説明
        .Range("D9:G9").Merge
        .Range("D9").Value = "※ %USERNAME% などの環境変数が使用可能"
        With .Range("D9")
            .Font.Name = "Meiryo UI"
            .Font.Size = 9
            .Font.Color = RGB(100, 100, 100)
            .Font.Italic = True
        End With

        ' 取得コミット数
        .Range("B10").Value = "取得件数:"
        With .Range("B10")
            .Font.Name = "Meiryo UI"
            .Font.Size = 11
            .Font.Bold = True
        End With

        .Range("D10").Value = 100
        With .Range("D10")
            .Interior.Color = RGB(255, 255, 230)
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
            .NumberFormat = "#,##0"
            .HorizontalAlignment = xlCenter
        End With
        With .Range("D10").Borders
            .LineStyle = xlContinuous
            .Color = RGB(200, 200, 200)
        End With

        .Range("E10:G10").Merge
        .Range("E10").Value = "件（最新から取得）"
        With .Range("E10")
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
            .Font.Color = RGB(100, 100, 100)
        End With

        ' =================================================================
        ' ボタンエリア (行13-14)
        ' =================================================================
        .Rows(12).RowHeight = 15
        .Rows(13).RowHeight = 50

        ' 実行ボタン - 角丸四角形
        Dim btnLeft As Double
        Dim btnTop As Double
        btnLeft = .Range("D13").Left
        btnTop = .Range("D13").Top + 5

        Set shp = .Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, 120, 40)
        With shp
            .Name = "btnExecute"
            .Placement = xlFreeFloating
            .Fill.ForeColor.RGB = RGB(76, 175, 80)
            .Line.ForeColor.RGB = RGB(56, 142, 60)
            .Line.Weight = 2
            .TextFrame2.TextRange.Characters.Text = "実行"
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextFrame2.TextRange.Font.Size = 14
            .TextFrame2.TextRange.Font.Bold = msoTrue
            .TextFrame2.TextRange.Font.Name = "Meiryo UI"
            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .TextFrame2.VerticalAnchor = msoAnchorMiddle
            .OnAction = "ShowBranchInfoBeforeRun"
        End With

        ' ブランチ切り替えボタン
        Set shp = .Shapes.AddShape(msoShapeRoundedRectangle, btnLeft + 140, btnTop, 140, 40)
        With shp
            .Name = "btnSwitchBranch"
            .Placement = xlFreeFloating
            .Fill.ForeColor.RGB = RGB(33, 150, 243)
            .Line.ForeColor.RGB = RGB(25, 118, 210)
            .Line.Weight = 2
            .TextFrame2.TextRange.Characters.Text = "ブランチ切替"
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextFrame2.TextRange.Font.Size = 14
            .TextFrame2.TextRange.Font.Bold = msoTrue
            .TextFrame2.TextRange.Font.Name = "Meiryo UI"
            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .TextFrame2.VerticalAnchor = msoAnchorMiddle
            .OnAction = "SelectAndSwitchBranch"
        End With

        ' =================================================================
        ' 出力シートセクション (行16-20)
        ' =================================================================
        .Range("B16:G16").Merge
        .Range("B16").Value = "出力シート"
        With .Range("B16")
            .Font.Name = "Meiryo UI"
            .Font.Size = 14
            .Font.Bold = True
            .Font.Color = RGB(68, 114, 196)
        End With
        With .Range("B16:G16").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(68, 114, 196)
            .Weight = xlMedium
        End With

        .Range("B18").Value = "ダッシュボード"
        With .Range("B18")
            .Font.Name = "Meiryo UI"
            .Font.Size = 11
            .Font.Bold = True
            .Font.Color = RGB(68, 114, 196)
        End With
        .Range("C18:G18").Merge
        .Range("C18").Value = "サマリー情報（総コミット数、作者数、変更量、作者別統計）"
        With .Range("C18")
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
        End With

        .Range("B19").Value = "コミット履歴"
        With .Range("B19")
            .Font.Name = "Meiryo UI"
            .Font.Size = 11
            .Font.Bold = True
            .Font.Color = RGB(68, 114, 196)
        End With
        .Range("C19:G19").Merge
        .Range("C19").Value = "コミット履歴の詳細一覧（ハッシュ、作者、日時、メッセージ、変更量等）"
        With .Range("C19")
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
        End With

        .Range("B20").Value = "ブランチグラフ"
        With .Range("B20")
            .Font.Name = "Meiryo UI"
            .Font.Size = 11
            .Font.Bold = True
            .Font.Color = RGB(68, 114, 196)
        End With
        .Range("C20:G20").Merge
        .Range("C20").Value = "ブランチ構造を視覚化（コミットノードと接続線）"
        With .Range("C20")
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
        End With

        ' =================================================================
        ' 色凡例セクション (行23-29)
        ' =================================================================
        .Range("B23:G23").Merge
        .Range("B23").Value = "ブランチグラフの色凡例"
        With .Range("B23")
            .Font.Name = "Meiryo UI"
            .Font.Size = 14
            .Font.Bold = True
            .Font.Color = RGB(68, 114, 196)
        End With
        With .Range("B23:G23").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(68, 114, 196)
            .Weight = xlMedium
        End With

        ' 初期コミット
        .Range("B25").Interior.Color = RGB(255, 0, 0)
        With .Range("B25").Borders
            .LineStyle = xlContinuous
            .Color = RGB(200, 200, 200)
        End With
        .Range("C25:E25").Merge
        .Range("C25").Value = "初期コミット（親コミットなし）"
        With .Range("C25")
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
        End With

        ' 通常コミット
        .Range("B26").Interior.Color = RGB(0, 128, 255)
        With .Range("B26").Borders
            .LineStyle = xlContinuous
            .Color = RGB(200, 200, 200)
        End With
        .Range("C26:E26").Merge
        .Range("C26").Value = "通常コミット（親コミット1つ）"
        With .Range("C26")
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
        End With

        ' マージコミット
        .Range("B27").Interior.Color = RGB(0, 255, 0)
        With .Range("B27").Borders
            .LineStyle = xlContinuous
            .Color = RGB(200, 200, 200)
        End With
        .Range("C27:E27").Merge
        .Range("C27").Value = "マージコミット（親コミット2つ以上）"
        With .Range("C27")
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
        End With

        ' =================================================================
        ' 列幅調整
        ' =================================================================
        .Columns("A").ColumnWidth = 3
        .Columns("B").ColumnWidth = 18
        .Columns("C").ColumnWidth = 12
        .Columns("D").ColumnWidth = 15
        .Columns("E").ColumnWidth = 15
        .Columns("F").ColumnWidth = 15
        .Columns("G").ColumnWidth = 15
        .Columns("H").ColumnWidth = 3

        .Range("A1").Select
    End With

    Application.ScreenUpdating = True
End Sub
