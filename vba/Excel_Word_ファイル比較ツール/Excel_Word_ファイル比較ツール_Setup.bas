'==============================================================================
' Excel/Word ファイル比較ツール - セットアップモジュール
' モジュール名: Excel_Word_ファイル比較ツール_Setup
'==============================================================================
' 概要:
'   Excel_Word_ファイル比較ツールの初期化とシートフォーマット機能を提供します。
'
' 含まれる機能:
'   - メインシート初期化
'   - シートフォーマット設定
'
' 作成日: 2025-12-17
'==============================================================================

Option Explicit

'==============================================================================
' 定数（差異ハイライト色）
'==============================================================================
Public Const COLOR_CHANGED As Long = 65535      ' 黄色: 値変更
Public Const COLOR_ADDED As Long = 5296274      ' 緑: 追加
Public Const COLOR_DELETED As Long = 13421823   ' 赤: 削除

'==============================================================================
' メインシート初期化
'==============================================================================
Public Sub InitializeExcelWordファイル比較ツール()
    Dim ws As Worksheet
    Dim mainSheetName As String

    mainSheetName = "メイン"

    On Error Resume Next
    Application.DisplayAlerts = False

    ' 既存のメインシートがあれば削除
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = mainSheetName Then
            ws.Delete
            Exit For
        End If
    Next ws

    Application.DisplayAlerts = True
    On Error GoTo 0

    ' 新しいシートを作成
    Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
    ws.Name = mainSheetName

    ' シートを初期化
    FormatMainSheet ws

    MsgBox "メインシートを初期化しました。", vbInformation, "初期化完了"
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
        .Range("B2:H2").Merge
        .Range("B2").Value = "Excel / Word ファイル比較ツール"
        With .Range("B2")
            .Font.Name = "Meiryo UI"
            .Font.Size = 20
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        .Range("B2:H3").Interior.Color = RGB(47, 84, 150)
        .Rows(2).RowHeight = 40
        .Rows(3).RowHeight = 5

        ' =================================================================
        ' 説明エリア (行5)
        ' =================================================================
        .Range("B5:H5").Merge
        .Range("B5").Value = "2つのExcelファイルまたはWordファイルを比較し、差異を一覧表示します。"
        With .Range("B5")
            .Font.Name = "Meiryo UI"
            .Font.Size = 11
            .Font.Color = RGB(64, 64, 64)
        End With

        ' =================================================================
        ' ボタンエリア (行8-10)
        ' =================================================================
        Dim btnLeft As Double
        Dim btnTop As Double
        Dim btnWidth As Double
        Dim btnHeight As Double
        Dim btnGap As Double

        .Rows(7).RowHeight = 15
        .Rows(8).RowHeight = 50

        ' ボタンサイズと位置の設定（固定値で統一）
        btnWidth = 130
        btnHeight = 40
        btnGap = 30
        btnLeft = .Range("B8").Left + 10
        btnTop = .Range("B8").Top + 5

        ' Excel比較ボタン（緑系）- 角丸四角形
        Set shp = .Shapes.AddShape(msoShapeRoundedRectangle, _
            btnLeft, btnTop, btnWidth, btnHeight)
        With shp
            .Name = "btnCompareExcel"
            .Placement = xlFreeFloating  ' セルサイズに連動しない
            .Fill.ForeColor.RGB = RGB(76, 175, 80)  ' 緑
            .Line.ForeColor.RGB = RGB(56, 142, 60)
            .Line.Weight = 2
            .TextFrame2.TextRange.Characters.Text = "Excel比較"
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextFrame2.TextRange.Font.Size = 14
            .TextFrame2.TextRange.Font.Bold = msoTrue
            .TextFrame2.TextRange.Font.Name = "Meiryo UI"
            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .TextFrame2.VerticalAnchor = msoAnchorMiddle
            .OnAction = "CompareExcelFiles"
        End With

        ' Word比較ボタン（青系）- 角丸四角形（Excel比較ボタンの右側に固定間隔で配置）
        Set shp = .Shapes.AddShape(msoShapeRoundedRectangle, _
            btnLeft + btnWidth + btnGap, btnTop, btnWidth, btnHeight)
        With shp
            .Name = "btnCompareWord"
            .Placement = xlFreeFloating  ' セルサイズに連動しない
            .Fill.ForeColor.RGB = RGB(33, 150, 243)  ' 青
            .Line.ForeColor.RGB = RGB(25, 118, 210)
            .Line.Weight = 2
            .TextFrame2.TextRange.Characters.Text = "Word比較"
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextFrame2.TextRange.Font.Size = 14
            .TextFrame2.TextRange.Font.Bold = msoTrue
            .TextFrame2.TextRange.Font.Name = "Meiryo UI"
            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .TextFrame2.VerticalAnchor = msoAnchorMiddle
            .OnAction = "CompareWordFiles"
        End With

        ' =================================================================
        ' 色凡例セクション (行11-16)
        ' =================================================================
        .Range("B11:H11").Merge
        .Range("B11").Value = "差異の色凡例"
        With .Range("B11")
            .Font.Name = "Meiryo UI"
            .Font.Size = 14
            .Font.Bold = True
            .Font.Color = RGB(47, 84, 150)
        End With

        ' セクション下線
        With .Range("B11:H11").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(47, 84, 150)
            .Weight = xlMedium
        End With

        ' 変更
        .Range("B13").Value = "変更"
        With .Range("B13")
            .Interior.Color = COLOR_CHANGED
            .Font.Name = "Meiryo UI"
            .Font.Size = 11
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
        With .Range("B13").Borders
            .LineStyle = xlContinuous
            .Color = RGB(200, 200, 200)
        End With
        .Range("C13:E13").Merge
        .Range("C13").Value = "値が変更された箇所（黄色）"
        With .Range("C13")
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
        End With

        ' 追加
        .Range("B14").Value = "追加"
        With .Range("B14")
            .Interior.Color = COLOR_ADDED
            .Font.Name = "Meiryo UI"
            .Font.Size = 11
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
        With .Range("B14").Borders
            .LineStyle = xlContinuous
            .Color = RGB(200, 200, 200)
        End With
        .Range("C14:E14").Merge
        .Range("C14").Value = "新ファイルで追加された箇所（緑）"
        With .Range("C14")
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
        End With

        ' 削除
        .Range("B15").Value = "削除"
        With .Range("B15")
            .Interior.Color = COLOR_DELETED
            .Font.Name = "Meiryo UI"
            .Font.Size = 11
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
        With .Range("B15").Borders
            .LineStyle = xlContinuous
            .Color = RGB(200, 200, 200)
        End With
        .Range("C15:E15").Merge
        .Range("C15").Value = "新ファイルで削除された箇所（ピンク）"
        With .Range("C15")
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
        End With

        ' =================================================================
        ' 設定セクション (行18-22)
        ' =================================================================
        .Range("B18:H18").Merge
        .Range("B18").Value = "現在の設定"
        With .Range("B18")
            .Font.Name = "Meiryo UI"
            .Font.Size = 14
            .Font.Bold = True
            .Font.Color = RGB(47, 84, 150)
        End With

        ' セクション下線
        With .Range("B18:H18").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(47, 84, 150)
            .Weight = xlMedium
        End With

        ' 設定値の表示
        .Range("B20").Value = "Excel比較:"
        .Range("D20").Value = "使用範囲のみ比較（制限なし）"
        .Range("B21").Value = "Word比較:"
        .Range("D21").Value = "WinMerge方式（LCSアルゴリズム）"

        .Range("B20:B21").Font.Name = "Meiryo UI"
        .Range("B20:B21").Font.Size = 10
        .Range("D20:D21").Font.Name = "Meiryo UI"
        .Range("D20:D21").Font.Size = 10
        .Range("D20:D21").Font.Color = RGB(0, 128, 0)

        ' =================================================================
        ' 対応ファイル形式セクション (行24-28)
        ' =================================================================
        .Range("B24:H24").Merge
        .Range("B24").Value = "対応ファイル形式"
        With .Range("B24")
            .Font.Name = "Meiryo UI"
            .Font.Size = 14
            .Font.Bold = True
            .Font.Color = RGB(47, 84, 150)
        End With

        ' セクション下線
        With .Range("B24:H24").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(47, 84, 150)
            .Weight = xlMedium
        End With

        .Range("B26").Value = "Excel:"
        .Range("C26").Value = ".xlsx, .xlsm, .xls, .xlsb"
        .Range("B27").Value = "Word:"
        .Range("C27").Value = ".docx, .docm, .doc"

        .Range("B26:B27").Font.Name = "Meiryo UI"
        .Range("B26:B27").Font.Size = 10
        .Range("B26:B27").Font.Bold = True
        .Range("C26:C27").Font.Name = "Meiryo UI"
        .Range("C26:C27").Font.Size = 10

        ' =================================================================
        ' 使い方セクション (行30-35)
        ' =================================================================
        .Range("B30:H30").Merge
        .Range("B30").Value = "使い方"
        With .Range("B30")
            .Font.Name = "Meiryo UI"
            .Font.Size = 14
            .Font.Bold = True
            .Font.Color = RGB(47, 84, 150)
        End With

        ' セクション下線
        With .Range("B30:H30").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(47, 84, 150)
            .Weight = xlMedium
        End With

        .Range("B32").Value = "1."
        .Range("C32").Value = "「Excel比較」または「Word比較」ボタンをクリック"
        .Range("B33").Value = "2."
        .Range("C33").Value = "1つ目のファイルを選択"
        .Range("B34").Value = "3."
        .Range("C34").Value = "2つ目のファイルを選択"
        .Range("B35").Value = "4."
        .Range("C35").Value = "比較結果が「比較結果」シートに出力されます"

        .Range("B32:B35").Font.Name = "Meiryo UI"
        .Range("B32:B35").Font.Size = 10
        .Range("B32:B35").Font.Bold = True
        .Range("B32:B35").Font.Color = RGB(47, 84, 150)
        .Range("C32:C35").Font.Name = "Meiryo UI"
        .Range("C32:C35").Font.Size = 10

        ' =================================================================
        ' 列幅・行高の調整
        ' =================================================================
        .Columns("A").ColumnWidth = 3
        .Columns("B").ColumnWidth = 20
        .Columns("C").ColumnWidth = 15
        .Columns("D").ColumnWidth = 12
        .Columns("E").ColumnWidth = 15
        .Columns("F").ColumnWidth = 12
        .Columns("G").ColumnWidth = 15
        .Columns("H").ColumnWidth = 12
        .Columns("I").ColumnWidth = 3

        ' セルA1を選択
        .Range("A1").Select
    End With

    Application.ScreenUpdating = True
End Sub
