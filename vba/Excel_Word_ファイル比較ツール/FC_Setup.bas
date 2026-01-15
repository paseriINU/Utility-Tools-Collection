Attribute VB_Name = "FC_Setup"
Option Explicit

'==============================================================================
' Excel/Word ファイル比較ツール - セットアップモジュール
' メインシート初期化、フォーマット機能を提供
' ※定数はFC_Configを参照。このモジュールは初期化後に削除可能
'==============================================================================

' ============================================================================
' メインシート初期化
' ============================================================================
Public Sub InitializeExcelWordファイル比較ツール()
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

    MsgBox "メインシートを初期化しました。", vbInformation, "初期化完了"
End Sub

' ============================================================================
' メインシートのフォーマット
' ============================================================================
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

        btnWidth = 130
        btnHeight = 40
        btnGap = 30
        btnLeft = .Range("B8").Left + 10
        btnTop = .Range("B8").Top + 5

        ' Excel比較ボタン
        Set shp = .Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, btnWidth, btnHeight)
        With shp
            .Name = "btnCompareExcel"
            .Placement = xlFreeFloating
            .Fill.ForeColor.RGB = RGB(76, 175, 80)
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

        ' Word比較ボタン
        Set shp = .Shapes.AddShape(msoShapeRoundedRectangle, btnLeft + btnWidth + btnGap, btnTop, btnWidth, btnHeight)
        With shp
            .Name = "btnCompareWord"
            .Placement = xlFreeFloating
            .Fill.ForeColor.RGB = RGB(33, 150, 243)
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
        .Range("B18").Value = "Word比較オプション"
        With .Range("B18")
            .Font.Name = "Meiryo UI"
            .Font.Size = 14
            .Font.Bold = True
            .Font.Color = RGB(47, 84, 150)
        End With

        With .Range("B18:H18").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(47, 84, 150)
            .Weight = xlMedium
        End With

        .Range("B20").Value = ""
        .Range("C20:F20").Merge
        .Range("C20").Value = "厳密比較（LCS）を使用する（処理時間が長くなります）"
        With .Range("C20")
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
        End With

        Dim chkBox As CheckBox
        Set chkBox = .CheckBoxes.Add(.Range("B20").Left + 5, .Range("B20").Top + 2, 15, 15)
        With chkBox
            .Name = "chkUseLCS"
            .Caption = ""
            .Value = xlOff
        End With

        .Range("C21:G21").Merge
        .Range("C21").Value = "　チェックなし: 簡易比較（高速、通常はこちらで十分）"
        With .Range("C21")
            .Font.Name = "Meiryo UI"
            .Font.Size = 9
            .Font.Color = RGB(100, 100, 100)
        End With

        .Range("C22:G22").Merge
        .Range("C22").Value = "　チェックあり: LCSアルゴリズム（大規模な構造変更に対応）"
        With .Range("C22")
            .Font.Name = "Meiryo UI"
            .Font.Size = 9
            .Font.Color = RGB(100, 100, 100)
        End With

        ' スタイル比較チェックボックス
        .Range("B24").Value = ""
        .Range("C24:F24").Merge
        .Range("C24").Value = "スタイル変更も検出する（フォント、サイズ等）"
        With .Range("C24")
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
        End With

        Dim chkStyleBox As CheckBox
        Set chkStyleBox = .CheckBoxes.Add(.Range("B24").Left + 5, .Range("B24").Top + 2, 15, 15)
        With chkStyleBox
            .Name = "chkCheckStyle"
            .Caption = ""
            .Value = xlOn
        End With

        .Range("C25:G25").Merge
        .Range("C25").Value = "　チェックなし: テキストの変更のみ検出（高速）"
        With .Range("C25")
            .Font.Name = "Meiryo UI"
            .Font.Size = 9
            .Font.Color = RGB(100, 100, 100)
        End With

        .Range("C26:G26").Merge
        .Range("C26").Value = "　チェックあり: スタイル変更も検出（書式の違いを検出）"
        With .Range("C26")
            .Font.Name = "Meiryo UI"
            .Font.Size = 9
            .Font.Color = RGB(100, 100, 100)
        End With

        ' =================================================================
        ' 対応ファイル形式セクション (行28-32)
        ' =================================================================
        .Range("B28:H28").Merge
        .Range("B28").Value = "対応ファイル形式"
        With .Range("B28")
            .Font.Name = "Meiryo UI"
            .Font.Size = 14
            .Font.Bold = True
            .Font.Color = RGB(47, 84, 150)
        End With

        With .Range("B28:H28").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(47, 84, 150)
            .Weight = xlMedium
        End With

        .Range("B30").Value = "Excel:"
        .Range("C30").Value = ".xlsx, .xlsm, .xls, .xlsb"
        .Range("B31").Value = "Word:"
        .Range("C31").Value = ".docx, .docm, .doc"

        .Range("B30:B31").Font.Name = "Meiryo UI"
        .Range("B30:B31").Font.Size = 10
        .Range("B30:B31").Font.Bold = True
        .Range("C30:C31").Font.Name = "Meiryo UI"
        .Range("C30:C31").Font.Size = 10

        ' =================================================================
        ' 使い方セクション (行34-39)
        ' =================================================================
        .Range("B34:H34").Merge
        .Range("B34").Value = "使い方"
        With .Range("B34")
            .Font.Name = "Meiryo UI"
            .Font.Size = 14
            .Font.Bold = True
            .Font.Color = RGB(47, 84, 150)
        End With

        With .Range("B34:H34").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(47, 84, 150)
            .Weight = xlMedium
        End With

        .Range("B36").Value = "1."
        .Range("C36").Value = "「Excel比較」または「Word比較」ボタンをクリック"
        .Range("B37").Value = "2."
        .Range("C37").Value = "1つ目のファイルを選択"
        .Range("B38").Value = "3."
        .Range("C38").Value = "2つ目のファイルを選択"
        .Range("B39").Value = "4."
        .Range("C39").Value = "比較結果が「比較結果」シートに出力されます"

        .Range("B36:B39").Font.Name = "Meiryo UI"
        .Range("B36:B39").Font.Size = 10
        .Range("B36:B39").Font.Bold = True
        .Range("B36:B39").Font.Color = RGB(47, 84, 150)
        .Range("C36:C39").Font.Name = "Meiryo UI"
        .Range("C36:C39").Font.Size = 10

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

        .Range("A1").Select
    End With

    Application.ScreenUpdating = True
End Sub
