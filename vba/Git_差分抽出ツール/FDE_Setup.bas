Attribute VB_Name = "FDE_Setup"
Option Explicit

'==============================================================================
' Git 差分ファイル抽出ツール（VBA版） - セットアップモジュール
' 初期化、シートフォーマット機能を提供
' ※このモジュールは初期化後に削除可能
'==============================================================================

'==============================================================================
' メインシート初期化
'==============================================================================
Public Sub InitializeGit差分抽出ツール()
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
           "リポジトリパスと比較対象を設定して、" & vbCrLf & _
           "「比較実行」ボタンをクリックしてください。", vbInformation, "初期化完了"
End Sub

'==============================================================================
' メインシートのフォーマット
'==============================================================================
Private Sub FormatMainSheet(ByRef ws As Worksheet)
    Dim shp As Shape
    Dim btnLeft As Double
    Dim btnTop As Double

    Application.ScreenUpdating = False

    With ws
        ' 全体の背景色を白に
        .Cells.Interior.Color = RGB(255, 255, 255)

        ' =================================================================
        ' タイトルエリア (行1-3)
        ' =================================================================
        .Range("B2:G2").Merge
        .Range("B2").Value = "Git 差分ファイル抽出ツール（VBA版）"
        With .Range("B2")
            .Font.Name = "Meiryo UI"
            .Font.Size = 18
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
        .Range("B5").Value = "Gitブランチ間/コミット間の差分ファイルを抽出し、履歴を残します。"
        With .Range("B5")
            .Font.Name = "Meiryo UI"
            .Font.Size = 11
            .Font.Color = RGB(64, 64, 64)
        End With

        ' =================================================================
        ' 設定セクション (行7-17)
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

        .Range("D8:F8").Merge
        .Range("D8").Value = "C:\Users\%USERNAME%\source\Git\project"
        With .Range("D8:F8")
            .Interior.Color = RGB(255, 255, 230)
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
        End With
        With .Range("D8:F8").Borders
            .LineStyle = xlContinuous
            .Color = RGB(200, 200, 200)
        End With

        ' 参照ボタン（リポジトリ）
        btnLeft = .Range("G8").Left
        btnTop = .Range("G8").Top
        Set shp = .Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, 60, 22)
        With shp
            .Name = "btnSelectRepo"
            .Placement = xlFreeFloating
            .Fill.ForeColor.RGB = RGB(200, 200, 200)
            .Line.ForeColor.RGB = RGB(150, 150, 150)
            .TextFrame2.TextRange.Characters.Text = "参照..."
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
            .TextFrame2.TextRange.Font.Size = 9
            .TextFrame2.TextRange.Font.Name = "Meiryo UI"
            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .TextFrame2.VerticalAnchor = msoAnchorMiddle
            .OnAction = "SelectRepoPath"
        End With

        ' 出力先フォルダ
        .Range("B10").Value = "出力先フォルダ:"
        With .Range("B10")
            .Font.Name = "Meiryo UI"
            .Font.Size = 11
            .Font.Bold = True
        End With

        .Range("D10:F10").Merge
        .Range("D10").Value = "%USERPROFILE%\Desktop\git_diff"
        With .Range("D10:F10")
            .Interior.Color = RGB(255, 255, 230)
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
        End With
        With .Range("D10:F10").Borders
            .LineStyle = xlContinuous
            .Color = RGB(200, 200, 200)
        End With

        ' 参照ボタン（出力先）
        btnLeft = .Range("G10").Left
        btnTop = .Range("G10").Top
        Set shp = .Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, 60, 22)
        With shp
            .Name = "btnSelectOutput"
            .Placement = xlFreeFloating
            .Fill.ForeColor.RGB = RGB(200, 200, 200)
            .Line.ForeColor.RGB = RGB(150, 150, 150)
            .TextFrame2.TextRange.Characters.Text = "参照..."
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
            .TextFrame2.TextRange.Font.Size = 9
            .TextFrame2.TextRange.Font.Name = "Meiryo UI"
            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .TextFrame2.VerticalAnchor = msoAnchorMiddle
            .OnAction = "SelectOutputFolder"
        End With

        ' =================================================================
        ' 比較対象セクション (行12-17)
        ' =================================================================
        .Range("B12:G12").Merge
        .Range("B12").Value = "比較対象"
        With .Range("B12")
            .Font.Name = "Meiryo UI"
            .Font.Size = 14
            .Font.Bold = True
            .Font.Color = RGB(68, 114, 196)
        End With
        With .Range("B12:G12").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(68, 114, 196)
            .Weight = xlMedium
        End With

        ' 比較元
        .Range("B14").Value = "比較元（修正前）:"
        With .Range("B14")
            .Font.Name = "Meiryo UI"
            .Font.Size = 11
            .Font.Bold = True
        End With

        .Range("D14:F14").Merge
        .Range("D14").Value = "main"
        With .Range("D14:F14")
            .Interior.Color = RGB(255, 230, 230)
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
        End With
        With .Range("D14:F14").Borders
            .LineStyle = xlContinuous
            .Color = RGB(200, 200, 200)
        End With

        ' 選択ボタン（比較元）
        btnLeft = .Range("G14").Left
        btnTop = .Range("G14").Top
        Set shp = .Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, 60, 22)
        With shp
            .Name = "btnSelectBase"
            .Placement = xlFreeFloating
            .Fill.ForeColor.RGB = RGB(200, 200, 200)
            .Line.ForeColor.RGB = RGB(150, 150, 150)
            .TextFrame2.TextRange.Characters.Text = "選択..."
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
            .TextFrame2.TextRange.Font.Size = 9
            .TextFrame2.TextRange.Font.Name = "Meiryo UI"
            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .TextFrame2.VerticalAnchor = msoAnchorMiddle
            .OnAction = "SelectBaseRef"
        End With

        ' 比較先
        .Range("B16").Value = "比較先（修正後）:"
        With .Range("B16")
            .Font.Name = "Meiryo UI"
            .Font.Size = 11
            .Font.Bold = True
        End With

        .Range("D16:F16").Merge
        .Range("D16").Value = "HEAD"
        With .Range("D16:F16")
            .Interior.Color = RGB(230, 255, 230)
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
        End With
        With .Range("D16:F16").Borders
            .LineStyle = xlContinuous
            .Color = RGB(200, 200, 200)
        End With

        ' 選択ボタン（比較先）
        btnLeft = .Range("G16").Left
        btnTop = .Range("G16").Top
        Set shp = .Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, 60, 22)
        With shp
            .Name = "btnSelectTarget"
            .Placement = xlFreeFloating
            .Fill.ForeColor.RGB = RGB(200, 200, 200)
            .Line.ForeColor.RGB = RGB(150, 150, 150)
            .TextFrame2.TextRange.Characters.Text = "選択..."
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
            .TextFrame2.TextRange.Font.Size = 9
            .TextFrame2.TextRange.Font.Name = "Meiryo UI"
            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .TextFrame2.VerticalAnchor = msoAnchorMiddle
            .OnAction = "SelectTargetRef"
        End With

        ' =================================================================
        ' 実行ボタンエリア (行18-20)
        ' =================================================================
        .Rows(18).RowHeight = 15
        .Rows(19).RowHeight = 50

        ' 比較実行ボタン
        btnLeft = .Range("C19").Left
        btnTop = .Range("C19").Top + 5

        Set shp = .Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, 120, 40)
        With shp
            .Name = "btnExecute"
            .Placement = xlFreeFloating
            .Fill.ForeColor.RGB = RGB(76, 175, 80)
            .Line.ForeColor.RGB = RGB(56, 142, 60)
            .Line.Weight = 2
            .TextFrame2.TextRange.Characters.Text = "比較実行"
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextFrame2.TextRange.Font.Size = 14
            .TextFrame2.TextRange.Font.Bold = msoTrue
            .TextFrame2.TextRange.Font.Name = "Meiryo UI"
            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .TextFrame2.VerticalAnchor = msoAnchorMiddle
            .OnAction = "ExecuteCompare"
        End With

        ' 差分ファイル抽出ボタン
        Set shp = .Shapes.AddShape(msoShapeRoundedRectangle, btnLeft + 140, btnTop, 140, 40)
        With shp
            .Name = "btnExtract"
            .Placement = xlFreeFloating
            .Fill.ForeColor.RGB = RGB(33, 150, 243)
            .Line.ForeColor.RGB = RGB(25, 118, 210)
            .Line.Weight = 2
            .TextFrame2.TextRange.Characters.Text = "差分ファイル抽出"
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextFrame2.TextRange.Font.Size = 14
            .TextFrame2.TextRange.Font.Bold = msoTrue
            .TextFrame2.TextRange.Font.Name = "Meiryo UI"
            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .TextFrame2.VerticalAnchor = msoAnchorMiddle
            .OnAction = "ExtractDiffFiles"
        End With

        ' =================================================================
        ' 説明セクション (行22-)
        ' =================================================================
        .Range("B22:G22").Merge
        .Range("B22").Value = "使い方"
        With .Range("B22")
            .Font.Name = "Meiryo UI"
            .Font.Size = 14
            .Font.Bold = True
            .Font.Color = RGB(68, 114, 196)
        End With
        With .Range("B22:G22").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(68, 114, 196)
            .Weight = xlMedium
        End With

        .Range("B24").Value = "1."
        .Range("C24:G24").Merge
        .Range("C24").Value = "リポジトリパスにGitプロジェクトのパスを入力"
        With .Range("C24")
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
        End With

        .Range("B25").Value = "2."
        .Range("C25:G25").Merge
        .Range("C25").Value = "比較元と比較先にブランチ名またはコミットハッシュを入力"
        With .Range("C25")
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
        End With

        .Range("B26").Value = "3."
        .Range("C26:G26").Merge
        .Range("C26").Value = "「比較実行」をクリック → 差分ファイル一覧がシートに出力"
        With .Range("C26")
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
        End With

        .Range("B27").Value = "4."
        .Range("C27:G27").Merge
        .Range("C27").Value = "「差分ファイル抽出」をクリック → 出力先にファイルをコピー"
        With .Range("C27")
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
        End With

        ' =================================================================
        ' 出力シート説明
        ' =================================================================
        .Range("B29:G29").Merge
        .Range("B29").Value = "出力フォルダ構成"
        With .Range("B29")
            .Font.Name = "Meiryo UI"
            .Font.Size = 14
            .Font.Bold = True
            .Font.Color = RGB(68, 114, 196)
        End With
        With .Range("B29:G29").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(68, 114, 196)
            .Weight = xlMedium
        End With

        .Range("B31:G31").Merge
        .Range("B31").Value = "  出力先フォルダ/"
        .Range("B32:G32").Merge
        .Range("B32").Value = "    01_修正前/     ← 比較元（修正前）のファイル"
        .Range("B33:G33").Merge
        .Range("B33").Value = "    02_修正後/     ← 比較先（修正後）のファイル"
        .Range("B34:G34").Merge
        .Range("B34").Value = "    diff_report.txt ← 差分レポート"

        Dim rng As Range
        For Each rng In .Range("B31:B34")
            rng.Font.Name = "Consolas"
            rng.Font.Size = 10
            rng.Font.Color = RGB(80, 80, 80)
        Next rng

        ' =================================================================
        ' 列幅調整
        ' =================================================================
        .Columns("A").ColumnWidth = 3
        .Columns("B").ColumnWidth = 18
        .Columns("C").ColumnWidth = 3
        .Columns("D").ColumnWidth = 20
        .Columns("E").ColumnWidth = 15
        .Columns("F").ColumnWidth = 15
        .Columns("G").ColumnWidth = 12
        .Columns("H").ColumnWidth = 3

        .Range("A1").Select
    End With

    Application.ScreenUpdating = True
End Sub

