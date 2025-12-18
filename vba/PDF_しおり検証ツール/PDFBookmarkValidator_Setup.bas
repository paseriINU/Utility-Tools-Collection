Attribute VB_Name = "PDFBookmarkValidator_Setup"
Option Explicit

'==============================================================================
' PDF しおり検証ツール - 初期化モジュール
'   - シート作成・フォーマット処理
'   - 初回セットアップ時のみ実行
'==============================================================================

' シート名定数（Publicで共有）
Public Const SHEET_SETTINGS As String = "設定"
Public Const SHEET_RESULT As String = "検証結果"

' 設定セル位置（設定シート）- Publicで共有
Public Const ROW_PDF_PATH As Long = 7
Public Const ROW_CHECK_PAGE As Long = 9
Public Const ROW_CHECK_TEXT As Long = 10
Public Const ROW_TEXT_MATCH_RATIO As Long = 11
Public Const COL_SETTING_LABEL As Long = 1
Public Const COL_SETTING_VALUE As Long = 3

' 検証結果シートの列位置 - Publicで共有
Public Const COL_NO As Long = 1
Public Const COL_BOOKMARK_NAME As Long = 2
Public Const COL_BOOKMARK_LEVEL As Long = 3
Public Const COL_LINK_PAGE As Long = 4
Public Const COL_PAGE_TEXT As Long = 5
Public Const COL_TEXT_MATCH As Long = 6
Public Const COL_MATCH_RATIO As Long = 7
Public Const COL_STATUS As Long = 8
Public Const ROW_RESULT_HEADER As Long = 3
Public Const ROW_RESULT_DATA_START As Long = 4

'==============================================================================
' 初期化（メインエントリポイント）
'==============================================================================
Public Sub InitializePDFBookmarkValidator()
    Application.ScreenUpdating = False

    ' シート作成
    CreateSheet SHEET_SETTINGS
    CreateSheet SHEET_RESULT

    ' 設定シートのフォーマット
    FormatSettingsSheet

    ' 検証結果シートのフォーマット
    FormatResultSheet

    ' 設定シートをアクティブに
    Worksheets(SHEET_SETTINGS).Activate

    Application.ScreenUpdating = True

    MsgBox "初期化が完了しました。" & vbCrLf & vbCrLf & _
           "【初回セットアップ】" & vbCrLf & _
           "1. iTextSharp.dll を lib フォルダに配置してください" & vbCrLf & _
           "   （詳細はREADME.mdを参照）" & vbCrLf & vbCrLf & _
           "【使い方】" & vbCrLf & _
           "1. PDFファイルを選択" & vbCrLf & _
           "2. 検証オプションを設定" & vbCrLf & _
           "3. 「しおり検証」ボタンをクリック", _
           vbInformation, "PDF しおり検証ツール"
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
        .Value = "PDF しおり検証ツール"
        .Font.Size = 16
        .Font.Bold = True
        .Interior.Color = RGB(112, 48, 160)  ' 紫
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .RowHeight = 30
    End With

    ' 説明
    ws.Range("A2").Value = "PDFのしおり（ブックマーク）がリンク先ページと正しく対応しているかを検証します。"

    ' ボタン配置用スペース
    ws.Rows(3).RowHeight = 40

    ' ボタン追加
    AddButton ws, 20, 55, 130, 32, "SelectPDFFile", "PDFファイル選択", RGB(112, 48, 160)
    AddButton ws, 160, 55, 130, 32, "ValidateBookmarks", "しおり検証", RGB(0, 176, 80)

    ' 設定セクション
    ws.Range("A5").Value = "■ PDFファイル"
    ws.Range("A5").Font.Bold = True

    ws.Cells(ROW_PDF_PATH, COL_SETTING_LABEL).Value = "PDFファイルパス"
    ws.Cells(ROW_PDF_PATH, COL_SETTING_VALUE).Value = ""
    ws.Cells(ROW_PDF_PATH, 5).Value = "※「PDFファイル選択」ボタンで選択"
    ws.Cells(ROW_PDF_PATH, 5).Font.Color = RGB(128, 128, 128)

    ' 検証オプションセクション
    ws.Range("A8").Value = "■ 検証オプション"
    ws.Range("A8").Font.Bold = True

    ws.Cells(ROW_CHECK_PAGE, COL_SETTING_LABEL).Value = "ページ番号確認"
    ws.Cells(ROW_CHECK_PAGE, COL_SETTING_VALUE).Value = "はい"
    AddDropdown ws, ws.Cells(ROW_CHECK_PAGE, COL_SETTING_VALUE), "はい,いいえ"
    ws.Cells(ROW_CHECK_PAGE, 5).Value = "※しおりのリンク先ページ番号を表示"
    ws.Cells(ROW_CHECK_PAGE, 5).Font.Color = RGB(128, 128, 128)

    ws.Cells(ROW_CHECK_TEXT, COL_SETTING_LABEL).Value = "テキスト一致確認"
    ws.Cells(ROW_CHECK_TEXT, COL_SETTING_VALUE).Value = "はい"
    AddDropdown ws, ws.Cells(ROW_CHECK_TEXT, COL_SETTING_VALUE), "はい,いいえ"
    ws.Cells(ROW_CHECK_TEXT, 5).Value = "※しおり名がリンク先ページに存在するかチェック"
    ws.Cells(ROW_CHECK_TEXT, 5).Font.Color = RGB(128, 128, 128)

    ws.Cells(ROW_TEXT_MATCH_RATIO, COL_SETTING_LABEL).Value = "一致判定しきい値（%）"
    ws.Cells(ROW_TEXT_MATCH_RATIO, COL_SETTING_VALUE).Value = 80
    ws.Cells(ROW_TEXT_MATCH_RATIO, 5).Value = "※しおり名とページテキストの類似度（0-100）"
    ws.Cells(ROW_TEXT_MATCH_RATIO, 5).Font.Color = RGB(128, 128, 128)

    ' 使い方セクション
    ws.Range("A13").Value = "■ 使い方"
    ws.Range("A13").Font.Bold = True
    ws.Range("A14").Value = "1. 「PDFファイル選択」ボタンでPDFを選択"
    ws.Range("A15").Value = "2. 検証オプションを必要に応じて変更"
    ws.Range("A16").Value = "3. 「しおり検証」ボタンをクリック"
    ws.Range("A17").Value = "4. 検証結果シートで結果を確認"

    ' 列幅調整
    ws.Columns("A").ColumnWidth = 22
    ws.Columns("B").ColumnWidth = 3
    ws.Columns("C").ColumnWidth = 50
    ws.Columns("D").ColumnWidth = 3
    ws.Columns("E").ColumnWidth = 45

    ' 入力セルの書式（設定セルを黄色背景に）
    Dim settingRows As Variant
    settingRows = Array(ROW_PDF_PATH, ROW_CHECK_PAGE, ROW_CHECK_TEXT, ROW_TEXT_MATCH_RATIO)
    Dim r As Variant
    For Each r In settingRows
        With ws.Cells(CLng(r), COL_SETTING_VALUE)
            .Interior.Color = RGB(255, 255, 204)
            .Borders.LineStyle = xlContinuous
        End With
    Next r

    ' PDFパスセルは幅広く
    ws.Range(ws.Cells(ROW_PDF_PATH, COL_SETTING_VALUE), ws.Cells(ROW_PDF_PATH, 4)).Merge
End Sub

'==============================================================================
' 検証結果シートのフォーマット
'==============================================================================
Private Sub FormatResultSheet()
    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_RESULT)

    ws.Cells.Clear

    ' タイトル
    With ws.Range("A1:H1")
        .Merge
        .Value = "しおり検証結果"
        .Font.Size = 14
        .Font.Bold = True
        .Interior.Color = RGB(0, 176, 80)  ' 緑
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .RowHeight = 25
    End With

    ' ボタン配置用スペース
    ws.Rows(2).RowHeight = 35

    ' ボタン追加
    AddButton ws, 20, 30, 130, 28, "ExportResultToCSV", "CSV出力", RGB(79, 129, 189)
    AddButton ws, 160, 30, 130, 28, "ClearResult", "結果クリア", RGB(192, 80, 77)

    ' ヘッダー
    ws.Cells(ROW_RESULT_HEADER, COL_NO).Value = "No"
    ws.Cells(ROW_RESULT_HEADER, COL_BOOKMARK_NAME).Value = "しおり名"
    ws.Cells(ROW_RESULT_HEADER, COL_BOOKMARK_LEVEL).Value = "階層"
    ws.Cells(ROW_RESULT_HEADER, COL_LINK_PAGE).Value = "リンク先ページ"
    ws.Cells(ROW_RESULT_HEADER, COL_PAGE_TEXT).Value = "ページ先頭テキスト"
    ws.Cells(ROW_RESULT_HEADER, COL_TEXT_MATCH).Value = "テキスト一致"
    ws.Cells(ROW_RESULT_HEADER, COL_MATCH_RATIO).Value = "一致率"
    ws.Cells(ROW_RESULT_HEADER, COL_STATUS).Value = "判定"

    With ws.Range(ws.Cells(ROW_RESULT_HEADER, COL_NO), ws.Cells(ROW_RESULT_HEADER, COL_STATUS))
        .Font.Bold = True
        .Interior.Color = RGB(79, 129, 189)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With

    ' 列幅調整
    ws.Columns(COL_NO).ColumnWidth = 5
    ws.Columns(COL_BOOKMARK_NAME).ColumnWidth = 40
    ws.Columns(COL_BOOKMARK_LEVEL).ColumnWidth = 6
    ws.Columns(COL_LINK_PAGE).ColumnWidth = 12
    ws.Columns(COL_PAGE_TEXT).ColumnWidth = 50
    ws.Columns(COL_TEXT_MATCH).ColumnWidth = 12
    ws.Columns(COL_MATCH_RATIO).ColumnWidth = 8
    ws.Columns(COL_STATUS).ColumnWidth = 8

    ' フィルター設定
    ws.Range(ws.Cells(ROW_RESULT_HEADER, COL_NO), ws.Cells(ROW_RESULT_HEADER, COL_STATUS)).AutoFilter
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
            .Fill.ForeColor.RGB = RGB(112, 48, 160)
        Else
            .Fill.ForeColor.RGB = fillColor
        End If

        ' 枠線
        .Line.ForeColor.RGB = RGB(80, 30, 120)
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
