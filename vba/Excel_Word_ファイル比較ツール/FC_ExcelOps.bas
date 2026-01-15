Attribute VB_Name = "FC_ExcelOps"
Option Explicit

'==============================================================================
' Excel/Word ファイル比較ツール - Excel操作モジュール
' Excel比較ロジック、結果シート作成機能を提供
'==============================================================================

' ============================================================================
' Excel比較の内部処理
' ============================================================================
Public Sub CompareExcelFilesInternal(ByVal file1Path As String, ByVal file2Path As String)
    Dim wb1 As Workbook
    Dim wb2 As Workbook
    Dim differences() As ExcelDiffInfo
    Dim diffCount As Long

    On Error GoTo ErrorHandler

    ' 処理開始
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    Debug.Print "========================================="
    Debug.Print "Excel ファイル比較を開始します"
    Debug.Print "旧ファイル: " & file1Path
    Debug.Print "新ファイル: " & file2Path
    Debug.Print "========================================="

    ' ファイルを開く
    Set wb1 = Workbooks.Open(file1Path, ReadOnly:=True)
    Set wb2 = Workbooks.Open(file2Path, ReadOnly:=True)

    ' 比較実行
    diffCount = 0
    ReDim differences(0 To 0)

    CompareWorkbooks wb1, wb2, differences, diffCount

    ' ファイルを閉じる
    wb1.Close SaveChanges:=False
    wb2.Close SaveChanges:=False

    ' 結果を出力
    If diffCount > 0 Then
        CreateExcelResultSheet differences, diffCount, file1Path, file2Path

        Debug.Print "========================================="
        Debug.Print "処理完了: " & diffCount & " 件の差異を検出"
        Debug.Print "========================================="

        MsgBox "比較が完了しました。" & vbCrLf & vbCrLf & _
               "検出された差異: " & diffCount & " 件" & vbCrLf & vbCrLf & _
               "結果は「比較結果」シートに出力されました。", _
               vbInformation, "処理完了"
    Else
        Debug.Print "========================================="
        Debug.Print "処理完了: 差異なし"
        Debug.Print "========================================="

        MsgBox "比較が完了しました。" & vbCrLf & vbCrLf & _
               "2つのファイルは同一です。差異はありませんでした。", _
               vbInformation, "処理完了"
    End If

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True

    ' 開いたワークブックを閉じる
    On Error Resume Next
    If Not wb1 Is Nothing Then wb1.Close SaveChanges:=False
    If Not wb2 Is Nothing Then wb2.Close SaveChanges:=False
    On Error GoTo 0

    MsgBox "エラーが発生しました: " & vbCrLf & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, vbCritical, "エラー"
End Sub

' ============================================================================
' ワークブックを比較（Excel）
' ============================================================================
Public Sub CompareWorkbooks(ByRef wb1 As Workbook, ByRef wb2 As Workbook, _
                            ByRef differences() As ExcelDiffInfo, ByRef diffCount As Long)
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim sheetNames1 As Object
    Dim sheetNames2 As Object
    Dim sheetName As Variant

    Set sheetNames1 = CreateObject("Scripting.Dictionary")
    Set sheetNames2 = CreateObject("Scripting.Dictionary")

    ' シート名を収集
    For Each ws1 In wb1.Worksheets
        sheetNames1.Add ws1.Name, ws1.Name
    Next ws1

    For Each ws2 In wb2.Worksheets
        sheetNames2.Add ws2.Name, ws2.Name
    Next ws2

    ' 両方に存在するシートを比較
    For Each sheetName In sheetNames1.Keys
        If sheetNames2.exists(sheetName) Then
            Debug.Print "シートを比較中: " & sheetName
            CompareSheets wb1.Worksheets(CStr(sheetName)), wb2.Worksheets(CStr(sheetName)), _
                          differences, diffCount
        Else
            ' wb2にないシート（削除されたシート）
            AddExcelDifference differences, diffCount, CStr(sheetName), "(シート全体)", _
                          "シート削除", "(存在)", "(削除済み)"
        End If
    Next sheetName

    ' wb2のみに存在するシート（追加されたシート）
    For Each sheetName In sheetNames2.Keys
        If Not sheetNames1.exists(sheetName) Then
            AddExcelDifference differences, diffCount, CStr(sheetName), "(シート全体)", _
                          "シート追加", "(なし)", "(追加済み)"
        End If
    Next sheetName
End Sub

' ============================================================================
' シートを比較（Excel）
' ============================================================================
Public Sub CompareSheets(ByRef ws1 As Worksheet, ByRef ws2 As Worksheet, _
                         ByRef differences() As ExcelDiffInfo, ByRef diffCount As Long)
    Dim lastRow1 As Long, lastCol1 As Long
    Dim lastRow2 As Long, lastCol2 As Long
    Dim maxRow As Long, maxCol As Long
    Dim r As Long, c As Long
    Dim val1 As Variant, val2 As Variant
    Dim cellAddr As String

    ' 使用範囲を取得
    lastRow1 = GetLastRow(ws1)
    lastCol1 = GetLastCol(ws1)
    lastRow2 = GetLastRow(ws2)
    lastCol2 = GetLastCol(ws2)

    ' 比較範囲を決定（使用範囲のみ比較、制限なし）
    maxRow = Application.WorksheetFunction.Max(lastRow1, lastRow2)
    maxCol = Application.WorksheetFunction.Max(lastCol1, lastCol2)

    Debug.Print "  比較範囲: " & maxRow & " 行 x " & maxCol & " 列"

    ' セル単位で比較
    For r = 1 To maxRow
        For c = 1 To maxCol
            val1 = ws1.Cells(r, c).Value
            val2 = ws2.Cells(r, c).Value

            ' 値が異なる場合
            If Not IsEqual(val1, val2) Then
                cellAddr = ws1.Cells(r, c).Address(False, False)

                ' 差異の種類を判定
                If IsEmpty(val1) And Not IsEmpty(val2) Then
                    ' 新ファイルで追加
                    AddExcelDifference differences, diffCount, ws1.Name, cellAddr, _
                                  "追加", "(空)", CStr(val2)
                ElseIf Not IsEmpty(val1) And IsEmpty(val2) Then
                    ' 新ファイルで削除
                    AddExcelDifference differences, diffCount, ws1.Name, cellAddr, _
                                  "削除", CStr(val1), "(空)"
                Else
                    ' 値の変更
                    AddExcelDifference differences, diffCount, ws1.Name, cellAddr, _
                                  "変更", CStr(val1), CStr(val2)
                End If
            End If
        Next c

        ' 進捗表示（100行ごと）
        If r Mod 100 = 0 Then
            Debug.Print "  " & ws1.Name & ": " & r & " / " & maxRow & " 行処理中..."
            DoEvents
        End If
    Next r
End Sub

' ============================================================================
' 最終行を取得
' ============================================================================
Public Function GetLastRow(ByRef ws As Worksheet) As Long
    On Error Resume Next
    GetLastRow = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
                              LookIn:=xlFormulas, LookAt:=xlPart, _
                              SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    If Err.Number <> 0 Or GetLastRow = 0 Then
        GetLastRow = 1
    End If
    On Error GoTo 0
End Function

' ============================================================================
' 最終列を取得
' ============================================================================
Public Function GetLastCol(ByRef ws As Worksheet) As Long
    On Error Resume Next
    GetLastCol = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
                              LookIn:=xlFormulas, LookAt:=xlPart, _
                              SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    If Err.Number <> 0 Or GetLastCol = 0 Then
        GetLastCol = 1
    End If
    On Error GoTo 0
End Function

' ============================================================================
' Excel差異を追加
' ============================================================================
Public Sub AddExcelDifference(ByRef differences() As ExcelDiffInfo, ByRef diffCount As Long, _
                         ByVal sheetName As String, ByVal cellAddr As String, _
                         ByVal diffType As String, ByVal oldVal As String, ByVal newVal As String)
    ' 配列を拡張
    If diffCount = 0 Then
        ReDim differences(0 To 0)
    Else
        ReDim Preserve differences(0 To diffCount)
    End If

    ' 差異情報を格納
    With differences(diffCount)
        .SheetName = sheetName
        .CellAddress = cellAddr
        .DiffType = diffType
        .OldValue = Left(oldVal, 255)  ' 長すぎる値を切り詰め
        .NewValue = Left(newVal, 255)
    End With

    diffCount = diffCount + 1
End Sub

' ============================================================================
' Excel結果シートを作成
' ============================================================================
Public Sub CreateExcelResultSheet(ByRef differences() As ExcelDiffInfo, ByVal diffCount As Long, _
                             ByVal file1Path As String, ByVal file2Path As String)
    Dim ws As Worksheet
    Dim i As Long
    Dim row As Long
    Dim hyperlinkAddr1 As String
    Dim hyperlinkAddr2 As String

    ' 既存の結果シートがあれば削除
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(SHEET_RESULT).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' 新しいシートを作成
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = SHEET_RESULT

    With ws
        ' タイトル
        .Range("A1").Value = "Excel ファイル比較結果"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True

        ' ファイル情報
        .Range("A3").Value = "旧ファイル（比較元）:"
        .Range("B3").Value = file1Path
        .Range("A4").Value = "新ファイル（比較先）:"
        .Range("B4").Value = file2Path
        .Range("A5").Value = "比較日時:"
        .Range("B5").Value = Now
        .Range("B5").NumberFormat = "yyyy/mm/dd hh:mm:ss"
        .Range("A6").Value = "検出差異数:"
        .Range("B6").Value = diffCount

        ' 凡例
        .Range("A8").Value = "凡例："
        .Range("B8").Value = "変更"
        .Range("B8").Interior.Color = COLOR_CHANGED
        .Range("C8").Value = "追加"
        .Range("C8").Interior.Color = COLOR_ADDED
        .Range("D8").Value = "削除"
        .Range("D8").Interior.Color = COLOR_DELETED

        ' ヘッダー
        .Range("A10").Value = "No"
        .Range("B10").Value = "シート名"
        .Range("C10").Value = "セル"
        .Range("D10").Value = "差異タイプ"
        .Range("E10").Value = "旧ファイルの値"
        .Range("F10").Value = "新ファイルの値"
        .Range("G10").Value = "旧ファイル"
        .Range("H10").Value = "新ファイル"

        ' ヘッダー書式
        With .Range("A10:H10")
            .Font.Bold = True
            .Interior.Color = RGB(68, 114, 196)
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
        End With

        ' データ行
        For i = 0 To diffCount - 1
            row = i + 11

            .Cells(row, 1).Value = i + 1
            .Cells(row, 2).Value = differences(i).SheetName
            .Cells(row, 3).Value = differences(i).CellAddress
            .Cells(row, 4).Value = differences(i).DiffType
            .Cells(row, 5).Value = differences(i).OldValue
            .Cells(row, 6).Value = differences(i).NewValue

            ' シート全体の差異でない場合はハイパーリンクを追加
            If differences(i).CellAddress <> "(シート全体)" Then
                ' 旧ファイルへのハイパーリンク
                hyperlinkAddr1 = file1Path & "#'" & differences(i).SheetName & "'!" & differences(i).CellAddress
                .Hyperlinks.Add Anchor:=.Cells(row, 7), Address:=hyperlinkAddr1, TextToDisplay:="移動"
                With .Cells(row, 7)
                    .Font.Color = RGB(0, 102, 204)
                    .Font.Underline = xlUnderlineStyleSingle
                    .HorizontalAlignment = xlCenter
                End With

                ' 新ファイルへのハイパーリンク
                hyperlinkAddr2 = file2Path & "#'" & differences(i).SheetName & "'!" & differences(i).CellAddress
                .Hyperlinks.Add Anchor:=.Cells(row, 8), Address:=hyperlinkAddr2, TextToDisplay:="移動"
                With .Cells(row, 8)
                    .Font.Color = RGB(0, 102, 204)
                    .Font.Underline = xlUnderlineStyleSingle
                    .HorizontalAlignment = xlCenter
                End With
            Else
                .Cells(row, 7).Value = "-"
                .Cells(row, 8).Value = "-"
                .Cells(row, 7).HorizontalAlignment = xlCenter
                .Cells(row, 8).HorizontalAlignment = xlCenter
            End If

            ' 差異タイプによって行に色を付ける
            Select Case differences(i).DiffType
                Case "変更"
                    .Range(.Cells(row, 1), .Cells(row, 8)).Interior.Color = COLOR_CHANGED
                Case "追加", "シート追加"
                    .Range(.Cells(row, 1), .Cells(row, 8)).Interior.Color = COLOR_ADDED
                Case "削除", "シート削除"
                    .Range(.Cells(row, 1), .Cells(row, 8)).Interior.Color = COLOR_DELETED
            End Select
        Next i

        ' 列幅調整
        .Columns("A").ColumnWidth = 8
        .Columns("B").ColumnWidth = 20
        .Columns("C").ColumnWidth = 10
        .Columns("D").ColumnWidth = 12
        .Columns("E").ColumnWidth = 30
        .Columns("F").ColumnWidth = 30
        .Columns("G").ColumnWidth = 10
        .Columns("H").ColumnWidth = 10

        ' フィルターを設定
        .Range("A10:H10").AutoFilter

        ' ウィンドウ枠の固定
        .Rows(11).Select
        ActiveWindow.FreezePanes = True

        ' セルA1を選択
        .Range("A1").Select
    End With
End Sub
