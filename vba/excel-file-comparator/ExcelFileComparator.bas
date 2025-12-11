'==============================================================================
' Excel ファイル比較ツール
' モジュール名: ExcelFileComparator
'==============================================================================
' 概要:
'   2つのExcelファイルを比較し、差異を一覧表示するツールです。
'
' 機能:
'   - ファイル選択ダイアログで2つのExcelファイルを指定
'   - シート単位・セル単位での差異検出
'   - 差異の種類を識別（値変更、追加、削除）
'   - 結果を新しいシートに出力
'   - 差異セルのハイライト表示
'
' 必要な環境:
'   - Microsoft Excel 2010以降
'
' 作成日: 2025-12-11
'==============================================================================

Option Explicit

'==============================================================================
' 設定: ここを編集してください
'==============================================================================
' 比較する最大行数（パフォーマンス対策）
Private Const MAX_ROWS As Long = 10000

' 比較する最大列数
Private Const MAX_COLS As Long = 256

' 差異ハイライト色
Private Const COLOR_CHANGED As Long = 65535      ' 黄色: 値変更
Private Const COLOR_ADDED As Long = 5296274      ' 緑: 追加
Private Const COLOR_DELETED As Long = 13421823   ' 赤: 削除

'==============================================================================
' データ構造
'==============================================================================
Private Type DifferenceInfo
    SheetName As String      ' シート名
    CellAddress As String    ' セルアドレス
    DiffType As String       ' 差異タイプ（変更/追加/削除）
    OldValue As String       ' 旧ファイルの値
    NewValue As String       ' 新ファイルの値
End Type

'==============================================================================
' メインプロシージャ: Excel ファイルを比較
'==============================================================================
Public Sub CompareExcelFiles()
    Dim file1Path As String
    Dim file2Path As String
    Dim wb1 As Workbook
    Dim wb2 As Workbook
    Dim differences() As DifferenceInfo
    Dim diffCount As Long

    On Error GoTo ErrorHandler

    ' ファイル選択
    MsgBox "2つのExcelファイルを比較します。" & vbCrLf & vbCrLf & _
           "まず、1つ目のファイル（旧ファイル）を選択してください。", _
           vbInformation, "Excel ファイル比較ツール"

    file1Path = SelectExcelFile("1つ目のファイル（旧ファイル）を選択")
    If file1Path = "" Then
        MsgBox "ファイル選択がキャンセルされました。", vbExclamation
        Exit Sub
    End If

    MsgBox "次に、2つ目のファイル（新ファイル）を選択してください。", _
           vbInformation, "Excel ファイル比較ツール"

    file2Path = SelectExcelFile("2つ目のファイル（新ファイル）を選択")
    If file2Path = "" Then
        MsgBox "ファイル選択がキャンセルされました。", vbExclamation
        Exit Sub
    End If

    ' 同じファイルが選択された場合
    If LCase(file1Path) = LCase(file2Path) Then
        MsgBox "同じファイルが選択されました。異なるファイルを選択してください。", vbExclamation
        Exit Sub
    End If

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
        CreateResultSheet differences, diffCount, file1Path, file2Path

        Debug.Print "========================================="
        Debug.Print "処理完了: " & diffCount & " 件の差異を検出"
        Debug.Print "========================================="

        MsgBox "比較が完了しました。" & vbCrLf & vbCrLf & _
               "検出された差異: " & diffCount & " 件" & vbCrLf & vbCrLf & _
               "結果は「CompareResult」シートに出力されました。", _
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

'==============================================================================
' Excelファイル選択ダイアログ
'==============================================================================
Private Function SelectExcelFile(ByVal dialogTitle As String) As String
    Dim fd As Object

    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Title = dialogTitle
        .Filters.Clear
        .Filters.Add "Excel ファイル", "*.xlsx;*.xlsm;*.xls;*.xlsb"
        .Filters.Add "すべてのファイル", "*.*"
        .FilterIndex = 1
        .AllowMultiSelect = False

        If .Show = -1 Then
            SelectExcelFile = .SelectedItems(1)
        Else
            SelectExcelFile = ""
        End If
    End With
End Function

'==============================================================================
' ワークブックを比較
'==============================================================================
Private Sub CompareWorkbooks(ByRef wb1 As Workbook, ByRef wb2 As Workbook, _
                             ByRef differences() As DifferenceInfo, ByRef diffCount As Long)
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim sheetNames1 As Object
    Dim sheetNames2 As Object
    Dim sheetName As Variant
    Dim found As Boolean

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
            AddDifference differences, diffCount, CStr(sheetName), "(シート全体)", _
                          "シート削除", "(存在)", "(削除済み)"
        End If
    Next sheetName

    ' wb2のみに存在するシート（追加されたシート）
    For Each sheetName In sheetNames2.Keys
        If Not sheetNames1.exists(sheetName) Then
            AddDifference differences, diffCount, CStr(sheetName), "(シート全体)", _
                          "シート追加", "(なし)", "(追加済み)"
        End If
    Next sheetName
End Sub

'==============================================================================
' シートを比較
'==============================================================================
Private Sub CompareSheets(ByRef ws1 As Worksheet, ByRef ws2 As Worksheet, _
                          ByRef differences() As DifferenceInfo, ByRef diffCount As Long)
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

    ' 比較範囲を決定
    maxRow = Application.WorksheetFunction.Max(lastRow1, lastRow2)
    maxCol = Application.WorksheetFunction.Max(lastCol1, lastCol2)

    ' 最大値を制限
    If maxRow > MAX_ROWS Then maxRow = MAX_ROWS
    If maxCol > MAX_COLS Then maxCol = MAX_COLS

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
                    AddDifference differences, diffCount, ws1.Name, cellAddr, _
                                  "追加", "(空)", CStr(val2)
                ElseIf Not IsEmpty(val1) And IsEmpty(val2) Then
                    ' 新ファイルで削除
                    AddDifference differences, diffCount, ws1.Name, cellAddr, _
                                  "削除", CStr(val1), "(空)"
                Else
                    ' 値の変更
                    AddDifference differences, diffCount, ws1.Name, cellAddr, _
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

'==============================================================================
' 値の比較（数値の微小差異を考慮）
'==============================================================================
Private Function IsEqual(ByVal val1 As Variant, ByVal val2 As Variant) As Boolean
    ' 両方Empty
    If IsEmpty(val1) And IsEmpty(val2) Then
        IsEqual = True
        Exit Function
    End If

    ' 片方がEmpty
    If IsEmpty(val1) Or IsEmpty(val2) Then
        IsEqual = False
        Exit Function
    End If

    ' 両方数値の場合、浮動小数点誤差を考慮
    If IsNumeric(val1) And IsNumeric(val2) Then
        If Abs(CDbl(val1) - CDbl(val2)) < 0.0000001 Then
            IsEqual = True
        Else
            IsEqual = False
        End If
        Exit Function
    End If

    ' 文字列比較
    IsEqual = (CStr(val1) = CStr(val2))
End Function

'==============================================================================
' 最終行を取得
'==============================================================================
Private Function GetLastRow(ByRef ws As Worksheet) As Long
    On Error Resume Next
    GetLastRow = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
                              LookIn:=xlFormulas, LookAt:=xlPart, _
                              SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    If Err.Number <> 0 Or GetLastRow = 0 Then
        GetLastRow = 1
    End If
    On Error GoTo 0
End Function

'==============================================================================
' 最終列を取得
'==============================================================================
Private Function GetLastCol(ByRef ws As Worksheet) As Long
    On Error Resume Next
    GetLastCol = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
                              LookIn:=xlFormulas, LookAt:=xlPart, _
                              SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    If Err.Number <> 0 Or GetLastCol = 0 Then
        GetLastCol = 1
    End If
    On Error GoTo 0
End Function

'==============================================================================
' 差異を追加
'==============================================================================
Private Sub AddDifference(ByRef differences() As DifferenceInfo, ByRef diffCount As Long, _
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

'==============================================================================
' 結果シートを作成
'==============================================================================
Private Sub CreateResultSheet(ByRef differences() As DifferenceInfo, ByVal diffCount As Long, _
                              ByVal file1Path As String, ByVal file2Path As String)
    Dim ws As Worksheet
    Dim i As Long
    Dim row As Long

    ' 既存の結果シートがあれば削除
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("CompareResult").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' 新しいシートを作成
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = "CompareResult"

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

        ' ヘッダー書式
        With .Range("A10:F10")
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

            ' 差異タイプによって行に色を付ける
            Select Case differences(i).DiffType
                Case "変更"
                    .Range(.Cells(row, 1), .Cells(row, 6)).Interior.Color = COLOR_CHANGED
                Case "追加", "シート追加"
                    .Range(.Cells(row, 1), .Cells(row, 6)).Interior.Color = COLOR_ADDED
                Case "削除", "シート削除"
                    .Range(.Cells(row, 1), .Cells(row, 6)).Interior.Color = COLOR_DELETED
            End Select
        Next i

        ' 列幅調整
        .Columns("A").ColumnWidth = 6
        .Columns("B").ColumnWidth = 20
        .Columns("C").ColumnWidth = 10
        .Columns("D").ColumnWidth = 12
        .Columns("E").ColumnWidth = 30
        .Columns("F").ColumnWidth = 30

        ' フィルターを設定
        .Range("A10:F10").AutoFilter

        ' ウィンドウ枠の固定
        .Rows(11).Select
        ActiveWindow.FreezePanes = True

        ' セルA1を選択
        .Range("A1").Select
    End With
End Sub

'==============================================================================
' クイック比較: 選択セル範囲のみを比較
'==============================================================================
Public Sub QuickCompareSelectedRange()
    Dim rng As Range
    Dim file2Path As String
    Dim wb2 As Workbook
    Dim ws2 As Worksheet
    Dim cell As Range
    Dim val1 As Variant, val2 As Variant
    Dim diffCount As Long
    Dim currentSheetName As String

    On Error GoTo ErrorHandler

    ' 選択範囲を取得
    Set rng = Selection
    If rng Is Nothing Then
        MsgBox "比較するセル範囲を選択してください。", vbExclamation
        Exit Sub
    End If

    currentSheetName = ActiveSheet.Name

    ' 比較対象ファイルを選択
    MsgBox "選択した範囲を別のExcelファイルと比較します。" & vbCrLf & vbCrLf & _
           "比較対象のファイルを選択してください。" & vbCrLf & _
           "（同じシート名・同じセル位置で比較します）", _
           vbInformation, "クイック比較"

    file2Path = SelectExcelFile("比較対象ファイルを選択")
    If file2Path = "" Then
        MsgBox "ファイル選択がキャンセルされました。", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' ファイルを開く
    Set wb2 = Workbooks.Open(file2Path, ReadOnly:=True)

    ' 同名シートの存在確認
    On Error Resume Next
    Set ws2 = wb2.Worksheets(currentSheetName)
    On Error GoTo ErrorHandler

    If ws2 Is Nothing Then
        wb2.Close SaveChanges:=False
        MsgBox "比較対象ファイルに「" & currentSheetName & "」シートが見つかりません。", vbExclamation
        GoTo Cleanup
    End If

    ' 差異をハイライト
    diffCount = 0
    For Each cell In rng
        val1 = cell.Value
        val2 = ws2.Range(cell.Address).Value

        If Not IsEqual(val1, val2) Then
            ' 差異があるセルをハイライト
            If IsEmpty(val1) And Not IsEmpty(val2) Then
                cell.Interior.Color = COLOR_ADDED
            ElseIf Not IsEmpty(val1) And IsEmpty(val2) Then
                cell.Interior.Color = COLOR_DELETED
            Else
                cell.Interior.Color = COLOR_CHANGED
            End If

            ' コメントで差異を表示
            On Error Resume Next
            cell.ClearComments
            cell.AddComment "旧: " & CStr(val1) & vbCrLf & "新: " & CStr(val2)
            On Error GoTo ErrorHandler

            diffCount = diffCount + 1
        End If
    Next cell

    ' ファイルを閉じる
    wb2.Close SaveChanges:=False

    MsgBox "クイック比較が完了しました。" & vbCrLf & vbCrLf & _
           "検出された差異: " & diffCount & " 件" & vbCrLf & vbCrLf & _
           "差異のあるセルはハイライト表示されています。" & vbCrLf & _
           "セルにマウスを合わせるとコメントで詳細を確認できます。", _
           vbInformation, "処理完了"

Cleanup:
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True

    On Error Resume Next
    If Not wb2 Is Nothing Then wb2.Close SaveChanges:=False
    On Error GoTo 0

    MsgBox "エラーが発生しました: " & vbCrLf & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, vbCritical, "エラー"
End Sub

'==============================================================================
' ハイライトをクリア
'==============================================================================
Public Sub ClearHighlight()
    Dim ws As Worksheet
    Dim cell As Range

    Set ws = ActiveSheet

    ' 背景色をクリア
    ws.Cells.Interior.ColorIndex = xlNone

    ' コメントをクリア
    On Error Resume Next
    ws.Cells.ClearComments
    On Error GoTo 0

    MsgBox "ハイライトとコメントをクリアしました。", vbInformation, "処理完了"
End Sub

'==============================================================================
' テスト用プロシージャ
'==============================================================================
Public Sub TestCompareExcelFiles()
    CompareExcelFiles
End Sub
