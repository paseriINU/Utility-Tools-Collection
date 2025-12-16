'==============================================================================
' Excel/Word ファイル比較ツール
' モジュール名: FileComparator
'==============================================================================
' 概要:
'   2つのExcelファイルまたはWordファイルを比較し、差異を一覧表示するツールです。
'   1つ目のファイル選択で自動的にファイルタイプを判定し、
'   2つ目は同じタイプのファイルのみ選択可能です。
'
' 機能:
'   - 1つ目のファイル選択でExcel/Wordを自動判定
'   - 2つ目は同じタイプのファイルのみ選択可能
'   - Excel: シート単位・セル単位での差異検出
'   - Word: 段落単位での差異検出
'   - 差異の種類を識別（値変更、追加、削除）
'   - 結果を新しいシートに出力
'   - 差異セルのハイライト表示
'
' 必要な環境:
'   - Microsoft Excel 2010以降
'   - Microsoft Word 2010以降（Word比較を使用する場合）
'
' 作成日: 2025-12-11
'==============================================================================

Option Explicit

'==============================================================================
' 設定: ここを編集してください
'==============================================================================
' 差異ハイライト色
Private Const COLOR_CHANGED As Long = 65535      ' 黄色: 値変更
Private Const COLOR_ADDED As Long = 5296274      ' 緑: 追加
Private Const COLOR_DELETED As Long = 13421823   ' 赤: 削除

'==============================================================================
' データ構造: Excel比較用
'==============================================================================
Private Type ExcelDifferenceInfo
    SheetName As String      ' シート名
    CellAddress As String    ' セルアドレス
    DiffType As String       ' 差異タイプ（変更/追加/削除）
    OldValue As String       ' 旧ファイルの値
    NewValue As String       ' 新ファイルの値
End Type

'==============================================================================
' データ構造: Word比較用
'==============================================================================
Private Type WordDifferenceInfo
    ParagraphNo As Long      ' 段落番号
    DiffType As String       ' 差異タイプ（変更/追加/削除/スタイル変更）
    OldText As String        ' 旧ファイルのテキスト
    NewText As String        ' 新ファイルのテキスト
    OldStyle As String       ' 旧ファイルのスタイル情報
    NewStyle As String       ' 新ファイルのスタイル情報
End Type

'==============================================================================
' Excel専用比較プロシージャ（ボタン用）
'==============================================================================
Public Sub CompareExcelFiles()
    Dim file1Path As String
    Dim file2Path As String

    On Error GoTo ErrorHandler

    ' 1つ目のExcelファイル選択
    MsgBox "2つのExcelファイルを比較します。" & vbCrLf & vbCrLf & _
           "まず、1つ目のExcelファイル（旧ファイル）を選択してください。", _
           vbInformation, "Excel ファイル比較"

    file1Path = SelectExcelFile("1つ目のExcelファイル（旧ファイル）を選択")
    If file1Path = "" Then
        MsgBox "ファイル選択がキャンセルされました。", vbExclamation
        Exit Sub
    End If

    ' 2つ目のExcelファイル選択
    MsgBox "次に、2つ目のExcelファイル（新ファイル）を選択してください。", _
           vbInformation, "Excel ファイル比較"

    file2Path = SelectExcelFile("2つ目のExcelファイル（新ファイル）を選択")
    If file2Path = "" Then
        MsgBox "ファイル選択がキャンセルされました。", vbExclamation
        Exit Sub
    End If

    ' 同じファイルが選択された場合
    If LCase(file1Path) = LCase(file2Path) Then
        MsgBox "同じファイルが選択されました。異なるファイルを選択してください。", vbExclamation
        Exit Sub
    End If

    ' Excel比較を実行
    CompareExcelFilesInternal file1Path, file2Path

    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & vbCrLf & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, vbCritical, "エラー"
End Sub

'==============================================================================
' Word専用比較プロシージャ（ボタン用）
'==============================================================================
Public Sub CompareWordFiles()
    Dim file1Path As String
    Dim file2Path As String

    On Error GoTo ErrorHandler

    ' 1つ目のWordファイル選択
    MsgBox "2つのWordファイルを比較します。" & vbCrLf & vbCrLf & _
           "まず、1つ目のWordファイル（旧ファイル）を選択してください。", _
           vbInformation, "Word ファイル比較"

    file1Path = SelectWordFile("1つ目のWordファイル（旧ファイル）を選択")
    If file1Path = "" Then
        MsgBox "ファイル選択がキャンセルされました。", vbExclamation
        Exit Sub
    End If

    ' 2つ目のWordファイル選択
    MsgBox "次に、2つ目のWordファイル（新ファイル）を選択してください。", _
           vbInformation, "Word ファイル比較"

    file2Path = SelectWordFile("2つ目のWordファイル（新ファイル）を選択")
    If file2Path = "" Then
        MsgBox "ファイル選択がキャンセルされました。", vbExclamation
        Exit Sub
    End If

    ' 同じファイルが選択された場合
    If LCase(file1Path) = LCase(file2Path) Then
        MsgBox "同じファイルが選択されました。異なるファイルを選択してください。", vbExclamation
        Exit Sub
    End If

    ' Word比較を実行
    CompareWordFilesInternal file1Path, file2Path

    Exit Sub

ErrorHandler:
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
' Wordファイル選択ダイアログ
'==============================================================================
Private Function SelectWordFile(ByVal dialogTitle As String) As String
    Dim fd As Object

    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Title = dialogTitle
        .Filters.Clear
        .Filters.Add "Word ファイル", "*.docx;*.docm;*.doc"
        .Filters.Add "すべてのファイル", "*.*"
        .FilterIndex = 1
        .AllowMultiSelect = False

        If .Show = -1 Then
            SelectWordFile = .SelectedItems(1)
        Else
            SelectWordFile = ""
        End If
    End With
End Function

'==============================================================================
' Excel比較の内部処理
'==============================================================================
Private Sub CompareExcelFilesInternal(ByVal file1Path As String, ByVal file2Path As String)
    Dim wb1 As Workbook
    Dim wb2 As Workbook
    Dim differences() As ExcelDifferenceInfo
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
' Word比較の内部処理
'==============================================================================
Private Sub CompareWordFilesInternal(ByVal file1Path As String, ByVal file2Path As String)
    Dim wordApp As Object
    Dim doc1 As Object
    Dim doc2 As Object
    Dim differences() As WordDifferenceInfo
    Dim diffCount As Long
    Dim wordWasRunning As Boolean

    On Error GoTo ErrorHandler

    ' 処理開始
    Application.ScreenUpdating = False

    Debug.Print "========================================="
    Debug.Print "Word ファイル比較を開始します"
    Debug.Print "旧ファイル: " & file1Path
    Debug.Print "新ファイル: " & file2Path
    Debug.Print "========================================="

    ' Wordアプリケーションを取得または起動
    On Error Resume Next
    Set wordApp = GetObject(, "Word.Application")
    If wordApp Is Nothing Then
        Set wordApp = CreateObject("Word.Application")
        wordWasRunning = False
    Else
        wordWasRunning = True
    End If
    On Error GoTo ErrorHandler

    wordApp.Visible = False
    wordApp.DisplayAlerts = False

    ' ファイルを開く
    Set doc1 = wordApp.Documents.Open(file1Path, ReadOnly:=True)
    Set doc2 = wordApp.Documents.Open(file2Path, ReadOnly:=True)

    ' 比較実行
    diffCount = 0
    ReDim differences(0 To 0)

    CompareWordDocuments doc1, doc2, differences, diffCount

    ' ドキュメントを閉じる
    doc1.Close SaveChanges:=False
    doc2.Close SaveChanges:=False

    ' Wordを終了（元々起動していなかった場合のみ）
    If Not wordWasRunning Then
        wordApp.Quit
    End If

    Set doc1 = Nothing
    Set doc2 = Nothing
    Set wordApp = Nothing

    ' 結果を出力
    If diffCount > 0 Then
        CreateWordResultSheet differences, diffCount, file1Path, file2Path

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
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True

    ' 開いたドキュメントを閉じる
    On Error Resume Next
    If Not doc1 Is Nothing Then doc1.Close SaveChanges:=False
    If Not doc2 Is Nothing Then doc2.Close SaveChanges:=False
    If Not wordApp Is Nothing And Not wordWasRunning Then wordApp.Quit
    On Error GoTo 0

    MsgBox "エラーが発生しました: " & vbCrLf & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, vbCritical, "エラー"
End Sub

'==============================================================================
' ワークブックを比較（Excel）
'==============================================================================
Private Sub CompareWorkbooks(ByRef wb1 As Workbook, ByRef wb2 As Workbook, _
                             ByRef differences() As ExcelDifferenceInfo, ByRef diffCount As Long)
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

'==============================================================================
' シートを比較（Excel）
'==============================================================================
Private Sub CompareSheets(ByRef ws1 As Worksheet, ByRef ws2 As Worksheet, _
                          ByRef differences() As ExcelDifferenceInfo, ByRef diffCount As Long)
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

'==============================================================================
' Word文書を段落単位で比較（テキスト＋スタイル）
'==============================================================================
Private Sub CompareWordDocuments(ByRef doc1 As Object, ByRef doc2 As Object, _
                                  ByRef differences() As WordDifferenceInfo, ByRef diffCount As Long)
    Dim paraCount1 As Long
    Dim paraCount2 As Long
    Dim maxParas As Long
    Dim i As Long
    Dim text1 As String
    Dim text2 As String
    Dim style1 As String
    Dim style2 As String

    paraCount1 = doc1.Paragraphs.Count
    paraCount2 = doc2.Paragraphs.Count

    Debug.Print "旧ファイル段落数: " & paraCount1
    Debug.Print "新ファイル段落数: " & paraCount2

    ' 比較範囲を決定（使用範囲のみ比較、制限なし）
    maxParas = Application.WorksheetFunction.Max(paraCount1, paraCount2)

    Debug.Print "比較範囲: " & maxParas & " 段落"

    ' 段落単位で比較
    For i = 1 To maxParas
        ' 旧ファイルのテキストとスタイル取得
        If i <= paraCount1 Then
            text1 = CleanText(doc1.Paragraphs(i).Range.Text)
            style1 = GetParagraphStyleInfo(doc1.Paragraphs(i))
        Else
            text1 = ""
            style1 = ""
        End If

        ' 新ファイルのテキストとスタイル取得
        If i <= paraCount2 Then
            text2 = CleanText(doc2.Paragraphs(i).Range.Text)
            style2 = GetParagraphStyleInfo(doc2.Paragraphs(i))
        Else
            text2 = ""
            style2 = ""
        End If

        ' テキスト比較
        If text1 <> text2 Then
            If Len(text1) = 0 And Len(text2) > 0 Then
                ' 追加
                AddWordDifferenceWithStyle differences, diffCount, i, "追加", "(空)", text2, "", style2
            ElseIf Len(text1) > 0 And Len(text2) = 0 Then
                ' 削除
                AddWordDifferenceWithStyle differences, diffCount, i, "削除", text1, "(空)", style1, ""
            Else
                ' 変更
                AddWordDifferenceWithStyle differences, diffCount, i, "変更", text1, text2, style1, style2
            End If
        ' テキストが同じでもスタイルが異なる場合
        ElseIf style1 <> style2 And Len(text1) > 0 Then
            AddWordDifferenceWithStyle differences, diffCount, i, "スタイル変更", text1, text2, style1, style2
        End If

        ' 進捗表示（100段落ごと）
        If i Mod 100 = 0 Then
            Debug.Print "  " & i & " / " & maxParas & " 段落処理中..."
            DoEvents
        End If
    Next i

    ' 段落数の違いを報告
    If paraCount1 <> paraCount2 Then
        If paraCount1 > paraCount2 Then
            For i = paraCount2 + 1 To paraCount1
                text1 = CleanText(doc1.Paragraphs(i).Range.Text)
                style1 = GetParagraphStyleInfo(doc1.Paragraphs(i))
                If Len(text1) > 0 Then
                    AddWordDifferenceWithStyle differences, diffCount, i, "削除", text1, "(段落なし)", style1, ""
                End If
            Next i
        Else
            For i = paraCount1 + 1 To paraCount2
                text2 = CleanText(doc2.Paragraphs(i).Range.Text)
                style2 = GetParagraphStyleInfo(doc2.Paragraphs(i))
                If Len(text2) > 0 Then
                    AddWordDifferenceWithStyle differences, diffCount, i, "追加", "(段落なし)", text2, "", style2
                End If
            Next i
        End If
    End If
End Sub

'==============================================================================
' 段落のスタイル情報を取得
'==============================================================================
Private Function GetParagraphStyleInfo(ByRef para As Object) As String
    Dim styleInfo As String
    Dim fontName As String
    Dim fontSize As Single
    Dim isBold As Boolean
    Dim isItalic As Boolean
    Dim styleName As String

    On Error Resume Next

    ' スタイル名
    styleName = para.Style.NameLocal
    If Err.Number <> 0 Then styleName = "(不明)"
    Err.Clear

    ' フォント情報（段落の最初の文字から取得）
    fontName = para.Range.Font.Name
    If Err.Number <> 0 Or fontName = "" Then fontName = "(混在)"
    Err.Clear

    fontSize = para.Range.Font.Size
    If Err.Number <> 0 Or fontSize = 9999999 Then
        fontSize = 0
    End If
    Err.Clear

    ' 太字・斜体（wdUndefined=-9999999の場合は混在）
    isBold = (para.Range.Font.Bold = True)
    isItalic = (para.Range.Font.Italic = True)

    On Error GoTo 0

    ' スタイル情報を文字列化
    styleInfo = "[" & styleName & "] " & fontName & " " & Format(fontSize, "0.0") & "pt"
    If isBold Then styleInfo = styleInfo & " 太字"
    If isItalic Then styleInfo = styleInfo & " 斜体"

    GetParagraphStyleInfo = styleInfo
End Function

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
' テキストをクリーンアップ（改行・特殊文字を除去）
'==============================================================================
Private Function CleanText(ByVal txt As String) As String
    ' 改行・段落記号を除去
    txt = Replace(txt, vbCr, "")
    txt = Replace(txt, vbLf, "")
    txt = Replace(txt, Chr(13), "")
    txt = Replace(txt, Chr(11), " ")  ' 行区切り
    txt = Replace(txt, Chr(7), "")    ' セル終端記号

    ' 前後の空白を除去
    CleanText = Trim(txt)
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
' Excel差異を追加
'==============================================================================
Private Sub AddExcelDifference(ByRef differences() As ExcelDifferenceInfo, ByRef diffCount As Long, _
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
' Word差異を追加
'==============================================================================
Private Sub AddWordDifference(ByRef differences() As WordDifferenceInfo, ByRef diffCount As Long, _
                               ByVal paraNo As Long, ByVal diffType As String, _
                               ByVal oldText As String, ByVal newText As String)
    ' スタイル情報なしで追加
    AddWordDifferenceWithStyle differences, diffCount, paraNo, diffType, oldText, newText, "", ""
End Sub

'==============================================================================
' Word差異を追加（スタイル情報付き）
'==============================================================================
Private Sub AddWordDifferenceWithStyle(ByRef differences() As WordDifferenceInfo, ByRef diffCount As Long, _
                               ByVal paraNo As Long, ByVal diffType As String, _
                               ByVal oldText As String, ByVal newText As String, _
                               ByVal oldStyle As String, ByVal newStyle As String)
    ' 配列を拡張
    If diffCount = 0 Then
        ReDim differences(0 To 0)
    Else
        ReDim Preserve differences(0 To diffCount)
    End If

    ' 差異情報を格納
    With differences(diffCount)
        .ParagraphNo = paraNo
        .DiffType = diffType
        .OldText = Left(oldText, 500)  ' 長すぎるテキストを切り詰め
        .NewText = Left(newText, 500)
        .OldStyle = oldStyle
        .NewStyle = newStyle
    End With

    diffCount = diffCount + 1
End Sub

'==============================================================================
' Excel結果シートを作成
'==============================================================================
Private Sub CreateExcelResultSheet(ByRef differences() As ExcelDifferenceInfo, ByVal diffCount As Long, _
                              ByVal file1Path As String, ByVal file2Path As String)
    Dim ws As Worksheet
    Dim i As Long
    Dim row As Long
    Dim hyperlinkAddr1 As String
    Dim hyperlinkAddr2 As String

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
        .Columns("A").ColumnWidth = 6
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

'==============================================================================
' Word結果シートを作成
'==============================================================================
Private Sub CreateWordResultSheet(ByRef differences() As WordDifferenceInfo, ByVal diffCount As Long, _
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
        .Range("A1").Value = "Word ファイル比較結果"
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
        .Range("E8").Value = "スタイル変更"
        .Range("E8").Interior.Color = RGB(204, 153, 255)  ' 薄紫

        ' ヘッダー
        .Range("A10").Value = "No"
        .Range("B10").Value = "段落番号"
        .Range("C10").Value = "差異タイプ"
        .Range("D10").Value = "旧ファイルのテキスト"
        .Range("E10").Value = "新ファイルのテキスト"
        .Range("F10").Value = "旧スタイル"
        .Range("G10").Value = "新スタイル"

        ' ヘッダー書式
        With .Range("A10:G10")
            .Font.Bold = True
            .Interior.Color = RGB(68, 114, 196)
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
        End With

        ' データ行
        For i = 0 To diffCount - 1
            row = i + 11

            .Cells(row, 1).Value = i + 1
            .Cells(row, 2).Value = differences(i).ParagraphNo
            .Cells(row, 3).Value = differences(i).DiffType
            .Cells(row, 4).Value = differences(i).OldText
            .Cells(row, 5).Value = differences(i).NewText
            .Cells(row, 6).Value = differences(i).OldStyle
            .Cells(row, 7).Value = differences(i).NewStyle

            ' テキストを折り返し
            .Cells(row, 4).WrapText = True
            .Cells(row, 5).WrapText = True

            ' 差異タイプによって行に色を付ける
            Select Case differences(i).DiffType
                Case "変更"
                    .Range(.Cells(row, 1), .Cells(row, 7)).Interior.Color = COLOR_CHANGED
                Case "追加"
                    .Range(.Cells(row, 1), .Cells(row, 7)).Interior.Color = COLOR_ADDED
                Case "削除"
                    .Range(.Cells(row, 1), .Cells(row, 7)).Interior.Color = COLOR_DELETED
                Case "スタイル変更"
                    .Range(.Cells(row, 1), .Cells(row, 7)).Interior.Color = RGB(204, 153, 255)  ' 薄紫
            End Select
        Next i

        ' 列幅調整
        .Columns("A").ColumnWidth = 6
        .Columns("B").ColumnWidth = 10
        .Columns("C").ColumnWidth = 14
        .Columns("D").ColumnWidth = 40
        .Columns("E").ColumnWidth = 40
        .Columns("F").ColumnWidth = 30
        .Columns("G").ColumnWidth = 30

        ' フィルターを設定
        .Range("A10:G10").AutoFilter

        ' ウィンドウ枠の固定
        .Rows(11).Select
        ActiveWindow.FreezePanes = True

        ' セルA1を選択
        .Range("A1").Select
    End With
End Sub

'==============================================================================
' メインシート初期化
'==============================================================================
Public Sub InitializeFileComparator()
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
        .Rows(7).RowHeight = 15
        .Rows(8).RowHeight = 50

        ' Excel比較ボタン（緑系）- 角丸四角形
        Set shp = .Shapes.AddShape(msoShapeRoundedRectangle, _
            .Range("B8").Left + 5, .Range("B8").Top + 5, 120, 40)
        With shp
            .Name = "btnCompareExcel"
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

        ' Word比較ボタン（青系）- 角丸四角形
        Set shp = .Shapes.AddShape(msoShapeRoundedRectangle, _
            .Range("E8").Left + 5, .Range("E8").Top + 5, 120, 40)
        With shp
            .Name = "btnCompareWord"
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
        .Range("D21").Value = "全段落・スタイル比較（制限なし）"

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
        .Range("C35").Value = "比較結果が「CompareResult」シートに出力されます"

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
