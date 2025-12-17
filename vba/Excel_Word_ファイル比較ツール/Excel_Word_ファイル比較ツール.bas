'==============================================================================
' Excel/Word ファイル比較ツール
' モジュール名: Excel_Word_ファイル比較ツール
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
'   - Word: WinMerge方式（LCSアルゴリズム）での差異検出
'     * 1行追加/削除があっても以降の行がすべて差異にならない
'     * 実際に変更・追加・削除された行のみを正確に検出
'     * 旧ファイルと新ファイルの行番号を両方表示
'   - 差異の種類を識別（値変更、追加、削除、スタイル変更）
'   - 結果を新しいシートに出力
'   - 差異セルのハイライト表示
'
' 必要な環境:
'   - Microsoft Excel 2010以降
'   - Microsoft Word 2010以降（Word比較を使用する場合）
'
' 注意:
'   - 初期化とシートフォーマット機能は Excel_Word_ファイル比較ツール_Setup.bas に分離されています
'   - COLOR_* 定数は Setup モジュールで定義されています
'
' 作成日: 2025-12-11
' 更新日: 2025-12-17 - Word比較のパフォーマンス最適化
'          - スタイル情報の遅延取得（差分段落のみ取得）
'          - ハッシュによる同一テキスト事前マッチング
'          - DoEvents呼び出し頻度の最適化
'==============================================================================

Option Explicit

'==============================================================================
' モジュールレベル変数: テキスト一致段落のスタイル比較用
'==============================================================================
Private g_MatchedOld() As Long    ' 旧ファイルの段落番号
Private g_MatchedNew() As Long    ' 新ファイルの段落番号
Private g_MatchedCount As Long    ' ペア数

'==============================================================================
' 進捗表示用ヘルパー関数
'==============================================================================
Private Sub ShowProgress(ByVal phase As String, ByVal current As Long, ByVal total As Long)
    Dim pct As Long
    Dim progressBar As String
    Dim barLength As Long
    Dim filledLength As Long
    Dim i As Long

    If total > 0 Then
        pct = CLng((current / total) * 100)
    Else
        pct = 0
    End If

    ' プログレスバー（20文字幅）
    barLength = 20
    filledLength = CLng(barLength * current / IIf(total > 0, total, 1))
    progressBar = ""
    For i = 1 To filledLength
        progressBar = progressBar & ChrW(&H2588)  ' █
    Next i
    For i = filledLength + 1 To barLength
        progressBar = progressBar & ChrW(&H2591)  ' ░
    Next i

    Application.StatusBar = phase & " " & progressBar & " " & pct & "% (" & current & "/" & total & ")"
    DoEvents
End Sub

Private Sub ClearProgress()
    Application.StatusBar = False
End Sub

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
' データ構造: Word比較用（WinMerge方式：旧/新両方の行番号を保持）
'==============================================================================
Private Type WordDifferenceInfo
    OldParagraphNo As Long   ' 旧ファイルの段落番号（0は該当なし）
    NewParagraphNo As Long   ' 新ファイルの段落番号（0は該当なし）
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
    ClearProgress
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    ClearProgress

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
' Word文書を段落単位で比較（WinMerge方式：LCSアルゴリズム使用）
' 【最適化版】スタイル情報は差分検出後に必要な段落のみ取得
'==============================================================================
Private Sub CompareWordDocuments(ByRef doc1 As Object, ByRef doc2 As Object, _
                                  ByRef differences() As WordDifferenceInfo, ByRef diffCount As Long)
    Dim paraCount1 As Long
    Dim paraCount2 As Long
    Dim texts1() As String
    Dim texts2() As String
    Dim i As Long

    paraCount1 = doc1.Paragraphs.Count
    paraCount2 = doc2.Paragraphs.Count

    Debug.Print "旧ファイル段落数: " & paraCount1
    Debug.Print "新ファイル段落数: " & paraCount2
    Debug.Print "WinMerge方式（LCSアルゴリズム）で比較します..."
    Debug.Print "【最適化モード】スタイル情報は差分段落のみ取得"

    ' 段落テキストのみを配列に取得（スタイルは後で必要な段落のみ取得）
    ReDim texts1(1 To paraCount1)
    For i = 1 To paraCount1
        texts1(i) = CleanText(doc1.Paragraphs(i).Range.Text)
        If i Mod 50 = 0 Or i = paraCount1 Then
            ShowProgress "[1/4] 旧ファイル読込", i, paraCount1
        End If
    Next i

    ReDim texts2(1 To paraCount2)
    For i = 1 To paraCount2
        texts2(i) = CleanText(doc2.Paragraphs(i).Range.Text)
        If i Mod 50 = 0 Or i = paraCount2 Then
            ShowProgress "[2/4] 新ファイル読込", i, paraCount2
        End If
    Next i

    ' LCSベースの差分検出を実行（スタイル情報なしで高速計算）
    Debug.Print "差分を計算中..."
    ComputeLCSDiffOptimized texts1, texts2, paraCount1, paraCount2, differences, diffCount

    ' 差分が検出された段落のみスタイル情報を取得（遅延評価）
    Debug.Print "差分段落のスタイル情報を取得中..."
    FetchStylesForDifferences doc1, doc2, differences, diffCount

    Debug.Print "差分計算完了: " & diffCount & " 件の差異を検出"
End Sub

'==============================================================================
' 差分が検出された段落のみスタイル情報を取得（遅延評価）
' また、テキスト一致段落のスタイル比較も実行
'==============================================================================
Private Sub FetchStylesForDifferences(ByRef doc1 As Object, ByRef doc2 As Object, _
                                       ByRef differences() As WordDifferenceInfo, ByRef diffCount As Long)
    Dim i As Long
    Dim oldParaNo As Long
    Dim newParaNo As Long
    Dim fetchedCount As Long
    Dim oldStyle As String
    Dim newStyle As String
    Dim oldText As String

    Dim totalStyleWork As Long
    totalStyleWork = diffCount + g_MatchedCount
    Dim styleProgress As Long
    styleProgress = 0

    ' 1. 既存の差分にスタイル情報を追加
    If diffCount > 0 Then
        fetchedCount = 0
        For i = 0 To diffCount - 1
            oldParaNo = differences(i).OldParagraphNo
            newParaNo = differences(i).NewParagraphNo

            ' 旧ファイルの段落スタイルを取得
            If oldParaNo > 0 And oldParaNo <= doc1.Paragraphs.Count Then
                differences(i).OldStyle = GetParagraphStyleInfo(doc1.Paragraphs(oldParaNo))
            End If

            ' 新ファイルの段落スタイルを取得
            If newParaNo > 0 And newParaNo <= doc2.Paragraphs.Count Then
                differences(i).NewStyle = GetParagraphStyleInfo(doc2.Paragraphs(newParaNo))
            End If

            fetchedCount = fetchedCount + 1
            styleProgress = styleProgress + 1
            If fetchedCount Mod 20 = 0 Or fetchedCount = diffCount Then
                ShowProgress "[4/4] スタイル取得", styleProgress, totalStyleWork
            End If
        Next i
    End If

    ' 2. テキスト一致段落のスタイル比較（スタイル変更の検出）
    If g_MatchedCount > 0 Then
        Debug.Print "  テキスト一致段落のスタイル比較: " & g_MatchedCount & " 件"

        Dim styleCheckCount As Long
        styleCheckCount = 0

        For i = 0 To g_MatchedCount - 1
            oldParaNo = g_MatchedOld(i)
            newParaNo = g_MatchedNew(i)

            ' 両方の段落が有効範囲内かチェック
            If oldParaNo > 0 And oldParaNo <= doc1.Paragraphs.Count And _
               newParaNo > 0 And newParaNo <= doc2.Paragraphs.Count Then

                oldStyle = GetParagraphStyleInfo(doc1.Paragraphs(oldParaNo))
                newStyle = GetParagraphStyleInfo(doc2.Paragraphs(newParaNo))

                ' スタイルが異なる場合は差分として追加
                If oldStyle <> newStyle Then
                    oldText = CleanText(doc1.Paragraphs(oldParaNo).Range.Text)
                    AddWordDiffNew differences, diffCount, oldParaNo, newParaNo, "スタイル変更", _
                        oldText, oldText, oldStyle, newStyle
                End If
            End If

            styleCheckCount = styleCheckCount + 1
            styleProgress = styleProgress + 1
            If styleCheckCount Mod 50 = 0 Or styleCheckCount = g_MatchedCount Then
                ShowProgress "[4/4] スタイル比較", styleProgress, totalStyleWork
            End If
        Next i

        ' モジュールレベル変数をクリア
        g_MatchedCount = 0
        Erase g_MatchedOld
        Erase g_MatchedNew
    End If

    ' 進捗表示をクリア
    ClearProgress
End Sub

'==============================================================================
' LCSベースの差分検出（最適化版：スタイル情報なし、ハッシュ事前マッチング付き）
' WinMergeのような行単位の差分を検出します
'==============================================================================
Private Sub ComputeLCSDiffOptimized(ByRef texts1() As String, ByRef texts2() As String, _
                                     ByVal n1 As Long, ByVal n2 As Long, _
                                     ByRef differences() As WordDifferenceInfo, ByRef diffCount As Long)
    Dim lcsMatrix() As Long
    Dim i As Long, j As Long
    Dim maxLen As Long
    Dim textHash1 As Object  ' Dictionary: テキスト -> 段落番号の配列
    Dim textHash2 As Object
    Dim uniqueTexts1 As Long, uniqueTexts2 As Long
    Dim commonTexts As Long

    ' ハッシュマップを作成して同一テキストを事前に特定
    Set textHash1 = CreateObject("Scripting.Dictionary")
    Set textHash2 = CreateObject("Scripting.Dictionary")

    ' 旧ファイルのテキストをハッシュ化
    For i = 1 To n1
        If Not textHash1.exists(texts1(i)) Then
            textHash1.Add texts1(i), i
        End If
    Next i
    uniqueTexts1 = textHash1.Count

    ' 新ファイルのテキストをハッシュ化
    For i = 1 To n2
        If Not textHash2.exists(texts2(i)) Then
            textHash2.Add texts2(i), i
        End If
    Next i
    uniqueTexts2 = textHash2.Count

    ' 共通テキストの数をカウント
    commonTexts = 0
    For Each i In textHash1.Keys
        If textHash2.exists(i) Then
            commonTexts = commonTexts + 1
        End If
    Next i

    Debug.Print "  ユニークテキスト数: 旧=" & uniqueTexts1 & ", 新=" & uniqueTexts2 & ", 共通=" & commonTexts

    ' LCS行列を計算（メモリ効率のため、大きなファイルでは制限）
    maxLen = Application.WorksheetFunction.Max(n1, n2)
    If maxLen > 5000 Then
        Debug.Print "警告: 段落数が多いため、簡易比較モードを使用します"
        ComputeSimpleDiffOptimized texts1, texts2, n1, n2, differences, diffCount
        Exit Sub
    End If

    ' LCS行列を初期化 (0-indexed: 0 to n)
    ReDim lcsMatrix(0 To n1, 0 To n2)

    ' LCS行列を構築（ハッシュを活用した高速比較）
    For i = 1 To n1
        For j = 1 To n2
            If texts1(i) = texts2(j) Then
                lcsMatrix(i, j) = lcsMatrix(i - 1, j - 1) + 1
            Else
                If lcsMatrix(i - 1, j) >= lcsMatrix(i, j - 1) Then
                    lcsMatrix(i, j) = lcsMatrix(i - 1, j)
                Else
                    lcsMatrix(i, j) = lcsMatrix(i, j - 1)
                End If
            End If
        Next j

        ' 進捗表示
        If i Mod 100 = 0 Or i = n1 Then
            ShowProgress "[3/4] 差分計算(LCS)", i, n1
        End If
    Next i

    ' バックトラックして差分を抽出（スタイル情報なし）
    BacktrackLCSOptimized lcsMatrix, texts1, texts2, n1, n2, differences, diffCount

    Set textHash1 = Nothing
    Set textHash2 = Nothing
End Sub

'==============================================================================
' LCSベースの差分検出（旧版：互換性のために残す）
'==============================================================================
Private Sub ComputeLCSDiff(ByRef texts1() As String, ByRef texts2() As String, _
                           ByRef styles1() As String, ByRef styles2() As String, _
                           ByVal n1 As Long, ByVal n2 As Long, _
                           ByRef differences() As WordDifferenceInfo, ByRef diffCount As Long)
    Dim lcsMatrix() As Long
    Dim i As Long, j As Long
    Dim maxLen As Long

    ' LCS行列を計算（メモリ効率のため、大きなファイルでは制限）
    maxLen = Application.WorksheetFunction.Max(n1, n2)
    If maxLen > 5000 Then
        Debug.Print "警告: 段落数が多いため、簡易比較モードを使用します"
        ComputeSimpleDiff texts1, texts2, styles1, styles2, n1, n2, differences, diffCount
        Exit Sub
    End If

    ' LCS行列を初期化 (0-indexed: 0 to n)
    ReDim lcsMatrix(0 To n1, 0 To n2)

    ' LCS行列を構築
    For i = 1 To n1
        For j = 1 To n2
            If texts1(i) = texts2(j) Then
                lcsMatrix(i, j) = lcsMatrix(i - 1, j - 1) + 1
            Else
                If lcsMatrix(i - 1, j) >= lcsMatrix(i, j - 1) Then
                    lcsMatrix(i, j) = lcsMatrix(i - 1, j)
                Else
                    lcsMatrix(i, j) = lcsMatrix(i, j - 1)
                End If
            End If
        Next j

        ' 進捗表示
        If i Mod 100 = 0 Then
            Debug.Print "  LCS計算中: " & i & " / " & n1
            DoEvents
        End If
    Next i

    ' バックトラックして差分を抽出
    BacktrackLCS lcsMatrix, texts1, texts2, styles1, styles2, n1, n2, differences, diffCount
End Sub

'==============================================================================
' LCS行列をバックトラックして差分を抽出（最適化版：スタイル情報なし）
' matchedPairs: テキストが一致した段落ペアを格納（後でスタイル比較用）
'==============================================================================
Private Sub BacktrackLCSOptimized(ByRef lcsMatrix() As Long, _
                                   ByRef texts1() As String, ByRef texts2() As String, _
                                   ByVal n1 As Long, ByVal n2 As Long, _
                                   ByRef differences() As WordDifferenceInfo, ByRef diffCount As Long)
    Dim i As Long, j As Long
    Dim tempDiffs() As WordDifferenceInfo
    Dim tempCount As Long
    Dim k As Long

    ' テキスト一致した段落ペアを記録（後でスタイル比較用）
    Dim matchedOld() As Long
    Dim matchedNew() As Long
    Dim matchedCount As Long

    matchedCount = 0
    ReDim matchedOld(0 To 0)
    ReDim matchedNew(0 To 0)

    ' 一時的な差分配列（逆順で格納される）
    tempCount = 0
    ReDim tempDiffs(0 To 0)

    i = n1
    j = n2

    ' バックトラック
    Do While i > 0 Or j > 0
        If i > 0 And j > 0 And texts1(i) = texts2(j) Then
            ' 一致：スタイル比較用に段落ペアを記録（空行以外）
            If Len(texts1(i)) > 0 Then
                If matchedCount = 0 Then
                    ReDim matchedOld(0 To 0)
                    ReDim matchedNew(0 To 0)
                Else
                    ReDim Preserve matchedOld(0 To matchedCount)
                    ReDim Preserve matchedNew(0 To matchedCount)
                End If
                matchedOld(matchedCount) = i
                matchedNew(matchedCount) = j
                matchedCount = matchedCount + 1
            End If
            i = i - 1
            j = j - 1
        ElseIf j > 0 And (i = 0 Or lcsMatrix(i, j - 1) >= lcsMatrix(i - 1, j)) Then
            ' 新ファイルで追加された行
            If Len(texts2(j)) > 0 Then
                AddTempWordDiff tempDiffs, tempCount, 0, j, "追加", _
                    "", texts2(j), "", ""
            End If
            j = j - 1
        ElseIf i > 0 And (j = 0 Or lcsMatrix(i - 1, j) > lcsMatrix(i, j - 1)) Then
            ' 旧ファイルから削除された行
            If Len(texts1(i)) > 0 Then
                AddTempWordDiff tempDiffs, tempCount, i, 0, "削除", _
                    texts1(i), "", "", ""
            End If
            i = i - 1
        Else
            ' 両方とも0の場合は終了
            Exit Do
        End If
    Loop

    ' 逆順を正順に変換して結果配列に格納
    diffCount = tempCount
    If tempCount > 0 Then
        ReDim differences(0 To tempCount - 1)
        For k = 0 To tempCount - 1
            differences(k) = tempDiffs(tempCount - 1 - k)
        Next k
    Else
        ReDim differences(0 To 0)
    End If

    ' 隣接する削除と追加を「変更」にマージ
    MergeAdjacentChanges differences, diffCount

    ' テキスト一致段落のスタイル比較用ペアをモジュールレベル変数に保存
    g_MatchedCount = matchedCount
    If matchedCount > 0 Then
        ReDim g_MatchedOld(0 To matchedCount - 1)
        ReDim g_MatchedNew(0 To matchedCount - 1)
        For k = 0 To matchedCount - 1
            g_MatchedOld(k) = matchedOld(k)
            g_MatchedNew(k) = matchedNew(k)
        Next k
    End If
End Sub

'==============================================================================
' LCS行列をバックトラックして差分を抽出（旧版：互換性のために残す）
'==============================================================================
Private Sub BacktrackLCS(ByRef lcsMatrix() As Long, _
                         ByRef texts1() As String, ByRef texts2() As String, _
                         ByRef styles1() As String, ByRef styles2() As String, _
                         ByVal n1 As Long, ByVal n2 As Long, _
                         ByRef differences() As WordDifferenceInfo, ByRef diffCount As Long)
    Dim i As Long, j As Long
    Dim tempDiffs() As WordDifferenceInfo
    Dim tempCount As Long
    Dim k As Long

    ' 一時的な差分配列（逆順で格納される）
    tempCount = 0
    ReDim tempDiffs(0 To 0)

    i = n1
    j = n2

    ' バックトラック
    Do While i > 0 Or j > 0
        If i > 0 And j > 0 And texts1(i) = texts2(j) Then
            ' 一致：スタイルの違いのみチェック
            If styles1(i) <> styles2(j) And Len(texts1(i)) > 0 Then
                AddTempWordDiff tempDiffs, tempCount, i, j, "スタイル変更", _
                    texts1(i), texts2(j), styles1(i), styles2(j)
            End If
            i = i - 1
            j = j - 1
        ElseIf j > 0 And (i = 0 Or lcsMatrix(i, j - 1) >= lcsMatrix(i - 1, j)) Then
            ' 新ファイルで追加された行
            If Len(texts2(j)) > 0 Then
                AddTempWordDiff tempDiffs, tempCount, 0, j, "追加", _
                    "", texts2(j), "", styles2(j)
            End If
            j = j - 1
        ElseIf i > 0 And (j = 0 Or lcsMatrix(i - 1, j) > lcsMatrix(i, j - 1)) Then
            ' 旧ファイルから削除された行
            If Len(texts1(i)) > 0 Then
                AddTempWordDiff tempDiffs, tempCount, i, 0, "削除", _
                    texts1(i), "", styles1(i), ""
            End If
            i = i - 1
        Else
            ' 両方とも0の場合は終了
            Exit Do
        End If
    Loop

    ' 逆順を正順に変換して結果配列に格納
    diffCount = tempCount
    If tempCount > 0 Then
        ReDim differences(0 To tempCount - 1)
        For k = 0 To tempCount - 1
            differences(k) = tempDiffs(tempCount - 1 - k)
        Next k
    Else
        ReDim differences(0 To 0)
    End If

    ' 隣接する削除と追加を「変更」にマージ
    MergeAdjacentChanges differences, diffCount
End Sub

'==============================================================================
' 隣接する削除と追加を「変更」にマージ
' （同じ位置での削除→追加は実質的に「変更」）
'==============================================================================
Private Sub MergeAdjacentChanges(ByRef differences() As WordDifferenceInfo, ByRef diffCount As Long)
    Dim i As Long
    Dim newDiffs() As WordDifferenceInfo
    Dim newCount As Long
    Dim merged As Boolean

    If diffCount <= 1 Then Exit Sub

    newCount = 0
    ReDim newDiffs(0 To diffCount - 1)

    i = 0
    Do While i < diffCount
        merged = False

        ' 削除の次に追加があり、位置が近い場合は「変更」にマージ
        If i < diffCount - 1 Then
            If differences(i).DiffType = "削除" And differences(i + 1).DiffType = "追加" Then
                ' 削除と追加が隣接している場合
                If Abs(differences(i).OldParagraphNo - differences(i + 1).NewParagraphNo) <= 1 Or _
                   (differences(i).OldParagraphNo > 0 And differences(i + 1).NewParagraphNo > 0) Then
                    ' マージして「変更」として記録
                    newDiffs(newCount).OldParagraphNo = differences(i).OldParagraphNo
                    newDiffs(newCount).NewParagraphNo = differences(i + 1).NewParagraphNo
                    newDiffs(newCount).DiffType = "変更"
                    newDiffs(newCount).OldText = differences(i).OldText
                    newDiffs(newCount).NewText = differences(i + 1).NewText
                    newDiffs(newCount).OldStyle = differences(i).OldStyle
                    newDiffs(newCount).NewStyle = differences(i + 1).NewStyle
                    newCount = newCount + 1
                    i = i + 2
                    merged = True
                End If
            End If
        End If

        If Not merged Then
            newDiffs(newCount) = differences(i)
            newCount = newCount + 1
            i = i + 1
        End If
    Loop

    ' 結果を元の配列にコピー
    diffCount = newCount
    If newCount > 0 Then
        ReDim differences(0 To newCount - 1)
        For i = 0 To newCount - 1
            differences(i) = newDiffs(i)
        Next i
    End If
End Sub

'==============================================================================
' 一時差分配列に追加（バックトラック用）
'==============================================================================
Private Sub AddTempWordDiff(ByRef tempDiffs() As WordDifferenceInfo, ByRef tempCount As Long, _
                            ByVal oldParaNo As Long, ByVal newParaNo As Long, _
                            ByVal diffType As String, ByVal oldText As String, ByVal newText As String, _
                            ByVal oldStyle As String, ByVal newStyle As String)
    ' 配列を拡張
    If tempCount = 0 Then
        ReDim tempDiffs(0 To 0)
    Else
        ReDim Preserve tempDiffs(0 To tempCount)
    End If

    ' 差異情報を格納
    With tempDiffs(tempCount)
        .OldParagraphNo = oldParaNo
        .NewParagraphNo = newParaNo
        .DiffType = diffType
        .OldText = Left(oldText, 500)
        .NewText = Left(newText, 500)
        .OldStyle = oldStyle
        .NewStyle = newStyle
    End With

    tempCount = tempCount + 1
End Sub

'==============================================================================
' 簡易差分検出（大きなファイル用：ブロック単位で比較）
'==============================================================================
Private Sub ComputeSimpleDiff(ByRef texts1() As String, ByRef texts2() As String, _
                              ByRef styles1() As String, ByRef styles2() As String, _
                              ByVal n1 As Long, ByVal n2 As Long, _
                              ByRef differences() As WordDifferenceInfo, ByRef diffCount As Long)
    Dim i1 As Long, i2 As Long
    Dim matchFound As Boolean
    Dim lookAhead As Long
    Dim j As Long

    diffCount = 0
    ReDim differences(0 To 0)

    i1 = 1
    i2 = 1
    lookAhead = 50  ' 前方探索の範囲

    Do While i1 <= n1 Or i2 <= n2
        ' 両方に残りがある場合
        If i1 <= n1 And i2 <= n2 Then
            If texts1(i1) = texts2(i2) Then
                ' 一致：スタイルの違いのみチェック
                If styles1(i1) <> styles2(i2) And Len(texts1(i1)) > 0 Then
                    AddWordDiffNew differences, diffCount, i1, i2, "スタイル変更", _
                        texts1(i1), texts2(i2), styles1(i1), styles2(i2)
                End If
                i1 = i1 + 1
                i2 = i2 + 1
            Else
                ' 不一致：前方探索で同期点を探す
                matchFound = False

                ' 新ファイルで追加された行を探す
                For j = i2 + 1 To Application.WorksheetFunction.Min(i2 + lookAhead, n2)
                    If texts1(i1) = texts2(j) Then
                        ' i2 から j-1 までが追加
                        Do While i2 < j
                            If Len(texts2(i2)) > 0 Then
                                AddWordDiffNew differences, diffCount, 0, i2, "追加", _
                                    "", texts2(i2), "", styles2(i2)
                            End If
                            i2 = i2 + 1
                        Loop
                        matchFound = True
                        Exit For
                    End If
                Next j

                If Not matchFound Then
                    ' 旧ファイルから削除された行を探す
                    For j = i1 + 1 To Application.WorksheetFunction.Min(i1 + lookAhead, n1)
                        If texts1(j) = texts2(i2) Then
                            ' i1 から j-1 までが削除
                            Do While i1 < j
                                If Len(texts1(i1)) > 0 Then
                                    AddWordDiffNew differences, diffCount, i1, 0, "削除", _
                                        texts1(i1), "", styles1(i1), ""
                                End If
                                i1 = i1 + 1
                            Loop
                            matchFound = True
                            Exit For
                        End If
                    Next j
                End If

                If Not matchFound Then
                    ' 同期点が見つからない：変更として記録
                    If Len(texts1(i1)) > 0 Or Len(texts2(i2)) > 0 Then
                        AddWordDiffNew differences, diffCount, i1, i2, "変更", _
                            texts1(i1), texts2(i2), styles1(i1), styles2(i2)
                    End If
                    i1 = i1 + 1
                    i2 = i2 + 1
                End If
            End If
        ' 旧ファイルのみ残り
        ElseIf i1 <= n1 Then
            If Len(texts1(i1)) > 0 Then
                AddWordDiffNew differences, diffCount, i1, 0, "削除", _
                    texts1(i1), "", styles1(i1), ""
            End If
            i1 = i1 + 1
        ' 新ファイルのみ残り
        Else
            If Len(texts2(i2)) > 0 Then
                AddWordDiffNew differences, diffCount, 0, i2, "追加", _
                    "", texts2(i2), "", styles2(i2)
            End If
            i2 = i2 + 1
        End If

        ' 進捗表示
        If (i1 + i2) Mod 200 = 0 Then
            Debug.Print "  簡易比較中: 旧=" & i1 & "/" & n1 & ", 新=" & i2 & "/" & n2
            DoEvents
        End If
    Loop
End Sub

'==============================================================================
' 簡易差分検出（最適化版：スタイル情報なし）
'==============================================================================
Private Sub ComputeSimpleDiffOptimized(ByRef texts1() As String, ByRef texts2() As String, _
                                        ByVal n1 As Long, ByVal n2 As Long, _
                                        ByRef differences() As WordDifferenceInfo, ByRef diffCount As Long)
    Dim i1 As Long, i2 As Long
    Dim matchFound As Boolean
    Dim lookAhead As Long
    Dim j As Long

    ' テキスト一致した段落ペアを記録（後でスタイル比較用）
    Dim matchedOld() As Long
    Dim matchedNew() As Long
    Dim matchedCount As Long

    matchedCount = 0
    ReDim matchedOld(0 To 0)
    ReDim matchedNew(0 To 0)

    diffCount = 0
    ReDim differences(0 To 0)

    i1 = 1
    i2 = 1
    lookAhead = 50  ' 前方探索の範囲

    Do While i1 <= n1 Or i2 <= n2
        ' 両方に残りがある場合
        If i1 <= n1 And i2 <= n2 Then
            If texts1(i1) = texts2(i2) Then
                ' 一致：スタイル比較用に段落ペアを記録（空行以外）
                If Len(texts1(i1)) > 0 Then
                    If matchedCount = 0 Then
                        ReDim matchedOld(0 To 0)
                        ReDim matchedNew(0 To 0)
                    Else
                        ReDim Preserve matchedOld(0 To matchedCount)
                        ReDim Preserve matchedNew(0 To matchedCount)
                    End If
                    matchedOld(matchedCount) = i1
                    matchedNew(matchedCount) = i2
                    matchedCount = matchedCount + 1
                End If
                i1 = i1 + 1
                i2 = i2 + 1
            Else
                ' 不一致：前方探索で同期点を探す
                matchFound = False

                ' 新ファイルで追加された行を探す
                For j = i2 + 1 To Application.WorksheetFunction.Min(i2 + lookAhead, n2)
                    If texts1(i1) = texts2(j) Then
                        ' i2 から j-1 までが追加
                        Do While i2 < j
                            If Len(texts2(i2)) > 0 Then
                                AddWordDiffNew differences, diffCount, 0, i2, "追加", _
                                    "", texts2(i2), "", ""
                            End If
                            i2 = i2 + 1
                        Loop
                        matchFound = True
                        Exit For
                    End If
                Next j

                If Not matchFound Then
                    ' 旧ファイルから削除された行を探す
                    For j = i1 + 1 To Application.WorksheetFunction.Min(i1 + lookAhead, n1)
                        If texts1(j) = texts2(i2) Then
                            ' i1 から j-1 までが削除
                            Do While i1 < j
                                If Len(texts1(i1)) > 0 Then
                                    AddWordDiffNew differences, diffCount, i1, 0, "削除", _
                                        texts1(i1), "", "", ""
                                End If
                                i1 = i1 + 1
                            Loop
                            matchFound = True
                            Exit For
                        End If
                    Next j
                End If

                If Not matchFound Then
                    ' 同期点が見つからない：変更として記録
                    If Len(texts1(i1)) > 0 Or Len(texts2(i2)) > 0 Then
                        AddWordDiffNew differences, diffCount, i1, i2, "変更", _
                            texts1(i1), texts2(i2), "", ""
                    End If
                    i1 = i1 + 1
                    i2 = i2 + 1
                End If
            End If
        ' 旧ファイルのみ残り
        ElseIf i1 <= n1 Then
            If Len(texts1(i1)) > 0 Then
                AddWordDiffNew differences, diffCount, i1, 0, "削除", _
                    texts1(i1), "", "", ""
            End If
            i1 = i1 + 1
        ' 新ファイルのみ残り
        Else
            If Len(texts2(i2)) > 0 Then
                AddWordDiffNew differences, diffCount, 0, i2, "追加", _
                    "", texts2(i2), "", ""
            End If
            i2 = i2 + 1
        End If

        ' 進捗表示
        If (i1 + i2) Mod 100 = 0 Then
            ShowProgress "[3/4] 差分計算(簡易)", i1 + i2, n1 + n2
        End If
    Loop
    ShowProgress "[3/4] 差分計算(簡易)", n1 + n2, n1 + n2

    ' テキスト一致段落のスタイル比較用ペアをモジュールレベル変数に保存
    g_MatchedCount = matchedCount
    If matchedCount > 0 Then
        ReDim g_MatchedOld(0 To matchedCount - 1)
        ReDim g_MatchedNew(0 To matchedCount - 1)
        Dim k As Long
        For k = 0 To matchedCount - 1
            g_MatchedOld(k) = matchedOld(k)
            g_MatchedNew(k) = matchedNew(k)
        Next k
    End If
End Sub

'==============================================================================
' Word差分を追加（新形式：旧/新段落番号両方）
'==============================================================================
Private Sub AddWordDiffNew(ByRef differences() As WordDifferenceInfo, ByRef diffCount As Long, _
                           ByVal oldParaNo As Long, ByVal newParaNo As Long, _
                           ByVal diffType As String, ByVal oldText As String, ByVal newText As String, _
                           ByVal oldStyle As String, ByVal newStyle As String)
    ' 配列を拡張
    If diffCount = 0 Then
        ReDim differences(0 To 0)
    Else
        ReDim Preserve differences(0 To diffCount)
    End If

    ' 差異情報を格納
    With differences(diffCount)
        .OldParagraphNo = oldParaNo
        .NewParagraphNo = newParaNo
        .DiffType = diffType
        .OldText = Left(oldText, 500)
        .NewText = Left(newText, 500)
        .OldStyle = oldStyle
        .NewStyle = newStyle
    End With

    diffCount = diffCount + 1
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

' 旧式のWord差異追加関数は削除されました
' 新しいAddWordDiffNew関数を使用してください（LCSベース比較で使用）

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
    ThisWorkbook.Worksheets("比較結果").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' 新しいシートを作成
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = "比較結果"

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

'==============================================================================
' Word結果シートを作成（WinMerge方式：旧/新行番号を両方表示）
'==============================================================================
Private Sub CreateWordResultSheet(ByRef differences() As WordDifferenceInfo, ByVal diffCount As Long, _
                                   ByVal file1Path As String, ByVal file2Path As String)
    Dim ws As Worksheet
    Dim i As Long
    Dim row As Long
    Dim oldParaStr As String
    Dim newParaStr As String
    Dim shp As Shape
    Dim btnLeft As Double
    Dim btnTop As Double
    Dim btnWidth As Double
    Dim btnHeight As Double

    ' 既存の結果シートがあれば削除
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("比較結果").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' 新しいシートを作成
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = "比較結果"

    With ws
        ' タイトル
        .Range("A1").Value = "Word ファイル比較結果（WinMerge方式）"
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
        .Range("A7").Value = "比較方式:"
        .Range("B7").Value = "LCS（最長共通部分列）アルゴリズム"

        ' 検索ボタンの説明
        .Range("F3").Value = "差分箇所を検索:"
        .Range("F3").Font.Bold = True
        .Range("F4").Value = "※データ行を選択してからボタンをクリック"
        .Range("F4").Font.Size = 9
        .Range("F4").Font.Color = RGB(128, 128, 128)

        ' 検索ボタンの配置
        btnWidth = 100
        btnHeight = 28
        btnLeft = .Range("G3").Left + 5
        btnTop = .Range("G3").Top + 2

        ' 旧ファイル検索ボタン
        Set shp = .Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, btnWidth, btnHeight)
        With shp
            .Name = "btnSearchOld"
            .Placement = xlFreeFloating
            .Fill.ForeColor.RGB = RGB(255, 152, 0)  ' オレンジ
            .Line.ForeColor.RGB = RGB(230, 126, 0)
            .Line.Weight = 1.5
            .TextFrame2.TextRange.Characters.Text = "旧ファイル検索"
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextFrame2.TextRange.Font.Size = 10
            .TextFrame2.TextRange.Font.Bold = msoTrue
            .TextFrame2.TextRange.Font.Name = "Meiryo UI"
            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .TextFrame2.VerticalAnchor = msoAnchorMiddle
            .OnAction = "SearchInOldWordFile"
        End With

        ' 新ファイル検索ボタン
        Set shp = .Shapes.AddShape(msoShapeRoundedRectangle, btnLeft + btnWidth + 10, btnTop, btnWidth, btnHeight)
        With shp
            .Name = "btnSearchNew"
            .Placement = xlFreeFloating
            .Fill.ForeColor.RGB = RGB(33, 150, 243)  ' 青
            .Line.ForeColor.RGB = RGB(25, 118, 210)
            .Line.Weight = 1.5
            .TextFrame2.TextRange.Characters.Text = "新ファイル検索"
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextFrame2.TextRange.Font.Size = 10
            .TextFrame2.TextRange.Font.Bold = msoTrue
            .TextFrame2.TextRange.Font.Name = "Meiryo UI"
            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .TextFrame2.VerticalAnchor = msoAnchorMiddle
            .OnAction = "SearchInNewWordFile"
        End With

        ' 凡例
        .Range("A9").Value = "凡例："
        .Range("B9").Value = "変更"
        .Range("B9").Interior.Color = COLOR_CHANGED
        .Range("C9").Value = "追加"
        .Range("C9").Interior.Color = COLOR_ADDED
        .Range("D9").Value = "削除"
        .Range("D9").Interior.Color = COLOR_DELETED
        .Range("E9").Value = "スタイル変更"
        .Range("E9").Interior.Color = RGB(204, 153, 255)  ' 薄紫

        ' ヘッダー（列を1つ追加：旧行番号と新行番号を分離）
        .Range("A11").Value = "No"
        .Range("B11").Value = "旧行番号"
        .Range("C11").Value = "新行番号"
        .Range("D11").Value = "差異タイプ"
        .Range("E11").Value = "旧ファイルのテキスト"
        .Range("F11").Value = "新ファイルのテキスト"
        .Range("G11").Value = "旧スタイル"
        .Range("H11").Value = "新スタイル"

        ' ヘッダー書式
        With .Range("A11:H11")
            .Font.Bold = True
            .Interior.Color = RGB(68, 114, 196)
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
        End With

        ' データ行
        For i = 0 To diffCount - 1
            row = i + 12

            .Cells(row, 1).Value = i + 1

            ' 旧行番号の表示（0の場合は「-」）
            If differences(i).OldParagraphNo > 0 Then
                oldParaStr = CStr(differences(i).OldParagraphNo)
            Else
                oldParaStr = "-"
            End If
            .Cells(row, 2).Value = oldParaStr
            .Cells(row, 2).HorizontalAlignment = xlCenter

            ' 新行番号の表示（0の場合は「-」）
            If differences(i).NewParagraphNo > 0 Then
                newParaStr = CStr(differences(i).NewParagraphNo)
            Else
                newParaStr = "-"
            End If
            .Cells(row, 3).Value = newParaStr
            .Cells(row, 3).HorizontalAlignment = xlCenter

            .Cells(row, 4).Value = differences(i).DiffType
            .Cells(row, 5).Value = differences(i).OldText
            .Cells(row, 6).Value = differences(i).NewText
            .Cells(row, 7).Value = differences(i).OldStyle
            .Cells(row, 8).Value = differences(i).NewStyle

            ' テキストを折り返し
            .Cells(row, 5).WrapText = True
            .Cells(row, 6).WrapText = True

            ' 差異タイプによって行に色を付ける
            Select Case differences(i).DiffType
                Case "変更"
                    .Range(.Cells(row, 1), .Cells(row, 8)).Interior.Color = COLOR_CHANGED
                Case "追加"
                    .Range(.Cells(row, 1), .Cells(row, 8)).Interior.Color = COLOR_ADDED
                Case "削除"
                    .Range(.Cells(row, 1), .Cells(row, 8)).Interior.Color = COLOR_DELETED
                Case "スタイル変更"
                    .Range(.Cells(row, 1), .Cells(row, 8)).Interior.Color = RGB(204, 153, 255)  ' 薄紫
            End Select
        Next i

        ' 列幅調整
        .Columns("A").ColumnWidth = 8
        .Columns("B").ColumnWidth = 10
        .Columns("C").ColumnWidth = 10
        .Columns("D").ColumnWidth = 14
        .Columns("E").ColumnWidth = 40
        .Columns("F").ColumnWidth = 40
        .Columns("G").ColumnWidth = 25
        .Columns("H").ColumnWidth = 25

        ' フィルターを設定
        .Range("A11:H11").AutoFilter

        ' シートをアクティブにしてからウィンドウ枠を固定
        .Activate
        .Rows(12).Select
        ActiveWindow.FreezePanes = True

        ' セルA1を選択
        .Range("A1").Select
    End With
End Sub

'==============================================================================
' 選択行のWord差分を旧ファイルで検索して開く
'==============================================================================
Public Sub SearchInOldWordFile()
    SearchWordDifference True
End Sub

'==============================================================================
' 選択行のWord差分を新ファイルで検索して開く
'==============================================================================
Public Sub SearchInNewWordFile()
    SearchWordDifference False
End Sub

'==============================================================================
' Word差分を検索して開く（内部処理）
'==============================================================================
Private Sub SearchWordDifference(ByVal isOldFile As Boolean)
    Dim ws As Worksheet
    Dim selectedRow As Long
    Dim filePath As String
    Dim searchText As String
    Dim wordApp As Object
    Dim doc As Object

    On Error GoTo ErrorHandler

    ' 比較結果シートを取得
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("比較結果")
    On Error GoTo ErrorHandler

    If ws Is Nothing Then
        MsgBox "比較結果シートが見つかりません。" & vbCrLf & _
               "先にWord比較を実行してください。", vbExclamation, "エラー"
        Exit Sub
    End If

    ' 選択されている行を取得
    selectedRow = Selection.row

    ' ヘッダー行以下かチェック（Word比較結果は12行目からデータ）
    If selectedRow < 12 Then
        MsgBox "差異データの行を選択してください。" & vbCrLf & _
               "（12行目以降のデータ行を選択）", vbExclamation, "行選択エラー"
        Exit Sub
    End If

    ' ファイルパスを取得（B3=旧ファイル、B4=新ファイル）
    If isOldFile Then
        filePath = ws.Range("B3").Value
        searchText = ws.Cells(selectedRow, 5).Value  ' E列：旧ファイルのテキスト
    Else
        filePath = ws.Range("B4").Value
        searchText = ws.Cells(selectedRow, 6).Value  ' F列：新ファイルのテキスト
    End If

    ' 検索テキストが空の場合
    If Len(Trim(searchText)) = 0 Then
        MsgBox "検索するテキストがありません。" & vbCrLf & _
               IIf(isOldFile, "旧ファイル側", "新ファイル側") & "にテキストがない差異です。", _
               vbExclamation, "検索エラー"
        Exit Sub
    End If

    ' ファイルの存在確認
    If Dir(filePath) = "" Then
        MsgBox "ファイルが見つかりません: " & vbCrLf & filePath, vbCritical, "ファイルエラー"
        Exit Sub
    End If

    ' 検索テキストを最初の100文字に制限（長すぎると検索に失敗する可能性）
    If Len(searchText) > 100 Then
        searchText = Left(searchText, 100)
    End If

    ' Wordアプリケーションを取得または起動
    On Error Resume Next
    Set wordApp = GetObject(, "Word.Application")
    If wordApp Is Nothing Then
        Set wordApp = CreateObject("Word.Application")
    End If
    On Error GoTo ErrorHandler

    wordApp.Visible = True

    ' ファイルを開く
    Set doc = wordApp.Documents.Open(filePath, ReadOnly:=True)

    ' 検索を実行
    With doc.Content.Find
        .ClearFormatting
        .Text = searchText
        .Forward = True
        .Wrap = 1  ' wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False

        If .Execute Then
            ' 見つかった場合、その位置を選択
            doc.ActiveWindow.ScrollIntoView doc.Content.Find.Parent
            doc.Content.Find.Parent.Select
            MsgBox "テキストが見つかりました。", vbInformation, "検索完了"
        Else
            MsgBox "テキストが見つかりませんでした。" & vbCrLf & vbCrLf & _
                   "検索テキスト: " & Left(searchText, 50) & IIf(Len(searchText) > 50, "...", ""), _
                   vbExclamation, "検索結果"
        End If
    End With

    ' Wordをアクティブにする
    wordApp.Activate

    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & vbCrLf & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, vbCritical, "エラー"
End Sub

