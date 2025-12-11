'==============================================================================
' Word ファイル比較ツール
' モジュール名: WordFileComparator
'==============================================================================
' 概要:
'   2つのWordファイルを比較し、差異を一覧表示するツールです。
'   ExcelからWordを操作して比較を実行します。
'
' 機能:
'   - ファイル選択ダイアログで2つのWordファイルを指定
'   - 段落単位での差異検出
'   - Wordの組み込み比較機能を使用した詳細比較
'   - 結果をExcelシートに出力
'
' 必要な環境:
'   - Microsoft Excel 2010以降
'   - Microsoft Word 2010以降（参照設定不要、遅延バインディング使用）
'
' 作成日: 2025-12-11
'==============================================================================

Option Explicit

'==============================================================================
' 設定
'==============================================================================
' 比較する最大段落数（パフォーマンス対策）
Private Const MAX_PARAGRAPHS As Long = 5000

' 差異ハイライト色
Private Const COLOR_CHANGED As Long = 65535      ' 黄色: 変更
Private Const COLOR_ADDED As Long = 5296274      ' 緑: 追加
Private Const COLOR_DELETED As Long = 13421823   ' 赤: 削除

'==============================================================================
' データ構造
'==============================================================================
Private Type WordDifferenceInfo
    ParagraphNo As Long      ' 段落番号
    DiffType As String       ' 差異タイプ（変更/追加/削除）
    OldText As String        ' 旧ファイルのテキスト
    NewText As String        ' 新ファイルのテキスト
End Type

'==============================================================================
' メインプロシージャ: Word ファイルを比較（段落単位）
'==============================================================================
Public Sub CompareWordFiles()
    Dim file1Path As String
    Dim file2Path As String
    Dim wordApp As Object
    Dim doc1 As Object
    Dim doc2 As Object
    Dim differences() As WordDifferenceInfo
    Dim diffCount As Long
    Dim wordWasRunning As Boolean

    On Error GoTo ErrorHandler

    ' ファイル選択
    MsgBox "2つのWordファイルを比較します。" & vbCrLf & vbCrLf & _
           "まず、1つ目のファイル（旧ファイル）を選択してください。", _
           vbInformation, "Word ファイル比較ツール"

    file1Path = SelectWordFile("1つ目のファイル（旧ファイル）を選択")
    If file1Path = "" Then
        MsgBox "ファイル選択がキャンセルされました。", vbExclamation
        Exit Sub
    End If

    MsgBox "次に、2つ目のファイル（新ファイル）を選択してください。", _
           vbInformation, "Word ファイル比較ツール"

    file2Path = SelectWordFile("2つ目のファイル（新ファイル）を選択")
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
               "結果は「WordCompareResult」シートに出力されました。", _
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
' Wordの組み込み比較機能を使用した詳細比較
'==============================================================================
Public Sub CompareWordFilesDetailed()
    Dim file1Path As String
    Dim file2Path As String
    Dim wordApp As Object
    Dim doc1 As Object
    Dim compDoc As Object
    Dim wordWasRunning As Boolean

    On Error GoTo ErrorHandler

    ' ファイル選択
    MsgBox "Wordの組み込み比較機能を使用して2つのファイルを比較します。" & vbCrLf & vbCrLf & _
           "比較結果はWordで表示されます。" & vbCrLf & vbCrLf & _
           "まず、1つ目のファイル（旧ファイル）を選択してください。", _
           vbInformation, "Word 詳細比較ツール"

    file1Path = SelectWordFile("1つ目のファイル（旧ファイル）を選択")
    If file1Path = "" Then
        MsgBox "ファイル選択がキャンセルされました。", vbExclamation
        Exit Sub
    End If

    MsgBox "次に、2つ目のファイル（新ファイル）を選択してください。", _
           vbInformation, "Word 詳細比較ツール"

    file2Path = SelectWordFile("2つ目のファイル（新ファイル）を選択")
    If file2Path = "" Then
        MsgBox "ファイル選択がキャンセルされました。", vbExclamation
        Exit Sub
    End If

    ' 同じファイルが選択された場合
    If LCase(file1Path) = LCase(file2Path) Then
        MsgBox "同じファイルが選択されました。異なるファイルを選択してください。", vbExclamation
        Exit Sub
    End If

    Debug.Print "========================================="
    Debug.Print "Word 詳細比較を開始します"
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

    wordApp.Visible = True

    ' 旧ファイルを開く
    Set doc1 = wordApp.Documents.Open(file1Path, ReadOnly:=True)

    ' Wordの比較機能を使用
    ' Document.Compare メソッド: 2つの文書を比較して差異を表示
    ' wdCompareTargetNew = 2 (新しい文書に比較結果を作成)
    Set compDoc = wordApp.CompareDocuments( _
        OriginalDocument:=doc1, _
        RevisedDocument:=wordApp.Documents.Open(file2Path, ReadOnly:=True), _
        Destination:=2, _
        Granularity:=1, _
        CompareFormatting:=True, _
        CompareCaseChanges:=True, _
        CompareWhitespace:=True, _
        CompareTables:=True, _
        CompareHeaders:=True, _
        CompareFootnotes:=True, _
        CompareTextboxes:=True, _
        CompareFields:=True, _
        CompareComments:=True)

    ' 比較結果文書をアクティブに
    compDoc.Activate

    ' 元のドキュメントを閉じる
    doc1.Close SaveChanges:=False

    MsgBox "Wordで比較結果が表示されています。" & vbCrLf & vbCrLf & _
           "変更箇所は変更履歴として表示されます。" & vbCrLf & _
           "「校閲」タブで変更内容を確認できます。", _
           vbInformation, "処理完了"

    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not doc1 Is Nothing Then doc1.Close SaveChanges:=False
    If Not wordApp Is Nothing And Not wordWasRunning Then wordApp.Quit
    On Error GoTo 0

    MsgBox "エラーが発生しました: " & vbCrLf & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, vbCritical, "エラー"
End Sub

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
' Word文書を段落単位で比較
'==============================================================================
Private Sub CompareWordDocuments(ByRef doc1 As Object, ByRef doc2 As Object, _
                                  ByRef differences() As WordDifferenceInfo, ByRef diffCount As Long)
    Dim paraCount1 As Long
    Dim paraCount2 As Long
    Dim maxParas As Long
    Dim i As Long
    Dim text1 As String
    Dim text2 As String

    paraCount1 = doc1.Paragraphs.Count
    paraCount2 = doc2.Paragraphs.Count

    Debug.Print "旧ファイル段落数: " & paraCount1
    Debug.Print "新ファイル段落数: " & paraCount2

    ' 比較範囲を決定
    maxParas = Application.WorksheetFunction.Max(paraCount1, paraCount2)
    If maxParas > MAX_PARAGRAPHS Then maxParas = MAX_PARAGRAPHS

    ' 段落単位で比較
    For i = 1 To maxParas
        ' 旧ファイルのテキスト取得
        If i <= paraCount1 Then
            text1 = CleanText(doc1.Paragraphs(i).Range.Text)
        Else
            text1 = ""
        End If

        ' 新ファイルのテキスト取得
        If i <= paraCount2 Then
            text2 = CleanText(doc2.Paragraphs(i).Range.Text)
        Else
            text2 = ""
        End If

        ' 比較
        If text1 <> text2 Then
            If Len(text1) = 0 And Len(text2) > 0 Then
                ' 追加
                AddWordDifference differences, diffCount, i, "追加", "(空)", text2
            ElseIf Len(text1) > 0 And Len(text2) = 0 Then
                ' 削除
                AddWordDifference differences, diffCount, i, "削除", text1, "(空)"
            Else
                ' 変更
                AddWordDifference differences, diffCount, i, "変更", text1, text2
            End If
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
            For i = paraCount2 + 1 To Application.WorksheetFunction.Min(paraCount1, MAX_PARAGRAPHS)
                text1 = CleanText(doc1.Paragraphs(i).Range.Text)
                If Len(text1) > 0 Then
                    AddWordDifference differences, diffCount, i, "削除", text1, "(段落なし)"
                End If
            Next i
        Else
            For i = paraCount1 + 1 To Application.WorksheetFunction.Min(paraCount2, MAX_PARAGRAPHS)
                text2 = CleanText(doc2.Paragraphs(i).Range.Text)
                If Len(text2) > 0 Then
                    AddWordDifference differences, diffCount, i, "追加", "(段落なし)", text2
                End If
            Next i
        End If
    End If
End Sub

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
' 差異を追加
'==============================================================================
Private Sub AddWordDifference(ByRef differences() As WordDifferenceInfo, ByRef diffCount As Long, _
                               ByVal paraNo As Long, ByVal diffType As String, _
                               ByVal oldText As String, ByVal newText As String)
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
    End With

    diffCount = diffCount + 1
End Sub

'==============================================================================
' 結果シートを作成
'==============================================================================
Private Sub CreateWordResultSheet(ByRef differences() As WordDifferenceInfo, ByVal diffCount As Long, _
                                   ByVal file1Path As String, ByVal file2Path As String)
    Dim ws As Worksheet
    Dim i As Long
    Dim row As Long

    ' 既存の結果シートがあれば削除
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("WordCompareResult").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' 新しいシートを作成
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = "WordCompareResult"

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

        ' ヘッダー
        .Range("A10").Value = "No"
        .Range("B10").Value = "段落番号"
        .Range("C10").Value = "差異タイプ"
        .Range("D10").Value = "旧ファイルのテキスト"
        .Range("E10").Value = "新ファイルのテキスト"

        ' ヘッダー書式
        With .Range("A10:E10")
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

            ' テキストを折り返し
            .Cells(row, 4).WrapText = True
            .Cells(row, 5).WrapText = True

            ' 差異タイプによって行に色を付ける
            Select Case differences(i).DiffType
                Case "変更"
                    .Range(.Cells(row, 1), .Cells(row, 5)).Interior.Color = COLOR_CHANGED
                Case "追加"
                    .Range(.Cells(row, 1), .Cells(row, 5)).Interior.Color = COLOR_ADDED
                Case "削除"
                    .Range(.Cells(row, 1), .Cells(row, 5)).Interior.Color = COLOR_DELETED
            End Select
        Next i

        ' 列幅調整
        .Columns("A").ColumnWidth = 6
        .Columns("B").ColumnWidth = 10
        .Columns("C").ColumnWidth = 12
        .Columns("D").ColumnWidth = 50
        .Columns("E").ColumnWidth = 50

        ' フィルターを設定
        .Range("A10:E10").AutoFilter

        ' ウィンドウ枠の固定
        .Rows(11).Select
        ActiveWindow.FreezePanes = True

        ' セルA1を選択
        .Range("A1").Select
    End With
End Sub

'==============================================================================
' テスト用プロシージャ
'==============================================================================
Public Sub TestCompareWordFiles()
    CompareWordFiles
End Sub

Public Sub TestCompareWordFilesDetailed()
    CompareWordFilesDetailed
End Sub
