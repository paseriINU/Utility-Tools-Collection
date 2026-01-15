Attribute VB_Name = "FC_WordOps"
Option Explicit

'==============================================================================
' Excel/Word ファイル比較ツール - Word操作モジュール
' Word比較ロジック、LCS/簡易比較、結果シート作成機能を提供
'==============================================================================

' ============================================================================
' Word比較の内部処理
' ============================================================================
Public Sub CompareWordFilesInternal(ByVal file1Path As String, ByVal file2Path As String)
    Dim wordApp As Object
    Dim doc1 As Object
    Dim doc2 As Object
    Dim differences() As WordDiffInfo
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

    ' バックグラウンドで処理（画面に表示しない）
    wordApp.Visible = False
    wordApp.DisplayAlerts = False
    wordApp.ScreenUpdating = False

    ' ファイルを開く（非表示で）
    Set doc1 = wordApp.Documents.Open(FileName:=file1Path, ReadOnly:=True, Visible:=False)
    Set doc2 = wordApp.Documents.Open(FileName:=file2Path, ReadOnly:=True, Visible:=False)

    ' 比較実行
    diffCount = 0
    ReDim differences(0 To 0)

    CompareWordDocuments doc1, doc2, differences, diffCount

    ' ドキュメントを閉じる
    doc1.Close SaveChanges:=False
    doc2.Close SaveChanges:=False

    ' Wordの設定を復元してから終了
    If Not wordWasRunning Then
        wordApp.Quit
    Else
        wordApp.ScreenUpdating = True
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
    If Not wordApp Is Nothing Then
        If Not wordWasRunning Then
            wordApp.Quit
        Else
            wordApp.ScreenUpdating = True
        End If
    End If
    On Error GoTo 0

    MsgBox "エラーが発生しました: " & vbCrLf & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, vbCritical, "エラー"
End Sub

' ============================================================================
' Word文書を段落単位で比較（WinMerge方式：LCSアルゴリズム使用）
' ============================================================================
Private Sub CompareWordDocuments(ByRef doc1 As Object, ByRef doc2 As Object, _
                                 ByRef differences() As WordDiffInfo, ByRef diffCount As Long)
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

    ' 段落テキストのみを配列に取得
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

    ' LCSベースの差分検出を実行
    Debug.Print "差分を計算中..."
    ComputeLCSDiffOptimized texts1, texts2, paraCount1, paraCount2, differences, diffCount

    ' 差分が検出された段落のみスタイル情報を取得
    Debug.Print "差分段落のスタイル情報を取得中..."
    FetchStylesForDifferences doc1, doc2, differences, diffCount

    Debug.Print "差分計算完了: " & diffCount & " 件の差異を検出"
End Sub

' ============================================================================
' 差分が検出された段落のみスタイル情報を取得
' ============================================================================
Private Sub FetchStylesForDifferences(ByRef doc1 As Object, ByRef doc2 As Object, _
                                      ByRef differences() As WordDiffInfo, ByRef diffCount As Long)
    Dim i As Long
    Dim oldParaNo As Long
    Dim newParaNo As Long
    Dim fetchedCount As Long
    Dim oldStyle As String
    Dim newStyle As String
    Dim oldText As String
    Dim checkStyleMode As Boolean

    ' スタイル比較モードを取得
    checkStyleMode = GetCheckStyleMode()

    ' スタイル比較がオフの場合はスキップ
    If Not checkStyleMode Then
        Debug.Print "  スタイル比較: スキップ（チェックボックスがオフ）"
        g_MatchedCount = 0
        Erase g_MatchedOld
        Erase g_MatchedNew
        ClearProgress
        Exit Sub
    End If

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

    ' 2. テキスト一致段落のスタイル比較
    If g_MatchedCount > 0 Then
        Debug.Print "  テキスト一致段落のスタイル比較: " & g_MatchedCount & " 件"

        Dim styleCheckCount As Long
        styleCheckCount = 0

        For i = 0 To g_MatchedCount - 1
            oldParaNo = g_MatchedOld(i)
            newParaNo = g_MatchedNew(i)

            If oldParaNo > 0 And oldParaNo <= doc1.Paragraphs.Count And _
               newParaNo > 0 And newParaNo <= doc2.Paragraphs.Count Then

                oldStyle = GetParagraphStyleInfo(doc1.Paragraphs(oldParaNo))
                newStyle = GetParagraphStyleInfo(doc2.Paragraphs(newParaNo))

                ' スタイルが異なる場合は差分として追加
                If oldStyle <> newStyle Then
                    oldText = CleanText(doc1.Paragraphs(oldParaNo).Range.Text)
                    AddWordDiff differences, diffCount, oldParaNo, newParaNo, "スタイル変更", _
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

    ClearProgress
End Sub

' ============================================================================
' LCSベースの差分検出（最適化版）
' ============================================================================
Private Sub ComputeLCSDiffOptimized(ByRef texts1() As String, ByRef texts2() As String, _
                                    ByVal n1 As Long, ByVal n2 As Long, _
                                    ByRef differences() As WordDiffInfo, ByRef diffCount As Long)
    Dim lcsMatrix() As Long
    Dim i As Long, j As Long
    Dim maxLen As Long
    Dim textHash1 As Object
    Dim textHash2 As Object
    Dim uniqueTexts1 As Long, uniqueTexts2 As Long
    Dim commonTexts As Long
    Dim useLCS As Boolean
    Dim key As Variant

    ' チェックボックスの状態を取得
    useLCS = GetUseLCSMode()

    ' ハッシュマップを作成して同一テキストを事前に特定
    Set textHash1 = CreateObject("Scripting.Dictionary")
    Set textHash2 = CreateObject("Scripting.Dictionary")

    For i = 1 To n1
        If Not textHash1.exists(texts1(i)) Then
            textHash1.Add texts1(i), i
        End If
    Next i
    uniqueTexts1 = textHash1.Count

    For i = 1 To n2
        If Not textHash2.exists(texts2(i)) Then
            textHash2.Add texts2(i), i
        End If
    Next i
    uniqueTexts2 = textHash2.Count

    commonTexts = 0
    For Each key In textHash1.Keys
        If textHash2.exists(key) Then
            commonTexts = commonTexts + 1
        End If
    Next key

    Debug.Print "  ユニークテキスト数: 旧=" & uniqueTexts1 & ", 新=" & uniqueTexts2 & ", 共通=" & commonTexts

    Set textHash1 = Nothing
    Set textHash2 = Nothing

    maxLen = Application.WorksheetFunction.Max(n1, n2)

    If Not useLCS Then
        Debug.Print "比較モード: 簡易比較（高速）"
        ComputeSimpleDiffOptimized texts1, texts2, n1, n2, differences, diffCount
        Exit Sub
    End If

    Debug.Print "比較モード: LCSアルゴリズム（厳密）"

    If maxLen > 10000 Then
        Dim result As VbMsgBoxResult
        result = MsgBox("段落数が " & maxLen & " あります。" & vbCrLf & vbCrLf & _
                        "LCSアルゴリズムは大量のメモリと時間を使用します。" & vbCrLf & _
                        "処理に数分～数十分かかる可能性があります。" & vbCrLf & vbCrLf & _
                        "続行しますか？", vbYesNo + vbExclamation, "警告")
        If result = vbNo Then
            Debug.Print "ユーザーがキャンセル。簡易比較モードにフォールバック"
            ComputeSimpleDiffOptimized texts1, texts2, n1, n2, differences, diffCount
            Exit Sub
        End If
    End If

    ' LCS行列を初期化
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

        If i Mod 100 = 0 Or i = n1 Then
            ShowProgress "[3/4] 差分計算(LCS)", i, n1
        End If
    Next i

    ' バックトラックして差分を抽出
    BacktrackLCSOptimized lcsMatrix, texts1, texts2, n1, n2, differences, diffCount
End Sub

' ============================================================================
' LCS行列をバックトラックして差分を抽出
' ============================================================================
Private Sub BacktrackLCSOptimized(ByRef lcsMatrix() As Long, _
                                  ByRef texts1() As String, ByRef texts2() As String, _
                                  ByVal n1 As Long, ByVal n2 As Long, _
                                  ByRef differences() As WordDiffInfo, ByRef diffCount As Long)
    Dim i As Long, j As Long
    Dim tempDiffs() As WordDiffInfo
    Dim tempCount As Long
    Dim k As Long
    Dim matchedOld() As Long
    Dim matchedNew() As Long
    Dim matchedCount As Long

    matchedCount = 0
    ReDim matchedOld(0 To 0)
    ReDim matchedNew(0 To 0)

    tempCount = 0
    ReDim tempDiffs(0 To 0)

    i = n1
    j = n2

    Do While i > 0 Or j > 0
        If i > 0 And j > 0 And texts1(i) = texts2(j) Then
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
            If Len(texts2(j)) > 0 Then
                AddTempWordDiff tempDiffs, tempCount, 0, j, "追加", "", texts2(j), "", ""
            End If
            j = j - 1
        ElseIf i > 0 And (j = 0 Or lcsMatrix(i - 1, j) > lcsMatrix(i, j - 1)) Then
            If Len(texts1(i)) > 0 Then
                AddTempWordDiff tempDiffs, tempCount, i, 0, "削除", texts1(i), "", "", ""
            End If
            i = i - 1
        Else
            Exit Do
        End If
    Loop

    ' 逆順を正順に変換
    diffCount = tempCount
    If tempCount > 0 Then
        ReDim differences(0 To tempCount - 1)
        For k = 0 To tempCount - 1
            differences(k) = tempDiffs(tempCount - 1 - k)
        Next k
    Else
        ReDim differences(0 To 0)
    End If

    MergeAdjacentChanges differences, diffCount

    ' テキスト一致段落をモジュールレベル変数に保存
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

' ============================================================================
' 隣接する削除と追加を「変更」にマージ
' ============================================================================
Private Sub MergeAdjacentChanges(ByRef differences() As WordDiffInfo, ByRef diffCount As Long)
    Dim i As Long
    Dim newDiffs() As WordDiffInfo
    Dim newCount As Long
    Dim merged As Boolean

    If diffCount <= 1 Then Exit Sub

    newCount = 0
    ReDim newDiffs(0 To diffCount - 1)

    i = 0
    Do While i < diffCount
        merged = False

        If i < diffCount - 1 Then
            If differences(i).DiffType = "削除" And differences(i + 1).DiffType = "追加" Then
                If Abs(differences(i).OldParagraphNo - differences(i + 1).NewParagraphNo) <= 1 Or _
                   (differences(i).OldParagraphNo > 0 And differences(i + 1).NewParagraphNo > 0) Then
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

    diffCount = newCount
    If newCount > 0 Then
        ReDim differences(0 To newCount - 1)
        For i = 0 To newCount - 1
            differences(i) = newDiffs(i)
        Next i
    End If
End Sub

' ============================================================================
' 一時差分配列に追加
' ============================================================================
Private Sub AddTempWordDiff(ByRef tempDiffs() As WordDiffInfo, ByRef tempCount As Long, _
                           ByVal oldParaNo As Long, ByVal newParaNo As Long, _
                           ByVal diffType As String, ByVal oldText As String, ByVal newText As String, _
                           ByVal oldStyle As String, ByVal newStyle As String)
    If tempCount = 0 Then
        ReDim tempDiffs(0 To 0)
    Else
        ReDim Preserve tempDiffs(0 To tempCount)
    End If

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

' ============================================================================
' 簡易差分検出（大きなファイル用）
' ============================================================================
Private Sub ComputeSimpleDiffOptimized(ByRef texts1() As String, ByRef texts2() As String, _
                                       ByVal n1 As Long, ByVal n2 As Long, _
                                       ByRef differences() As WordDiffInfo, ByRef diffCount As Long)
    Dim i1 As Long, i2 As Long
    Dim matchFound As Boolean
    Dim lookAhead As Long
    Dim j As Long
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
    lookAhead = 50

    Do While i1 <= n1 Or i2 <= n2
        If i1 <= n1 And i2 <= n2 Then
            If texts1(i1) = texts2(i2) Then
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
                matchFound = False

                For j = i2 + 1 To Application.WorksheetFunction.Min(i2 + lookAhead, n2)
                    If texts1(i1) = texts2(j) Then
                        Do While i2 < j
                            If Len(texts2(i2)) > 0 Then
                                AddWordDiff differences, diffCount, 0, i2, "追加", "", texts2(i2), "", ""
                            End If
                            i2 = i2 + 1
                        Loop
                        matchFound = True
                        Exit For
                    End If
                Next j

                If Not matchFound Then
                    For j = i1 + 1 To Application.WorksheetFunction.Min(i1 + lookAhead, n1)
                        If texts1(j) = texts2(i2) Then
                            Do While i1 < j
                                If Len(texts1(i1)) > 0 Then
                                    AddWordDiff differences, diffCount, i1, 0, "削除", texts1(i1), "", "", ""
                                End If
                                i1 = i1 + 1
                            Loop
                            matchFound = True
                            Exit For
                        End If
                    Next j
                End If

                If Not matchFound Then
                    If Len(texts1(i1)) > 0 Or Len(texts2(i2)) > 0 Then
                        AddWordDiff differences, diffCount, i1, i2, "変更", texts1(i1), texts2(i2), "", ""
                    End If
                    i1 = i1 + 1
                    i2 = i2 + 1
                End If
            End If
        ElseIf i1 <= n1 Then
            If Len(texts1(i1)) > 0 Then
                AddWordDiff differences, diffCount, i1, 0, "削除", texts1(i1), "", "", ""
            End If
            i1 = i1 + 1
        Else
            If Len(texts2(i2)) > 0 Then
                AddWordDiff differences, diffCount, 0, i2, "追加", "", texts2(i2), "", ""
            End If
            i2 = i2 + 1
        End If

        If (i1 + i2) Mod 100 = 0 Then
            ShowProgress "[3/4] 差分計算(簡易)", i1 + i2, n1 + n2
        End If
    Loop
    ShowProgress "[3/4] 差分計算(簡易)", n1 + n2, n1 + n2

    ' テキスト一致段落をモジュールレベル変数に保存
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

' ============================================================================
' Word差分を追加
' ============================================================================
Public Sub AddWordDiff(ByRef differences() As WordDiffInfo, ByRef diffCount As Long, _
                      ByVal oldParaNo As Long, ByVal newParaNo As Long, _
                      ByVal diffType As String, ByVal oldText As String, ByVal newText As String, _
                      ByVal oldStyle As String, ByVal newStyle As String)
    If diffCount = 0 Then
        ReDim differences(0 To 0)
    Else
        ReDim Preserve differences(0 To diffCount)
    End If

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

' ============================================================================
' 段落のスタイル情報を取得
' ============================================================================
Private Function GetParagraphStyleInfo(ByRef para As Object) As String
    Dim styleInfo As String
    Dim fontName As String
    Dim fontSize As Single
    Dim isBold As Boolean
    Dim isItalic As Boolean
    Dim styleName As String

    On Error Resume Next

    styleName = para.Style.NameLocal
    If Err.Number <> 0 Then styleName = "(不明)"
    Err.Clear

    fontName = para.Range.Font.Name
    If Err.Number <> 0 Or fontName = "" Then fontName = "(混在)"
    Err.Clear

    fontSize = para.Range.Font.Size
    If Err.Number <> 0 Or fontSize = 9999999 Then
        fontSize = 0
    End If
    Err.Clear

    isBold = (para.Range.Font.Bold = True)
    isItalic = (para.Range.Font.Italic = True)

    On Error GoTo 0

    styleInfo = "[" & styleName & "] " & fontName & " " & Format(fontSize, "0.0") & "pt"
    If isBold Then styleInfo = styleInfo & " 太字"
    If isItalic Then styleInfo = styleInfo & " 斜体"

    GetParagraphStyleInfo = styleInfo
End Function

' ============================================================================
' Word結果シートを作成
' ============================================================================
Public Sub CreateWordResultSheet(ByRef differences() As WordDiffInfo, ByVal diffCount As Long, _
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
    ThisWorkbook.Worksheets(SHEET_RESULT).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' 新しいシートを作成
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = SHEET_RESULT

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
            .Fill.ForeColor.RGB = RGB(255, 152, 0)
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
            .Fill.ForeColor.RGB = RGB(33, 150, 243)
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
        .Range("E9").Interior.Color = COLOR_STYLE

        ' ヘッダー
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

            If differences(i).OldParagraphNo > 0 Then
                oldParaStr = CStr(differences(i).OldParagraphNo)
            Else
                oldParaStr = "-"
            End If
            .Cells(row, 2).Value = oldParaStr
            .Cells(row, 2).HorizontalAlignment = xlCenter

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

            .Cells(row, 5).WrapText = True
            .Cells(row, 6).WrapText = True

            Select Case differences(i).DiffType
                Case "変更"
                    .Range(.Cells(row, 1), .Cells(row, 8)).Interior.Color = COLOR_CHANGED
                Case "追加"
                    .Range(.Cells(row, 1), .Cells(row, 8)).Interior.Color = COLOR_ADDED
                Case "削除"
                    .Range(.Cells(row, 1), .Cells(row, 8)).Interior.Color = COLOR_DELETED
                Case "スタイル変更"
                    .Range(.Cells(row, 1), .Cells(row, 8)).Interior.Color = COLOR_STYLE
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

        .Range("A11:H11").AutoFilter

        .Activate
        .Rows(12).Select
        ActiveWindow.FreezePanes = True

        .Range("A1").Select
    End With
End Sub

' ============================================================================
' 選択行のWord差分を旧ファイルで検索して開く
' ============================================================================
Public Sub SearchInOldWordFile()
    SearchWordDifference True
End Sub

' ============================================================================
' 選択行のWord差分を新ファイルで検索して開く
' ============================================================================
Public Sub SearchInNewWordFile()
    SearchWordDifference False
End Sub

' ============================================================================
' Word差分を検索して開く（内部処理）
' ============================================================================
Private Sub SearchWordDifference(ByVal isOldFile As Boolean)
    Dim ws As Worksheet
    Dim selectedRow As Long
    Dim filePath As String
    Dim searchText As String
    Dim wordApp As Object
    Dim doc As Object

    On Error GoTo ErrorHandler

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_RESULT)
    On Error GoTo ErrorHandler

    If ws Is Nothing Then
        MsgBox "比較結果シートが見つかりません。" & vbCrLf & _
               "先にWord比較を実行してください。", vbExclamation, "エラー"
        Exit Sub
    End If

    selectedRow = Selection.row

    If selectedRow < 12 Then
        MsgBox "差異データの行を選択してください。" & vbCrLf & _
               "（12行目以降のデータ行を選択）", vbExclamation, "行選択エラー"
        Exit Sub
    End If

    If isOldFile Then
        filePath = ws.Range("B3").Value
        searchText = ws.Cells(selectedRow, 5).Value
    Else
        filePath = ws.Range("B4").Value
        searchText = ws.Cells(selectedRow, 6).Value
    End If

    If Len(Trim(searchText)) = 0 Then
        MsgBox "検索するテキストがありません。" & vbCrLf & _
               IIf(isOldFile, "旧ファイル側", "新ファイル側") & "にテキストがない差異です。", _
               vbExclamation, "検索エラー"
        Exit Sub
    End If

    If Dir(filePath) = "" Then
        MsgBox "ファイルが見つかりません: " & vbCrLf & filePath, vbCritical, "ファイルエラー"
        Exit Sub
    End If

    If Len(searchText) > 100 Then
        searchText = Left(searchText, 100)
    End If

    On Error Resume Next
    Set wordApp = GetObject(, "Word.Application")
    If wordApp Is Nothing Then
        Set wordApp = CreateObject("Word.Application")
    End If
    On Error GoTo ErrorHandler

    wordApp.Visible = True

    Set doc = wordApp.Documents.Open(filePath, ReadOnly:=True)

    With doc.Content.Find
        .ClearFormatting
        .Text = searchText
        .Forward = True
        .Wrap = 1
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False

        If .Execute Then
            doc.ActiveWindow.ScrollIntoView doc.Content.Find.Parent
            doc.Content.Find.Parent.Select
            MsgBox "テキストが見つかりました。", vbInformation, "検索完了"
        Else
            MsgBox "テキストが見つかりませんでした。" & vbCrLf & vbCrLf & _
                   "検索テキスト: " & Left(searchText, 50) & IIf(Len(searchText) > 50, "...", ""), _
                   vbExclamation, "検索結果"
        End If
    End With

    wordApp.Activate

    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & vbCrLf & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, vbCritical, "エラー"
End Sub
