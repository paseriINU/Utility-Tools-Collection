Attribute VB_Name = "FS_Main"
Option Explicit

'==============================================================================
' Excel フィルター検索ツール - メインモジュール
' OR条件で複数キーワードのフィルターを適用（テーブル専用）
'==============================================================================

' ============================================================================
' 設定（必要に応じて変更してください）
' ============================================================================
Private Const TARGET_SHEET_NAME As String = "テスト"    ' 対象シート名
Private Const FILTER_COLUMN_A As Long = 9               ' フィルター対象列1（I列=9）
Private Const FILTER_COLUMN_B As Long = 10              ' フィルター対象列2（J列=10）

' ============================================================================
' 公開プロシージャ
' ============================================================================

' フォームを表示
Public Sub ShowFilterSearchForm()
    FSearchForm.Show vbModeless
End Sub

' OR条件でフィルターを適用（テーブル専用）
Public Sub ApplyOrFilter(keywords() As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim keywordCount As Long
    Dim matchingValues() As String
    Dim matchCount As Long

    On Error GoTo ErrorHandler

    ' 対象シートを取得
    Set ws = GetTargetSheet()
    If ws Is Nothing Then Exit Sub

    ' テーブルを取得
    Set tbl = GetTargetTable(ws)
    If tbl Is Nothing Then
        MsgBox "テーブルが見つかりません。", vbExclamation, "エラー"
        Exit Sub
    End If

    ' 既存のフィルターをクリア
    If Not tbl.AutoFilter Is Nothing Then
        If tbl.AutoFilter.FilterMode Then
            tbl.AutoFilter.ShowAllData
        End If
    End If

    ' キーワード数を取得
    keywordCount = UBound(keywords) - LBound(keywords) + 1

    ' キーワードに一致する値を収集（2列をOR検索）
    matchCount = CollectMatchingValuesFromTwoColumns(tbl, FILTER_COLUMN_A, FILTER_COLUMN_B, _
        keywords, matchingValues)

    If matchCount = 0 Then
        MsgBox "一致するデータがありません。", vbInformation, "検索結果"
        ws.Activate
        Exit Sub
    End If

    ' フィルターを適用（FILTER_COLUMN_Aでフィルター）
    ApplyTableFilter tbl, FILTER_COLUMN_A, matchingValues, matchCount

    ' 対象シートをアクティブにする
    ws.Activate

    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub

' フィルターをクリア
Public Sub ClearFilter()
    Dim ws As Worksheet
    Dim tbl As ListObject

    On Error GoTo ErrorHandler

    Set ws = GetTargetSheet()
    If ws Is Nothing Then Exit Sub

    ' テーブルのフィルターをクリア
    For Each tbl In ws.ListObjects
        If Not tbl.AutoFilter Is Nothing Then
            If tbl.AutoFilter.FilterMode Then
                tbl.AutoFilter.ShowAllData
            End If
        End If
    Next tbl

    ' 対象シートをアクティブにする
    ws.Activate

    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub

' ============================================================================
' 内部プロシージャ
' ============================================================================

' 対象シートを取得
Private Function GetTargetSheet() As Worksheet
    On Error Resume Next
    Set GetTargetSheet = ThisWorkbook.Worksheets(TARGET_SHEET_NAME)
    On Error GoTo 0

    If GetTargetSheet Is Nothing Then
        MsgBox "シート「" & TARGET_SHEET_NAME & "」が見つかりません。", vbCritical, "エラー"
    End If
End Function

' 対象テーブルを取得（シート内の最初のテーブル）
Private Function GetTargetTable(ws As Worksheet) As ListObject
    If ws.ListObjects.Count > 0 Then
        Set GetTargetTable = ws.ListObjects(1)
    End If
End Function

' 2列をOR検索してキーワードに部分一致する行の値を収集
Private Function CollectMatchingValuesFromTwoColumns(tbl As ListObject, _
    colA As Long, colB As Long, _
    keywords() As String, ByRef matchingValues() As String) As Long

    Dim rowCount As Long
    Dim r As Long
    Dim i As Long
    Dim valueA As String
    Dim valueB As String
    Dim isMatch As Boolean
    Dim dict As Object

    ' 重複排除用のDictionary
    Set dict = CreateObject("Scripting.Dictionary")

    ' データ範囲を取得（ヘッダー除く）
    If tbl.DataBodyRange Is Nothing Then
        CollectMatchingValuesFromTwoColumns = 0
        Exit Function
    End If

    rowCount = tbl.DataBodyRange.Rows.Count

    ' 各行をチェック
    For r = 1 To rowCount
        valueA = CStr(tbl.DataBodyRange.Cells(r, colA).Value)
        valueB = CStr(tbl.DataBodyRange.Cells(r, colB).Value)

        ' いずれかのキーワードがcolA列またはcolB列に部分一致するかチェック
        isMatch = False
        For i = LBound(keywords) To UBound(keywords)
            If InStr(1, valueA, keywords(i), vbTextCompare) > 0 Or _
               InStr(1, valueB, keywords(i), vbTextCompare) > 0 Then
                isMatch = True
                Exit For
            End If
        Next i

        ' 一致した場合、colA列の値をDictionaryに追加（重複排除）
        If isMatch Then
            If Not dict.Exists(valueA) Then
                dict.Add valueA, True
            End If
        End If
    Next r

    ' 結果を配列に変換
    Dim matchCount As Long
    matchCount = dict.Count
    If matchCount > 0 Then
        ReDim matchingValues(1 To matchCount)
        i = 1
        Dim key As Variant
        For Each key In dict.Keys
            matchingValues(i) = CStr(key)
            i = i + 1
        Next key
    End If

    CollectMatchingValuesFromTwoColumns = matchCount
End Function

' テーブルにフィルターを適用
Private Sub ApplyTableFilter(tbl As ListObject, colNum As Long, _
    matchingValues() As String, matchCount As Long)

    Dim filterArray() As String
    Dim i As Long

    ' フィルター用の配列を作成
    ReDim filterArray(1 To matchCount)
    For i = 1 To matchCount
        filterArray(i) = matchingValues(i)
    Next i

    ' xlFilterValuesでフィルターを適用（完全一致だが、収集時に部分一致済み）
    tbl.Range.AutoFilter Field:=colNum, _
        Criteria1:=filterArray, _
        Operator:=xlFilterValues
End Sub

