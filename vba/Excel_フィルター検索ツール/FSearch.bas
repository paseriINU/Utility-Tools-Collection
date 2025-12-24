Option Explicit

' ========================================
' フィルター検索モジュール
' OR条件で複数キーワードのフィルターを適用
' ========================================

' ----------------------------------------
' 設定（必要に応じて変更してください）
' ----------------------------------------
Private Const TARGET_SHEET_NAME As String = "テスト"    ' 対象シート名
Private Const HEADER_ROW As Long = 3                    ' ヘッダー行（タイトル行）
Private Const DATA_START_COL As Long = 1                ' データ開始列（A列=1）
Private Const DATA_END_COL As Long = 18                 ' データ終了列（R列=18）
Private Const FILTER_COLUMN_A As Long = 1               ' フィルター列1（A列=1）
Private Const FILTER_COLUMN_B As Long = 2               ' フィルター列2（B列=2）

' ========================================
' 公開プロシージャ
' ========================================

' フォームを表示
Public Sub ShowFilterSearchForm()
    FSearchForm.Show vbModeless
End Sub

' OR条件でフィルターを適用
Public Sub ApplyOrFilter(keywords() As String)
    Dim ws As Worksheet
    Dim dataRange As Range
    Dim lastRow As Long
    Dim lastCol As Long
    Dim criteria() As String
    Dim i As Long
    Dim keywordCount As Long
    Dim tbl As ListObject
    Dim useTable As Boolean

    On Error GoTo ErrorHandler

    ' 対象シートを取得
    Set ws = GetTargetSheet()
    If ws Is Nothing Then Exit Sub

    ' データ範囲を取得（A列を下から上に見て最終行を取得）
    lastRow = ws.Cells(ws.Rows.Count, DATA_START_COL).End(xlUp).Row
    lastCol = DATA_END_COL  ' R列固定

    If lastRow <= HEADER_ROW Then
        MsgBox "データがありません。", vbExclamation, "エラー"
        Exit Sub
    End If

    ' データ範囲（ヘッダー行から最終行まで）
    Set dataRange = ws.Range(ws.Cells(HEADER_ROW, DATA_START_COL), ws.Cells(lastRow, lastCol))

    ' テーブル（ListObject）が存在するかチェック
    useTable = False
    For Each tbl In ws.ListObjects
        If Not Intersect(tbl.Range, dataRange) Is Nothing Then
            ' テーブルが見つかった場合、テーブルのフィルターを使用
            useTable = True
            Set dataRange = tbl.Range
            Exit For
        End If
    Next tbl

    ' 既存のフィルターをクリア
    If useTable Then
        ' テーブルのフィルターをクリア
        If tbl.AutoFilter.FilterMode Then
            tbl.AutoFilter.ShowAllData
        End If
    Else
        ' 通常のオートフィルターをクリア
        If ws.AutoFilterMode Then
            ws.AutoFilterMode = False
        End If
        ' オートフィルターを有効化
        dataRange.AutoFilter
    End If

    ' キーワード数を取得
    keywordCount = UBound(keywords) - LBound(keywords) + 1

    ' フィルター条件を作成（ワイルドカード付き）
    ReDim criteria(1 To keywordCount * 2)
    For i = LBound(keywords) To UBound(keywords)
        criteria((i - LBound(keywords)) * 2 + 1) = "=*" & keywords(i) & "*"
        If i < UBound(keywords) Then
            criteria((i - LBound(keywords)) * 2 + 2) = "=*" & keywords(i) & "*"
        End If
    Next i

    ' A列にフィルター適用
    ApplyColumnFilter ws, dataRange, FILTER_COLUMN_A, keywords

    ' B列にフィルター適用（A列の結果に追加）
    ' 注意: 複数列のOR条件は標準AutoFilterでは難しいため、
    ' AdvancedFilterまたは別のアプローチを使用

    MsgBox keywordCount & "個のキーワードでフィルターを適用しました。", vbInformation, "完了"

    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub

' フィルターをクリア
Public Sub ClearFilter()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim tableCleared As Boolean

    On Error GoTo ErrorHandler

    Set ws = GetTargetSheet()
    If ws Is Nothing Then Exit Sub

    tableCleared = False

    ' テーブル（ListObject）のフィルターをクリア
    For Each tbl In ws.ListObjects
        If tbl.AutoFilter.FilterMode Then
            tbl.AutoFilter.ShowAllData
            tableCleared = True
        End If
    Next tbl

    ' 通常のオートフィルターをクリア
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If

    MsgBox "フィルターをクリアしました。", vbInformation, "完了"

    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub

' ========================================
' 内部プロシージャ
' ========================================

' 対象シートを取得
Private Function GetTargetSheet() As Worksheet
    On Error Resume Next
    Set GetTargetSheet = ThisWorkbook.Worksheets(TARGET_SHEET_NAME)
    On Error GoTo 0

    If GetTargetSheet Is Nothing Then
        MsgBox "シート「" & TARGET_SHEET_NAME & "」が見つかりません。", vbCritical, "エラー"
    End If
End Function

' 指定列にOR条件フィルターを適用
Private Sub ApplyColumnFilter(ws As Worksheet, dataRange As Range, colNum As Long, keywords() As String)
    Dim criteria As Variant
    Dim i As Long
    Dim keywordCount As Long

    keywordCount = UBound(keywords) - LBound(keywords) + 1

    ' AutoFilterのCriteria1は最大2つまでしか指定できないため、
    ' 3つ以上の場合は配列でOperator:=xlFilterValuesを使用
    ' ただしxlFilterValuesは完全一致のため、部分一致にはAdvancedFilterが必要

    ' 2つ以下の場合は標準的なOR条件を使用
    If keywordCount = 1 Then
        dataRange.AutoFilter Field:=colNum, _
            Criteria1:="=*" & keywords(1) & "*", _
            Operator:=xlOr, _
            Criteria2:="=*" & keywords(1) & "*"

    ElseIf keywordCount = 2 Then
        dataRange.AutoFilter Field:=colNum, _
            Criteria1:="=*" & keywords(1) & "*", _
            Operator:=xlOr, _
            Criteria2:="=*" & keywords(2) & "*"

    Else
        ' 3つ以上の場合はAdvancedFilterを使用
        ApplyAdvancedFilter ws, colNum, keywords
    End If
End Sub

' AdvancedFilterで複数条件のOR検索を実行
Private Sub ApplyAdvancedFilter(ws As Worksheet, colNum As Long, keywords() As String)
    Dim criteriaRange As Range
    Dim dataRange As Range
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long
    Dim tempSheet As Worksheet
    Dim criteriaSheetName As String
    Dim tbl As ListObject

    ' 一時的な条件範囲を作成
    criteriaSheetName = "_FilterCriteria_"

    ' 既存の一時シートを削除
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(criteriaSheetName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' 条件シートを作成
    Set tempSheet = ThisWorkbook.Worksheets.Add
    tempSheet.Name = criteriaSheetName
    tempSheet.Visible = xlSheetVeryHidden

    ' 条件を設定（ヘッダー + 各キーワード）
    ' A列とB列の両方をOR条件にするため、列を並べる
    tempSheet.Cells(1, 1).Value = ws.Cells(HEADER_ROW, FILTER_COLUMN_A).Value
    tempSheet.Cells(1, 2).Value = ws.Cells(HEADER_ROW, FILTER_COLUMN_B).Value

    For i = LBound(keywords) To UBound(keywords)
        ' A列の条件
        tempSheet.Cells(i - LBound(keywords) + 2, 1).Value = "*" & keywords(i) & "*"
        ' B列の条件（同じ行に入れるとAND、別の行に入れるとOR）
    Next i

    ' B列用の条件も追加（OR条件なので別の行に）
    For i = LBound(keywords) To UBound(keywords)
        tempSheet.Cells(UBound(keywords) + i - LBound(keywords) + 2, 2).Value = "*" & keywords(i) & "*"
    Next i

    ' データ範囲（A列を下から上に見て最終行を取得）
    lastRow = ws.Cells(ws.Rows.Count, DATA_START_COL).End(xlUp).Row
    lastCol = DATA_END_COL  ' R列固定
    Set dataRange = ws.Range(ws.Cells(HEADER_ROW, DATA_START_COL), ws.Cells(lastRow, lastCol))

    ' 条件範囲
    Dim criteriaLastRow As Long
    criteriaLastRow = (UBound(keywords) - LBound(keywords) + 1) * 2 + 1
    Set criteriaRange = tempSheet.Range(tempSheet.Cells(1, 1), tempSheet.Cells(criteriaLastRow, 2))

    ' 既存のフィルターをクリア（テーブル対応）
    For Each tbl In ws.ListObjects
        If Not Intersect(tbl.Range, dataRange) Is Nothing Then
            If tbl.AutoFilter.FilterMode Then
                tbl.AutoFilter.ShowAllData
            End If
        End If
    Next tbl

    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If

    ' AdvancedFilterを適用（その場でフィルター）
    dataRange.AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:=criteriaRange, Unique:=False

    ' 一時シートを削除
    Application.DisplayAlerts = False
    tempSheet.Delete
    Application.DisplayAlerts = True
End Sub
