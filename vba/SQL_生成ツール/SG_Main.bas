Attribute VB_Name = "SG_Main"
Option Explicit

'==============================================================================
' Oracle SELECT文生成ツール - メインモジュール
' エントリーポイント、ユーザー操作機能を提供
'==============================================================================

'==============================================================================
' SQL生成メインプロシージャ
'==============================================================================
Public Sub GenerateSQL()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = Sheets(SHEET_MAIN)

    Dim sql As String
    Dim withClause As String
    Dim selectClause As String
    Dim fromClause As String
    Dim whereClause As String
    Dim groupByClause As String
    Dim havingClause As String
    Dim orderByClause As String
    Dim limitClause As String
    Dim unionClause As String

    ' WITH句の生成
    If ws.Range("E" & ROW_OPTIONS).Value = "使用する" Then
        withClause = SG_Generator.GenerateWithClause()
    End If

    ' SELECT句の生成
    selectClause = SG_Generator.GenerateSelectClause(ws)
    If selectClause = "" Then
        MsgBox "取得カラムを1つ以上指定してください。", vbExclamation, "入力エラー"
        Exit Sub
    End If

    ' FROM句とJOIN句の生成
    fromClause = SG_Generator.GenerateFromClause(ws)
    If fromClause = "" Then
        MsgBox "メインテーブルを指定してください。", vbExclamation, "入力エラー"
        Exit Sub
    End If

    ' WHERE句の生成
    whereClause = SG_Generator.GenerateWhereClause(ws)

    ' GROUP BY句の生成
    groupByClause = SG_Generator.GenerateGroupByClause(ws)

    ' HAVING句の生成
    havingClause = SG_Generator.GenerateHavingClause(ws)

    ' ORDER BY句の生成
    orderByClause = SG_Generator.GenerateOrderByClause(ws)

    ' 件数制限の生成
    limitClause = SG_Generator.GenerateLimitClause(ws)

    ' UNION句の生成
    If ws.Range("H" & ROW_OPTIONS).Value = "使用する" Then
        unionClause = SG_Generator.GenerateUnionClause()
    End If

    ' SQLを組み立て
    sql = ""

    If withClause <> "" Then
        sql = withClause & vbCrLf
    End If

    sql = sql & selectClause & vbCrLf
    sql = sql & fromClause

    If whereClause <> "" Then
        sql = sql & vbCrLf & whereClause
    End If

    If groupByClause <> "" Then
        sql = sql & vbCrLf & groupByClause
    End If

    If havingClause <> "" Then
        sql = sql & vbCrLf & havingClause
    End If

    If orderByClause <> "" Then
        sql = sql & vbCrLf & orderByClause
    End If

    If limitClause <> "" Then
        sql = sql & vbCrLf & limitClause
    End If

    If unionClause <> "" Then
        sql = sql & vbCrLf & unionClause
    End If

    sql = sql & ";"

    ' 結果を出力
    ws.Range("A" & ROW_SQL_OUTPUT + 1).Value = sql

    MsgBox "SQLを生成しました。", vbInformation, "完了"

    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub

'==============================================================================
' メインシートをクリア
'==============================================================================
Public Sub ClearMainSheet()
    Dim ws As Worksheet
    Set ws = Sheets(SHEET_MAIN)

    Dim i As Long

    ' オプション
    ws.Range("B" & ROW_OPTIONS).Value = ""
    ws.Range("E" & ROW_OPTIONS).Value = ""
    ws.Range("H" & ROW_OPTIONS).Value = ""

    ' メインテーブル
    ws.Range("B" & ROW_MAIN_TABLE + 1).Value = ""
    ws.Range("E" & ROW_MAIN_TABLE + 1).Value = ""

    ' JOIN
    For i = ROW_JOIN_START + 2 To ROW_JOIN_END
        ws.Range("B" & i).Value = ""
        ws.Range("C" & i).Value = ""
        ws.Range("D" & i).Value = ""
        ws.Range("E" & i).Value = ""
    Next i

    ' カラム
    For i = ROW_COLUMNS_START To ROW_COLUMNS_END
        ws.Range("B" & i).Value = ""
        ws.Range("C" & i).Value = ""
        ws.Range("D" & i).Value = ""
        ws.Range("E" & i).Value = ""
        ws.Range("F" & i).Value = ""
    Next i

    ' WHERE
    For i = ROW_WHERE_START To ROW_WHERE_END
        ws.Range("B" & i).Value = ""
        ws.Range("C" & i).Value = ""
        ws.Range("D" & i).Value = ""
        ws.Range("E" & i).Value = ""
        ws.Range("F" & i).Value = ""
        ws.Range("G" & i).Value = ""
        ws.Range("H" & i).Value = ""
    Next i

    ' GROUP BY
    ws.Range("B" & ROW_GROUPBY + 1).Value = ""

    ' HAVING
    For i = ROW_HAVING_START To ROW_HAVING_END
        ws.Range("B" & i).Value = ""
        ws.Range("C" & i).Value = ""
    Next i

    ' ORDER BY
    For i = ROW_ORDERBY_START To ROW_ORDERBY_END
        ws.Range("B" & i).Value = ""
        ws.Range("C" & i).Value = ""
        ws.Range("D" & i).Value = ""
        ws.Range("E" & i).Value = ""
    Next i

    ' 件数制限
    ws.Range("B" & ROW_LIMIT + 1).Value = ""
    ws.Range("D" & ROW_LIMIT + 1).Value = "100"
    ws.Range("F" & ROW_LIMIT + 1).Value = "FETCH FIRST"

    ' SQL出力
    ws.Range("A" & ROW_SQL_OUTPUT + 1).Value = ""

    MsgBox "入力内容をクリアしました。", vbInformation, "クリア完了"
End Sub

'==============================================================================
' 生成したSQLを履歴に保存
'==============================================================================
Public Sub SaveToHistory()
    Dim wsMain As Worksheet
    Dim wsHistory As Worksheet
    Dim sql As String
    Dim description As String
    Dim nextRow As Long

    Set wsMain = Sheets(SHEET_MAIN)
    Set wsHistory = Sheets(SHEET_HISTORY)

    sql = Trim(wsMain.Range("A" & ROW_SQL_OUTPUT + 1).Value)

    If sql = "" Then
        MsgBox "保存するSQLがありません。先にSQLを生成してください。", vbExclamation, "エラー"
        Exit Sub
    End If

    description = InputBox("このSQLの説明を入力してください（省略可）:", "履歴保存")

    ' 次の空き行を探す
    nextRow = wsHistory.Cells(wsHistory.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 4 Then nextRow = 4

    wsHistory.Range("A" & nextRow).Value = nextRow - 3
    wsHistory.Range("B" & nextRow).Value = Now
    wsHistory.Range("B" & nextRow).NumberFormat = "yyyy/mm/dd hh:mm:ss"
    wsHistory.Range("C" & nextRow).Value = description
    wsHistory.Range("D" & nextRow).Value = sql

    MsgBox "SQLを履歴に保存しました。" & vbCrLf & "No: " & (nextRow - 3), vbInformation, "保存完了"
End Sub

'==============================================================================
' 生成したSQLをクリップボードにコピー
'==============================================================================
Public Sub CopySQL()
    Dim ws As Worksheet
    Dim sql As String
    Dim dataObj As Object

    Set ws = Sheets(SHEET_MAIN)
    sql = Trim(ws.Range("A" & ROW_SQL_OUTPUT + 1).Value)

    If sql = "" Then
        MsgBox "コピーするSQLがありません。先にSQLを生成してください。", vbExclamation, "エラー"
        Exit Sub
    End If

    ' クリップボードにコピー
    On Error Resume Next
    Set dataObj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    dataObj.SetText sql
    dataObj.PutInClipboard
    On Error GoTo 0

    If dataObj Is Nothing Then
        ' MSFormsが使えない場合は手動コピーを促す
        ws.Range("A" & ROW_SQL_OUTPUT + 1).Select
        Selection.Copy
        MsgBox "SQLをコピーしました。" & vbCrLf & "(セル選択状態でCtrl+Cでもコピーできます)", vbInformation, "コピー完了"
    Else
        MsgBox "SQLをクリップボードにコピーしました。", vbInformation, "コピー完了"
    End If
End Sub

'==============================================================================
' テーブル定義からプルダウンを更新
'==============================================================================
Public Sub UpdateDropdownsFromTableDef()
    On Error GoTo ErrorHandler

    Dim wsMain As Worksheet
    Dim wsDef As Worksheet
    Dim tableList As String
    Dim i As Long

    Set wsMain = Sheets(SHEET_MAIN)
    Set wsDef = Sheets(SHEET_TABLE_DEF)

    ' テーブル一覧を取得
    tableList = SG_Generator.GetTableList()

    If tableList = "" Then
        MsgBox "テーブル定義シートにテーブルが登録されていません。" & vbCrLf & _
               "「テーブル定義」シートのB列にテーブル名を登録してください。", vbExclamation, "確認"
        Exit Sub
    End If

    ' メインテーブルのプルダウンを更新（テーブル一覧用プレフィックス）
    SG_Setup.AddDropdown wsMain, "B" & ROW_MAIN_TABLE + 1, tableList, "TableList"

    ' JOINテーブルのプルダウンを更新（テーブル一覧用プレフィックス）
    For i = ROW_JOIN_START + 2 To ROW_JOIN_END
        SG_Setup.AddDropdown wsMain, "C" & i, tableList, "TableList"
    Next i

    ' カラム選択のプルダウンを更新（全テーブルの全カラム、カラム一覧用プレフィックス）
    Dim columnList As String
    columnList = SG_Generator.GetAllColumnList()

    For i = ROW_COLUMNS_START To ROW_COLUMNS_END
        If columnList <> "" Then
            SG_Setup.AddDropdown wsMain, "C" & i, columnList, "ColumnList"
        End If
    Next i

    ' WHERE句のカラムプルダウンを更新（カラム一覧用プレフィックス）
    For i = ROW_WHERE_START To ROW_WHERE_END
        If columnList <> "" Then
            SG_Setup.AddDropdown wsMain, "E" & i, columnList, "ColumnList"
        End If
    Next i

    ' ORDER BY句のカラムプルダウンを更新（カラム一覧用プレフィックス）
    For i = ROW_ORDERBY_START To ROW_ORDERBY_END
        If columnList <> "" Then
            SG_Setup.AddDropdown wsMain, "C" & i, columnList, "ColumnList"
        End If
    Next i

    ' テーブル別名のプルダウンを更新（入力済みの別名から取得）
    Dim aliasList As String
    aliasList = SG_Generator.GetAliasListFromMain()

    If aliasList <> "" Then
        For i = ROW_COLUMNS_START To ROW_COLUMNS_END
            SG_Setup.AddDropdown wsMain, "B" & i, aliasList
        Next i
        For i = ROW_WHERE_START To ROW_WHERE_END
            SG_Setup.AddDropdown wsMain, "D" & i, aliasList
        Next i
        For i = ROW_ORDERBY_START To ROW_ORDERBY_END
            SG_Setup.AddDropdown wsMain, "B" & i, aliasList
        Next i
    End If

    MsgBox "プルダウンを更新しました。" & vbCrLf & vbCrLf & _
           "テーブル数: " & UBound(Split(tableList, ",")) + 1, vbInformation, "更新完了"

    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub

'==============================================================================
' 選択テーブルに基づいてカラムドロップダウンを更新（ボタン用）
'==============================================================================
Public Sub RefreshColumnDropdownsByTable()
    On Error GoTo ErrorHandler

    Dim wsMain As Worksheet
    Dim columnList As String
    Dim i As Long
    Dim tableCount As Long
    Dim tableName As String

    Set wsMain = Sheets(SHEET_MAIN)

    ' 選択されたテーブルのカラム一覧を取得
    columnList = SG_Generator.GetColumnListForSelectedTables()

    If columnList = "" Or columnList = "*" Then
        MsgBox "テーブルが選択されていません。" & vbCrLf & _
               "メインテーブルまたはJOINテーブルを選択してから実行してください。", vbExclamation, "確認"
        Exit Sub
    End If

    ' カラム選択のプルダウンを更新（カラム一覧用プレフィックス）
    For i = ROW_COLUMNS_START To ROW_COLUMNS_END
        SG_Setup.AddDropdown wsMain, "C" & i, columnList, "ColumnList"
    Next i

    ' WHERE句のカラムプルダウンを更新（カラム一覧用プレフィックス）
    For i = ROW_WHERE_START To ROW_WHERE_END
        SG_Setup.AddDropdown wsMain, "E" & i, columnList, "ColumnList"
    Next i

    ' ORDER BY句のカラムプルダウンを更新（カラム一覧用プレフィックス）
    For i = ROW_ORDERBY_START To ROW_ORDERBY_END
        SG_Setup.AddDropdown wsMain, "C" & i, columnList, "ColumnList"
    Next i

    ' 選択されたテーブル数をカウント
    tableCount = 0
    tableName = SG_Generator.ExtractTableName(Trim(wsMain.Range("B" & ROW_MAIN_TABLE + 1).Value))
    If tableName <> "" Then tableCount = tableCount + 1

    For i = ROW_JOIN_START + 2 To ROW_JOIN_END
        tableName = SG_Generator.ExtractTableName(Trim(wsMain.Range("C" & i).Value))
        If tableName <> "" Then tableCount = tableCount + 1
    Next i

    MsgBox "カラムプルダウンを更新しました。" & vbCrLf & vbCrLf & _
           "対象テーブル数: " & tableCount, vbInformation, "更新完了"

    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub

'==============================================================================
' テーブル選択時の自動カラム更新（Worksheet_Changeイベントから呼び出し）
' 引数: changedRange - 変更されたセル範囲
'==============================================================================
Public Sub OnTableSelectionChanged(ByVal changedRange As Range)
    On Error Resume Next

    Dim wsMain As Worksheet
    Dim targetRow As Long
    Dim targetCol As String
    Dim isTableCell As Boolean
    Dim i As Long
    Dim nm As Name

    Set wsMain = Sheets(SHEET_MAIN)

    ' 変更されたセルがメインシートでない場合は終了
    If changedRange.Worksheet.Name <> SHEET_MAIN Then Exit Sub

    ' 変更されたセルがテーブル選択セルかチェック
    isTableCell = False

    ' メインテーブル（B7）
    If changedRange.Row = ROW_MAIN_TABLE + 1 And changedRange.Column = 2 Then
        isTableCell = True
    End If

    ' JOINテーブル（C11～C18）
    If changedRange.Column = 3 Then
        For i = ROW_JOIN_START + 2 To ROW_JOIN_END
            If changedRange.Row = i Then
                isTableCell = True
                Exit For
            End If
        Next i
    End If

    ' テーブル選択セルでない場合は終了
    If Not isTableCell Then Exit Sub

    ' カラムドロップダウンを自動更新（メッセージなし）
    Dim columnList As String
    columnList = SG_Generator.GetColumnListForSelectedTables()

    ' テーブルが選択されていない場合は終了
    If columnList = "" Or columnList = "*" Then Exit Sub

    Application.EnableEvents = False

    ' 既存のColumnList_*名前付き範囲を削除（カラム用のみ、テーブル用は残す）
    For Each nm In ThisWorkbook.Names
        If Left(nm.Name, 11) = "ColumnList_" Then
            nm.Delete
        End If
    Next nm

    ' カラム選択のプルダウンを更新（カラム一覧用プレフィックス）
    For i = ROW_COLUMNS_START To ROW_COLUMNS_END
        SG_Setup.AddDropdown wsMain, "C" & i, columnList, "ColumnList"
    Next i

    ' WHERE句のカラムプルダウンを更新（カラム一覧用プレフィックス）
    For i = ROW_WHERE_START To ROW_WHERE_END
        SG_Setup.AddDropdown wsMain, "E" & i, columnList, "ColumnList"
    Next i

    ' ORDER BY句のカラムプルダウンを更新（カラム一覧用プレフィックス）
    For i = ROW_ORDERBY_START To ROW_ORDERBY_END
        SG_Setup.AddDropdown wsMain, "C" & i, columnList, "ColumnList"
    Next i

    Application.EnableEvents = True
End Sub

'==============================================================================
' 別名更新ボタン用マクロ（テーブル別名を入力後に実行）
'==============================================================================
Public Sub RefreshAliasDropdowns()
    On Error GoTo ErrorHandler

    Dim wsMain As Worksheet
    Dim aliasList As String
    Dim i As Long

    Set wsMain = Sheets(SHEET_MAIN)

    ' 入力済みの別名を取得
    aliasList = SG_Generator.GetAliasListFromMain()

    If aliasList = "" Or aliasList = "," Then
        MsgBox "テーブル別名が入力されていません。" & vbCrLf & _
               "メインテーブルやJOINテーブルに別名を入力してから実行してください。", vbExclamation, "確認"
        Exit Sub
    End If

    ' 各セクションの「テーブル別名」プルダウンを更新
    For i = ROW_COLUMNS_START To ROW_COLUMNS_END
        SG_Setup.AddDropdown wsMain, "B" & i, aliasList
    Next i

    For i = ROW_WHERE_START To ROW_WHERE_END
        SG_Setup.AddDropdown wsMain, "D" & i, aliasList
    Next i

    For i = ROW_ORDERBY_START To ROW_ORDERBY_END
        SG_Setup.AddDropdown wsMain, "B" & i, aliasList
    Next i

    MsgBox "テーブル別名のプルダウンを更新しました。" & vbCrLf & _
           "別名: " & Replace(aliasList, ",", ", "), vbInformation, "更新完了"

    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub

'==============================================================================
' テスト実行用
'==============================================================================
Public Sub TestGenerateSQL()
    SG_Setup.InitializeSQL生成ツール
End Sub

'==============================================================================
' テーブル定義書インポート機能
'==============================================================================
Public Sub ImportTableDefinitions()
    On Error GoTo ErrorHandler

    Dim wsDef As Worksheet
    Dim folderPath As String
    Dim fileName As String
    Dim sourceWb As Workbook
    Dim sourceWs As Worksheet
    Dim tableNameCell As String
    Dim tableDescCell As String
    Dim columnStartRow As Long
    Dim colNumber As String
    Dim colItemName As String
    Dim colName As String
    Dim colDataType As String
    Dim colLength As String
    Dim colNullable As String
    Dim tableName As String
    Dim tableDesc As String
    Dim tableCount As Long
    Dim columnCount As Long
    Dim nextTableRow As Long
    Dim nextColumnRow As Long
    Dim sheetIdx As Long
    Dim currentRow As Long
    Dim colNumberValue As Variant
    Dim colNameValue As String
    Dim colItemNameValue As String
    Dim colTypeValue As String
    Dim colLengthValue As String
    Dim colNullableValue As String
    Dim importedTables As String
    Dim presetPath As String
    Dim exceptionDbName As String
    Dim colOffset As Long
    Dim actualColItemName As String
    Dim actualColName As String
    Dim actualColDataType As String
    Dim actualColLength As String
    Dim actualColNullable As String
    Dim actualTableNameCell As String
    Dim actualTableDescCell As String
    Dim emptyCount As Long
    Dim checkRow As Long

    Set wsDef = Sheets(SHEET_TABLE_DEF)

    ' 設定を取得
    tableNameCell = SG_Generator.GetImportSetting(wsDef, "テーブル名セル", DEFAULT_TABLE_NAME_CELL)
    tableDescCell = SG_Generator.GetImportSetting(wsDef, "テーブル名称セル", DEFAULT_TABLE_DESC_CELL)
    columnStartRow = CLng(SG_Generator.GetImportSetting(wsDef, "カラム開始行", CStr(DEFAULT_COLUMN_START_ROW)))
    colNumber = SG_Generator.GetImportSetting(wsDef, "カラム番号列", DEFAULT_COL_NUMBER)
    colItemName = SG_Generator.GetImportSetting(wsDef, "項目名列", DEFAULT_COL_ITEM_NAME)
    colName = SG_Generator.GetImportSetting(wsDef, "カラム名列", DEFAULT_COL_NAME)
    colDataType = SG_Generator.GetImportSetting(wsDef, "データ型列", DEFAULT_COL_DATATYPE)
    colLength = SG_Generator.GetImportSetting(wsDef, "桁数列", DEFAULT_COL_LENGTH)
    colNullable = SG_Generator.GetImportSetting(wsDef, "NULL列", DEFAULT_COL_NULLABLE)

    ' フォルダパス設定を取得
    presetPath = Trim(CStr(wsDef.Range("K15").Value))
    exceptionDbName = Trim(CStr(wsDef.Range("K16").Value))

    ' %USERNAME%を展開
    If InStr(presetPath, "%USERNAME%") > 0 Then
        presetPath = Replace(presetPath, "%USERNAME%", Environ("USERNAME"))
    End If

    ' フォルダパスの決定（設定済みならそのまま使用、なければダイアログ表示）
    If presetPath <> "" Then
        folderPath = presetPath
    Else
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = "テーブル定義書フォルダを選択"
            .AllowMultiSelect = False
            If .Show = -1 Then
                folderPath = .SelectedItems(1)
            Else
                MsgBox "フォルダ選択がキャンセルされました。", vbInformation, "キャンセル"
                Exit Sub
            End If
        End With
    End If

    ' フォルダパスの末尾にバックスラッシュを追加
    If Right(folderPath, 1) <> "\" And Right(folderPath, 1) <> "/" Then
        folderPath = folderPath & "\"
    End If

    ' フォルダ存在チェック
    If Dir(folderPath, vbDirectory) = "" Then
        MsgBox "フォルダが見つかりません: " & folderPath, vbExclamation, "エラー"
        Exit Sub
    End If

    ' 既存データをクリア（常に置換）
    If wsDef.Range("B6").Value <> "" Then
        wsDef.Range("A6:C" & wsDef.Cells(wsDef.Rows.Count, "B").End(xlUp).Row).ClearContents
    End If
    If wsDef.Range("E6").Value <> "" Then
        wsDef.Range("E6:H" & wsDef.Cells(wsDef.Rows.Count, "E").End(xlUp).Row).ClearContents
    End If
    nextTableRow = 6
    nextColumnRow = 6

    tableCount = 0
    columnCount = 0
    importedTables = ""

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' フォルダ内のExcelファイルを処理
    fileName = Dir(folderPath & "*.xls*")

    Do While fileName <> ""
        ' 自分自身をスキップ
        If folderPath & fileName <> ThisWorkbook.FullName Then
            ' ファイルを開く
            Set sourceWb = Workbooks.Open(folderPath & fileName, ReadOnly:=True, UpdateLinks:=0)

            ' 全シートを処理
            For sheetIdx = 1 To sourceWb.Sheets.Count
                Set sourceWs = sourceWb.Sheets(sheetIdx)

                ' シート名に例外DB名がカッコ内に含まれているか判定してオフセットを決定
                colOffset = 0
                If exceptionDbName <> "" Then
                    If InStr(sourceWs.Name, "（" & exceptionDbName & "）") > 0 Or _
                       InStr(sourceWs.Name, "(" & exceptionDbName & ")") > 0 Then
                        colOffset = 1
                    End If
                End If

                ' オフセットを適用した列名を計算
                actualColItemName = Chr(Asc(colItemName) + colOffset)
                actualColName = Chr(Asc(colName) + colOffset)
                actualColDataType = Chr(Asc(colDataType) + colOffset)
                actualColLength = Chr(Asc(colLength) + colOffset)
                actualColNullable = Chr(Asc(colNullable) + colOffset)

                ' テーブル名セルにもオフセットを適用
                actualTableNameCell = SG_Generator.OffsetCellColumn(tableNameCell, colOffset)
                actualTableDescCell = SG_Generator.OffsetCellColumn(tableDescCell, colOffset)

                ' テーブル名を取得
                tableName = Trim(CStr(sourceWs.Range(actualTableNameCell).Value))
                tableDesc = Trim(CStr(sourceWs.Range(actualTableDescCell).Value))

                If tableName <> "" Then
                    ' テーブル一覧に追加
                    wsDef.Range("A" & nextTableRow).Value = tableCount + 1
                    wsDef.Range("B" & nextTableRow).Value = tableName
                    wsDef.Range("C" & nextTableRow).Value = tableDesc
                    nextTableRow = nextTableRow + 1
                    tableCount = tableCount + 1

                    If importedTables <> "" Then importedTables = importedTables & ", "
                    importedTables = importedTables & tableName

                    ' カラム定義を取得
                    currentRow = columnStartRow
                    Do While True
                        ' カラム番号列の値を取得
                        colNumberValue = sourceWs.Range(colNumber & currentRow).Value

                        ' カラム番号が数値でない場合はスキップ
                        If Not IsNumeric(colNumberValue) Then
                            currentRow = currentRow + 1
                            ' 安全制限
                            If currentRow > 1000 Then Exit Do
                            ' 空行が続いたら終了（10行連続で空なら終了）
                            If currentRow > columnStartRow + 10 Then
                                emptyCount = 0
                                For checkRow = currentRow - 10 To currentRow - 1
                                    If Trim(CStr(sourceWs.Range(actualColName & checkRow).Value)) = "" Then
                                        emptyCount = emptyCount + 1
                                    End If
                                Next checkRow
                                If emptyCount >= 10 Then Exit Do
                            End If
                            GoTo NextRow
                        End If

                        colNameValue = Trim(CStr(sourceWs.Range(actualColName & currentRow).Value))

                        ' カラム名が空なら終了
                        If colNameValue = "" Then Exit Do

                        colItemNameValue = Trim(CStr(sourceWs.Range(actualColItemName & currentRow).Value))
                        colTypeValue = Trim(CStr(sourceWs.Range(actualColDataType & currentRow).Value))
                        colLengthValue = Trim(CStr(sourceWs.Range(actualColLength & currentRow).Value))
                        colNullableValue = Trim(CStr(sourceWs.Range(actualColNullable & currentRow).Value))

                        ' カラム一覧に追加
                        wsDef.Range("E" & nextColumnRow).Value = tableName
                        wsDef.Range("F" & nextColumnRow).Value = colNameValue
                        wsDef.Range("G" & nextColumnRow).Value = colTypeValue & IIf(colLengthValue <> "", "(" & colLengthValue & ")", "")
                        wsDef.Range("H" & nextColumnRow).Value = colItemNameValue

                        nextColumnRow = nextColumnRow + 1
                        columnCount = columnCount + 1

NextRow:
                        currentRow = currentRow + 1

                        ' 安全制限
                        If currentRow > 1000 Then Exit Do
                    Loop
                End If
            Next sheetIdx

            ' ファイルを閉じる
            sourceWb.Close SaveChanges:=False
        End If

        fileName = Dir()
    Loop

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    If tableCount = 0 Then
        MsgBox "インポートできるテーブル定義書が見つかりませんでした。" & vbCrLf & vbCrLf & _
               "確認事項:" & vbCrLf & _
               "・フォルダにExcelファイル(.xls/.xlsx/.xlsm)が存在するか" & vbCrLf & _
               "・テーブル名セル(" & tableNameCell & ")に値があるか", _
               vbExclamation, "インポート結果"
    Else
        MsgBox "テーブル定義のインポートが完了しました。" & vbCrLf & vbCrLf & _
               "インポートしたテーブル数: " & tableCount & vbCrLf & _
               "インポートしたカラム数: " & columnCount & vbCrLf & vbCrLf & _
               "テーブル: " & importedTables, _
               vbInformation, "インポート完了"

        ' プルダウンを自動更新
        UpdateDropdownsFromTableDef
    End If

    Exit Sub

ErrorHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    If Not sourceWb Is Nothing Then
        sourceWb.Close SaveChanges:=False
    End If

    MsgBox "エラーが発生しました: " & Err.Description & vbCrLf & _
           "ファイル: " & fileName, vbCritical, "エラー"
End Sub

