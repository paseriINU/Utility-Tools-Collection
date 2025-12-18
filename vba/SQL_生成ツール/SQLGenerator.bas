'==============================================================================
' Oracle SELECT文生成ツール - メインモジュール (ビジネスロジック)
' モジュール名: SQL_生成ツール
'==============================================================================
' 概要:
'   ExcelからOracle用のSELECT文を対話的に生成するツールです。
'   複雑なJOIN、サブクエリ、UNION、WITH句にも対応しています。
'
'   ※このモジュールはSQL_生成ツール_Setupモジュールと組み合わせて使用します。
'   ※定数と初期化関数はSQL_生成ツール_Setupモジュールに定義されています。
'
' 機能:
'   - SQL生成（SELECT, FROM, WHERE, GROUP BY, HAVING, ORDER BY, LIMIT, UNION, WITH句）
'   - テーブル定義連動機能
'   - テーブル定義書インポート機能
'   - 生成したSQLの履歴保存・クリップボードコピー
'
' 必要な環境:
'   - Microsoft Excel 2010以降
'
' 使い方:
'   1. このモジュールとSQL_生成ツール_Setup.basの両方をVBAエディタにインポート
'   2. InitializeSQL生成ツール マクロを実行してシートを初期化
'   3. 各シートに必要な情報を入力
'   4. GenerateSQL マクロを実行してSQLを生成
'
' 作成日: 2025-12-17（Setupモジュールとの分離）
'==============================================================================

Option Explicit

'==============================================================================
' 注意: 定数はSQL_生成ツール_Setupモジュールで定義されています
'==============================================================================
' 以下の定数はSQL_生成ツール_SetupモジュールでPublicとして定義されており、
' このモジュールから直接参照できます:
'   - SHEET_MAIN, SHEET_TABLE_DEF, SHEET_HISTORY, SHEET_SUBQUERY, SHEET_CTE, SHEET_UNION, SHEET_HELP
'   - ROW_TITLE, ROW_OPTIONS, ROW_MAIN_TABLE, ROW_JOIN_START, ROW_JOIN_END
'   - ROW_COLUMNS_LABEL, ROW_COLUMNS_START, ROW_COLUMNS_END
'   - ROW_WHERE_LABEL, ROW_WHERE_START, ROW_WHERE_END
'   - ROW_GROUPBY, ROW_HAVING_LABEL, ROW_HAVING_START, ROW_HAVING_END
'   - ROW_ORDERBY_LABEL, ROW_ORDERBY_START, ROW_ORDERBY_END
'   - ROW_LIMIT, ROW_SQL_OUTPUT
'   - DEFAULT_TABLE_NAME_CELL, DEFAULT_TABLE_DESC_CELL, DEFAULT_COLUMN_START_ROW
'   - DEFAULT_COL_NUMBER, DEFAULT_COL_ITEM_NAME, DEFAULT_COL_NAME
'   - DEFAULT_COL_DATATYPE, DEFAULT_COL_LENGTH, DEFAULT_COL_NULLABLE

'==============================================================================
' ユーティリティ関数: セルアドレスの列をオフセット分ずらす
' 例: "J2", 1 → "K2"  /  "D2", 1 → "E2"
'==============================================================================
Private Function OffsetCellColumn(ByVal cellAddr As String, ByVal colOffset As Long) As String
    Dim colPart As String
    Dim rowPart As String
    Dim i As Long
    Dim c As String

    ' オフセットが0なら元のアドレスをそのまま返す
    If colOffset = 0 Then
        OffsetCellColumn = cellAddr
        Exit Function
    End If

    ' 列部分と行部分を分離
    colPart = ""
    rowPart = ""
    For i = 1 To Len(cellAddr)
        c = Mid(cellAddr, i, 1)
        If c >= "A" And c <= "Z" Then
            colPart = colPart & c
        ElseIf c >= "a" And c <= "z" Then
            colPart = colPart & UCase(c)
        Else
            rowPart = Mid(cellAddr, i)
            Exit For
        End If
    Next i

    ' 列番号に変換してオフセットを適用
    Dim colNum As Long
    colNum = 0
    For i = 1 To Len(colPart)
        colNum = colNum * 26 + (Asc(Mid(colPart, i, 1)) - Asc("A") + 1)
    Next i
    colNum = colNum + colOffset

    ' 列番号を文字に変換
    Dim newColPart As String
    newColPart = ""
    Do While colNum > 0
        newColPart = Chr(((colNum - 1) Mod 26) + Asc("A")) & newColPart
        colNum = (colNum - 1) \ 26
    Loop

    OffsetCellColumn = newColPart & rowPart
End Function

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
        withClause = GenerateWithClause()
    End If

    ' SELECT句の生成
    selectClause = GenerateSelectClause(ws)
    If selectClause = "" Then
        MsgBox "取得カラムを1つ以上指定してください。", vbExclamation, "入力エラー"
        Exit Sub
    End If

    ' FROM句とJOIN句の生成
    fromClause = GenerateFromClause(ws)
    If fromClause = "" Then
        MsgBox "メインテーブルを指定してください。", vbExclamation, "入力エラー"
        Exit Sub
    End If

    ' WHERE句の生成
    whereClause = GenerateWhereClause(ws)

    ' GROUP BY句の生成
    groupByClause = GenerateGroupByClause(ws)

    ' HAVING句の生成
    havingClause = GenerateHavingClause(ws)

    ' ORDER BY句の生成
    orderByClause = GenerateOrderByClause(ws)

    ' 件数制限の生成
    limitClause = GenerateLimitClause(ws)

    ' UNION句の生成
    If ws.Range("H" & ROW_OPTIONS).Value = "使用する" Then
        unionClause = GenerateUnionClause()
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
' SELECT句の生成
'==============================================================================
Private Function GenerateSelectClause(ByVal ws As Worksheet) As String
    Dim result As String
    Dim columns As String
    Dim i As Long
    Dim tableAlias As String
    Dim columnName As String
    Dim colAlias As String
    Dim aggFunc As String
    Dim subqueryNo As String
    Dim colExpr As String
    Dim isDistinct As Boolean

    isDistinct = (ws.Range("B" & ROW_OPTIONS).Value = "DISTINCT")

    columns = ""

    For i = ROW_COLUMNS_START To ROW_COLUMNS_END
        tableAlias = Trim(ws.Range("B" & i).Value)
        columnName = ExtractTableName(Trim(ws.Range("C" & i).Value))
        colAlias = Trim(ws.Range("D" & i).Value)
        aggFunc = ExtractTableName(Trim(ws.Range("E" & i).Value))
        subqueryNo = Trim(ws.Range("F" & i).Value)

        If columnName <> "" Or subqueryNo <> "" Then
            ' サブクエリの場合
            If subqueryNo <> "" Then
                colExpr = GetSubquery(subqueryNo)
                If colExpr = "" Then
                    colExpr = "(サブクエリ" & subqueryNo & "が見つかりません)"
                Else
                    colExpr = "(" & vbCrLf & "    " & Replace(colExpr, vbCrLf, vbCrLf & "    ") & vbCrLf & ")"
                End If
            Else
                ' 通常のカラム
                If tableAlias <> "" Then
                    colExpr = tableAlias & "." & columnName
                Else
                    colExpr = columnName
                End If

                ' 集計関数
                If aggFunc <> "" Then
                    If aggFunc = "COUNT(DISTINCT)" Then
                        colExpr = "COUNT(DISTINCT " & colExpr & ")"
                    Else
                        colExpr = aggFunc & "(" & colExpr & ")"
                    End If
                End If
            End If

            ' 別名
            If colAlias <> "" Then
                colExpr = colExpr & " AS " & colAlias
            End If

            If columns <> "" Then
                columns = columns & "," & vbCrLf & "       "
            End If
            columns = columns & colExpr
        End If
    Next i

    If columns = "" Then
        GenerateSelectClause = ""
        Exit Function
    End If

    If isDistinct Then
        result = "SELECT DISTINCT " & columns
    Else
        result = "SELECT " & columns
    End If

    GenerateSelectClause = result
End Function

'==============================================================================
' FROM句とJOIN句の生成
'==============================================================================
Private Function GenerateFromClause(ByVal ws As Worksheet) As String
    Dim result As String
    Dim mainTable As String
    Dim mainAlias As String
    Dim i As Long
    Dim joinType As String
    Dim joinTable As String
    Dim joinAlias As String
    Dim joinCondition As String

    mainTable = ExtractTableName(Trim(ws.Range("B" & ROW_MAIN_TABLE + 1).Value))
    mainAlias = Trim(ws.Range("E" & ROW_MAIN_TABLE + 1).Value)

    If mainTable = "" Then
        GenerateFromClause = ""
        Exit Function
    End If

    result = "FROM " & mainTable
    If mainAlias <> "" Then
        result = result & " " & mainAlias
    End If

    ' JOIN句
    For i = ROW_JOIN_START + 2 To ROW_JOIN_END
        joinType = ExtractTableName(Trim(ws.Range("B" & i).Value))
        joinTable = ExtractTableName(Trim(ws.Range("C" & i).Value))
        joinAlias = Trim(ws.Range("D" & i).Value)
        joinCondition = Trim(ws.Range("E" & i).Value)

        If joinType <> "" And joinTable <> "" Then
            result = result & vbCrLf & joinType & " " & joinTable
            If joinAlias <> "" Then
                result = result & " " & joinAlias
            End If
            If joinCondition <> "" And InStr(joinType, "CROSS") = 0 Then
                result = result & " ON " & joinCondition
            End If
        End If
    Next i

    GenerateFromClause = result
End Function

'==============================================================================
' WHERE句の生成
'==============================================================================
Private Function GenerateWhereClause(ByVal ws As Worksheet) As String
    Dim result As String
    Dim i As Long
    Dim andOr As String
    Dim openParen As String
    Dim tableAlias As String
    Dim columnName As String
    Dim operator As String
    Dim value As String
    Dim closeParen As String
    Dim condition As String
    Dim isFirst As Boolean

    result = ""
    isFirst = True

    For i = ROW_WHERE_START To ROW_WHERE_END
        andOr = ExtractTableName(Trim(ws.Range("B" & i).Value))
        openParen = Trim(ws.Range("C" & i).Value)
        tableAlias = Trim(ws.Range("D" & i).Value)
        columnName = ExtractTableName(Trim(ws.Range("E" & i).Value))
        operator = ExtractTableName(Trim(ws.Range("F" & i).Value))
        value = Trim(ws.Range("G" & i).Value)
        closeParen = Trim(ws.Range("H" & i).Value)

        If columnName <> "" And operator <> "" Then
            ' カラム式を構築
            If tableAlias <> "" Then
                condition = tableAlias & "." & columnName
            Else
                condition = columnName
            End If

            ' 演算子と値を追加
            Select Case operator
                Case "IS NULL", "IS NOT NULL"
                    condition = condition & " " & operator
                Case "IN", "NOT IN"
                    ' サブクエリかリストか判定
                    If Left(value, 3) = "SUB" Then
                        Dim subSql As String
                        subSql = GetSubquery(value)
                        If subSql <> "" Then
                            condition = condition & " " & operator & " (" & vbCrLf & "    " & _
                                        Replace(subSql, vbCrLf, vbCrLf & "    ") & vbCrLf & ")"
                        Else
                            condition = condition & " " & operator & " (" & value & ")"
                        End If
                    Else
                        condition = condition & " " & operator & " (" & value & ")"
                    End If
                Case "EXISTS", "NOT EXISTS"
                    If Left(value, 3) = "SUB" Then
                        subSql = GetSubquery(value)
                        If subSql <> "" Then
                            condition = operator & " (" & vbCrLf & "    " & _
                                        Replace(subSql, vbCrLf, vbCrLf & "    ") & vbCrLf & ")"
                        Else
                            condition = operator & " (" & value & ")"
                        End If
                    Else
                        condition = operator & " (" & value & ")"
                    End If
                Case "BETWEEN"
                    condition = condition & " BETWEEN " & value
                Case "LIKE", "NOT LIKE"
                    condition = condition & " " & operator & " '" & value & "'"
                Case Else
                    ' 数値かどうか判定
                    If IsNumeric(value) Then
                        condition = condition & " " & operator & " " & value
                    ElseIf UCase(value) = "NULL" Or UCase(value) = "SYSDATE" Or _
                           Left(UCase(value), 7) = "SYSDATE" Or Left(value, 3) = "SUB" Then
                        ' サブクエリの場合
                        If Left(value, 3) = "SUB" Then
                            subSql = GetSubquery(value)
                            If subSql <> "" Then
                                condition = condition & " " & operator & " (" & vbCrLf & "    " & _
                                            Replace(subSql, vbCrLf, vbCrLf & "    ") & vbCrLf & ")"
                            Else
                                condition = condition & " " & operator & " " & value
                            End If
                        Else
                            condition = condition & " " & operator & " " & value
                        End If
                    Else
                        condition = condition & " " & operator & " '" & value & "'"
                    End If
            End Select

            ' 括弧を追加
            If openParen <> "" Then
                condition = openParen & condition
            End If
            If closeParen <> "" Then
                condition = condition & closeParen
            End If

            ' AND/OR を追加
            If isFirst Then
                result = "WHERE " & condition
                isFirst = False
            Else
                result = result & vbCrLf & "  " & andOr & " " & condition
            End If
        End If
    Next i

    GenerateWhereClause = result
End Function

'==============================================================================
' GROUP BY句の生成
'==============================================================================
Private Function GenerateGroupByClause(ByVal ws As Worksheet) As String
    Dim groupByColumns As String

    groupByColumns = Trim(ws.Range("B" & ROW_GROUPBY + 1).Value)

    If groupByColumns = "" Then
        GenerateGroupByClause = ""
    Else
        GenerateGroupByClause = "GROUP BY " & groupByColumns
    End If
End Function

'==============================================================================
' HAVING句の生成
'==============================================================================
Private Function GenerateHavingClause(ByVal ws As Worksheet) As String
    Dim result As String
    Dim i As Long
    Dim andOr As String
    Dim condition As String
    Dim isFirst As Boolean

    result = ""
    isFirst = True

    For i = ROW_HAVING_START To ROW_HAVING_END
        andOr = Trim(ws.Range("B" & i).Value)
        condition = Trim(ws.Range("C" & i).Value)

        If condition <> "" Then
            If isFirst Then
                result = "HAVING " & condition
                isFirst = False
            Else
                result = result & vbCrLf & "  " & andOr & " " & condition
            End If
        End If
    Next i

    GenerateHavingClause = result
End Function

'==============================================================================
' ORDER BY句の生成
'==============================================================================
Private Function GenerateOrderByClause(ByVal ws As Worksheet) As String
    Dim result As String
    Dim i As Long
    Dim tableAlias As String
    Dim columnName As String
    Dim sortOrder As String
    Dim nullsOrder As String
    Dim orderExpr As String
    Dim isFirst As Boolean

    result = ""
    isFirst = True

    For i = ROW_ORDERBY_START To ROW_ORDERBY_END
        tableAlias = Trim(ws.Range("B" & i).Value)
        columnName = ExtractTableName(Trim(ws.Range("C" & i).Value))
        sortOrder = ExtractTableName(Trim(ws.Range("D" & i).Value))
        nullsOrder = ExtractTableName(Trim(ws.Range("E" & i).Value))

        If columnName <> "" Then
            If tableAlias <> "" Then
                orderExpr = tableAlias & "." & columnName
            Else
                orderExpr = columnName
            End If

            If sortOrder <> "" Then
                orderExpr = orderExpr & " " & sortOrder
            End If

            If nullsOrder <> "" Then
                orderExpr = orderExpr & " " & nullsOrder
            End If

            If isFirst Then
                result = "ORDER BY " & orderExpr
                isFirst = False
            Else
                result = result & ", " & orderExpr
            End If
        End If
    Next i

    GenerateOrderByClause = result
End Function

'==============================================================================
' 件数制限句の生成
'==============================================================================
Private Function GenerateLimitClause(ByVal ws As Worksheet) As String
    Dim isEnabled As String
    Dim limitCount As String
    Dim limitType As String

    isEnabled = Trim(ws.Range("B" & ROW_LIMIT + 1).Value)

    If isEnabled <> "有効" Then
        GenerateLimitClause = ""
        Exit Function
    End If

    limitCount = Trim(ws.Range("D" & ROW_LIMIT + 1).Value)
    limitType = Trim(ws.Range("F" & ROW_LIMIT + 1).Value)

    If limitCount = "" Then limitCount = "100"
    If limitType = "" Then limitType = "FETCH FIRST"

    If limitType = "FETCH FIRST" Then
        GenerateLimitClause = "FETCH FIRST " & limitCount & " ROWS ONLY"
    Else
        ' ROWNUM方式の場合はWHERE句に追加する必要があるため、コメントで出力
        GenerateLimitClause = "-- ROWNUM <= " & limitCount & " (WHERE句に追加してください)"
    End If
End Function

'==============================================================================
' WITH句の生成
'==============================================================================
Private Function GenerateWithClause() As String
    Dim ws As Worksheet
    Dim result As String
    Dim i As Long
    Dim cteName As String
    Dim cteColumns As String
    Dim cteSql As String
    Dim isFirst As Boolean

    On Error Resume Next
    Set ws = Sheets(SHEET_CTE)
    If ws Is Nothing Then
        GenerateWithClause = ""
        Exit Function
    End If
    On Error GoTo 0

    result = ""
    isFirst = True

    For i = 5 To 9 ' CTE入力行
        cteName = Trim(ws.Range("A" & i).Value)
        cteColumns = Trim(ws.Range("B" & i).Value)
        cteSql = Trim(ws.Range("C" & i).Value)

        If cteName <> "" And cteSql <> "" Then
            If isFirst Then
                result = "WITH "
                isFirst = False
            Else
                result = result & "," & vbCrLf & "     "
            End If

            result = result & cteName
            If cteColumns <> "" Then
                result = result & " (" & cteColumns & ")"
            End If
            result = result & " AS (" & vbCrLf
            result = result & "    " & Replace(cteSql, vbCrLf, vbCrLf & "    ") & vbCrLf
            result = result & ")"
        End If
    Next i

    GenerateWithClause = result
End Function

'==============================================================================
' UNION句の生成
'==============================================================================
Private Function GenerateUnionClause() As String
    Dim ws As Worksheet
    Dim result As String
    Dim i As Long
    Dim unionType As String
    Dim unionSql As String

    On Error Resume Next
    Set ws = Sheets(SHEET_UNION)
    If ws Is Nothing Then
        GenerateUnionClause = ""
        Exit Function
    End If
    On Error GoTo 0

    result = ""

    For i = 5 To 9 ' UNION入力行
        unionType = Trim(ws.Range("B" & i).Value)
        unionSql = Trim(ws.Range("C" & i).Value)

        If unionType <> "" And unionSql <> "" Then
            result = result & vbCrLf & unionType & vbCrLf & unionSql
        End If
    Next i

    GenerateUnionClause = result
End Function

'==============================================================================
' サブクエリを取得
'==============================================================================
Private Function GetSubquery(ByVal subqueryNo As String) As String
    Dim ws As Worksheet
    Dim i As Long
    Dim cellNo As String
    Dim cellSql As String

    On Error Resume Next
    Set ws = Sheets(SHEET_SUBQUERY)
    If ws Is Nothing Then
        GetSubquery = ""
        Exit Function
    End If
    On Error GoTo 0

    For i = 5 To 14 ' サブクエリ入力行
        cellNo = Trim(ws.Range("A" & i).Value)
        cellSql = Trim(ws.Range("C" & i).Value)

        If cellNo = subqueryNo And cellSql <> "" Then
            GetSubquery = cellSql
            Exit Function
        End If
    Next i

    GetSubquery = ""
End Function

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
    tableList = GetTableList()

    If tableList = "" Then
        MsgBox "テーブル定義シートにテーブルが登録されていません。" & vbCrLf & _
               "「テーブル定義」シートのB列にテーブル名を登録してください。", vbExclamation, "確認"
        Exit Sub
    End If

    ' メインテーブルのプルダウンを更新（テーブル一覧用プレフィックス）
    AddDropdown wsMain, "B" & ROW_MAIN_TABLE + 1, tableList, "TableList"

    ' JOINテーブルのプルダウンを更新（テーブル一覧用プレフィックス）
    For i = ROW_JOIN_START + 2 To ROW_JOIN_END
        AddDropdown wsMain, "C" & i, tableList, "TableList"
    Next i

    ' カラム選択のプルダウンを更新（全テーブルの全カラム、カラム一覧用プレフィックス）
    Dim columnList As String
    columnList = GetAllColumnList()

    For i = ROW_COLUMNS_START To ROW_COLUMNS_END
        If columnList <> "" Then
            AddDropdown wsMain, "C" & i, columnList, "ColumnList"
        End If
    Next i

    ' WHERE句のカラムプルダウンを更新（カラム一覧用プレフィックス）
    For i = ROW_WHERE_START To ROW_WHERE_END
        If columnList <> "" Then
            AddDropdown wsMain, "E" & i, columnList, "ColumnList"
        End If
    Next i

    ' ORDER BY句のカラムプルダウンを更新（カラム一覧用プレフィックス）
    For i = ROW_ORDERBY_START To ROW_ORDERBY_END
        If columnList <> "" Then
            AddDropdown wsMain, "C" & i, columnList, "ColumnList"
        End If
    Next i

    ' テーブル別名のプルダウンを更新（入力済みの別名から取得）
    Dim aliasList As String
    aliasList = GetAliasListFromMain()

    If aliasList <> "" Then
        For i = ROW_COLUMNS_START To ROW_COLUMNS_END
            AddDropdown wsMain, "B" & i, aliasList
        Next i
        For i = ROW_WHERE_START To ROW_WHERE_END
            AddDropdown wsMain, "D" & i, aliasList
        Next i
        For i = ROW_ORDERBY_START To ROW_ORDERBY_END
            AddDropdown wsMain, "B" & i, aliasList
        Next i
    End If

    MsgBox "プルダウンを更新しました。" & vbCrLf & vbCrLf & _
           "テーブル数: " & UBound(Split(tableList, ",")) + 1, vbInformation, "更新完了"

    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub

'==============================================================================
' テーブル定義シートからテーブル一覧を取得
' 形式: テーブル名(テーブル名称)
'==============================================================================
Private Function GetTableList() As String
    Dim ws As Worksheet
    Dim result As String
    Dim i As Long
    Dim lastRow As Long
    Dim tableName As String
    Dim tableDesc As String
    Dim displayName As String

    On Error Resume Next
    Set ws = Sheets(SHEET_TABLE_DEF)
    If ws Is Nothing Then
        GetTableList = ""
        Exit Function
    End If
    On Error GoTo 0

    result = ""

    ' B列の最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    ' B列（テーブル名）、C列（テーブル名称）を読み取り（6行目から開始）
    i = 6
    Do While ws.Range("B" & i).Value <> ""
        tableName = Trim(ws.Range("B" & i).Value)
        tableDesc = Trim(ws.Range("C" & i).Value)
        If tableName <> "" Then
            ' テーブル名称がある場合はカッコで追加
            If tableDesc <> "" Then
                displayName = tableName & "(" & tableDesc & ")"
            Else
                displayName = tableName
            End If
            If result = "" Then
                result = displayName
            Else
                result = result & "," & displayName
            End If
        End If
        i = i + 1
        If i > lastRow Then Exit Do ' 最終行を超えたら終了
    Loop

    GetTableList = result
End Function

'==============================================================================
' 文字列からカッコ部分を除去（テーブル名、JOINタイプ等で使用）
' 例: "USER_MASTER(ユーザー)" → "USER_MASTER"
'     "INNER JOIN(両方に存在)" → "INNER JOIN"
'==============================================================================
Private Function ExtractTableName(ByVal displayName As String) As String
    Dim pos As Long
    pos = InStr(displayName, "(")
    If pos > 0 Then
        ExtractTableName = Left(displayName, pos - 1)
    Else
        ExtractTableName = displayName
    End If
End Function

'==============================================================================
' テーブル定義シートから全カラム一覧を取得
' 形式: カラム名(項目名)
'==============================================================================
Private Function GetAllColumnList() As String
    Dim ws As Worksheet
    Dim result As String
    Dim i As Long
    Dim lastRowE As Long
    Dim lastRowF As Long
    Dim lastRow As Long
    Dim tableName As String
    Dim columnName As String
    Dim itemName As String
    Dim displayName As String
    Dim dict As Object

    On Error Resume Next
    Set ws = Sheets(SHEET_TABLE_DEF)
    If ws Is Nothing Then
        GetAllColumnList = ""
        Exit Function
    End If
    On Error GoTo 0

    Set dict = CreateObject("Scripting.Dictionary")
    result = ""

    ' E列とF列の最終行を取得し、大きい方を使用
    lastRowE = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    lastRowF = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
    If lastRowE > lastRowF Then
        lastRow = lastRowE
    Else
        lastRow = lastRowF
    End If

    ' E列（テーブル名）、F列（カラム名）、H列（項目名）を読み取り（6行目から開始）
    i = 6
    Do While ws.Range("E" & i).Value <> "" Or ws.Range("F" & i).Value <> ""
        tableName = Trim(ws.Range("E" & i).Value)
        columnName = Trim(ws.Range("F" & i).Value)
        itemName = Trim(ws.Range("H" & i).Value)

        If columnName <> "" Then
            ' 重複チェック
            If Not dict.exists(columnName) Then
                dict.Add columnName, True
                ' 項目名がある場合はカッコで追加
                If itemName <> "" Then
                    displayName = columnName & "(" & itemName & ")"
                Else
                    displayName = columnName
                End If
                If result = "" Then
                    result = displayName
                Else
                    result = result & "," & displayName
                End If
            End If
        End If
        i = i + 1
        If i > lastRow Then Exit Do ' 最終行を超えたら終了
    Loop

    ' 特殊なカラム名を追加
    result = "*," & result

    GetAllColumnList = result
End Function

'==============================================================================
' 指定テーブルのカラム一覧を取得（項目名付き）
'==============================================================================
Private Function GetColumnListForTable(ByVal targetTable As String) As String
    Dim ws As Worksheet
    Dim result As String
    Dim i As Long
    Dim lastRowE As Long
    Dim lastRowF As Long
    Dim lastRow As Long
    Dim tableName As String
    Dim columnName As String
    Dim itemName As String
    Dim displayName As String

    On Error Resume Next
    Set ws = Sheets(SHEET_TABLE_DEF)
    If ws Is Nothing Then
        GetColumnListForTable = ""
        Exit Function
    End If
    On Error GoTo 0

    result = ""

    ' E列とF列の最終行を取得し、大きい方を使用
    lastRowE = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    lastRowF = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
    If lastRowE > lastRowF Then
        lastRow = lastRowE
    Else
        lastRow = lastRowF
    End If

    ' E列（テーブル名）、F列（カラム名）、H列（項目名）を読み取り（6行目から開始）
    i = 6
    Do While ws.Range("E" & i).Value <> "" Or ws.Range("F" & i).Value <> ""
        tableName = Trim(ws.Range("E" & i).Value)
        columnName = Trim(ws.Range("F" & i).Value)
        itemName = Trim(ws.Range("H" & i).Value)

        If UCase(tableName) = UCase(targetTable) And columnName <> "" Then
            ' 項目名がある場合はカッコで追加
            If itemName <> "" Then
                displayName = columnName & "(" & itemName & ")"
            Else
                displayName = columnName
            End If
            If result = "" Then
                result = displayName
            Else
                result = result & "," & displayName
            End If
        End If
        i = i + 1
        If i > lastRow Then Exit Do ' 最終行を超えたら終了
    Loop

    GetColumnListForTable = result
End Function

'==============================================================================
' 選択されたテーブル（メイン＋JOIN）のカラム一覧を取得
'==============================================================================
Private Function GetColumnListForSelectedTables() As String
    Dim wsMain As Worksheet
    Dim result As String
    Dim tableName As String
    Dim columnList As String
    Dim dict As Object
    Dim i As Long
    Dim cols() As String
    Dim c As Long

    Set wsMain = Sheets(SHEET_MAIN)
    Set dict = CreateObject("Scripting.Dictionary")
    result = ""

    ' メインテーブルのカラムを取得
    tableName = ExtractTableName(Trim(wsMain.Range("B" & ROW_MAIN_TABLE + 1).Value))
    If tableName <> "" Then
        columnList = GetColumnListForTable(tableName)
        If columnList <> "" Then
            cols = Split(columnList, ",")
            For c = LBound(cols) To UBound(cols)
                If Not dict.exists(cols(c)) Then
                    dict.Add cols(c), True
                    If result = "" Then
                        result = cols(c)
                    Else
                        result = result & "," & cols(c)
                    End If
                End If
            Next c
        End If
    End If

    ' JOINテーブルのカラムを取得
    For i = ROW_JOIN_START + 2 To ROW_JOIN_END
        tableName = ExtractTableName(Trim(wsMain.Range("C" & i).Value))
        If tableName <> "" Then
            columnList = GetColumnListForTable(tableName)
            If columnList <> "" Then
                cols = Split(columnList, ",")
                For c = LBound(cols) To UBound(cols)
                    If Not dict.exists(cols(c)) Then
                        dict.Add cols(c), True
                        If result = "" Then
                            result = cols(c)
                        Else
                            result = result & "," & cols(c)
                        End If
                    End If
                Next c
            End If
        End If
    Next i

    ' 特殊なカラム名を先頭に追加
    If result <> "" Then
        result = "*," & result
    End If

    GetColumnListForSelectedTables = result
End Function

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
    columnList = GetColumnListForSelectedTables()

    If columnList = "" Or columnList = "*" Then
        MsgBox "テーブルが選択されていません。" & vbCrLf & _
               "メインテーブルまたはJOINテーブルを選択してから実行してください。", vbExclamation, "確認"
        Exit Sub
    End If

    ' カラム選択のプルダウンを更新（カラム一覧用プレフィックス）
    For i = ROW_COLUMNS_START To ROW_COLUMNS_END
        AddDropdown wsMain, "C" & i, columnList, "ColumnList"
    Next i

    ' WHERE句のカラムプルダウンを更新（カラム一覧用プレフィックス）
    For i = ROW_WHERE_START To ROW_WHERE_END
        AddDropdown wsMain, "E" & i, columnList, "ColumnList"
    Next i

    ' ORDER BY句のカラムプルダウンを更新（カラム一覧用プレフィックス）
    For i = ROW_ORDERBY_START To ROW_ORDERBY_END
        AddDropdown wsMain, "C" & i, columnList, "ColumnList"
    Next i

    ' 選択されたテーブル数をカウント
    tableCount = 0
    tableName = ExtractTableName(Trim(wsMain.Range("B" & ROW_MAIN_TABLE + 1).Value))
    If tableName <> "" Then tableCount = tableCount + 1

    For i = ROW_JOIN_START + 2 To ROW_JOIN_END
        tableName = ExtractTableName(Trim(wsMain.Range("C" & i).Value))
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
    columnList = GetColumnListForSelectedTables()

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
        AddDropdown wsMain, "C" & i, columnList, "ColumnList"
    Next i

    ' WHERE句のカラムプルダウンを更新（カラム一覧用プレフィックス）
    For i = ROW_WHERE_START To ROW_WHERE_END
        AddDropdown wsMain, "E" & i, columnList, "ColumnList"
    Next i

    ' ORDER BY句のカラムプルダウンを更新（カラム一覧用プレフィックス）
    For i = ROW_ORDERBY_START To ROW_ORDERBY_END
        AddDropdown wsMain, "C" & i, columnList, "ColumnList"
    Next i

    Application.EnableEvents = True
End Sub

'==============================================================================
' メインシートから入力済みの別名一覧を取得
'==============================================================================
Private Function GetAliasListFromMain() As String
    Dim ws As Worksheet
    Dim result As String
    Dim alias As String
    Dim dict As Object
    Dim i As Long

    Set ws = Sheets(SHEET_MAIN)
    Set dict = CreateObject("Scripting.Dictionary")
    result = ""

    ' メインテーブルの別名
    alias = Trim(ws.Range("E" & ROW_MAIN_TABLE + 1).Value)
    If alias <> "" And Not dict.exists(alias) Then
        dict.Add alias, True
        result = alias
    End If

    ' JOINテーブルの別名
    For i = ROW_JOIN_START + 2 To ROW_JOIN_END
        alias = Trim(ws.Range("D" & i).Value)
        If alias <> "" And Not dict.exists(alias) Then
            dict.Add alias, True
            If result = "" Then
                result = alias
            Else
                result = result & "," & alias
            End If
        End If
    Next i

    ' 空白を先頭に追加（未選択用）
    If result <> "" Then
        result = "," & result
    End If

    GetAliasListFromMain = result
End Function

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
    aliasList = GetAliasListFromMain()

    If aliasList = "" Or aliasList = "," Then
        MsgBox "テーブル別名が入力されていません。" & vbCrLf & _
               "メインテーブルやJOINテーブルに別名を入力してから実行してください。", vbExclamation, "確認"
        Exit Sub
    End If

    ' 各セクションの「テーブル別名」プルダウンを更新
    For i = ROW_COLUMNS_START To ROW_COLUMNS_END
        AddDropdown wsMain, "B" & i, aliasList
    Next i

    For i = ROW_WHERE_START To ROW_WHERE_END
        AddDropdown wsMain, "D" & i, aliasList
    Next i

    For i = ROW_ORDERBY_START To ROW_ORDERBY_END
        AddDropdown wsMain, "B" & i, aliasList
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
    InitializeSQL生成ツール
End Sub

'==============================================================================
' テーブル定義書インポート機能
'==============================================================================
' 外部のテーブル定義書（Excelファイル）からテーブル・カラム情報をインポート
'
' 定義書フォーマット（デフォルト設定）:
'   - テーブル名: E4セル
'   - カラム定義: A10行から開始
'     - A列: カラム番号
'     - B列: カラム名
'     - C列: 項目名
'     - D列: カラム名
'     - E列: データ型
'     - F列: 桁数
'     - H列: NULL許可
'
' フォルダパス対応:
'   - テーブル定義シートのK15にフォルダパスを設定可能
'   - %USERNAME%を環境変数として展開
'   - 1ファイル内の全シートを読み込み
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
    tableNameCell = GetImportSetting(wsDef, "テーブル名セル", DEFAULT_TABLE_NAME_CELL)
    tableDescCell = GetImportSetting(wsDef, "テーブル名称セル", DEFAULT_TABLE_DESC_CELL)
    columnStartRow = CLng(GetImportSetting(wsDef, "カラム開始行", CStr(DEFAULT_COLUMN_START_ROW)))
    colNumber = GetImportSetting(wsDef, "カラム番号列", DEFAULT_COL_NUMBER)
    colItemName = GetImportSetting(wsDef, "項目名列", DEFAULT_COL_ITEM_NAME)
    colName = GetImportSetting(wsDef, "カラム名列", DEFAULT_COL_NAME)
    colDataType = GetImportSetting(wsDef, "データ型列", DEFAULT_COL_DATATYPE)
    colLength = GetImportSetting(wsDef, "桁数列", DEFAULT_COL_LENGTH)
    colNullable = GetImportSetting(wsDef, "NULL列", DEFAULT_COL_NULLABLE)

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
                ' 例: シート名「テーブル定義（AAA）」、例外DB名「AAA」→ 一致
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
                actualTableNameCell = OffsetCellColumn(tableNameCell, colOffset)
                actualTableDescCell = OffsetCellColumn(tableDescCell, colOffset)

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

'==============================================================================
' インポート設定を取得
'==============================================================================
Private Function GetImportSetting(ByVal ws As Worksheet, ByVal settingName As String, ByVal defaultValue As String) As String
    Dim i As Long
    Dim settingRow As Long

    ' 設定エリアを検索（J列〜K列）
    settingRow = 0
    For i = 4 To 20
        If Trim(CStr(ws.Range("J" & i).Value)) = settingName Then
            settingRow = i
            Exit For
        End If
    Next i

    If settingRow > 0 Then
        Dim val As String
        val = Trim(CStr(ws.Range("K" & settingRow).Value))
        If val <> "" Then
            GetImportSetting = val
        Else
            GetImportSetting = defaultValue
        End If
    Else
        GetImportSetting = defaultValue
    End If
End Function

'==============================================================================
' ソースシートからテーブル説明を取得（E4の下や近辺から推測）
'==============================================================================
Private Function GetTableDescription(ByVal ws As Worksheet) As String
    Dim desc As String

    ' F4に説明があることを想定
    desc = Trim(CStr(ws.Range("F4").Value))

    ' なければE5を確認
    If desc = "" Then
        desc = Trim(CStr(ws.Range("E5").Value))
    End If

    ' それでもなければ空を返す
    GetTableDescription = desc
End Function

