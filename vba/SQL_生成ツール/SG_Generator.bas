Attribute VB_Name = "SG_Generator"
Option Explicit

'==============================================================================
' Oracle SELECT文生成ツール - SQL生成モジュール
' SELECT文の各句を生成するロジックを提供
'==============================================================================

'==============================================================================
' ユーティリティ関数: セルアドレスの列をオフセット分ずらす
' 例: "J2", 1 → "K2"  /  "D2", 1 → "E2"
'==============================================================================
Public Function OffsetCellColumn(ByVal cellAddr As String, ByVal colOffset As Long) As String
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
' SELECT句の生成
'==============================================================================
Public Function GenerateSelectClause(ByVal ws As Worksheet) As String
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
Public Function GenerateFromClause(ByVal ws As Worksheet) As String
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
Public Function GenerateWhereClause(ByVal ws As Worksheet) As String
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
    Dim subSql As String

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
Public Function GenerateGroupByClause(ByVal ws As Worksheet) As String
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
Public Function GenerateHavingClause(ByVal ws As Worksheet) As String
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
Public Function GenerateOrderByClause(ByVal ws As Worksheet) As String
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
Public Function GenerateLimitClause(ByVal ws As Worksheet) As String
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
Public Function GenerateWithClause() As String
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
Public Function GenerateUnionClause() As String
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
Public Function GetSubquery(ByVal subqueryNo As String) As String
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
' 文字列からカッコ部分を除去（テーブル名、JOINタイプ等で使用）
' 例: "USER_MASTER(ユーザー)" → "USER_MASTER"
'     "INNER JOIN(両方に存在)" → "INNER JOIN"
'==============================================================================
Public Function ExtractTableName(ByVal displayName As String) As String
    Dim pos As Long
    pos = InStr(displayName, "(")
    If pos > 0 Then
        ExtractTableName = Left(displayName, pos - 1)
    Else
        ExtractTableName = displayName
    End If
End Function

'==============================================================================
' テーブル定義シートからテーブル一覧を取得
' 形式: テーブル名(テーブル名称)
'==============================================================================
Public Function GetTableList() As String
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
' テーブル定義シートから全カラム一覧を取得
' 形式: カラム名(項目名)
'==============================================================================
Public Function GetAllColumnList() As String
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
Public Function GetColumnListForTable(ByVal targetTable As String) As String
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
Public Function GetColumnListForSelectedTables() As String
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
' メインシートから入力済みの別名一覧を取得
'==============================================================================
Public Function GetAliasListFromMain() As String
    Dim ws As Worksheet
    Dim result As String
    Dim aliasName As String
    Dim dict As Object
    Dim i As Long

    Set ws = Sheets(SHEET_MAIN)
    Set dict = CreateObject("Scripting.Dictionary")
    result = ""

    ' メインテーブルの別名
    aliasName = Trim(ws.Range("E" & ROW_MAIN_TABLE + 1).Value)
    If aliasName <> "" And Not dict.exists(aliasName) Then
        dict.Add aliasName, True
        result = aliasName
    End If

    ' JOINテーブルの別名
    For i = ROW_JOIN_START + 2 To ROW_JOIN_END
        aliasName = Trim(ws.Range("D" & i).Value)
        If aliasName <> "" And Not dict.exists(aliasName) Then
            dict.Add aliasName, True
            If result = "" Then
                result = aliasName
            Else
                result = result & "," & aliasName
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
' インポート設定を取得
'==============================================================================
Public Function GetImportSetting(ByVal ws As Worksheet, ByVal settingName As String, ByVal defaultValue As String) As String
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
Public Function GetTableDescription(ByVal ws As Worksheet) As String
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

'==============================================================================
' プルダウンリスト追加ヘルパー
' ※Setupモジュール削除後も動作するようSG_Generatorに配置
'==============================================================================
Public Sub AddDropdown(ByVal ws As Worksheet, ByVal cellAddr As String, ByVal listItems As String, Optional ByVal namePrefix As String = "DropList")
    On Error Resume Next

    With ws.Range(cellAddr).Validation
        .Delete
        If Len(listItems) > 0 Then
            ' 直接設定（255文字以下の場合）
            If Len(listItems) <= 255 Then
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=listItems
            Else
                ' 255文字を超える場合は先頭255文字で切り詰め（名前付き範囲は使用しない）
                ' ※名前付き範囲を使うと複雑になるため、シンプルに切り詰める
                Dim truncatedList As String
                Dim items() As String
                Dim i As Long

                items = Split(listItems, ",")
                truncatedList = ""

                For i = LBound(items) To UBound(items)
                    If truncatedList = "" Then
                        truncatedList = Trim(items(i))
                    ElseIf Len(truncatedList) + Len(items(i)) + 1 <= 255 Then
                        truncatedList = truncatedList & "," & Trim(items(i))
                    Else
                        Exit For
                    End If
                Next i

                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=truncatedList
            End If
            .IgnoreBlank = True
            .InCellDropdown = True
        End If
    End With

    On Error GoTo 0
End Sub

