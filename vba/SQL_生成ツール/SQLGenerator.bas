'==============================================================================
' Oracle SELECT文生成ツール
' モジュール名: SQLGenerator
'==============================================================================
' 概要:
'   ExcelからOracle用のSELECT文を対話的に生成するツールです。
'   複雑なJOIN、サブクエリ、UNION、WITH句にも対応しています。
'
' 機能:
'   - 基本的なSELECT文生成
'   - JOIN（INNER, LEFT, RIGHT, FULL OUTER, CROSS）
'   - WHERE条件（各種演算子対応）
'   - GROUP BY / HAVING
'   - ORDER BY
'   - DISTINCT
'   - 件数制限（ROWNUM / FETCH FIRST）
'   - サブクエリ（SELECT句、WHERE句）
'   - UNION / UNION ALL
'   - WITH句（共通テーブル式）
'
' 必要な環境:
'   - Microsoft Excel 2010以降
'
' 使い方:
'   1. このモジュールをExcelのVBAエディタにインポート
'   2. InitializeSQLGenerator マクロを実行してシートを初期化
'   3. 各シートに必要な情報を入力
'   4. GenerateSQL マクロを実行してSQLを生成
'
' 作成日: 2025-12-12
'==============================================================================

Option Explicit

'==============================================================================
' 定数定義
'==============================================================================
Private Const SHEET_MAIN As String = "メイン"
Private Const SHEET_TABLE_DEF As String = "テーブル定義"
Private Const SHEET_HISTORY As String = "生成履歴"
Private Const SHEET_SUBQUERY As String = "サブクエリ"
Private Const SHEET_CTE As String = "WITH句"
Private Const SHEET_UNION As String = "UNION"
Private Const SHEET_HELP As String = "SQLヘルプ"

' メインシートの行位置
Private Const ROW_TITLE As Long = 1
Private Const ROW_OPTIONS As Long = 3
Private Const ROW_MAIN_TABLE As Long = 6
Private Const ROW_JOIN_START As Long = 9
Private Const ROW_JOIN_END As Long = 18
Private Const ROW_COLUMNS_LABEL As Long = 20
Private Const ROW_COLUMNS_START As Long = 22
Private Const ROW_COLUMNS_END As Long = 41
Private Const ROW_WHERE_LABEL As Long = 43
Private Const ROW_WHERE_START As Long = 45
Private Const ROW_WHERE_END As Long = 59
Private Const ROW_GROUPBY As Long = 61
Private Const ROW_HAVING_LABEL As Long = 63
Private Const ROW_HAVING_START As Long = 65
Private Const ROW_HAVING_END As Long = 69
Private Const ROW_ORDERBY_LABEL As Long = 71
Private Const ROW_ORDERBY_START As Long = 73
Private Const ROW_ORDERBY_END As Long = 82
Private Const ROW_LIMIT As Long = 84
Private Const ROW_SQL_OUTPUT As Long = 88

'==============================================================================
' 初期化: シートを作成してフォーマットを設定
'==============================================================================
Public Sub InitializeSQLGenerator()
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False

    ' シートを作成
    CreateSheet SHEET_MAIN
    CreateSheet SHEET_TABLE_DEF
    CreateSheet SHEET_HISTORY
    CreateSheet SHEET_SUBQUERY
    CreateSheet SHEET_CTE
    CreateSheet SHEET_UNION
    CreateSheet SHEET_HELP

    ' 各シートをフォーマット
    FormatMainSheet
    FormatTableDefSheet
    FormatHistorySheet
    FormatSubquerySheet
    FormatCTESheet
    FormatUnionSheet
    FormatHelpSheet

    ' メインシートをアクティブに
    Sheets(SHEET_MAIN).Activate

    Application.ScreenUpdating = True

    MsgBox "SQL生成ツールの初期化が完了しました。" & vbCrLf & vbCrLf & _
           "【使い方】" & vbCrLf & _
           "1. 「テーブル定義」シートにテーブル・カラム情報を登録" & vbCrLf & _
           "2. 「メイン」シートで条件を入力" & vbCrLf & _
           "3. 「SQL生成」ボタンをクリック", vbInformation, "初期化完了"

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub

'==============================================================================
' シート作成ヘルパー
'==============================================================================
Private Sub CreateSheet(ByVal sheetName As String)
    Dim ws As Worksheet
    Dim exists As Boolean

    exists = False
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    If Not ws Is Nothing Then exists = True
    Err.Clear
    On Error GoTo 0

    If exists Then
        ws.Cells.Clear
        ws.Cells.Interior.ColorIndex = xlNone
    Else
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = sheetName
    End If
End Sub

'==============================================================================
' メインシートのフォーマット
'==============================================================================
Private Sub FormatMainSheet()
    Dim ws As Worksheet
    Set ws = Sheets(SHEET_MAIN)

    With ws
        ' タイトル
        .Range("A" & ROW_TITLE).Value = "Oracle SELECT文生成ツール"
        .Range("A" & ROW_TITLE).Font.Size = 18
        .Range("A" & ROW_TITLE).Font.Bold = True
        .Range("A" & ROW_TITLE & ":J" & ROW_TITLE).Merge
        .Range("A" & ROW_TITLE).Interior.Color = RGB(68, 114, 196)
        .Range("A" & ROW_TITLE).Font.Color = RGB(255, 255, 255)

        ' オプション行
        .Range("A" & ROW_OPTIONS).Value = "DISTINCT:"
        .Range("B" & ROW_OPTIONS).Value = ""
        AddDropdown ws, "B" & ROW_OPTIONS, "なし,DISTINCT"

        .Range("D" & ROW_OPTIONS).Value = "WITH句使用:"
        .Range("E" & ROW_OPTIONS).Value = ""
        AddDropdown ws, "E" & ROW_OPTIONS, "なし,使用する"

        .Range("G" & ROW_OPTIONS).Value = "UNION使用:"
        .Range("H" & ROW_OPTIONS).Value = ""
        AddDropdown ws, "H" & ROW_OPTIONS, "なし,使用する"

        ' メインテーブル
        .Range("A" & ROW_MAIN_TABLE).Value = "メインテーブル (FROM句)"
        .Range("A" & ROW_MAIN_TABLE).Font.Bold = True
        .Range("A" & ROW_MAIN_TABLE).Font.Size = 12
        .Range("A" & ROW_MAIN_TABLE).Interior.Color = RGB(221, 235, 247)
        .Range("A" & ROW_MAIN_TABLE & ":J" & ROW_MAIN_TABLE).Merge

        ' メインテーブルの説明
        .Range("A" & ROW_MAIN_TABLE + 1).Value = "テーブル名:"
        .Range("B" & ROW_MAIN_TABLE + 1).Value = ""
        .Range("D" & ROW_MAIN_TABLE + 1).Value = "別名:"
        .Range("E" & ROW_MAIN_TABLE + 1).Value = ""
        .Range("G" & ROW_MAIN_TABLE + 1).Value = "※データを取得する主となるテーブル。別名を付けると短く参照できます"
        .Range("G" & ROW_MAIN_TABLE + 1).Font.Color = RGB(128, 128, 128)
        .Range("G" & ROW_MAIN_TABLE + 1).Font.Size = 9

        ' 結合テーブル（JOIN）
        .Range("A" & ROW_JOIN_START).Value = "結合テーブル (JOIN) - 複数テーブルを連結してデータを取得"
        .Range("A" & ROW_JOIN_START).Font.Bold = True
        .Range("A" & ROW_JOIN_START).Font.Size = 12
        .Range("A" & ROW_JOIN_START).Interior.Color = RGB(221, 235, 247)
        .Range("A" & ROW_JOIN_START & ":J" & ROW_JOIN_START).Merge

        ' JOINヘッダー
        .Range("A" & ROW_JOIN_START + 1).Value = "No"
        .Range("B" & ROW_JOIN_START + 1).Value = "結合種別"
        .Range("C" & ROW_JOIN_START + 1).Value = "テーブル名"
        .Range("D" & ROW_JOIN_START + 1).Value = "別名"
        .Range("E" & ROW_JOIN_START + 1).Value = "結合条件 (ON句)"
        .Range("F" & ROW_JOIN_START + 1).Value = "説明"
        .Range("A" & ROW_JOIN_START + 1 & ":E" & ROW_JOIN_START + 1).Font.Bold = True
        .Range("A" & ROW_JOIN_START + 1 & ":E" & ROW_JOIN_START + 1).Interior.Color = RGB(180, 198, 231)
        .Range("F" & ROW_JOIN_START + 1).Font.Color = RGB(128, 128, 128)
        .Range("F" & ROW_JOIN_START + 1).Font.Size = 9
        .Range("F" & ROW_JOIN_START + 1).Value = "※INNER=両方に存在, LEFT=左テーブル全部, RIGHT=右テーブル全部, FULL=両方全部"

        ' JOIN行
        Dim i As Long
        For i = 1 To 8
            .Range("A" & ROW_JOIN_START + 1 + i).Value = i
            AddDropdown ws, "B" & ROW_JOIN_START + 1 + i, ",INNER JOIN,LEFT JOIN,RIGHT JOIN,FULL OUTER JOIN,CROSS JOIN"
        Next i

        ' 取得カラム
        .Range("A" & ROW_COLUMNS_LABEL).Value = "取得カラム (SELECT句) - 表示したい列を指定"
        .Range("A" & ROW_COLUMNS_LABEL).Font.Bold = True
        .Range("A" & ROW_COLUMNS_LABEL).Font.Size = 12
        .Range("A" & ROW_COLUMNS_LABEL).Interior.Color = RGB(221, 235, 247)
        .Range("A" & ROW_COLUMNS_LABEL & ":J" & ROW_COLUMNS_LABEL).Merge

        ' カラムヘッダー
        .Range("A" & ROW_COLUMNS_LABEL + 1).Value = "No"
        .Range("B" & ROW_COLUMNS_LABEL + 1).Value = "テーブル別名"
        .Range("C" & ROW_COLUMNS_LABEL + 1).Value = "カラム名/式"
        .Range("D" & ROW_COLUMNS_LABEL + 1).Value = "別名 (AS)"
        .Range("E" & ROW_COLUMNS_LABEL + 1).Value = "集計関数"
        .Range("F" & ROW_COLUMNS_LABEL + 1).Value = "サブクエリNo"
        .Range("G" & ROW_COLUMNS_LABEL + 1).Value = "※集計関数: COUNT=件数, SUM=合計, AVG=平均, MAX=最大, MIN=最小"
        .Range("A" & ROW_COLUMNS_LABEL + 1 & ":F" & ROW_COLUMNS_LABEL + 1).Font.Bold = True
        .Range("A" & ROW_COLUMNS_LABEL + 1 & ":F" & ROW_COLUMNS_LABEL + 1).Interior.Color = RGB(180, 198, 231)
        .Range("G" & ROW_COLUMNS_LABEL + 1).Font.Color = RGB(128, 128, 128)
        .Range("G" & ROW_COLUMNS_LABEL + 1).Font.Size = 9

        ' カラム行
        For i = 1 To 20
            .Range("A" & ROW_COLUMNS_START + i - 1).Value = i
            AddDropdown ws, "E" & ROW_COLUMNS_START + i - 1, ",COUNT,SUM,AVG,MAX,MIN,COUNT(DISTINCT)"
        Next i

        ' WHERE条件
        .Range("A" & ROW_WHERE_LABEL).Value = "抽出条件 (WHERE句) - データを絞り込む条件を指定"
        .Range("A" & ROW_WHERE_LABEL).Font.Bold = True
        .Range("A" & ROW_WHERE_LABEL).Font.Size = 12
        .Range("A" & ROW_WHERE_LABEL).Interior.Color = RGB(221, 235, 247)
        .Range("A" & ROW_WHERE_LABEL & ":J" & ROW_WHERE_LABEL).Merge

        ' WHEREヘッダー
        .Range("A" & ROW_WHERE_LABEL + 1).Value = "No"
        .Range("B" & ROW_WHERE_LABEL + 1).Value = "AND/OR"
        .Range("C" & ROW_WHERE_LABEL + 1).Value = "("
        .Range("D" & ROW_WHERE_LABEL + 1).Value = "テーブル別名"
        .Range("E" & ROW_WHERE_LABEL + 1).Value = "カラム名/式"
        .Range("F" & ROW_WHERE_LABEL + 1).Value = "演算子"
        .Range("G" & ROW_WHERE_LABEL + 1).Value = "値/サブクエリNo"
        .Range("H" & ROW_WHERE_LABEL + 1).Value = ")"
        .Range("I" & ROW_WHERE_LABEL + 1).Value = "※AND=両方満たす, OR=どちらか満たす, ()で優先順位指定"
        .Range("A" & ROW_WHERE_LABEL + 1 & ":H" & ROW_WHERE_LABEL + 1).Font.Bold = True
        .Range("A" & ROW_WHERE_LABEL + 1 & ":H" & ROW_WHERE_LABEL + 1).Interior.Color = RGB(180, 198, 231)
        .Range("I" & ROW_WHERE_LABEL + 1).Font.Color = RGB(128, 128, 128)
        .Range("I" & ROW_WHERE_LABEL + 1).Font.Size = 9

        ' WHERE行
        For i = 1 To 15
            .Range("A" & ROW_WHERE_START + i - 1).Value = i
            If i = 1 Then
                AddDropdown ws, "B" & ROW_WHERE_START + i - 1, ""
            Else
                AddDropdown ws, "B" & ROW_WHERE_START + i - 1, ",AND,OR"
            End If
            AddDropdown ws, "C" & ROW_WHERE_START + i - 1, ",("
            AddDropdown ws, "F" & ROW_WHERE_START + i - 1, ",=,<>,>,<,>=,<=,LIKE,NOT LIKE,IN,NOT IN,IS NULL,IS NOT NULL,BETWEEN,EXISTS,NOT EXISTS"
            AddDropdown ws, "H" & ROW_WHERE_START + i - 1, ",)"
        Next i

        ' GROUP BY
        .Range("A" & ROW_GROUPBY).Value = "グループ化 (GROUP BY句) - 同じ値でまとめて集計"
        .Range("A" & ROW_GROUPBY).Font.Bold = True
        .Range("A" & ROW_GROUPBY).Font.Size = 12
        .Range("A" & ROW_GROUPBY).Interior.Color = RGB(221, 235, 247)
        .Range("A" & ROW_GROUPBY & ":J" & ROW_GROUPBY).Merge
        .Range("A" & ROW_GROUPBY + 1).Value = "カラム:"
        .Range("B" & ROW_GROUPBY + 1 & ":F" & ROW_GROUPBY + 1).Merge
        .Range("G" & ROW_GROUPBY + 1).Value = "※例: u.USER_ID, u.USER_NAME (集計関数以外のSELECTカラムを指定)"
        .Range("G" & ROW_GROUPBY + 1).Font.Color = RGB(128, 128, 128)
        .Range("G" & ROW_GROUPBY + 1).Font.Size = 9

        ' HAVING
        .Range("A" & ROW_HAVING_LABEL).Value = "グループ条件 (HAVING句) - 集計結果を絞り込む"
        .Range("A" & ROW_HAVING_LABEL).Font.Bold = True
        .Range("A" & ROW_HAVING_LABEL).Font.Size = 12
        .Range("A" & ROW_HAVING_LABEL).Interior.Color = RGB(221, 235, 247)
        .Range("A" & ROW_HAVING_LABEL & ":J" & ROW_HAVING_LABEL).Merge

        ' HAVINGヘッダー
        .Range("A" & ROW_HAVING_LABEL + 1).Value = "No"
        .Range("B" & ROW_HAVING_LABEL + 1).Value = "AND/OR"
        .Range("C" & ROW_HAVING_LABEL + 1).Value = "条件式"
        .Range("D" & ROW_HAVING_LABEL + 1).Value = "※例: SUM(o.AMOUNT) > 10000, COUNT(*) >= 5 (集計後の条件)"
        .Range("A" & ROW_HAVING_LABEL + 1 & ":C" & ROW_HAVING_LABEL + 1).Font.Bold = True
        .Range("A" & ROW_HAVING_LABEL + 1 & ":C" & ROW_HAVING_LABEL + 1).Interior.Color = RGB(180, 198, 231)
        .Range("D" & ROW_HAVING_LABEL + 1).Font.Color = RGB(128, 128, 128)
        .Range("D" & ROW_HAVING_LABEL + 1).Font.Size = 9

        ' HAVING行
        For i = 1 To 5
            .Range("A" & ROW_HAVING_START + i - 1).Value = i
            If i = 1 Then
                AddDropdown ws, "B" & ROW_HAVING_START + i - 1, ""
            Else
                AddDropdown ws, "B" & ROW_HAVING_START + i - 1, ",AND,OR"
            End If
            .Range("C" & ROW_HAVING_START + i - 1 & ":J" & ROW_HAVING_START + i - 1).Merge
        Next i

        ' ORDER BY
        .Range("A" & ROW_ORDERBY_LABEL).Value = "並び順 (ORDER BY句) - 結果の並べ替え"
        .Range("A" & ROW_ORDERBY_LABEL).Font.Bold = True
        .Range("A" & ROW_ORDERBY_LABEL).Font.Size = 12
        .Range("A" & ROW_ORDERBY_LABEL).Interior.Color = RGB(221, 235, 247)
        .Range("A" & ROW_ORDERBY_LABEL & ":J" & ROW_ORDERBY_LABEL).Merge

        ' ORDER BYヘッダー
        .Range("A" & ROW_ORDERBY_LABEL + 1).Value = "No"
        .Range("B" & ROW_ORDERBY_LABEL + 1).Value = "テーブル別名"
        .Range("C" & ROW_ORDERBY_LABEL + 1).Value = "カラム名/式"
        .Range("D" & ROW_ORDERBY_LABEL + 1).Value = "昇順/降順"
        .Range("E" & ROW_ORDERBY_LABEL + 1).Value = "NULLS"
        .Range("F" & ROW_ORDERBY_LABEL + 1).Value = "※ASC=昇順(小→大), DESC=降順(大→小), NULLS=NULL値の位置"
        .Range("A" & ROW_ORDERBY_LABEL + 1 & ":E" & ROW_ORDERBY_LABEL + 1).Font.Bold = True
        .Range("A" & ROW_ORDERBY_LABEL + 1 & ":E" & ROW_ORDERBY_LABEL + 1).Interior.Color = RGB(180, 198, 231)
        .Range("F" & ROW_ORDERBY_LABEL + 1).Font.Color = RGB(128, 128, 128)
        .Range("F" & ROW_ORDERBY_LABEL + 1).Font.Size = 9

        ' ORDER BY行
        For i = 1 To 10
            .Range("A" & ROW_ORDERBY_START + i - 1).Value = i
            AddDropdown ws, "D" & ROW_ORDERBY_START + i - 1, ",ASC,DESC"
            AddDropdown ws, "E" & ROW_ORDERBY_START + i - 1, ",NULLS FIRST,NULLS LAST"
        Next i

        ' 件数制限
        .Range("A" & ROW_LIMIT).Value = "件数制限 - 取得する行数を制限"
        .Range("A" & ROW_LIMIT).Font.Bold = True
        .Range("A" & ROW_LIMIT).Font.Size = 12
        .Range("A" & ROW_LIMIT).Interior.Color = RGB(221, 235, 247)
        .Range("A" & ROW_LIMIT & ":J" & ROW_LIMIT).Merge

        .Range("A" & ROW_LIMIT + 1).Value = "有効:"
        AddDropdown ws, "B" & ROW_LIMIT + 1, "なし,有効"
        .Range("C" & ROW_LIMIT + 1).Value = "件数:"
        .Range("D" & ROW_LIMIT + 1).Value = "100"
        .Range("E" & ROW_LIMIT + 1).Value = "方式:"
        AddDropdown ws, "F" & ROW_LIMIT + 1, "FETCH FIRST,ROWNUM"
        .Range("G" & ROW_LIMIT + 1).Value = "※FETCH FIRST=Oracle12c以降推奨, ROWNUM=旧方式"
        .Range("G" & ROW_LIMIT + 1).Font.Color = RGB(128, 128, 128)
        .Range("G" & ROW_LIMIT + 1).Font.Size = 9

        ' SQL出力エリア
        .Range("A" & ROW_SQL_OUTPUT).Value = "生成されたSQL"
        .Range("A" & ROW_SQL_OUTPUT).Font.Bold = True
        .Range("A" & ROW_SQL_OUTPUT).Font.Size = 12
        .Range("A" & ROW_SQL_OUTPUT).Interior.Color = RGB(198, 224, 180)
        .Range("A" & ROW_SQL_OUTPUT & ":J" & ROW_SQL_OUTPUT).Merge

        ' SQL出力セル
        .Range("A" & ROW_SQL_OUTPUT + 1 & ":J" & ROW_SQL_OUTPUT + 20).Merge
        .Range("A" & ROW_SQL_OUTPUT + 1).Font.Name = "Consolas"
        .Range("A" & ROW_SQL_OUTPUT + 1).Font.Size = 10
        .Range("A" & ROW_SQL_OUTPUT + 1).VerticalAlignment = xlTop
        .Range("A" & ROW_SQL_OUTPUT + 1).WrapText = True
        With .Range("A" & ROW_SQL_OUTPUT + 1 & ":J" & ROW_SQL_OUTPUT + 20).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With

        ' 列幅設定
        .Columns("A").ColumnWidth = 12
        .Columns("B").ColumnWidth = 14
        .Columns("C").ColumnWidth = 18
        .Columns("D").ColumnWidth = 12
        .Columns("E").ColumnWidth = 30
        .Columns("F").ColumnWidth = 16
        .Columns("G").ColumnWidth = 20
        .Columns("H").ColumnWidth = 5
        .Columns("I").ColumnWidth = 10
        .Columns("J").ColumnWidth = 10

        ' ボタン追加
        AddButton ws, "K" & ROW_TITLE, "SQL生成", "GenerateSQL"
        AddButton ws, "K" & ROW_OPTIONS, "クリア", "ClearMainSheet"
        AddButton ws, "K" & ROW_MAIN_TABLE, "履歴に保存", "SaveToHistory"
        AddButton ws, "K" & ROW_JOIN_START, "コピー", "CopySQL"
    End With
End Sub

'==============================================================================
' テーブル定義シートのフォーマット
'==============================================================================
Private Sub FormatTableDefSheet()
    Dim ws As Worksheet
    Set ws = Sheets(SHEET_TABLE_DEF)

    With ws
        ' タイトル
        .Range("A1").Value = "テーブル定義"
        .Range("A1").Font.Size = 18
        .Range("A1").Font.Bold = True
        .Range("A1:F1").Merge
        .Range("A1").Interior.Color = RGB(68, 114, 196)
        .Range("A1").Font.Color = RGB(255, 255, 255)

        ' 説明
        .Range("A2").Value = "※ここにテーブルとカラムの情報を登録してください。プルダウン選択で使用できます。"
        .Range("A2:F2").Merge
        .Range("A2").Font.Color = RGB(128, 128, 128)

        ' テーブル一覧
        .Range("A4").Value = "テーブル一覧"
        .Range("A4").Font.Bold = True
        .Range("A4").Font.Size = 12
        .Range("A4").Interior.Color = RGB(221, 235, 247)
        .Range("A4:C4").Merge

        ' テーブルヘッダー
        .Range("A5").Value = "No"
        .Range("B5").Value = "テーブル名"
        .Range("C5").Value = "説明"
        .Range("A5:C5").Font.Bold = True
        .Range("A5:C5").Interior.Color = RGB(180, 198, 231)

        ' サンプルデータ
        .Range("A6").Value = 1
        .Range("B6").Value = "USERS"
        .Range("C6").Value = "ユーザーマスタ"
        .Range("A7").Value = 2
        .Range("B7").Value = "ORDERS"
        .Range("C7").Value = "注文テーブル"
        .Range("A8").Value = 3
        .Range("B8").Value = "PRODUCTS"
        .Range("C8").Value = "商品マスタ"
        .Range("A9").Value = 4
        .Range("B9").Value = "ORDER_DETAILS"
        .Range("C9").Value = "注文明細"

        ' カラム一覧
        .Range("E4").Value = "カラム一覧"
        .Range("E4").Font.Bold = True
        .Range("E4").Font.Size = 12
        .Range("E4").Interior.Color = RGB(221, 235, 247)
        .Range("E4:H4").Merge

        ' カラムヘッダー
        .Range("E5").Value = "テーブル名"
        .Range("F5").Value = "カラム名"
        .Range("G5").Value = "データ型"
        .Range("H5").Value = "説明"
        .Range("E5:H5").Font.Bold = True
        .Range("E5:H5").Interior.Color = RGB(180, 198, 231)

        ' サンプルカラムデータ
        Dim sampleData As Variant
        sampleData = Array( _
            Array("USERS", "USER_ID", "NUMBER", "ユーザーID"), _
            Array("USERS", "USER_NAME", "VARCHAR2(100)", "ユーザー名"), _
            Array("USERS", "EMAIL", "VARCHAR2(200)", "メールアドレス"), _
            Array("USERS", "STATUS", "NUMBER", "ステータス"), _
            Array("USERS", "CREATED_AT", "DATE", "作成日時"), _
            Array("ORDERS", "ORDER_ID", "NUMBER", "注文ID"), _
            Array("ORDERS", "USER_ID", "NUMBER", "ユーザーID"), _
            Array("ORDERS", "ORDER_DATE", "DATE", "注文日"), _
            Array("ORDERS", "TOTAL_AMOUNT", "NUMBER", "合計金額"), _
            Array("PRODUCTS", "PRODUCT_ID", "NUMBER", "商品ID"), _
            Array("PRODUCTS", "PRODUCT_NAME", "VARCHAR2(200)", "商品名"), _
            Array("PRODUCTS", "PRICE", "NUMBER", "価格"), _
            Array("ORDER_DETAILS", "DETAIL_ID", "NUMBER", "明細ID"), _
            Array("ORDER_DETAILS", "ORDER_ID", "NUMBER", "注文ID"), _
            Array("ORDER_DETAILS", "PRODUCT_ID", "NUMBER", "商品ID"), _
            Array("ORDER_DETAILS", "QUANTITY", "NUMBER", "数量") _
        )

        Dim i As Long
        For i = 0 To UBound(sampleData)
            .Range("E" & (6 + i)).Value = sampleData(i)(0)
            .Range("F" & (6 + i)).Value = sampleData(i)(1)
            .Range("G" & (6 + i)).Value = sampleData(i)(2)
            .Range("H" & (6 + i)).Value = sampleData(i)(3)
        Next i

        ' 列幅設定
        .Columns("A").ColumnWidth = 5
        .Columns("B").ColumnWidth = 20
        .Columns("C").ColumnWidth = 25
        .Columns("D").ColumnWidth = 3
        .Columns("E").ColumnWidth = 20
        .Columns("F").ColumnWidth = 20
        .Columns("G").ColumnWidth = 15
        .Columns("H").ColumnWidth = 25
    End With
End Sub

'==============================================================================
' 生成履歴シートのフォーマット
'==============================================================================
Private Sub FormatHistorySheet()
    Dim ws As Worksheet
    Set ws = Sheets(SHEET_HISTORY)

    With ws
        ' タイトル
        .Range("A1").Value = "SQL生成履歴"
        .Range("A1").Font.Size = 18
        .Range("A1").Font.Bold = True
        .Range("A1:D1").Merge
        .Range("A1").Interior.Color = RGB(68, 114, 196)
        .Range("A1").Font.Color = RGB(255, 255, 255)

        ' ヘッダー
        .Range("A3").Value = "No"
        .Range("B3").Value = "生成日時"
        .Range("C3").Value = "説明"
        .Range("D3").Value = "SQL"
        .Range("A3:D3").Font.Bold = True
        .Range("A3:D3").Interior.Color = RGB(180, 198, 231)

        ' 列幅設定
        .Columns("A").ColumnWidth = 5
        .Columns("B").ColumnWidth = 18
        .Columns("C").ColumnWidth = 30
        .Columns("D").ColumnWidth = 100
    End With
End Sub

'==============================================================================
' サブクエリシートのフォーマット
'==============================================================================
Private Sub FormatSubquerySheet()
    Dim ws As Worksheet
    Set ws = Sheets(SHEET_SUBQUERY)

    With ws
        ' タイトル
        .Range("A1").Value = "サブクエリ定義"
        .Range("A1").Font.Size = 18
        .Range("A1").Font.Bold = True
        .Range("A1:F1").Merge
        .Range("A1").Interior.Color = RGB(68, 114, 196)
        .Range("A1").Font.Color = RGB(255, 255, 255)

        ' 説明
        .Range("A2").Value = "※SELECT句やWHERE句で使用するサブクエリを定義します。「サブクエリNo」でメインシートから参照できます。"
        .Range("A2:F2").Merge
        .Range("A2").Font.Color = RGB(128, 128, 128)

        ' ヘッダー
        .Range("A4").Value = "サブクエリNo"
        .Range("B4").Value = "説明"
        .Range("C4").Value = "サブクエリSQL"
        .Range("A4:C4").Font.Bold = True
        .Range("A4:C4").Interior.Color = RGB(180, 198, 231)

        ' サブクエリ入力行
        Dim i As Long
        For i = 1 To 10
            .Range("A" & (4 + i)).Value = "SUB" & i
            .Range("C" & (4 + i) & ":F" & (4 + i)).Merge
        Next i

        ' 列幅設定
        .Columns("A").ColumnWidth = 15
        .Columns("B").ColumnWidth = 25
        .Columns("C").ColumnWidth = 80
    End With
End Sub

'==============================================================================
' WITH句シートのフォーマット
'==============================================================================
Private Sub FormatCTESheet()
    Dim ws As Worksheet
    Set ws = Sheets(SHEET_CTE)

    With ws
        ' タイトル
        .Range("A1").Value = "WITH句 (共通テーブル式) 定義"
        .Range("A1").Font.Size = 18
        .Range("A1").Font.Bold = True
        .Range("A1:F1").Merge
        .Range("A1").Interior.Color = RGB(68, 114, 196)
        .Range("A1").Font.Color = RGB(255, 255, 255)

        ' 説明
        .Range("A2").Value = "※WITH句を使用する場合は、メインシートの「WITH句使用」を「使用する」に設定してください。"
        .Range("A2:F2").Merge
        .Range("A2").Font.Color = RGB(128, 128, 128)

        ' ヘッダー
        .Range("A4").Value = "CTE名"
        .Range("B4").Value = "カラム定義 (省略可)"
        .Range("C4").Value = "SELECT文"
        .Range("A4:C4").Font.Bold = True
        .Range("A4:C4").Interior.Color = RGB(180, 198, 231)

        ' CTE入力行
        Dim i As Long
        For i = 1 To 5
            .Range("C" & (4 + i) & ":F" & (4 + i)).Merge
        Next i

        ' サンプル
        .Range("A5").Value = "active_users"
        .Range("B5").Value = ""
        .Range("C5").Value = "SELECT USER_ID, USER_NAME FROM USERS WHERE STATUS = 1"

        ' 列幅設定
        .Columns("A").ColumnWidth = 20
        .Columns("B").ColumnWidth = 25
        .Columns("C").ColumnWidth = 80
    End With
End Sub

'==============================================================================
' UNIONシートのフォーマット
'==============================================================================
Private Sub FormatUnionSheet()
    Dim ws As Worksheet
    Set ws = Sheets(SHEET_UNION)

    With ws
        ' タイトル
        .Range("A1").Value = "UNION定義"
        .Range("A1").Font.Size = 18
        .Range("A1").Font.Bold = True
        .Range("A1:F1").Merge
        .Range("A1").Interior.Color = RGB(68, 114, 196)
        .Range("A1").Font.Color = RGB(255, 255, 255)

        ' 説明
        .Range("A2").Value = "※メインシートのSQLに追加でUNIONするSQLを定義します。メインシートの「UNION使用」を「使用する」に設定してください。"
        .Range("A2:F2").Merge
        .Range("A2").Font.Color = RGB(128, 128, 128)

        ' ヘッダー
        .Range("A4").Value = "No"
        .Range("B4").Value = "UNION種別"
        .Range("C4").Value = "SELECT文"
        .Range("A4:C4").Font.Bold = True
        .Range("A4:C4").Interior.Color = RGB(180, 198, 231)

        ' UNION入力行
        Dim i As Long
        For i = 1 To 5
            .Range("A" & (4 + i)).Value = i
            AddDropdown ws, "B" & (4 + i), ",UNION,UNION ALL"
            .Range("C" & (4 + i) & ":F" & (4 + i)).Merge
        Next i

        ' 列幅設定
        .Columns("A").ColumnWidth = 5
        .Columns("B").ColumnWidth = 15
        .Columns("C").ColumnWidth = 80
    End With
End Sub

'==============================================================================
' プルダウンリスト追加ヘルパー
'==============================================================================
Private Sub AddDropdown(ByVal ws As Worksheet, ByVal cellAddr As String, ByVal listItems As String)
    With ws.Range(cellAddr).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=listItems
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
End Sub

'==============================================================================
' ボタン追加ヘルパー
'==============================================================================
Private Sub AddButton(ByVal ws As Worksheet, ByVal cellAddr As String, ByVal caption As String, ByVal macroName As String)
    Dim btn As Object
    Dim rng As Range

    Set rng = ws.Range(cellAddr)
    Set btn = ws.Buttons.Add(rng.Left, rng.Top, 80, 25)
    btn.OnAction = macroName
    btn.caption = caption
    btn.Font.Size = 10
End Sub

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
        columnName = Trim(ws.Range("C" & i).Value)
        colAlias = Trim(ws.Range("D" & i).Value)
        aggFunc = Trim(ws.Range("E" & i).Value)
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

    mainTable = Trim(ws.Range("B" & ROW_MAIN_TABLE + 1).Value)
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
        joinType = Trim(ws.Range("B" & i).Value)
        joinTable = Trim(ws.Range("C" & i).Value)
        joinAlias = Trim(ws.Range("D" & i).Value)
        joinCondition = Trim(ws.Range("E" & i).Value)

        If joinType <> "" And joinTable <> "" Then
            result = result & vbCrLf & joinType & " " & joinTable
            If joinAlias <> "" Then
                result = result & " " & joinAlias
            End If
            If joinCondition <> "" And joinType <> "CROSS JOIN" Then
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
        andOr = Trim(ws.Range("B" & i).Value)
        openParen = Trim(ws.Range("C" & i).Value)
        tableAlias = Trim(ws.Range("D" & i).Value)
        columnName = Trim(ws.Range("E" & i).Value)
        operator = Trim(ws.Range("F" & i).Value)
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
        columnName = Trim(ws.Range("C" & i).Value)
        sortOrder = Trim(ws.Range("D" & i).Value)
        nullsOrder = Trim(ws.Range("E" & i).Value)

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
    nextRow = wsHistory.Cells(wsHistory.Rows.Count, 1).End(xlUp).row + 1
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
' SQLヘルプシートのフォーマット
'==============================================================================
Private Sub FormatHelpSheet()
    Dim ws As Worksheet
    Set ws = Sheets(SHEET_HELP)

    With ws
        ' タイトル
        .Range("A1").Value = "SQL構文ヘルプ - SELECT文の書き方"
        .Range("A1").Font.Size = 18
        .Range("A1").Font.Bold = True
        .Range("A1:H1").Merge
        .Range("A1").Interior.Color = RGB(68, 114, 196)
        .Range("A1").Font.Color = RGB(255, 255, 255)

        ' 基本構文
        .Range("A3").Value = "SELECT文の基本構文"
        .Range("A3").Font.Size = 14
        .Range("A3").Font.Bold = True
        .Range("A3").Interior.Color = RGB(221, 235, 247)
        .Range("A3:H3").Merge

        .Range("A4").Value = "SELECT [DISTINCT] カラム名" & vbCrLf & _
                             "FROM テーブル名" & vbCrLf & _
                             "[JOIN 結合テーブル ON 条件]" & vbCrLf & _
                             "[WHERE 抽出条件]" & vbCrLf & _
                             "[GROUP BY グループ化カラム]" & vbCrLf & _
                             "[HAVING 集計条件]" & vbCrLf & _
                             "[ORDER BY 並び順]"
        .Range("A4").Font.Name = "Consolas"
        .Range("A4:C10").Merge
        .Range("A4").VerticalAlignment = xlTop

        ' JOIN（結合）の説明
        .Range("A12").Value = "JOIN（結合）の種類"
        .Range("A12").Font.Size = 14
        .Range("A12").Font.Bold = True
        .Range("A12").Interior.Color = RGB(221, 235, 247)
        .Range("A12:H12").Merge

        ' JOINの表
        .Range("A13").Value = "結合種別"
        .Range("B13").Value = "説明"
        .Range("C13").Value = "使用例"
        .Range("A13:C13").Font.Bold = True
        .Range("A13:C13").Interior.Color = RGB(180, 198, 231)

        Dim joinData As Variant
        joinData = Array( _
            Array("INNER JOIN", "両方のテーブルに存在するデータのみ取得", "注文データと存在するユーザーのみ"), _
            Array("LEFT JOIN", "左テーブルの全データ＋右テーブルの一致データ", "全ユーザー＋注文があれば表示"), _
            Array("RIGHT JOIN", "右テーブルの全データ＋左テーブルの一致データ", "全注文＋ユーザー情報があれば表示"), _
            Array("FULL OUTER JOIN", "両方のテーブルの全データ（一致しなくてもOK）", "全ユーザーと全注文を表示"), _
            Array("CROSS JOIN", "全ての組み合わせ（直積）", "全商品×全店舗の組み合わせ") _
        )

        Dim i As Long
        For i = 0 To UBound(joinData)
            .Range("A" & (14 + i)).Value = joinData(i)(0)
            .Range("B" & (14 + i)).Value = joinData(i)(1)
            .Range("C" & (14 + i)).Value = joinData(i)(2)
        Next i

        ' JOIN図解
        .Range("E13").Value = "【INNER JOIN のイメージ】"
        .Range("E13").Font.Bold = True
        .Range("E14").Value = "テーブルA    テーブルB"
        .Range("E15").Value = "  [====共通部分====]"
        .Range("E16").Value = "  ↑この部分だけ取得"
        .Range("E14:E16").Font.Name = "Consolas"

        .Range("E18").Value = "【LEFT JOIN のイメージ】"
        .Range("E18").Font.Bold = True
        .Range("E19").Value = "[テーブルA全部][共通部分]"
        .Range("E20").Value = " ↑全部取得   ↑あれば取得"
        .Range("E19:E20").Font.Name = "Consolas"

        ' 演算子の説明
        .Range("A21").Value = "WHERE句の演算子"
        .Range("A21").Font.Size = 14
        .Range("A21").Font.Bold = True
        .Range("A21").Interior.Color = RGB(221, 235, 247)
        .Range("A21:H21").Merge

        .Range("A22").Value = "演算子"
        .Range("B22").Value = "説明"
        .Range("C22").Value = "使用例"
        .Range("A22:C22").Font.Bold = True
        .Range("A22:C22").Interior.Color = RGB(180, 198, 231)

        Dim opData As Variant
        opData = Array( _
            Array("=", "等しい", "STATUS = 1"), _
            Array("<>", "等しくない", "STATUS <> 0"), _
            Array(">、<、>=、<=", "大小比較", "AMOUNT > 1000"), _
            Array("LIKE", "パターン一致（%=任意、_=1文字）", "NAME LIKE '%田中%'"), _
            Array("IN", "リスト内に存在", "STATUS IN (1, 2, 3)"), _
            Array("BETWEEN", "範囲内", "AGE BETWEEN 20 AND 30"), _
            Array("IS NULL", "NULLかどうか", "DELETE_DATE IS NULL"), _
            Array("EXISTS", "サブクエリに結果が存在するか", "EXISTS (SELECT 1 FROM ...)") _
        )

        For i = 0 To UBound(opData)
            .Range("A" & (23 + i)).Value = opData(i)(0)
            .Range("B" & (23 + i)).Value = opData(i)(1)
            .Range("C" & (23 + i)).Value = opData(i)(2)
        Next i

        ' 集計関数の説明
        .Range("A32").Value = "集計関数"
        .Range("A32").Font.Size = 14
        .Range("A32").Font.Bold = True
        .Range("A32").Interior.Color = RGB(221, 235, 247)
        .Range("A32:H32").Merge

        .Range("A33").Value = "関数"
        .Range("B33").Value = "説明"
        .Range("C33").Value = "使用例"
        .Range("A33:C33").Font.Bold = True
        .Range("A33:C33").Interior.Color = RGB(180, 198, 231)

        Dim aggData As Variant
        aggData = Array( _
            Array("COUNT(*)", "行数をカウント", "COUNT(*) → 全行数"), _
            Array("COUNT(カラム)", "NULL以外の件数", "COUNT(EMAIL) → メールがある件数"), _
            Array("COUNT(DISTINCT カラム)", "重複を除いた件数", "COUNT(DISTINCT USER_ID)"), _
            Array("SUM(カラム)", "合計値", "SUM(AMOUNT) → 金額の合計"), _
            Array("AVG(カラム)", "平均値", "AVG(PRICE) → 価格の平均"), _
            Array("MAX(カラム)", "最大値", "MAX(ORDER_DATE) → 最新日"), _
            Array("MIN(カラム)", "最小値", "MIN(PRICE) → 最安値") _
        )

        For i = 0 To UBound(aggData)
            .Range("A" & (34 + i)).Value = aggData(i)(0)
            .Range("B" & (34 + i)).Value = aggData(i)(1)
            .Range("C" & (34 + i)).Value = aggData(i)(2)
        Next i

        ' サブクエリの説明
        .Range("A42").Value = "サブクエリ（副問い合わせ）"
        .Range("A42").Font.Size = 14
        .Range("A42").Font.Bold = True
        .Range("A42").Interior.Color = RGB(221, 235, 247)
        .Range("A42:H42").Merge

        .Range("A43").Value = "サブクエリとは、SELECT文の中に別のSELECT文を入れ子にする機能です。"
        .Range("A43:H43").Merge

        .Range("A45").Value = "使用場所"
        .Range("B45").Value = "説明"
        .Range("C45").Value = "例"
        .Range("A45:C45").Font.Bold = True
        .Range("A45:C45").Interior.Color = RGB(180, 198, 231)

        Dim subData As Variant
        subData = Array( _
            Array("SELECT句", "計算結果を列として表示", "(SELECT MAX(PRICE) FROM PRODUCTS) AS 最高価格"), _
            Array("WHERE IN", "リストの代わりにSELECT結果を使用", "WHERE USER_ID IN (SELECT USER_ID FROM VIP_USERS)"), _
            Array("WHERE EXISTS", "条件に合うデータが存在するか", "WHERE EXISTS (SELECT 1 FROM ORDERS WHERE ...)") _
        )

        For i = 0 To UBound(subData)
            .Range("A" & (46 + i)).Value = subData(i)(0)
            .Range("B" & (46 + i)).Value = subData(i)(1)
            .Range("C" & (46 + i)).Value = subData(i)(2)
        Next i

        ' WITH句の説明
        .Range("A50").Value = "WITH句（共通テーブル式 / CTE）"
        .Range("A50").Font.Size = 14
        .Range("A50").Font.Bold = True
        .Range("A50").Interior.Color = RGB(221, 235, 247)
        .Range("A50:H50").Merge

        .Range("A51").Value = "WITH句を使うと、複雑なクエリを分かりやすく整理できます。一時的な名前付きテーブルを作成するイメージです。"
        .Range("A51:H51").Merge

        .Range("A53").Value = "例："
        .Range("A53").Font.Bold = True
        .Range("A54").Value = "WITH active_users AS (" & vbCrLf & _
                              "    SELECT USER_ID, USER_NAME FROM USERS WHERE STATUS = 1" & vbCrLf & _
                              ")" & vbCrLf & _
                              "SELECT * FROM active_users WHERE ..."
        .Range("A54").Font.Name = "Consolas"
        .Range("A54:D57").Merge
        .Range("A54").VerticalAlignment = xlTop

        ' UNIONの説明
        .Range("A59").Value = "UNION（結果の結合）"
        .Range("A59").Font.Size = 14
        .Range("A59").Font.Bold = True
        .Range("A59").Interior.Color = RGB(221, 235, 247)
        .Range("A59:H59").Merge

        .Range("A60").Value = "種別"
        .Range("B60").Value = "説明"
        .Range("A60:B60").Font.Bold = True
        .Range("A60:B60").Interior.Color = RGB(180, 198, 231)

        .Range("A61").Value = "UNION"
        .Range("B61").Value = "2つのSELECT結果を結合（重複は除外）"
        .Range("A62").Value = "UNION ALL"
        .Range("B62").Value = "2つのSELECT結果を結合（重複も含む、高速）"

        .Range("A64").Value = "※UNIONを使う場合、両方のSELECTで列数と型が一致している必要があります。"
        .Range("A64").Font.Color = RGB(192, 0, 0)

        ' 列幅設定
        .Columns("A").ColumnWidth = 20
        .Columns("B").ColumnWidth = 45
        .Columns("C").ColumnWidth = 50
        .Columns("D").ColumnWidth = 5
        .Columns("E").ColumnWidth = 35
        .Columns("F").ColumnWidth = 15
        .Columns("G").ColumnWidth = 15
        .Columns("H").ColumnWidth = 15
    End With
End Sub

'==============================================================================
' テスト実行用
'==============================================================================
Public Sub TestGenerateSQL()
    InitializeSQLGenerator
End Sub
