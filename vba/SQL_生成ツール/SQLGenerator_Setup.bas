'==============================================================================
' Oracle SELECT文生成ツール - Setupモジュール
' モジュール名: SQLGenerator_Setup
'==============================================================================
' 概要:
'   SQLGeneratorの初期化とシートフォーマット機能を提供します。
'   メインモジュール（SQLGenerator）と組み合わせて使用します。
'
' 機能:
'   - ツール初期化（InitializeSQL生成ツール）
'   - シート作成・フォーマット
'   - ドロップダウンリスト設定
'   - ボタン追加
'   - Worksheet_Changeイベント設定
'
' 使い方:
'   1. このモジュールとSQLGenerator.basの両方をVBAエディタにインポート
'   2. InitializeSQL生成ツール マクロを実行してシートを初期化
'
' 作成日: 2025-12-17
'==============================================================================

Option Explicit

'==============================================================================
' 定数定義（Public - メインモジュールからも参照可能）
'==============================================================================
Public Const SHEET_MAIN As String = "メイン"
Public Const SHEET_TABLE_DEF As String = "テーブル定義"
Public Const SHEET_HISTORY As String = "生成履歴"
Public Const SHEET_SUBQUERY As String = "サブクエリ"
Public Const SHEET_CTE As String = "WITH句"
Public Const SHEET_UNION As String = "UNION"
Public Const SHEET_HELP As String = "SQLヘルプ"

' メインシートの行位置
Public Const ROW_TITLE As Long = 1

' テーブル定義書インポート設定（デフォルト値）
' ※メインシートの「設定」から変更可能
Public Const DEFAULT_TABLE_NAME_CELL As String = "J2"           ' テーブル名のセル位置
Public Const DEFAULT_TABLE_DESC_CELL As String = "D2"           ' テーブル名称のセル位置
Public Const DEFAULT_COLUMN_START_ROW As Long = 5                ' カラム定義開始行
Public Const DEFAULT_COL_NUMBER As String = "A"                  ' カラム番号の列
Public Const DEFAULT_COL_ITEM_NAME As String = "C"               ' 項目名の列
Public Const DEFAULT_COL_NAME As String = "D"                    ' カラム名の列
Public Const DEFAULT_COL_DATATYPE As String = "E"                ' データ型の列
Public Const DEFAULT_COL_LENGTH As String = "F"                  ' 桁数の列
Public Const DEFAULT_COL_NULLABLE As String = "H"                ' NULL許可の列
Public Const ROW_OPTIONS As Long = 3
Public Const ROW_MAIN_TABLE As Long = 6
Public Const ROW_JOIN_START As Long = 9
Public Const ROW_JOIN_END As Long = 18
Public Const ROW_COLUMNS_LABEL As Long = 20
Public Const ROW_COLUMNS_START As Long = 22
Public Const ROW_COLUMNS_END As Long = 41
Public Const ROW_WHERE_LABEL As Long = 43
Public Const ROW_WHERE_START As Long = 45
Public Const ROW_WHERE_END As Long = 59
Public Const ROW_GROUPBY As Long = 61
Public Const ROW_HAVING_LABEL As Long = 63
Public Const ROW_HAVING_START As Long = 65
Public Const ROW_HAVING_END As Long = 69
Public Const ROW_ORDERBY_LABEL As Long = 71
Public Const ROW_ORDERBY_START As Long = 73
Public Const ROW_ORDERBY_END As Long = 82
Public Const ROW_LIMIT As Long = 84
Public Const ROW_SQL_OUTPUT As Long = 88

'==============================================================================
' 初期化: シートを作成してフォーマットを設定
'==============================================================================
Public Sub InitializeSQL生成ツール()
    On Error GoTo ErrorHandler

    Dim autoColumnFilterEnabled As Boolean

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

    ' メインシートにWorksheet_Changeイベントを設定
    autoColumnFilterEnabled = SetupWorksheetChangeEvent()

    ' メインシートをアクティブに
    Sheets(SHEET_MAIN).Activate

    Application.ScreenUpdating = True

    If autoColumnFilterEnabled Then
        MsgBox "SQL生成ツールの初期化が完了しました。" & vbCrLf & vbCrLf & _
               "【使い方】" & vbCrLf & _
               "1. 「テーブル定義」シートにテーブル・カラム情報を登録" & vbCrLf & _
               "2. 「メイン」シートで条件を入力" & vbCrLf & _
               "3. 「SQL生成」ボタンをクリック" & vbCrLf & vbCrLf & _
               "※テーブル選択時にカラムが自動で絞り込まれます", vbInformation, "初期化完了"
    Else
        MsgBox "SQL生成ツールの初期化が完了しました。" & vbCrLf & vbCrLf & _
               "【使い方】" & vbCrLf & _
               "1. 「テーブル定義」シートにテーブル・カラム情報を登録" & vbCrLf & _
               "2. 「メイン」シートで条件を入力" & vbCrLf & _
               "3. 「SQL生成」ボタンをクリック" & vbCrLf & vbCrLf & _
               "【注意】カラム自動絞り込みを有効にするには：" & vbCrLf & _
               "テーブル選択後「カラム絞込」ボタンをクリックしてください。" & vbCrLf & vbCrLf & _
               "※自動化するにはVBAプロジェクトへのアクセス許可が必要です。" & vbCrLf & _
               "「ファイル」→「オプション」→「トラストセンター」→" & vbCrLf & _
               "「トラストセンターの設定」→「マクロの設定」→" & vbCrLf & _
               "「VBAプロジェクト オブジェクト モデルへのアクセスを信頼する」", _
               vbInformation, "初期化完了"
    End If

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
' メインシートにWorksheet_Changeイベントを設定
' ※VBAプロジェクトへのアクセス許可が必要
'   「ファイル」→「オプション」→「トラストセンター」→「トラストセンターの設定」
'   →「マクロの設定」→「VBAプロジェクト オブジェクト モデルへのアクセスを信頼する」
' 戻り値: True=成功, False=失敗（手動設定が必要）
'==============================================================================
Private Function SetupWorksheetChangeEvent() As Boolean
    On Error GoTo ManualSetup

    Dim ws As Worksheet
    Dim vbComp As Object
    Dim codeModule As Object
    Dim eventCode As String
    Dim lineNum As Long

    Set ws = Sheets(SHEET_MAIN)

    ' シートのCodeNameを取得してVBComponentにアクセス
    Set vbComp = ThisWorkbook.VBProject.VBComponents(ws.CodeName)
    Set codeModule = vbComp.CodeModule

    ' 既にWorksheet_Changeが存在するかチェック
    On Error Resume Next
    lineNum = 0
    lineNum = codeModule.ProcStartLine("Worksheet_Change", 0)
    On Error GoTo ManualSetup

    ' 既に存在する場合は成功
    If lineNum > 0 Then
        SetupWorksheetChangeEvent = True
        Exit Function
    End If

    ' イベントコードを追加
    eventCode = vbCrLf & _
        "Private Sub Worksheet_Change(ByVal Target As Range)" & vbCrLf & _
        "    ' テーブル選択時にカラムドロップダウンを自動更新" & vbCrLf & _
        "    SQLGenerator.OnTableSelectionChanged Target" & vbCrLf & _
        "End Sub" & vbCrLf

    codeModule.AddFromString eventCode
    SetupWorksheetChangeEvent = True

    Exit Function

ManualSetup:
    ' VBProjectへのアクセスが許可されていない場合
    SetupWorksheetChangeEvent = False
End Function

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

        ' メインテーブルのドロップダウン（テーブル一覧）
        Dim tableList As String
        tableList = GetTableList()
        If tableList <> "" Then
            AddDropdown ws, "B" & ROW_MAIN_TABLE + 1, tableList, "TableList"
        End If

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
            AddDropdown ws, "B" & ROW_JOIN_START + 1 + i, ",INNER JOIN(両方に存在),LEFT JOIN(左を全て),RIGHT JOIN(右を全て),FULL OUTER JOIN(両方全て),CROSS JOIN(全組合せ)"
            ' JOINテーブルのテーブル名ドロップダウン
            If tableList <> "" Then
                AddDropdown ws, "C" & ROW_JOIN_START + 1 + i, tableList, "TableList"
            End If
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
            AddDropdown ws, "E" & ROW_COLUMNS_START + i - 1, ",COUNT(件数),SUM(合計),AVG(平均),MAX(最大),MIN(最小),COUNT(DISTINCT)(重複除外件数)"
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
                AddDropdown ws, "B" & ROW_WHERE_START + i - 1, ",AND(かつ),OR(または)"
            End If
            AddDropdown ws, "C" & ROW_WHERE_START + i - 1, ",("
            AddDropdown ws, "F" & ROW_WHERE_START + i - 1, ",=(等しい),<>(等しくない),>(より大きい),<(より小さい),>=(以上),<=(以下),LIKE(部分一致),NOT LIKE(部分不一致),IN(いずれか),NOT IN(いずれでもない),IS NULL(空),IS NOT NULL(空でない),BETWEEN(範囲),EXISTS(存在する),NOT EXISTS(存在しない)"
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
                AddDropdown ws, "B" & ROW_HAVING_START + i - 1, ",AND(かつ),OR(または)"
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
            AddDropdown ws, "D" & ROW_ORDERBY_START + i - 1, ",ASC(昇順),DESC(降順)"
            AddDropdown ws, "E" & ROW_ORDERBY_START + i - 1, ",NULLS FIRST(NULL先頭),NULLS LAST(NULL末尾)"
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
        .Columns("K").ColumnWidth = 16
        .Columns("L").ColumnWidth = 16

        ' ボタン追加
        AddButton ws, "K" & ROW_TITLE, "SQL生成", "GenerateSQL"
        AddButton ws, "K" & ROW_OPTIONS, "クリア", "ClearMainSheet"
        AddButton ws, "K" & ROW_MAIN_TABLE, "履歴に保存", "SaveToHistory"
        AddButton ws, "K" & ROW_JOIN_START, "コピー", "CopySQL"
        AddButton ws, "L" & ROW_TITLE, "プルダウン更新", "UpdateDropdownsFromTableDef"
        AddButton ws, "L" & ROW_OPTIONS, "別名更新", "RefreshAliasDropdowns"
        AddButton ws, "L" & ROW_MAIN_TABLE, "カラム絞込", "RefreshColumnDropdownsByTable"
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

        ' インポート設定エリア
        .Range("J3").Value = "インポート設定"
        .Range("J3").Font.Bold = True
        .Range("J3").Font.Size = 12
        .Range("J3").Interior.Color = RGB(255, 242, 204)
        .Range("J3:K3").Merge

        ' 設定項目
        .Range("J4").Value = "テーブル名セル"
        .Range("K4").Value = DEFAULT_TABLE_NAME_CELL
        .Range("J5").Value = "テーブル名称セル"
        .Range("K5").Value = DEFAULT_TABLE_DESC_CELL
        .Range("J6").Value = "カラム開始行"
        .Range("K6").Value = DEFAULT_COLUMN_START_ROW
        .Range("J7").Value = "カラム番号列"
        .Range("K7").Value = DEFAULT_COL_NUMBER
        .Range("J8").Value = "項目名列"
        .Range("K8").Value = DEFAULT_COL_ITEM_NAME
        .Range("J9").Value = "カラム名列"
        .Range("K9").Value = DEFAULT_COL_NAME
        .Range("J10").Value = "データ型列"
        .Range("K10").Value = DEFAULT_COL_DATATYPE
        .Range("J11").Value = "桁数列"
        .Range("K11").Value = DEFAULT_COL_LENGTH
        .Range("J12").Value = "NULL列"
        .Range("K12").Value = DEFAULT_COL_NULLABLE

        ' ヘッダー色
        .Range("J4:J12").Font.Bold = True
        .Range("J4:J12").Interior.Color = RGB(255, 250, 230)

        ' フォルダパス設定
        .Range("J14").Value = "フォルダパス設定"
        .Range("J14").Font.Bold = True
        .Range("J14").Interior.Color = RGB(255, 242, 204)
        .Range("J14:K14").Merge

        .Range("J15").Value = "フォルダパス"
        .Range("K15").Value = ""
        .Range("J16").Value = "例外DB(+1列)"
        .Range("K16").Value = ""
        .Range("J15:J16").Font.Bold = True
        .Range("J15:J16").Interior.Color = RGB(255, 250, 230)

        ' 説明
        .Range("J18").Value = "※設定を変更することで、"
        .Range("J19").Value = "  異なるフォーマットの定義書に対応。"
        .Range("J20").Value = "※フォルダパスに%USERNAME%を使用可能。"
        .Range("J21").Value = "※1ファイル内の全シートを読み込みます。"
        .Range("J22").Value = "※例外DBはシート名に含まれる場合、列を+1。"
        .Range("J18:J22").Font.Size = 9
        .Range("J18:J22").Font.Color = RGB(128, 128, 128)

        ' 列幅
        .Columns("J").ColumnWidth = 16
        .Columns("K").ColumnWidth = 40

        ' ボタン追加
        AddButton ws, "J1", "定義書インポート", "ImportTableDefinitions"
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
' ※255文字を超える場合は名前付き範囲を使用
' namePrefix: 名前付き範囲のプレフィックス（TableList/ColumnList等）
'==============================================================================
Public Sub AddDropdown(ByVal ws As Worksheet, ByVal cellAddr As String, ByVal listItems As String, Optional ByVal namePrefix As String = "DropList")
    Dim items() As String
    Dim wsDef As Worksheet
    Dim rangeName As String
    Dim startRow As Long
    Dim i As Long
    Dim listRange As Range
    Dim listHash As Long
    Dim existingName As Name
    Dim targetCol As String

    With ws.Range(cellAddr).Validation
        .Delete
        If Len(listItems) > 0 Then
            ' 255文字以下の場合は直接設定
            If Len(listItems) <= 255 Then
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=listItems
            Else
                ' 255文字を超える場合は名前付き範囲を使用
                items = Split(listItems, ",")

                ' テーブル定義シートに一時リストを作成
                On Error Resume Next
                Set wsDef = Sheets(SHEET_TABLE_DEF)
                On Error GoTo 0

                If Not wsDef Is Nothing Then
                    ' プレフィックスに応じて異なる列を使用（データの上書きを防ぐ）
                    Select Case namePrefix
                        Case "TableList"
                            targetCol = "Z"
                        Case "ColumnList"
                            targetCol = "AA"
                        Case Else
                            targetCol = "AB"
                    End Select

                    ' リスト内容に基づいてユニークな名前を生成（同じリストは共有）
                    listHash = Len(listItems) + UBound(items) * 100
                    rangeName = namePrefix & "_" & listHash

                    ' 既存の名前付き範囲をチェック
                    Set existingName = Nothing
                    On Error Resume Next
                    Set existingName = ThisWorkbook.Names(rangeName)
                    On Error GoTo 0

                    ' 名前付き範囲が存在しない場合のみ作成
                    If existingName Is Nothing Then
                        ' 1行目から書き出し
                        startRow = 1

                        ' リストを縦方向に書き出し
                        For i = LBound(items) To UBound(items)
                            wsDef.Range(targetCol & (startRow + i)).Value = Trim(items(i))
                        Next i

                        ' 範囲に名前を付ける
                        Set listRange = wsDef.Range(targetCol & startRow & ":" & targetCol & (startRow + UBound(items)))
                        ThisWorkbook.Names.Add Name:=rangeName, RefersTo:=listRange
                    End If

                    ' 名前付き範囲を参照してドロップダウンを設定
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=" & rangeName
                Else
                    ' テーブル定義シートがない場合は直接設定を試みる（エラーになる可能性あり）
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=listItems
                End If
            End If
            .IgnoreBlank = True
            .InCellDropdown = True
        End If
    End With
End Sub

'==============================================================================
' ボタン追加ヘルパー
'==============================================================================
Private Sub AddButton(ByVal ws As Worksheet, ByVal cellAddr As String, ByVal caption As String, ByVal macroName As String)
    Dim btn As Object
    Dim rng As Range

    Set rng = ws.Range(cellAddr)
    Set btn = ws.Buttons.Add(rng.Left, rng.Top, 110, 28)
    btn.OnAction = macroName
    btn.caption = caption
    btn.Font.Size = 10
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
' インポート設定の初期化
'==============================================================================
Public Sub InitializeImportSettings()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = Sheets(SHEET_TABLE_DEF)

    ' 設定エリアのタイトル
    ws.Range("J3").Value = "インポート設定"
    ws.Range("J3").Font.Bold = True
    ws.Range("J3").Font.Size = 12
    ws.Range("J3").Interior.Color = RGB(255, 242, 204)
    ws.Range("J3:K3").Merge

    ' 設定項目
    ws.Range("J4").Value = "テーブル名セル"
    ws.Range("K4").Value = DEFAULT_TABLE_NAME_CELL
    ws.Range("J5").Value = "テーブル名称セル"
    ws.Range("K5").Value = DEFAULT_TABLE_DESC_CELL
    ws.Range("J6").Value = "カラム開始行"
    ws.Range("K6").Value = DEFAULT_COLUMN_START_ROW
    ws.Range("J7").Value = "カラム番号列"
    ws.Range("K7").Value = DEFAULT_COL_NUMBER
    ws.Range("J8").Value = "項目名列"
    ws.Range("K8").Value = DEFAULT_COL_ITEM_NAME
    ws.Range("J9").Value = "カラム名列"
    ws.Range("K9").Value = DEFAULT_COL_NAME
    ws.Range("J10").Value = "データ型列"
    ws.Range("K10").Value = DEFAULT_COL_DATATYPE
    ws.Range("J11").Value = "桁数列"
    ws.Range("K11").Value = DEFAULT_COL_LENGTH
    ws.Range("J12").Value = "NULL列"
    ws.Range("K12").Value = DEFAULT_COL_NULLABLE

    ' ヘッダー色
    ws.Range("J4:J12").Font.Bold = True
    ws.Range("J4:J12").Interior.Color = RGB(255, 250, 230)

    ' フォルダパス設定
    ws.Range("J14").Value = "フォルダパス設定"
    ws.Range("J14").Font.Bold = True
    ws.Range("J14").Interior.Color = RGB(255, 242, 204)
    ws.Range("J14:K14").Merge

    ws.Range("J15").Value = "フォルダパス"
    ws.Range("K15").Value = ""
    ws.Range("J16").Value = "例外DB(+1列)"
    ws.Range("K16").Value = ""
    ws.Range("J15:J16").Font.Bold = True
    ws.Range("J15:J16").Interior.Color = RGB(255, 250, 230)

    ' 説明
    ws.Range("J18").Value = "※設定を変更することで、"
    ws.Range("J19").Value = "  異なるフォーマットの定義書に対応。"
    ws.Range("J20").Value = "※フォルダパスに%USERNAME%を使用可能。"
    ws.Range("J21").Value = "※1ファイル内の全シートを読み込みます。"
    ws.Range("J22").Value = "※例外DBはシート名に含まれる場合、列を+1。"
    ws.Range("J18:J22").Font.Size = 9
    ws.Range("J18:J22").Font.Color = RGB(128, 128, 128)

    ' 列幅
    ws.Columns("J").ColumnWidth = 16
    ws.Columns("K").ColumnWidth = 40

    MsgBox "インポート設定を初期化しました。", vbInformation, "完了"

    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub
