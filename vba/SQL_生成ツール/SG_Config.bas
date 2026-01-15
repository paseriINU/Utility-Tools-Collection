Attribute VB_Name = "SG_Config"
Option Explicit

'==============================================================================
' Oracle SELECT文生成ツール - 設定モジュール
' 定数、設定を提供
' ※このモジュールはSetupモジュールを削除しても動作するよう設計
'==============================================================================

' ============================================================================
' シート名定数
' ============================================================================
Public Const SHEET_MAIN As String = "メイン"
Public Const SHEET_TABLE_DEF As String = "テーブル定義"
Public Const SHEET_HISTORY As String = "生成履歴"
Public Const SHEET_SUBQUERY As String = "サブクエリ"
Public Const SHEET_CTE As String = "WITH句"
Public Const SHEET_UNION As String = "UNION"
Public Const SHEET_HELP As String = "SQLヘルプ"

' ============================================================================
' メインシートの行位置
' ============================================================================
Public Const ROW_TITLE As Long = 1
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

' ============================================================================
' テーブル定義書インポート設定（デフォルト値）
' ※メインシートの「設定」から変更可能
' ============================================================================
Public Const DEFAULT_TABLE_NAME_CELL As String = "J2"           ' テーブル名のセル位置
Public Const DEFAULT_TABLE_DESC_CELL As String = "D2"           ' テーブル名称のセル位置
Public Const DEFAULT_COLUMN_START_ROW As Long = 5                ' カラム定義開始行
Public Const DEFAULT_COL_NUMBER As String = "A"                  ' カラム番号の列
Public Const DEFAULT_COL_ITEM_NAME As String = "C"               ' 項目名の列
Public Const DEFAULT_COL_NAME As String = "D"                    ' カラム名の列
Public Const DEFAULT_COL_DATATYPE As String = "E"                ' データ型の列
Public Const DEFAULT_COL_LENGTH As String = "F"                  ' 桁数の列
Public Const DEFAULT_COL_NULLABLE As String = "H"                ' NULL許可の列

