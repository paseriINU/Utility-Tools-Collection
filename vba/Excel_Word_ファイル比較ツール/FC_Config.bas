Attribute VB_Name = "FC_Config"
Option Explicit

'==============================================================================
' Excel/Word ファイル比較ツール - 設定モジュール
' 定数、型定義、設定取得機能を提供
'==============================================================================

' ============================================================================
' 定数（差異ハイライト色）
' ============================================================================
Public Const COLOR_CHANGED As Long = 65535      ' 黄色: 値変更
Public Const COLOR_ADDED As Long = 5296274      ' 緑: 追加
Public Const COLOR_DELETED As Long = 13421823   ' ピンク: 削除
Public Const COLOR_STYLE As Long = 13408614     ' 薄紫: スタイル変更

' ============================================================================
' シート名定数
' ============================================================================
Public Const SHEET_MAIN As String = "メイン"
Public Const SHEET_RESULT As String = "比較結果"

' ============================================================================
' データ構造: Excel比較用
' ============================================================================
Public Type ExcelDiffInfo
    SheetName As String      ' シート名
    CellAddress As String    ' セルアドレス
    DiffType As String       ' 差異タイプ（変更/追加/削除）
    OldValue As String       ' 旧ファイルの値
    NewValue As String       ' 新ファイルの値
End Type

' ============================================================================
' データ構造: Word比較用（WinMerge方式：旧/新両方の行番号を保持）
' ============================================================================
Public Type WordDiffInfo
    OldParagraphNo As Long   ' 旧ファイルの段落番号（0は該当なし）
    NewParagraphNo As Long   ' 新ファイルの段落番号（0は該当なし）
    DiffType As String       ' 差異タイプ（変更/追加/削除/スタイル変更）
    OldText As String        ' 旧ファイルのテキスト
    NewText As String        ' 新ファイルのテキスト
    OldStyle As String       ' 旧ファイルのスタイル情報
    NewStyle As String       ' 新ファイルのスタイル情報
End Type

' ============================================================================
' モジュールレベル変数: テキスト一致段落のスタイル比較用
' ============================================================================
Public g_MatchedOld() As Long    ' 旧ファイルの段落番号
Public g_MatchedNew() As Long    ' 新ファイルの段落番号
Public g_MatchedCount As Long    ' ペア数

' ============================================================================
' メインシートのチェックボックス状態を取得（LCSモード）
' True: LCSモード、False: 簡易モード
' ============================================================================
Public Function GetUseLCSMode() As Boolean
    Dim ws As Worksheet
    Dim chkBox As CheckBox

    On Error Resume Next

    ' メインシートを取得
    Set ws = ThisWorkbook.Worksheets(SHEET_MAIN)
    If ws Is Nothing Then
        GetUseLCSMode = False
        Exit Function
    End If

    ' チェックボックスを取得
    Set chkBox = ws.CheckBoxes("chkUseLCS")
    If chkBox Is Nothing Then
        GetUseLCSMode = False
        Exit Function
    End If

    ' チェック状態を返す
    GetUseLCSMode = (chkBox.Value = xlOn)

    On Error GoTo 0
End Function

' ============================================================================
' メインシートのスタイル比較チェックボックス状態を取得
' True: スタイル比較する、False: スタイル比較しない
' ============================================================================
Public Function GetCheckStyleMode() As Boolean
    Dim ws As Worksheet
    Dim chkBox As CheckBox

    On Error Resume Next

    ' メインシートを取得
    Set ws = ThisWorkbook.Worksheets(SHEET_MAIN)
    If ws Is Nothing Then
        GetCheckStyleMode = True
        Exit Function
    End If

    ' チェックボックスを取得
    Set chkBox = ws.CheckBoxes("chkCheckStyle")
    If chkBox Is Nothing Then
        GetCheckStyleMode = True
        Exit Function
    End If

    ' チェック状態を返す
    GetCheckStyleMode = (chkBox.Value = xlOn)

    On Error GoTo 0
End Function

' ============================================================================
' 進捗表示
' ============================================================================
Public Sub ShowProgress(ByVal phase As String, ByVal current As Long, ByVal total As Long)
    Dim pct As Long
    Dim progressBar As String
    Dim barLength As Long
    Dim filledLength As Long
    Dim i As Long

    If total > 0 Then
        pct = CLng((current / total) * 100)
    Else
        pct = 0
    End If

    ' プログレスバー（20文字幅）
    barLength = 20
    filledLength = CLng(barLength * current / IIf(total > 0, total, 1))
    progressBar = ""
    For i = 1 To filledLength
        progressBar = progressBar & ChrW(&H2588)  ' █
    Next i
    For i = filledLength + 1 To barLength
        progressBar = progressBar & ChrW(&H2591)  ' ░
    Next i

    Application.StatusBar = phase & " " & progressBar & " " & pct & "% (" & current & "/" & total & ")"
    DoEvents
End Sub

Public Sub ClearProgress()
    Application.StatusBar = False
End Sub

' ============================================================================
' テキストをクリーンアップ（改行・特殊文字を除去）
' ============================================================================
Public Function CleanText(ByVal txt As String) As String
    ' 改行・段落記号を除去
    txt = Replace(txt, vbCr, "")
    txt = Replace(txt, vbLf, "")
    txt = Replace(txt, Chr(13), "")
    txt = Replace(txt, Chr(11), " ")  ' 行区切り
    txt = Replace(txt, Chr(7), "")    ' セル終端記号

    ' 前後の空白を除去
    CleanText = Trim(txt)
End Function

' ============================================================================
' 値の比較（数値の微小差異を考慮）
' ============================================================================
Public Function IsEqual(ByVal val1 As Variant, ByVal val2 As Variant) As Boolean
    ' 両方Empty
    If IsEmpty(val1) And IsEmpty(val2) Then
        IsEqual = True
        Exit Function
    End If

    ' 片方がEmpty
    If IsEmpty(val1) Or IsEmpty(val2) Then
        IsEqual = False
        Exit Function
    End If

    ' 両方数値の場合、浮動小数点誤差を考慮
    If IsNumeric(val1) And IsNumeric(val2) Then
        If Abs(CDbl(val1) - CDbl(val2)) < 0.0000001 Then
            IsEqual = True
        Else
            IsEqual = False
        End If
        Exit Function
    End If

    ' 文字列比較
    IsEqual = (CStr(val1) = CStr(val2))
End Function
