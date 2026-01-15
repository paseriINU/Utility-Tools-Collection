Attribute VB_Name = "WBO_Config"
Option Explicit

' ============================================================================
' Word しおり整理ツール - 設定モジュール
' 定数、型定義、設定読み込み機能を提供
' ============================================================================

' === シート名定数 ===
Public Const SHEET_MAIN As String = "Word_しおり整理ツール"
Public Const SHEET_SETTINGS As String = "設定"

' === 設定シートのセル位置定数 ===
' フォルダ設定
Public Const SETTINGS_ROW_FOLDER_HEADER As Long = 2
Public Const SETTINGS_ROW_INPUT_FOLDER As Long = 3
Public Const SETTINGS_ROW_OUTPUT_FOLDER As Long = 4

' スタイル設定テーブル
Public Const SETTINGS_ROW_STYLE_HEADER As Long = 7
Public Const SETTINGS_ROW_STYLE_START As Long = 8   ' スタイル設定開始行

' オプション設定
Public Const SETTINGS_ROW_OPTION_HEADER As Long = 30
Public Const SETTINGS_ROW_PDF_OUTPUT As Long = 31

' 列定数
Public Const SETTINGS_COL_LABEL As Long = 2     ' B列
Public Const SETTINGS_COL_VALUE As Long = 3     ' C列
Public Const SETTINGS_COL_PATTERN As Long = 4   ' D列（パターン）
Public Const SETTINGS_COL_STYLE As Long = 5     ' E列（スタイル名）
Public Const SETTINGS_COL_NOTE As Long = 6      ' F列（備考）

' === デフォルトスタイル名定数（ヘッダーのSTYLEREF更新用） ===
Public Const DEFAULT_LEVEL1_STYLE As String = "表題1"
Public Const DEFAULT_LEVEL2_STYLE As String = "表題2"
Public Const DEFAULT_LEVEL3_STYLE As String = "表題3"
Public Const DEFAULT_LEVEL4_STYLE As String = "表題4"
Public Const DEFAULT_LEVEL5_STYLE As String = "表題5"

' === スタイル設定構造体（動的配列対応） ===
Public Type StyleSetting
    Category As String      ' 種別: パターン, 帳票, 特定, 例外
    Level As String         ' レベル: 1, 2, 3, 3-節, 4, 4-節, 5-節 など
    Pattern As String       ' パターン（正規表現）またはテキスト
    StyleName As String     ' 適用スタイル名
End Type

' ============================================================================
' 設定シートを取得
' ============================================================================
Public Function GetSettingsSheet() As Worksheet
    On Error Resume Next
    Set GetSettingsSheet = ThisWorkbook.Worksheets(SHEET_SETTINGS)
    If GetSettingsSheet Is Nothing Then
        MsgBox "エラー: 「設定」シートが見つかりません。" & vbCrLf & _
               "初期化マクロを実行してください。", vbCritical
    End If
    On Error GoTo 0
End Function

' ============================================================================
' フォルダパスを取得
' ============================================================================
Public Function GetInputFolder(ByRef wsSettings As Worksheet) As String
    Dim folder As String
    folder = CStr(wsSettings.Cells(SETTINGS_ROW_INPUT_FOLDER, SETTINGS_COL_VALUE).Value)
    If Right(folder, 1) <> "\" Then folder = folder & "\"
    GetInputFolder = folder
End Function

Public Function GetOutputFolder(ByRef wsSettings As Worksheet) As String
    Dim folder As String
    folder = CStr(wsSettings.Cells(SETTINGS_ROW_OUTPUT_FOLDER, SETTINGS_COL_VALUE).Value)
    If Right(folder, 1) <> "\" Then folder = folder & "\"
    GetOutputFolder = folder
End Function

' ============================================================================
' 設定シートから動的に設定を読み込み
' ============================================================================
Public Function LoadSettings(ByRef wsSettings As Worksheet, _
                             ByRef styleSettings() As StyleSetting, _
                             ByRef styleCount As Long, _
                             ByRef doPdfOutput As Boolean) As Boolean
    On Error GoTo ErrorHandler

    Dim row As Long
    Dim category As String
    Dim maxRows As Long
    Dim emptyRowCount As Long

    maxRows = 100
    styleCount = 0
    ReDim styleSettings(0 To maxRows - 1)

    row = SETTINGS_ROW_STYLE_START
    emptyRowCount = 0

    Do While row < SETTINGS_ROW_STYLE_START + maxRows
        category = Trim(CStr(wsSettings.Cells(row, SETTINGS_COL_LABEL).Value))

        If category = "" Then
            emptyRowCount = emptyRowCount + 1
            If emptyRowCount >= 3 Then Exit Do
            row = row + 1
            GoTo NextRow
        End If

        emptyRowCount = 0

        styleSettings(styleCount).Category = category
        styleSettings(styleCount).Level = Trim(CStr(wsSettings.Cells(row, SETTINGS_COL_VALUE).Value))
        styleSettings(styleCount).Pattern = Trim(CStr(wsSettings.Cells(row, SETTINGS_COL_PATTERN).Value))
        styleSettings(styleCount).StyleName = Trim(CStr(wsSettings.Cells(row, SETTINGS_COL_STYLE).Value))

        styleCount = styleCount + 1
        row = row + 1
NextRow:
    Loop

    If styleCount > 0 Then
        ReDim Preserve styleSettings(0 To styleCount - 1)
    End If

    doPdfOutput = (CStr(wsSettings.Cells(SETTINGS_ROW_PDF_OUTPUT, SETTINGS_COL_VALUE).Value) = "はい")

    LoadSettings = True
    Exit Function

ErrorHandler:
    LoadSettings = False
End Function

' ============================================================================
' フォルダの存在確認
' ============================================================================
Public Function ValidateFolders(ByVal inputDir As String, ByVal outputDir As String) As Boolean
    If Dir(inputDir, vbDirectory) = "" Then
        MsgBox "エラー: 入力フォルダが存在しません。" & vbCrLf & vbCrLf & _
               inputDir & vbCrLf & vbCrLf & _
               "設定シートのフォルダ設定を確認してください。", vbCritical, "フォルダエラー"
        ValidateFolders = False
        Exit Function
    End If

    If Dir(outputDir, vbDirectory) = "" Then
        MsgBox "エラー: 出力フォルダが存在しません。" & vbCrLf & vbCrLf & _
               outputDir & vbCrLf & vbCrLf & _
               "設定シートのフォルダ設定を確認してください。", vbCritical, "フォルダエラー"
        ValidateFolders = False
        Exit Function
    End If

    ValidateFolders = True
End Function

' ============================================================================
' Inputフォルダから処理対象のWord文書を選択
' ============================================================================
Public Function SelectWordFile(ByVal inputDir As String) As String
    Dim fileList() As String
    Dim fileCount As Long
    Dim fileName As String
    Dim i As Long
    Dim selectedIndex As Long
    Dim msg As String
    Dim isDuplicate As Boolean

    fileCount = 0
    ReDim fileList(0 To 99)

    fileName = Dir(inputDir & "*.docx")
    Do While fileName <> ""
        fileList(fileCount) = fileName
        fileCount = fileCount + 1
        fileName = Dir()
    Loop

    fileName = Dir(inputDir & "*.doc")
    Do While fileName <> ""
        isDuplicate = False
        For i = 0 To fileCount - 1
            If Left(fileList(i), Len(fileList(i)) - 1) = fileName Then
                isDuplicate = True
                Exit For
            End If
        Next i

        If Not isDuplicate Then
            fileList(fileCount) = fileName
            fileCount = fileCount + 1
        End If
        fileName = Dir()
    Loop

    If fileCount = 0 Then
        MsgBox "Inputフォルダ内にWord文書が見つかりません。" & vbCrLf & vbCrLf & _
               "フォルダ: " & inputDir, vbExclamation, "ファイルなし"
        SelectWordFile = ""
        Exit Function
    End If

    If fileCount = 1 Then
        SelectWordFile = inputDir & fileList(0)
        Exit Function
    End If

    msg = "処理するWord文書を選択してください:" & vbCrLf & vbCrLf
    For i = 0 To fileCount - 1
        msg = msg & (i + 1) & ". " & fileList(i) & vbCrLf
    Next i
    msg = msg & vbCrLf & "番号を入力 (1-" & fileCount & "):"

    Dim userInput As String
    userInput = InputBox(msg, "ファイル選択", "1")

    If userInput = "" Then
        SelectWordFile = ""
        Exit Function
    End If

    If Not IsNumeric(userInput) Then
        MsgBox "無効な入力です。", vbExclamation
        SelectWordFile = ""
        Exit Function
    End If

    selectedIndex = CLng(userInput) - 1
    If selectedIndex < 0 Or selectedIndex >= fileCount Then
        MsgBox "無効な番号です。", vbExclamation
        SelectWordFile = ""
        Exit Function
    End If

    SelectWordFile = inputDir & fileList(selectedIndex)
End Function
