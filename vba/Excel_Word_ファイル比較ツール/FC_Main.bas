Attribute VB_Name = "FC_Main"
Option Explicit

'==============================================================================
' Excel/Word ファイル比較ツール - メインモジュール
' エントリーポイント、ファイル選択機能を提供
'==============================================================================

' ============================================================================
' Excel専用比較プロシージャ（ボタン用）
' ============================================================================
Public Sub CompareExcelFiles()
    Dim file1Path As String
    Dim file2Path As String

    On Error GoTo ErrorHandler

    ' 1つ目のExcelファイル選択
    MsgBox "2つのExcelファイルを比較します。" & vbCrLf & vbCrLf & _
           "まず、1つ目のExcelファイル（旧ファイル）を選択してください。", _
           vbInformation, "Excel ファイル比較"

    file1Path = SelectExcelFile("1つ目のExcelファイル（旧ファイル）を選択")
    If file1Path = "" Then
        MsgBox "ファイル選択がキャンセルされました。", vbExclamation
        Exit Sub
    End If

    ' 2つ目のExcelファイル選択
    MsgBox "次に、2つ目のExcelファイル（新ファイル）を選択してください。", _
           vbInformation, "Excel ファイル比較"

    file2Path = SelectExcelFile("2つ目のExcelファイル（新ファイル）を選択")
    If file2Path = "" Then
        MsgBox "ファイル選択がキャンセルされました。", vbExclamation
        Exit Sub
    End If

    ' 同じファイルが選択された場合
    If LCase(file1Path) = LCase(file2Path) Then
        MsgBox "同じファイルが選択されました。異なるファイルを選択してください。", vbExclamation
        Exit Sub
    End If

    ' Excel比較を実行
    CompareExcelFilesInternal file1Path, file2Path

    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & vbCrLf & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, vbCritical, "エラー"
End Sub

' ============================================================================
' Word専用比較プロシージャ（ボタン用）
' ============================================================================
Public Sub CompareWordFiles()
    Dim file1Path As String
    Dim file2Path As String

    On Error GoTo ErrorHandler

    ' 1つ目のWordファイル選択
    MsgBox "2つのWordファイルを比較します。" & vbCrLf & vbCrLf & _
           "まず、1つ目のWordファイル（旧ファイル）を選択してください。", _
           vbInformation, "Word ファイル比較"

    file1Path = SelectWordFile("1つ目のWordファイル（旧ファイル）を選択")
    If file1Path = "" Then
        MsgBox "ファイル選択がキャンセルされました。", vbExclamation
        Exit Sub
    End If

    ' 2つ目のWordファイル選択
    MsgBox "次に、2つ目のWordファイル（新ファイル）を選択してください。", _
           vbInformation, "Word ファイル比較"

    file2Path = SelectWordFile("2つ目のWordファイル（新ファイル）を選択")
    If file2Path = "" Then
        MsgBox "ファイル選択がキャンセルされました。", vbExclamation
        Exit Sub
    End If

    ' 同じファイルが選択された場合
    If LCase(file1Path) = LCase(file2Path) Then
        MsgBox "同じファイルが選択されました。異なるファイルを選択してください。", vbExclamation
        Exit Sub
    End If

    ' Word比較を実行
    CompareWordFilesInternal file1Path, file2Path

    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & vbCrLf & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, vbCritical, "エラー"
End Sub

' ============================================================================
' Excelファイル選択ダイアログ
' ============================================================================
Private Function SelectExcelFile(ByVal dialogTitle As String) As String
    Dim fd As Object

    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Title = dialogTitle
        .Filters.Clear
        .Filters.Add "Excel ファイル", "*.xlsx;*.xlsm;*.xls;*.xlsb"
        .Filters.Add "すべてのファイル", "*.*"
        .FilterIndex = 1
        .AllowMultiSelect = False

        If .Show = -1 Then
            SelectExcelFile = .SelectedItems(1)
        Else
            SelectExcelFile = ""
        End If
    End With
End Function

' ============================================================================
' Wordファイル選択ダイアログ
' ============================================================================
Private Function SelectWordFile(ByVal dialogTitle As String) As String
    Dim fd As Object

    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Title = dialogTitle
        .Filters.Clear
        .Filters.Add "Word ファイル", "*.docx;*.docm;*.doc"
        .Filters.Add "すべてのファイル", "*.*"
        .FilterIndex = 1
        .AllowMultiSelect = False

        If .Show = -1 Then
            SelectWordFile = .SelectedItems(1)
        Else
            SelectWordFile = ""
        End If
    End With
End Function
