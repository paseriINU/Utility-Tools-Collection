Attribute VB_Name = "CommonUtils"
Option Explicit

' ============================================================================
' 汎用ユーティリティモジュール
' 他のVBAマクロでも再利用可能な共通関数を提供
' ============================================================================

' ============================================================================
' ファイル・フォルダ操作
' ============================================================================

' フォルダが存在するか確認
Public Function FolderExists(ByVal folderPath As String) As Boolean
    On Error Resume Next
    FolderExists = (Dir(folderPath, vbDirectory) <> "")
    On Error GoTo 0
End Function

' フォルダパスの末尾に\を追加（なければ）
Public Function EnsureTrailingBackslash(ByVal folderPath As String) As String
    If Right(folderPath, 1) <> "\" Then
        EnsureTrailingBackslash = folderPath & "\"
    Else
        EnsureTrailingBackslash = folderPath
    End If
End Function

' フォルダ内のファイル一覧を取得
Public Function GetFilesInFolder(ByVal folderPath As String, _
                                  ByVal pattern As String, _
                                  ByRef fileList() As String) As Long
    Dim fileCount As Long
    Dim fileName As String

    fileCount = 0
    ReDim fileList(0 To 99)

    folderPath = EnsureTrailingBackslash(folderPath)
    fileName = Dir(folderPath & pattern)

    Do While fileName <> ""
        If fileCount > UBound(fileList) Then
            ReDim Preserve fileList(0 To fileCount + 100)
        End If
        fileList(fileCount) = fileName
        fileCount = fileCount + 1
        fileName = Dir()
    Loop

    If fileCount > 0 Then
        ReDim Preserve fileList(0 To fileCount - 1)
    End If

    GetFilesInFolder = fileCount
End Function

' ファイル選択ダイアログを表示
Public Function ShowFileSelectionDialog(ByVal folderPath As String, _
                                         ByVal title As String, _
                                         ByVal filter As String) As String
    Dim fd As Object
    Set fd = Application.FileDialog(3) ' msoFileDialogFilePicker = 3

    With fd
        .Title = title
        .InitialFileName = EnsureTrailingBackslash(folderPath)
        .Filters.Clear
        .Filters.Add "対象ファイル", filter
        .AllowMultiSelect = False

        If .Show = -1 Then
            ShowFileSelectionDialog = .SelectedItems(1)
        Else
            ShowFileSelectionDialog = ""
        End If
    End With

    Set fd = Nothing
End Function

' フォルダ選択ダイアログを表示
Public Function ShowFolderSelectionDialog(ByVal title As String, _
                                           Optional ByVal initialPath As String = "") As String
    Dim fd As Object
    Set fd = Application.FileDialog(4) ' msoFileDialogFolderPicker = 4

    With fd
        .Title = title
        If initialPath <> "" Then
            .InitialFileName = EnsureTrailingBackslash(initialPath)
        End If

        If .Show = -1 Then
            ShowFolderSelectionDialog = .SelectedItems(1)
        Else
            ShowFolderSelectionDialog = ""
        End If
    End With

    Set fd = Nothing
End Function

' ============================================================================
' 文字列操作
' ============================================================================

' 全角を半角に変換（数字、アルファベット、記号）
Public Function ToHalfWidth(ByVal Text As String) As String
    Dim i As Long
    Dim char As String
    Dim result As String
    Dim charCode As Long

    result = ""
    For i = 1 To Len(Text)
        char = Mid(Text, i, 1)
        charCode = AscW(char)

        Select Case charCode
            Case &HFF10 To &HFF19  ' 全角数字 ０-９
                result = result & Chr(charCode - &HFF10 + Asc("0"))
            Case &HFF21 To &HFF3A  ' 全角大文字 Ａ-Ｚ
                result = result & Chr(charCode - &HFF21 + Asc("A"))
            Case &HFF41 To &HFF5A  ' 全角小文字 ａ-ｚ
                result = result & Chr(charCode - &HFF41 + Asc("a"))
            Case &HFF0D, &H2212, &H30FC  ' 全角ハイフン、マイナス、長音
                result = result & "-"
            Case &HFF0C  ' 全角カンマ
                result = result & ","
            Case &HFF0E  ' 全角ピリオド
                result = result & "."
            Case &HFF08  ' 全角左括弧
                result = result & "("
            Case &HFF09  ' 全角右括弧
                result = result & ")"
            Case &H3000  ' 全角スペース
                result = result & " "
            Case &HFF1A  ' 全角コロン
                result = result & ":"
            Case &HFF1B  ' 全角セミコロン
                result = result & ";"
            Case &HFF01  ' 全角感嘆符
                result = result & "!"
            Case &HFF1F  ' 全角疑問符
                result = result & "?"
            Case Else
                result = result & char
        End Select
    Next i

    ToHalfWidth = result
End Function

' 半角を全角に変換（数字、アルファベット）
Public Function ToFullWidth(ByVal Text As String) As String
    Dim i As Long
    Dim char As String
    Dim result As String
    Dim charCode As Long

    result = ""
    For i = 1 To Len(Text)
        char = Mid(Text, i, 1)
        charCode = Asc(char)

        Select Case charCode
            Case Asc("0") To Asc("9")  ' 半角数字
                result = result & ChrW(&HFF10 + charCode - Asc("0"))
            Case Asc("A") To Asc("Z")  ' 半角大文字
                result = result & ChrW(&HFF21 + charCode - Asc("A"))
            Case Asc("a") To Asc("z")  ' 半角小文字
                result = result & ChrW(&HFF41 + charCode - Asc("a"))
            Case Else
                result = result & char
        End Select
    Next i

    ToFullWidth = result
End Function

' 制御文字を除去（CR, LF, Tab, FormFeed, VerticalTab, BellなD）
Public Function RemoveControlChars(ByVal Text As String) As String
    Dim result As String
    result = Text
    result = Replace(result, vbCr, "")
    result = Replace(result, vbLf, "")
    result = Replace(result, vbTab, " ")
    result = Replace(result, Chr(12), "")  ' FormFeed
    result = Replace(result, Chr(11), "")  ' VerticalTab
    result = Replace(result, Chr(7), "")   ' Bell
    RemoveControlChars = Trim(result)
End Function

' 文字列が空または空白のみかチェック
Public Function IsNullOrWhitespace(ByVal Text As String) As Boolean
    IsNullOrWhitespace = (Trim(Text) = "")
End Function

' ============================================================================
' 正規表現
' ============================================================================

' 正規表現マッチ（True/False）
Public Function RegexMatch(ByVal Text As String, ByVal Pattern As String, _
                           Optional ByVal ignoreCase As Boolean = False) As Boolean
    On Error GoTo ErrorHandler

    If Text = "" Or Pattern = "" Then
        RegexMatch = False
        Exit Function
    End If

    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")

    With regex
        .Global = False
        .ignoreCase = ignoreCase
        .MultiLine = False
        .Pattern = Pattern
    End With

    RegexMatch = regex.Test(Text)
    Set regex = Nothing
    Exit Function

ErrorHandler:
    RegexMatch = False
End Function

' 正規表現で最初のマッチを取得
Public Function RegexFirstMatch(ByVal Text As String, ByVal Pattern As String, _
                                 Optional ByVal ignoreCase As Boolean = False) As String
    On Error GoTo ErrorHandler

    If Text = "" Or Pattern = "" Then
        RegexFirstMatch = ""
        Exit Function
    End If

    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")

    With regex
        .Global = False
        .ignoreCase = ignoreCase
        .MultiLine = False
        .Pattern = Pattern
    End With

    Dim matches As Object
    Set matches = regex.Execute(Text)

    If matches.Count > 0 Then
        RegexFirstMatch = matches(0).Value
    Else
        RegexFirstMatch = ""
    End If

    Set matches = Nothing
    Set regex = Nothing
    Exit Function

ErrorHandler:
    RegexFirstMatch = ""
End Function

' 正規表現で置換
Public Function RegexReplace(ByVal Text As String, ByVal Pattern As String, _
                              ByVal replacement As String, _
                              Optional ByVal ignoreCase As Boolean = False, _
                              Optional ByVal replaceAll As Boolean = True) As String
    On Error GoTo ErrorHandler

    If Text = "" Or Pattern = "" Then
        RegexReplace = Text
        Exit Function
    End If

    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")

    With regex
        .Global = replaceAll
        .ignoreCase = ignoreCase
        .MultiLine = False
        .Pattern = Pattern
    End With

    RegexReplace = regex.Replace(Text, replacement)
    Set regex = Nothing
    Exit Function

ErrorHandler:
    RegexReplace = Text
End Function

' ============================================================================
' コレクション操作
' ============================================================================

' コレクションにキーが存在するかチェック
Public Function CollectionKeyExists(ByRef col As Collection, ByVal Key As String) As Boolean
    On Error Resume Next
    Dim dummy As Variant
    dummy = col(Key)
    CollectionKeyExists = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

' 配列を文字列に結合（Join関数の拡張版）
Public Function JoinArray(ByRef arr() As String, _
                           Optional ByVal delimiter As String = ",", _
                           Optional ByVal startIndex As Long = -1, _
                           Optional ByVal endIndex As Long = -1) As String
    On Error GoTo ErrorHandler

    Dim i As Long
    Dim result As String
    Dim actualStart As Long
    Dim actualEnd As Long

    actualStart = IIf(startIndex = -1, LBound(arr), startIndex)
    actualEnd = IIf(endIndex = -1, UBound(arr), endIndex)

    result = ""
    For i = actualStart To actualEnd
        If i > actualStart Then
            result = result & delimiter
        End If
        result = result & arr(i)
    Next i

    JoinArray = result
    Exit Function

ErrorHandler:
    JoinArray = ""
End Function

' ============================================================================
' 日付・時刻
' ============================================================================

' 現在日時を指定フォーマットで取得
Public Function FormatNow(Optional ByVal formatStr As String = "yyyy-MM-dd HH:mm:ss") As String
    FormatNow = Format(Now, formatStr)
End Function

' タイムスタンプ文字列を取得（ファイル名用）
Public Function GetTimestamp(Optional ByVal includeTime As Boolean = True) As String
    If includeTime Then
        GetTimestamp = Format(Now, "yyyyMMdd_HHmmss")
    Else
        GetTimestamp = Format(Now, "yyyyMMdd")
    End If
End Function

' ============================================================================
' Officeアプリケーション操作
' ============================================================================

' Wordアプリケーションを取得または起動
Public Function GetWordApplication(Optional ByVal visible As Boolean = False) As Object
    On Error Resume Next
    Set GetWordApplication = GetObject(, "Word.Application")
    If GetWordApplication Is Nothing Then
        Set GetWordApplication = CreateObject("Word.Application")
    End If
    If Not GetWordApplication Is Nothing Then
        GetWordApplication.visible = visible
    End If
    On Error GoTo 0
End Function

' Word文書を開く
Public Function OpenWordDocument(ByRef wordApp As Object, _
                                  ByVal filePath As String, _
                                  Optional ByVal readOnly As Boolean = False) As Object
    On Error GoTo ErrorHandler

    If wordApp Is Nothing Then
        Set wordApp = GetWordApplication()
    End If

    Set OpenWordDocument = wordApp.Documents.Open( _
        fileName:=filePath, _
        readOnly:=readOnly)
    Exit Function

ErrorHandler:
    Set OpenWordDocument = Nothing
End Function

' Wordアプリケーションを安全に終了
Public Sub CloseWordApplication(ByRef wordApp As Object, _
                                 Optional ByRef wordDoc As Object = Nothing)
    On Error Resume Next
    If Not wordDoc Is Nothing Then
        wordDoc.Close SaveChanges:=False
        Set wordDoc = Nothing
    End If
    If Not wordApp Is Nothing Then
        wordApp.Quit
        Set wordApp = Nothing
    End If
    On Error GoTo 0
End Sub

' ============================================================================
' ログ出力
' ============================================================================

' デバッグログ出力（イミディエイトウィンドウ）
Public Sub LogDebug(ByVal message As String)
    Debug.Print "[" & FormatNow() & "] " & message
End Sub

' 区切り線付きログ出力
Public Sub LogSection(ByVal title As String)
    Debug.Print String(60, "=")
    Debug.Print title
    Debug.Print String(60, "=")
End Sub

' サブセクション区切り線付きログ出力
Public Sub LogSubSection(ByVal title As String)
    Debug.Print String(40, "-")
    Debug.Print title
End Sub

' ============================================================================
' エラーハンドリング
' ============================================================================

' エラーメッセージを整形
Public Function FormatErrorMessage(Optional ByVal prefix As String = "エラーが発生しました") As String
    FormatErrorMessage = prefix & vbCrLf & vbCrLf & _
                         "エラー番号: " & Err.Number & vbCrLf & _
                         "エラー内容: " & Err.Description
End Function

' エラーダイアログを表示
Public Sub ShowErrorDialog(Optional ByVal title As String = "エラー", _
                            Optional ByVal prefix As String = "エラーが発生しました")
    MsgBox FormatErrorMessage(prefix), vbCritical, title
End Sub

' ============================================================================
' シート操作
' ============================================================================

' シートが存在するかチェック
Public Function SheetExists(ByRef wb As Workbook, ByVal sheetName As String) As Boolean
    On Error Resume Next
    SheetExists = Not wb.Worksheets(sheetName) Is Nothing
    On Error GoTo 0
End Function

' シートを取得（存在しない場合はNothing）
Public Function GetSheet(ByRef wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheet = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

' シートを作成（存在する場合は既存シートを返す）
Public Function GetOrCreateSheet(ByRef wb As Workbook, ByVal sheetName As String) As Worksheet
    If SheetExists(wb, sheetName) Then
        Set GetOrCreateSheet = wb.Worksheets(sheetName)
    Else
        Set GetOrCreateSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        GetOrCreateSheet.Name = sheetName
    End If
End Function

' シートを削除（存在する場合のみ）
Public Sub DeleteSheetIfExists(ByRef wb As Workbook, ByVal sheetName As String)
    If SheetExists(wb, sheetName) Then
        Application.DisplayAlerts = False
        wb.Worksheets(sheetName).Delete
        Application.DisplayAlerts = True
    End If
End Sub
