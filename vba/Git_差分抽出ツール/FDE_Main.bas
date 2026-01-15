Attribute VB_Name = "FDE_Main"
Option Explicit

'==============================================================================
' Git 差分ファイル抽出ツール（VBA版） - メインモジュール
' 比較処理、ファイル抽出機能を提供
' ※バッチ版「Git_差分ファイル抽出ツール.bat」のVBA版
'==============================================================================

' ============================================================================
' 公開プロシージャ: 比較を実行
' ============================================================================
Public Sub ExecuteCompare()
    Dim wsMain As Worksheet
    Dim wsResult As Worksheet
    Dim repoPath As String
    Dim baseRef As String
    Dim targetRef As String
    Dim diffFiles() As DiffFileInfo
    Dim diffCount As Long
    Dim i As Long

    On Error GoTo ErrorHandler

    Set wsMain = ThisWorkbook.Worksheets(SHEET_MAIN)

    ' 入力値を取得
    repoPath = Trim(wsMain.Range(CELL_REPO_PATH).Value)
    baseRef = Trim(wsMain.Range(CELL_BASE_REF).Value)
    targetRef = Trim(wsMain.Range(CELL_TARGET_REF).Value)

    ' 入力チェック
    If repoPath = "" Then
        MsgBox "リポジトリパスを指定してください。", vbExclamation, "入力エラー"
        Exit Sub
    End If

    If baseRef = "" Then
        MsgBox "比較元（修正前）を指定してください。", vbExclamation, "入力エラー"
        Exit Sub
    End If

    If targetRef = "" Then
        MsgBox "比較先（修正後）を指定してください。", vbExclamation, "入力エラー"
        Exit Sub
    End If

    ' 環境変数を展開
    repoPath = ExpandEnvironmentVariables(repoPath)

    ' リポジトリ存在確認
    If Not FolderExists(repoPath) Then
        MsgBox "リポジトリフォルダが存在しません:" & vbCrLf & repoPath, vbExclamation, "入力エラー"
        Exit Sub
    End If

    ' Gitリポジトリ確認
    If Not IsGitRepository(repoPath) Then
        MsgBox "指定されたフォルダはGitリポジトリではありません:" & vbCrLf & repoPath, vbExclamation, "入力エラー"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.StatusBar = "差分ファイルを取得中..."

    ' 差分ファイルを取得
    diffCount = GetDiffFiles(repoPath, baseRef, targetRef, diffFiles)

    If diffCount = 0 Then
        Application.StatusBar = False
        Application.ScreenUpdating = True
        MsgBox "差分ファイルが見つかりませんでした。" & vbCrLf & _
               "比較対象は同じ内容です。", vbInformation, "情報"
        Exit Sub
    End If

    ' 結果シートを作成/クリア
    Set wsResult = GetOrCreateResultSheet()

    ' 結果を出力
    Application.StatusBar = "結果を出力中..."
    OutputResults wsResult, diffFiles, diffCount, repoPath, baseRef, targetRef

    ' 結果シートをアクティブに
    wsResult.Activate
    wsResult.Range("A1").Select

    Application.StatusBar = False
    Application.ScreenUpdating = True

    MsgBox "比較が完了しました。" & vbCrLf & vbCrLf & _
           "差分ファイル数: " & diffCount & " 件" & vbCrLf & vbCrLf & _
           "「差分ファイル抽出」ボタンで出力できます。", vbInformation, "比較完了"

    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub

' ============================================================================
' 公開プロシージャ: 差分ファイルを抽出
' ============================================================================
Public Sub ExtractDiffFiles()
    Dim wsMain As Worksheet
    Dim wsResult As Worksheet
    Dim repoPath As String
    Dim outputFolder As String
    Dim baseRef As String
    Dim targetRef As String
    Dim fso As Object
    Dim i As Long
    Dim lastRow As Long
    Dim relativePath As String
    Dim changeType As String
    Dim baseContent As String
    Dim targetContent As String
    Dim copyCountBefore As Long
    Dim copyCountAfter As Long

    On Error GoTo ErrorHandler

    ' シートを取得
    On Error Resume Next
    Set wsResult = ThisWorkbook.Worksheets(SHEET_RESULT)
    On Error GoTo ErrorHandler

    If wsResult Is Nothing Then
        MsgBox "先に「比較実行」を行ってください。", vbExclamation, "エラー"
        Exit Sub
    End If

    Set wsMain = ThisWorkbook.Worksheets(SHEET_MAIN)

    ' フォルダパスを取得
    repoPath = ExpandEnvironmentVariables(Trim(wsMain.Range(CELL_REPO_PATH).Value))
    outputFolder = Trim(wsMain.Range(CELL_OUTPUT_FOLDER).Value)
    baseRef = Trim(wsMain.Range(CELL_BASE_REF).Value)
    targetRef = Trim(wsMain.Range(CELL_TARGET_REF).Value)

    If outputFolder = "" Then
        MsgBox "出力先フォルダを指定してください。", vbExclamation, "入力エラー"
        Exit Sub
    End If

    ' 環境変数を展開
    outputFolder = ExpandEnvironmentVariables(outputFolder)

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 出力先フォルダを作成
    CreateFolderRecursive fso, outputFolder & "\01_修正前"
    CreateFolderRecursive fso, outputFolder & "\02_修正後"

    Application.ScreenUpdating = False
    Application.StatusBar = "差分ファイルを抽出中..."

    lastRow = wsResult.Cells(wsResult.Rows.Count, COL_RELATIVE_PATH).End(xlUp).Row
    copyCountBefore = 0
    copyCountAfter = 0

    For i = 7 To lastRow
        relativePath = wsResult.Cells(i, COL_RELATIVE_PATH).Value
        changeType = wsResult.Cells(i, COL_CHANGE_TYPE).Value

        Application.StatusBar = "抽出中: " & relativePath

        ' 修正前ファイルを取得
        If changeType = CHANGE_MODIFIED Or changeType = CHANGE_DELETED Then
            baseContent = GetFileContent(repoPath, baseRef, relativePath)
            If baseContent <> "" Then
                WriteFileWithFolder fso, outputFolder & "\01_修正前\" & relativePath, baseContent
                copyCountBefore = copyCountBefore + 1
            End If
        End If

        ' 修正後ファイルを取得
        If changeType = CHANGE_MODIFIED Or changeType = CHANGE_ADDED Then
            targetContent = GetFileContent(repoPath, targetRef, relativePath)
            If targetContent <> "" Then
                WriteFileWithFolder fso, outputFolder & "\02_修正後\" & relativePath, targetContent
                copyCountAfter = copyCountAfter + 1
            End If
        End If
    Next i

    ' レポートを出力
    OutputDiffReport fso, outputFolder, wsResult, baseRef, targetRef

    Application.StatusBar = False
    Application.ScreenUpdating = True

    MsgBox "差分ファイルの抽出が完了しました。" & vbCrLf & vbCrLf & _
           "01_修正前: " & copyCountBefore & " 件" & vbCrLf & _
           "02_修正後: " & copyCountAfter & " 件" & vbCrLf & vbCrLf & _
           "出力先: " & outputFolder, vbInformation, "抽出完了"

    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub

' ============================================================================
' 公開プロシージャ: ブランチ選択（比較元）
' ============================================================================
Public Sub SelectBaseRef()
    SelectRef CELL_BASE_REF, "比較元（修正前）"
End Sub

' ============================================================================
' 公開プロシージャ: ブランチ選択（比較先）
' ============================================================================
Public Sub SelectTargetRef()
    SelectRef CELL_TARGET_REF, "比較先（修正後）"
End Sub

' ============================================================================
' 公開プロシージャ: リポジトリパス選択
' ============================================================================
Public Sub SelectRepoPath()
    Dim wsMain As Worksheet
    Dim selectedFolder As String

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Gitリポジトリフォルダを選択"
        .AllowMultiSelect = False

        If .Show = -1 Then
            selectedFolder = .SelectedItems(1)
            Set wsMain = ThisWorkbook.Worksheets(SHEET_MAIN)
            wsMain.Range(CELL_REPO_PATH).Value = selectedFolder
        End If
    End With
End Sub

' ============================================================================
' 公開プロシージャ: 出力先フォルダ選択
' ============================================================================
Public Sub SelectOutputFolder()
    Dim wsMain As Worksheet
    Dim selectedFolder As String

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "出力先フォルダを選択"
        .AllowMultiSelect = False

        If .Show = -1 Then
            selectedFolder = .SelectedItems(1)
            Set wsMain = ThisWorkbook.Worksheets(SHEET_MAIN)
            wsMain.Range(CELL_OUTPUT_FOLDER).Value = selectedFolder
        End If
    End With
End Sub

' ============================================================================
' 内部関数: ブランチ選択ダイアログ
' ============================================================================
Private Sub SelectRef(ByVal targetCell As String, ByVal title As String)
    Dim wsMain As Worksheet
    Dim repoPath As String
    Dim branches() As String
    Dim branchCount As Long
    Dim selectedRef As String
    Dim msg As String
    Dim i As Long

    Set wsMain = ThisWorkbook.Worksheets(SHEET_MAIN)
    repoPath = ExpandEnvironmentVariables(Trim(wsMain.Range(CELL_REPO_PATH).Value))

    If repoPath = "" Then
        MsgBox "先にリポジトリパスを指定してください。", vbExclamation, "入力エラー"
        Exit Sub
    End If

    If Not FolderExists(repoPath) Then
        MsgBox "リポジトリフォルダが存在しません。", vbExclamation, "入力エラー"
        Exit Sub
    End If

    ' ブランチ一覧を取得
    branchCount = GetBranches(repoPath, branches)

    If branchCount = 0 Then
        MsgBox "ブランチが見つかりませんでした。", vbExclamation, "エラー"
        Exit Sub
    End If

    ' 選択メッセージを作成
    msg = title & "を選択してください:" & vbCrLf & vbCrLf
    For i = 1 To branchCount
        msg = msg & i & ". " & branches(i) & vbCrLf
    Next i
    msg = msg & vbCrLf & "番号を入力 (1-" & branchCount & "):"

    Dim input As String
    input = InputBox(msg, title & "の選択", "1")

    If input = "" Then Exit Sub

    If IsNumeric(input) Then
        Dim num As Long
        num = CLng(input)
        If num >= 1 And num <= branchCount Then
            selectedRef = branches(num)
            wsMain.Range(targetCell).Value = selectedRef
        Else
            MsgBox "無効な番号です。", vbExclamation, "入力エラー"
        End If
    Else
        MsgBox "数字を入力してください。", vbExclamation, "入力エラー"
    End If
End Sub

' ============================================================================
' 内部関数: Gitコマンドを実行
' ============================================================================
Private Function RunGitCommand(ByVal repoPath As String, ByVal gitArgs As String) As String
    Dim wsh As Object
    Dim execObj As Object
    Dim output As String
    Dim cmd As String

    Set wsh = CreateObject("WScript.Shell")

    cmd = "cmd /c cd /d """ & repoPath & """ && " & GIT_COMMAND & " " & gitArgs

    On Error Resume Next
    Set execObj = wsh.Exec(cmd)

    If Not execObj Is Nothing Then
        Do While execObj.Status = 0
            DoEvents
        Loop
        output = execObj.StdOut.ReadAll
    End If
    On Error GoTo 0

    RunGitCommand = output
End Function

' ============================================================================
' 内部関数: Gitリポジトリかどうか確認
' ============================================================================
Private Function IsGitRepository(ByVal repoPath As String) As Boolean
    Dim output As String
    output = RunGitCommand(repoPath, "rev-parse --git-dir")
    IsGitRepository = (Len(Trim(output)) > 0)
End Function

' ============================================================================
' 内部関数: ブランチ一覧を取得
' ============================================================================
Private Function GetBranches(ByVal repoPath As String, ByRef branches() As String) As Long
    Dim output As String
    Dim lines() As String
    Dim i As Long
    Dim count As Long
    Dim branchName As String

    output = RunGitCommand(repoPath, "branch --format=""%(refname:short)""")

    If Len(Trim(output)) = 0 Then
        GetBranches = 0
        Exit Function
    End If

    lines = Split(output, vbLf)
    count = 0

    For i = LBound(lines) To UBound(lines)
        branchName = Trim(lines(i))
        If Len(branchName) > 0 Then
            count = count + 1
        End If
    Next i

    If count = 0 Then
        GetBranches = 0
        Exit Function
    End If

    ReDim branches(1 To count)
    count = 0

    For i = LBound(lines) To UBound(lines)
        branchName = Trim(lines(i))
        If Len(branchName) > 0 Then
            count = count + 1
            branches(count) = branchName
        End If
    Next i

    GetBranches = count
End Function

' ============================================================================
' 内部関数: 差分ファイルを取得
' ============================================================================
Private Function GetDiffFiles(ByVal repoPath As String, ByVal baseRef As String, ByVal targetRef As String, _
                              ByRef diffFiles() As DiffFileInfo) As Long
    Dim output As String
    Dim lines() As String
    Dim i As Long
    Dim count As Long
    Dim fileLine As String
    Dim parts() As String

    output = RunGitCommand(repoPath, "diff --name-status """ & baseRef & """..""" & targetRef & """")

    If Len(Trim(output)) = 0 Then
        GetDiffFiles = 0
        Exit Function
    End If

    lines = Split(output, vbLf)
    count = 0

    For i = LBound(lines) To UBound(lines)
        fileLine = Trim(lines(i))
        If Len(fileLine) > 0 Then
            count = count + 1
        End If
    Next i

    If count = 0 Then
        GetDiffFiles = 0
        Exit Function
    End If

    ReDim diffFiles(1 To count)
    count = 0

    For i = LBound(lines) To UBound(lines)
        fileLine = Trim(lines(i))
        If Len(fileLine) > 0 Then
            count = count + 1

            parts = Split(fileLine, vbTab)
            If UBound(parts) >= 1 Then
                diffFiles(count).Status = Trim(parts(0))
                diffFiles(count).RelativePath = Replace(Trim(parts(1)), "/", "\")
                diffFiles(count).FileName = GetFileName(diffFiles(count).RelativePath)

                Select Case Left(diffFiles(count).Status, 1)
                    Case STATUS_ADDED
                        diffFiles(count).ChangeType = CHANGE_ADDED
                        diffFiles(count).BaseExists = False
                        diffFiles(count).TargetExists = True
                    Case STATUS_MODIFIED
                        diffFiles(count).ChangeType = CHANGE_MODIFIED
                        diffFiles(count).BaseExists = True
                        diffFiles(count).TargetExists = True
                    Case STATUS_DELETED
                        diffFiles(count).ChangeType = CHANGE_DELETED
                        diffFiles(count).BaseExists = True
                        diffFiles(count).TargetExists = False
                    Case STATUS_RENAMED
                        diffFiles(count).ChangeType = CHANGE_RENAMED
                        diffFiles(count).BaseExists = True
                        diffFiles(count).TargetExists = True
                    Case Else
                        diffFiles(count).ChangeType = diffFiles(count).Status
                        diffFiles(count).BaseExists = True
                        diffFiles(count).TargetExists = True
                End Select
            End If
        End If
    Next i

    GetDiffFiles = count
End Function

' ============================================================================
' 内部関数: git show でファイル内容を取得
' ============================================================================
Private Function GetFileContent(ByVal repoPath As String, ByVal ref As String, ByVal filePath As String) As String
    Dim wsh As Object
    Dim execObj As Object
    Dim output As String
    Dim cmd As String
    Dim gitPath As String

    gitPath = Replace(filePath, "\", "/")

    Set wsh = CreateObject("WScript.Shell")
    cmd = "cmd /c cd /d """ & repoPath & """ && " & GIT_COMMAND & " show """ & ref & ":" & gitPath & """"

    On Error Resume Next
    Set execObj = wsh.Exec(cmd)

    If Not execObj Is Nothing Then
        Do While execObj.Status = 0
            DoEvents
        Loop
        output = execObj.StdOut.ReadAll
    End If
    On Error GoTo 0

    GetFileContent = output
End Function

' ============================================================================
' 内部関数: 結果シートを取得または作成
' ============================================================================
Private Function GetOrCreateResultSheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_RESULT)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SHEET_RESULT
    Else
        ws.Cells.Clear
    End If

    Set GetOrCreateResultSheet = ws
End Function

' ============================================================================
' 内部関数: 結果を出力
' ============================================================================
Private Sub OutputResults(ByRef ws As Worksheet, ByRef diffFiles() As DiffFileInfo, ByVal diffCount As Long, _
                          ByVal repoPath As String, ByVal baseRef As String, ByVal targetRef As String)
    Dim i As Long
    Dim row As Long

    With ws
        .Range("A1").Value = "Git 差分ファイル比較結果"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14

        .Range("A2").Value = "実行日時:"
        .Range("B2").Value = Now
        .Range("B2").NumberFormat = "yyyy/mm/dd hh:mm:ss"

        .Range("A3").Value = "リポジトリ:"
        .Range("B3").Value = repoPath

        .Range("A4").Value = "比較元:"
        .Range("B4").Value = baseRef

        .Range("A5").Value = "比較先:"
        .Range("B5").Value = targetRef

        row = 6
        .Cells(row, COL_DIFF_MARK).Value = "抽出"
        .Cells(row, COL_RELATIVE_PATH).Value = "相対パス"
        .Cells(row, COL_FILE_NAME).Value = "ファイル名"
        .Cells(row, COL_STATUS).Value = "状態"
        .Cells(row, COL_BASE_EXISTS).Value = "修正前"
        .Cells(row, COL_BASE_SIZE).Value = "修正前サイズ"
        .Cells(row, COL_TARGET_EXISTS).Value = "修正後"
        .Cells(row, COL_TARGET_SIZE).Value = "修正後サイズ"
        .Cells(row, COL_CHANGE_TYPE).Value = "変更種別"

        With .Range(.Cells(row, 1), .Cells(row, 9))
            .Font.Bold = True
            .Interior.Color = RGB(68, 114, 196)
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
        End With

        For i = 1 To diffCount
            row = row + 1

            .Cells(row, COL_DIFF_MARK).Value = "*"
            .Cells(row, COL_DIFF_MARK).Font.Bold = True
            .Cells(row, COL_DIFF_MARK).Font.Color = RGB(255, 0, 0)
            .Cells(row, COL_DIFF_MARK).HorizontalAlignment = xlCenter

            .Cells(row, COL_RELATIVE_PATH).Value = diffFiles(i).RelativePath
            .Cells(row, COL_FILE_NAME).Value = diffFiles(i).FileName
            .Cells(row, COL_STATUS).Value = diffFiles(i).Status

            If diffFiles(i).BaseExists Then
                .Cells(row, COL_BASE_EXISTS).Value = "○"
            Else
                .Cells(row, COL_BASE_EXISTS).Value = "-"
            End If
            .Cells(row, COL_BASE_EXISTS).HorizontalAlignment = xlCenter

            If diffFiles(i).TargetExists Then
                .Cells(row, COL_TARGET_EXISTS).Value = "○"
            Else
                .Cells(row, COL_TARGET_EXISTS).Value = "-"
            End If
            .Cells(row, COL_TARGET_EXISTS).HorizontalAlignment = xlCenter

            .Cells(row, COL_CHANGE_TYPE).Value = diffFiles(i).ChangeType

            Select Case diffFiles(i).ChangeType
                Case CHANGE_MODIFIED
                    .Range(.Cells(row, 1), .Cells(row, 9)).Interior.Color = RGB(255, 255, 200)
                Case CHANGE_ADDED
                    .Range(.Cells(row, 1), .Cells(row, 9)).Interior.Color = RGB(230, 255, 230)
                Case CHANGE_DELETED
                    .Range(.Cells(row, 1), .Cells(row, 9)).Interior.Color = RGB(255, 230, 230)
            End Select
        Next i

        .Columns(COL_DIFF_MARK).ColumnWidth = 5
        .Columns(COL_RELATIVE_PATH).ColumnWidth = 50
        .Columns(COL_FILE_NAME).ColumnWidth = 25
        .Columns(COL_STATUS).ColumnWidth = 6
        .Columns(COL_BASE_EXISTS).ColumnWidth = 8
        .Columns(COL_BASE_SIZE).ColumnWidth = 12
        .Columns(COL_TARGET_EXISTS).ColumnWidth = 8
        .Columns(COL_TARGET_SIZE).ColumnWidth = 12
        .Columns(COL_CHANGE_TYPE).ColumnWidth = 10

        .Range(.Cells(6, 1), .Cells(row, 9)).AutoFilter
    End With
End Sub

' ============================================================================
' 内部関数: 差異レポートを出力
' ============================================================================
Private Sub OutputDiffReport(ByRef fso As Object, ByVal outputFolder As String, ByRef wsResult As Worksheet, _
                             ByVal baseRef As String, ByVal targetRef As String)
    Dim reportPath As String
    Dim ts As Object
    Dim i As Long
    Dim lastRow As Long
    Dim relativePath As String
    Dim changeType As String

    reportPath = outputFolder & "\diff_report.txt"

    Set ts = fso.CreateTextFile(reportPath, True, True)

    ts.WriteLine "Git 差分ファイルレポート"
    ts.WriteLine "========================"
    ts.WriteLine ""
    ts.WriteLine "実行日時    : " & Format(Now, "yyyy/mm/dd hh:mm:ss")
    ts.WriteLine "リポジトリ  : " & wsResult.Range("B3").Value
    ts.WriteLine "比較元      : " & baseRef
    ts.WriteLine "比較先      : " & targetRef
    ts.WriteLine ""
    ts.WriteLine "差分ファイル一覧"
    ts.WriteLine "----------------"

    lastRow = wsResult.Cells(wsResult.Rows.Count, COL_RELATIVE_PATH).End(xlUp).Row

    For i = 7 To lastRow
        relativePath = wsResult.Cells(i, COL_RELATIVE_PATH).Value
        changeType = wsResult.Cells(i, COL_CHANGE_TYPE).Value
        ts.WriteLine "[" & changeType & "] " & relativePath
    Next i

    ts.Close
End Sub

' ============================================================================
' ユーティリティ関数
' ============================================================================
Private Function ExpandEnvironmentVariables(ByVal path As String) As String
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    ExpandEnvironmentVariables = wsh.ExpandEnvironmentStrings(path)
End Function

Private Function FolderExists(ByVal folderPath As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FolderExists = fso.FolderExists(folderPath)
End Function

Private Function GetFileName(ByVal filePath As String) As String
    Dim pos As Long
    pos = InStrRev(filePath, "\")
    If pos > 0 Then
        GetFileName = Mid(filePath, pos + 1)
    Else
        pos = InStrRev(filePath, "/")
        If pos > 0 Then
            GetFileName = Mid(filePath, pos + 1)
        Else
            GetFileName = filePath
        End If
    End If
End Function

Private Sub CreateFolderRecursive(ByRef fso As Object, ByVal folderPath As String)
    Dim parentPath As String

    If fso.FolderExists(folderPath) Then Exit Sub

    parentPath = fso.GetParentFolderName(folderPath)

    If Not fso.FolderExists(parentPath) Then
        CreateFolderRecursive fso, parentPath
    End If

    fso.CreateFolder folderPath
End Sub

Private Sub WriteFileWithFolder(ByRef fso As Object, ByVal destPath As String, ByVal content As String)
    Dim destFolder As String
    Dim ts As Object

    destFolder = fso.GetParentFolderName(destPath)

    If Not fso.FolderExists(destFolder) Then
        CreateFolderRecursive fso, destFolder
    End If

    Set ts = fso.CreateTextFile(destPath, True, True)
    ts.Write content
    ts.Close
End Sub

