Attribute VB_Name = "GLV_Parser"
Option Explicit

'==============================================================================
' Git Log 可視化ツール - パーサーモジュール
' Git出力のパース、データ取得機能を提供
'==============================================================================

'==============================================================================
' 環境変数を展開する (%USERNAME% など)
'==============================================================================
Public Function ExpandEnvironmentVariables(ByVal path As String) As String
    Dim result As String
    Dim startPos As Long
    Dim endPos As Long
    Dim varName As String
    Dim varValue As String

    result = path

    ' %VAR% 形式の環境変数をすべて展開
    startPos = InStr(result, "%")
    Do While startPos > 0
        endPos = InStr(startPos + 1, result, "%")
        If endPos > startPos + 1 Then
            varName = Mid(result, startPos + 1, endPos - startPos - 1)
            varValue = Environ(varName)
            If Len(varValue) > 0 Then
                result = Left(result, startPos - 1) & varValue & Mid(result, endPos + 1)
            Else
                ' 環境変数が見つからない場合はスキップして次を探す
                startPos = endPos
            End If
            startPos = InStr(startPos + Len(varValue), result, "%")
        Else
            ' 閉じる % がない場合は終了
            Exit Do
        End If
    Loop

    ExpandEnvironmentVariables = result
End Function

'==============================================================================
' メインシートから設定値を取得
'==============================================================================
Public Function GetRepoPathFromMainSheet() As String
    Dim rawPath As String

    On Error Resume Next
    rawPath = ThisWorkbook.Sheets(SHEET_MAIN).Range(CELL_REPO_PATH).Value
    If Err.Number <> 0 Then
        GetRepoPathFromMainSheet = ""
        Exit Function
    End If
    On Error GoTo 0

    ' 環境変数を展開
    GetRepoPathFromMainSheet = ExpandEnvironmentVariables(rawPath)
End Function

Public Function GetCommitCountFromMainSheet() As Long
    On Error Resume Next
    GetCommitCountFromMainSheet = CLng(ThisWorkbook.Sheets(SHEET_MAIN).Range(CELL_COMMIT_COUNT).Value)
    If Err.Number <> 0 Or GetCommitCountFromMainSheet <= 0 Then
        GetCommitCountFromMainSheet = 100
    End If
    On Error GoTo 0
End Function

'==============================================================================
' Gitリポジトリかどうかを確認
'==============================================================================
Public Function IsGitRepository(ByVal repoPath As String) As Boolean
    Dim wsh As Object
    Dim execObj As Object
    Dim command As String

    Set wsh = CreateObject("WScript.Shell")
    command = "cmd /c cd /d """ & repoPath & """ && " & GIT_COMMAND & " rev-parse --git-dir >nul 2>&1"
    Set execObj = wsh.exec(command)

    Do While execObj.Status = 0
        DoEvents
    Loop

    IsGitRepository = (execObj.ExitCode = 0)
End Function

'==============================================================================
' コマンドを実行して結果を返す
'==============================================================================
Public Function RunCommand(ByVal cmd As String) As String
    Dim wsh As Object
    Dim fso As Object
    Dim tempFile As String
    Dim output As String
    Dim stream As Object

    Set wsh = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")

    tempFile = fso.GetSpecialFolder(2) & "\cmd_" & fso.GetTempName & ".txt"

    ' コマンドを実行して結果を一時ファイルに出力
    wsh.Run "cmd /c chcp 65001 >nul && " & cmd & " > """ & tempFile & """ 2>&1", 0, True

    ' 結果を読み込み
    If fso.FileExists(tempFile) Then
        On Error Resume Next
        Set stream = CreateObject("ADODB.Stream")
        If stream Is Nothing Then
            output = fso.OpenTextFile(tempFile, 1, False, -1).ReadAll
        Else
            stream.Type = 2
            stream.Charset = "UTF-8"
            stream.Open
            stream.LoadFromFile tempFile
            output = stream.ReadText
            stream.Close
            Set stream = Nothing
        End If
        On Error GoTo 0

        On Error Resume Next
        fso.DeleteFile tempFile
        On Error GoTo 0
    Else
        output = ""
    End If

    RunCommand = output

    Set fso = Nothing
    Set wsh = Nothing
End Function

'==============================================================================
' Git Log を取得
'==============================================================================
Public Function GetGitLog(ByVal repoPath As String, ByVal maxCount As Long) As CommitInfo()
    Dim wsh As Object
    Dim fso As Object
    Dim command As String
    Dim output As String
    Dim commits() As CommitInfo
    Dim i As Long
    Dim commitIndex As Long
    Dim parts() As String
    Dim tempFile As String
    Dim stream As Object
    Dim commitBlocks() As String
    Dim block As String
    Dim headerLine As String
    Dim bodyLines As String

    Set wsh = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")

    tempFile = fso.GetSpecialFolder(2) & "\gitlog_" & fso.GetTempName & ".txt"

    ' コミット区切りマーカーを使用し、メッセージ全文（%B）を取得
    command = "cmd /c chcp 65001 >nul && cd /d """ & repoPath & """ && " & _
              GIT_COMMAND & " log --all -n " & maxCount & _
              " --pretty=format:""<<<COMMIT>>>%h|%H|%P|%an|%ae|%ai|%d<<<MSG>>>%B<<<END>>>"" --numstat > """ & tempFile & """ 2>&1"

    wsh.Run command, 0, True

    If Not fso.FileExists(tempFile) Then
        ReDim commits(0 To 0)
        GetGitLog = commits
        Exit Function
    End If

    On Error Resume Next
    Set stream = CreateObject("ADODB.Stream")
    If stream Is Nothing Then
        output = fso.OpenTextFile(tempFile, 1, False, -1).ReadAll
    Else
        stream.Type = 2
        stream.Charset = "UTF-8"
        stream.Open
        stream.LoadFromFile tempFile
        output = stream.ReadText
        stream.Close
        Set stream = Nothing
    End If
    On Error GoTo 0

    On Error Resume Next
    fso.DeleteFile tempFile
    On Error GoTo 0

    If Len(output) = 0 Then
        ReDim commits(0 To 0)
        GetGitLog = commits
        Exit Function
    End If

    output = Replace(output, vbCrLf, vbLf)
    output = Replace(output, vbCr, vbLf)

    ' <<<COMMIT>>> でコミットブロックを分割
    commitBlocks = Split(output, "<<<COMMIT>>>")

    commitIndex = 0
    ReDim commits(0 To maxCount - 1)

    For i = 1 To UBound(commitBlocks)  ' 最初の空要素をスキップ
        block = commitBlocks(i)

        ' <<<MSG>>> でヘッダーとメッセージを分離
        Dim msgPos As Long
        Dim endPos As Long
        msgPos = InStr(block, "<<<MSG>>>")
        endPos = InStr(block, "<<<END>>>")

        If msgPos > 0 And endPos > msgPos Then
            headerLine = Left(block, msgPos - 1)
            bodyLines = Mid(block, msgPos + 9, endPos - msgPos - 9)

            ' ヘッダーをパース
            parts = Split(headerLine, "|")

            If UBound(parts) >= 5 Then
                With commits(commitIndex)
                    .Hash = parts(0)
                    .FullHash = parts(1)
                    .ParentHashes = parts(2)
                    If Len(Trim(parts(2))) = 0 Then
                        .ParentCount = 0
                    Else
                        .ParentCount = UBound(Split(Trim(parts(2)), " ")) + 1
                    End If
                    .Author = parts(3)
                    .AuthorEmail = parts(4)
                    .CommitDate = ParseGitDate(parts(5))
                    If UBound(parts) >= 6 Then
                        .RefNames = Trim(Replace(Replace(parts(6), "(", ""), ")", ""))
                    Else
                        .RefNames = ""
                    End If

                    ' メッセージ全文（改行を保持）
                    .Subject = Trim(bodyLines)

                    ' numstat を解析（<<<END>>>以降）
                    .FilesChanged = 0
                    .Insertions = 0
                    .Deletions = 0

                    Dim afterEnd As String
                    afterEnd = Mid(block, endPos + 9)
                    Dim statLines() As String
                    statLines = Split(afterEnd, vbLf)

                    Dim j As Long
                    For j = 0 To UBound(statLines)
                        Dim statLine As String
                        statLine = Trim(statLines(j))

                        If Len(statLine) > 0 And InStr(statLine, vbTab) > 0 Then
                            Dim statParts() As String
                            statParts = Split(statLine, vbTab)

                            If UBound(statParts) >= 2 Then
                                .FilesChanged = .FilesChanged + 1
                                If IsNumeric(statParts(0)) Then
                                    .Insertions = .Insertions + CLng(statParts(0))
                                End If
                                If IsNumeric(statParts(1)) Then
                                    .Deletions = .Deletions + CLng(statParts(1))
                                End If
                            End If
                        End If
                    Next j
                End With

                commitIndex = commitIndex + 1
                If commitIndex >= maxCount Then Exit For
            End If
        End If
    Next i

    If commitIndex > 0 Then
        ReDim Preserve commits(0 To commitIndex - 1)
    Else
        ReDim commits(0 To 0)
    End If

    GetGitLog = commits
End Function

'==============================================================================
' Git の日付文字列をDateに変換
'==============================================================================
Public Function ParseGitDate(ByVal dateStr As String) As Date
    On Error Resume Next
    ParseGitDate = CDate(Left(dateStr, 19))
    If Err.Number <> 0 Then
        ParseGitDate = Now
        Err.Clear
    End If
    On Error GoTo 0
End Function

'==============================================================================
' 現在のブランチを取得
'==============================================================================
Public Function GetCurrentBranch(ByVal repoPath As String) As String
    Dim cmd As String
    Dim result As String

    cmd = "cd /d """ & repoPath & """ && " & GIT_COMMAND & " branch --show-current"
    result = RunCommand(cmd)

    GetCurrentBranch = Trim(Replace(Replace(result, vbCr, ""), vbLf, ""))
End Function

'==============================================================================
' ブランチ一覧を取得（ローカルブランチ）
'==============================================================================
Public Function GetBranchList(ByVal repoPath As String) As String()
    Dim cmd As String
    Dim result As String
    Dim lines() As String
    Dim branches() As String
    Dim i As Long
    Dim branchCount As Long
    Dim branchName As String

    cmd = "cd /d """ & repoPath & """ && " & GIT_COMMAND & " branch"
    result = RunCommand(cmd)

    ' 結果を行に分割
    lines = Split(result, vbLf)
    branchCount = 0
    ReDim branches(0 To UBound(lines))

    For i = 0 To UBound(lines)
        branchName = Trim(Replace(lines(i), vbCr, ""))
        ' 先頭の * を除去（現在のブランチを示す）
        If Left(branchName, 2) = "* " Then
            branchName = Mid(branchName, 3)
        End If
        branchName = Trim(branchName)

        If Len(branchName) > 0 Then
            branches(branchCount) = branchName
            branchCount = branchCount + 1
        End If
    Next i

    ' 配列サイズを調整
    If branchCount > 0 Then
        ReDim Preserve branches(0 To branchCount - 1)
    Else
        ReDim branches(0 To 0)
        branches(0) = ""
    End If

    GetBranchList = branches
End Function

'==============================================================================
' ブランチを切り替え
'==============================================================================
Public Function SwitchBranch(ByVal repoPath As String, ByVal branchName As String) As Boolean
    Dim cmd As String
    Dim result As String

    cmd = "cd /d """ & repoPath & """ && " & GIT_COMMAND & " checkout """ & branchName & """ 2>&1"
    result = RunCommand(cmd)

    ' エラーチェック
    If InStr(result, "error:") > 0 Or InStr(result, "fatal:") > 0 Then
        SwitchBranch = False
    Else
        SwitchBranch = True
    End If
End Function

