'==============================================================================
' Git Log 可視化ツール
' モジュール名: GitLogVisualizer
'==============================================================================
' 概要:
'   Excelから実行し、Gitリポジトリのコミット履歴を取得して、
'   表形式で視覚化するツールです。
'
' 機能:
'   - コミット履歴の一覧表示（詳細情報付き）
'   - ブランチグラフによる可視化
'   - メインシートで設定を入力可能
'
' 必要な環境:
'   - Microsoft Excel 2010以降
'   - Git がインストールされており、パスが通っていること
'
' 作成日: 2025-12-07
' 更新日: 2025-12-17 - メインシート追加、シート名日本語化、履歴シート改善
'==============================================================================

Option Explicit

'==============================================================================
' 定数
'==============================================================================
' Gitコマンドのパス（通常は "git" でOK。パスが通っていない場合はフルパス指定）
Private Const GIT_COMMAND As String = "git"

' シート名
Private Const SHEET_MAIN As String = "メイン"
Private Const SHEET_DASHBOARD As String = "ダッシュボード"
Private Const SHEET_HISTORY As String = "コミット履歴"
Private Const SHEET_BRANCH_GRAPH As String = "ブランチグラフ"

' メインシートのセル位置
Private Const CELL_REPO_PATH As String = "D8"
Private Const CELL_COMMIT_COUNT As String = "D10"

'==============================================================================
' データ構造
'==============================================================================
Private Type CommitInfo
    Hash As String          ' コミットハッシュ（短縮）
    FullHash As String      ' コミットハッシュ（フル）
    Author As String        ' 作者名
    AuthorEmail As String   ' 作者メール
    CommitDate As Date      ' コミット日時
    Subject As String       ' コミットメッセージ（件名）
    RefNames As String      ' ブランチ・タグ名
    ParentHashes As String  ' 親コミットハッシュ（スペース区切り）
    ParentCount As Long     ' 親コミット数（0=初期, 1=通常, 2+=マージ）
    FilesChanged As Long    ' 変更ファイル数
    Insertions As Long      ' 追加行数
    Deletions As Long       ' 削除行数
End Type

'==============================================================================
' メインシート初期化
'==============================================================================
Public Sub InitializeGitLogVisualizer()
    Dim ws As Worksheet

    On Error Resume Next
    Application.DisplayAlerts = False

    ' 既存のメインシートがあれば削除
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = SHEET_MAIN Then
            ws.Delete
            Exit For
        End If
    Next ws

    Application.DisplayAlerts = True
    On Error GoTo 0

    ' 新しいシートを作成
    Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
    ws.Name = SHEET_MAIN

    ' シートを初期化
    FormatMainSheet ws

    MsgBox "メインシートを初期化しました。" & vbCrLf & vbCrLf & _
           "リポジトリパスと取得件数を設定して、" & vbCrLf & _
           "「実行」ボタンをクリックしてください。", vbInformation, "初期化完了"
End Sub

'==============================================================================
' メインシートのフォーマット
'==============================================================================
Private Sub FormatMainSheet(ByRef ws As Worksheet)
    Dim shp As Shape

    Application.ScreenUpdating = False

    With ws
        ' 全体の背景色を白に
        .Cells.Interior.Color = RGB(255, 255, 255)

        ' =================================================================
        ' タイトルエリア (行1-3)
        ' =================================================================
        .Range("B2:G2").Merge
        .Range("B2").Value = "Git Log 可視化ツール"
        With .Range("B2")
            .Font.Name = "Meiryo UI"
            .Font.Size = 20
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        .Range("B2:G3").Interior.Color = RGB(68, 114, 196)
        .Rows(2).RowHeight = 40
        .Rows(3).RowHeight = 5

        ' =================================================================
        ' 説明エリア (行5)
        ' =================================================================
        .Range("B5:G5").Merge
        .Range("B5").Value = "Gitリポジトリのコミット履歴を取得して視覚化します。"
        With .Range("B5")
            .Font.Name = "Meiryo UI"
            .Font.Size = 11
            .Font.Color = RGB(64, 64, 64)
        End With

        ' =================================================================
        ' 設定セクション (行7-12)
        ' =================================================================
        .Range("B7:G7").Merge
        .Range("B7").Value = "設定"
        With .Range("B7")
            .Font.Name = "Meiryo UI"
            .Font.Size = 14
            .Font.Bold = True
            .Font.Color = RGB(68, 114, 196)
        End With
        With .Range("B7:G7").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(68, 114, 196)
            .Weight = xlMedium
        End With

        ' リポジトリパス
        .Range("B8").Value = "リポジトリパス:"
        With .Range("B8")
            .Font.Name = "Meiryo UI"
            .Font.Size = 11
            .Font.Bold = True
        End With

        .Range("D8:G8").Merge
        .Range("D8").Value = "C:\Users\%USERNAME%\source\Git\project"
        With .Range("D8:G8")
            .Interior.Color = RGB(255, 255, 230)
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
        End With
        With .Range("D8:G8").Borders
            .LineStyle = xlContinuous
            .Color = RGB(200, 200, 200)
        End With

        ' 環境変数の説明
        .Range("D9:G9").Merge
        .Range("D9").Value = "※ %USERNAME% などの環境変数が使用可能"
        With .Range("D9")
            .Font.Name = "Meiryo UI"
            .Font.Size = 9
            .Font.Color = RGB(100, 100, 100)
            .Font.Italic = True
        End With

        ' 取得コミット数
        .Range("B10").Value = "取得件数:"
        With .Range("B10")
            .Font.Name = "Meiryo UI"
            .Font.Size = 11
            .Font.Bold = True
        End With

        .Range("D10").Value = 100
        With .Range("D10")
            .Interior.Color = RGB(255, 255, 230)
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
            .NumberFormat = "#,##0"
            .HorizontalAlignment = xlCenter
        End With
        With .Range("D10").Borders
            .LineStyle = xlContinuous
            .Color = RGB(200, 200, 200)
        End With

        .Range("E10:G10").Merge
        .Range("E10").Value = "件（最新から取得）"
        With .Range("E10")
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
            .Font.Color = RGB(100, 100, 100)
        End With

        ' =================================================================
        ' ボタンエリア (行13-14)
        ' =================================================================
        .Rows(12).RowHeight = 15
        .Rows(13).RowHeight = 50

        ' 実行ボタン - 角丸四角形
        Dim btnLeft As Double
        Dim btnTop As Double
        btnLeft = .Range("D13").Left
        btnTop = .Range("D13").Top + 5

        Set shp = .Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, 120, 40)
        With shp
            .Name = "btnExecute"
            .Placement = xlFreeFloating
            .Fill.ForeColor.RGB = RGB(76, 175, 80)
            .Line.ForeColor.RGB = RGB(56, 142, 60)
            .Line.Weight = 2
            .TextFrame2.TextRange.Characters.Text = "実行"
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextFrame2.TextRange.Font.Size = 14
            .TextFrame2.TextRange.Font.Bold = msoTrue
            .TextFrame2.TextRange.Font.Name = "Meiryo UI"
            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .TextFrame2.VerticalAnchor = msoAnchorMiddle
            .OnAction = "VisualizeGitLog"
        End With

        ' =================================================================
        ' 出力シートセクション (行16-20)
        ' =================================================================
        .Range("B16:G16").Merge
        .Range("B16").Value = "出力シート"
        With .Range("B16")
            .Font.Name = "Meiryo UI"
            .Font.Size = 14
            .Font.Bold = True
            .Font.Color = RGB(68, 114, 196)
        End With
        With .Range("B16:G16").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(68, 114, 196)
            .Weight = xlMedium
        End With

        .Range("B18").Value = "ダッシュボード"
        With .Range("B18")
            .Font.Name = "Meiryo UI"
            .Font.Size = 11
            .Font.Bold = True
            .Font.Color = RGB(68, 114, 196)
        End With
        .Range("C18:G18").Merge
        .Range("C18").Value = "サマリー情報（総コミット数、作者数、変更量、作者別統計）"
        With .Range("C18")
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
        End With

        .Range("B19").Value = "コミット履歴"
        With .Range("B19")
            .Font.Name = "Meiryo UI"
            .Font.Size = 11
            .Font.Bold = True
            .Font.Color = RGB(68, 114, 196)
        End With
        .Range("C19:G19").Merge
        .Range("C19").Value = "コミット履歴の詳細一覧（ハッシュ、作者、日時、メッセージ、変更量等）"
        With .Range("C19")
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
        End With

        .Range("B20").Value = "ブランチグラフ"
        With .Range("B20")
            .Font.Name = "Meiryo UI"
            .Font.Size = 11
            .Font.Bold = True
            .Font.Color = RGB(68, 114, 196)
        End With
        .Range("C20:G20").Merge
        .Range("C20").Value = "ブランチ構造を視覚化（コミットノードと接続線）"
        With .Range("C20")
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
        End With

        ' =================================================================
        ' 色凡例セクション (行23-29)
        ' =================================================================
        .Range("B23:G23").Merge
        .Range("B23").Value = "ブランチグラフの色凡例"
        With .Range("B23")
            .Font.Name = "Meiryo UI"
            .Font.Size = 14
            .Font.Bold = True
            .Font.Color = RGB(68, 114, 196)
        End With
        With .Range("B23:G23").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(68, 114, 196)
            .Weight = xlMedium
        End With

        ' 初期コミット
        .Range("B25").Interior.Color = RGB(255, 0, 0)
        With .Range("B25").Borders
            .LineStyle = xlContinuous
            .Color = RGB(200, 200, 200)
        End With
        .Range("C25:E25").Merge
        .Range("C25").Value = "初期コミット（親コミットなし）"
        With .Range("C25")
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
        End With

        ' 通常コミット
        .Range("B26").Interior.Color = RGB(0, 128, 255)
        With .Range("B26").Borders
            .LineStyle = xlContinuous
            .Color = RGB(200, 200, 200)
        End With
        .Range("C26:E26").Merge
        .Range("C26").Value = "通常コミット（親コミット1つ）"
        With .Range("C26")
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
        End With

        ' マージコミット
        .Range("B27").Interior.Color = RGB(0, 255, 0)
        With .Range("B27").Borders
            .LineStyle = xlContinuous
            .Color = RGB(200, 200, 200)
        End With
        .Range("C27:E27").Merge
        .Range("C27").Value = "マージコミット（親コミット2つ以上）"
        With .Range("C27")
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
        End With

        ' =================================================================
        ' 列幅調整
        ' =================================================================
        .Columns("A").ColumnWidth = 3
        .Columns("B").ColumnWidth = 18
        .Columns("C").ColumnWidth = 12
        .Columns("D").ColumnWidth = 15
        .Columns("E").ColumnWidth = 15
        .Columns("F").ColumnWidth = 15
        .Columns("G").ColumnWidth = 15
        .Columns("H").ColumnWidth = 3

        .Range("A1").Select
    End With

    Application.ScreenUpdating = True
End Sub

'==============================================================================
' 環境変数を展開する (%USERNAME% など)
'==============================================================================
Private Function ExpandEnvironmentVariables(ByVal path As String) As String
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
Private Function GetRepoPathFromMainSheet() As String
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

Private Function GetCommitCountFromMainSheet() As Long
    On Error Resume Next
    GetCommitCountFromMainSheet = CLng(ThisWorkbook.Sheets(SHEET_MAIN).Range(CELL_COMMIT_COUNT).Value)
    If Err.Number <> 0 Or GetCommitCountFromMainSheet <= 0 Then
        GetCommitCountFromMainSheet = 100
    End If
    On Error GoTo 0
End Function

'==============================================================================
' メインプロシージャ: Git Log を可視化
'==============================================================================
Public Sub VisualizeGitLog()
    Dim commits() As CommitInfo
    Dim commitCount As Long
    Dim gitRepoPath As String
    Dim maxCommits As Long

    ' メインシートから設定を取得
    gitRepoPath = GetRepoPathFromMainSheet()
    maxCommits = GetCommitCountFromMainSheet()

    ' パスが空の場合
    If Len(Trim(gitRepoPath)) = 0 Then
        MsgBox "リポジトリパスが設定されていません。" & vbCrLf & _
               "メインシートのリポジトリパスを入力してください。", vbExclamation, "エラー"
        Exit Sub
    End If

    ' パスの存在確認
    If Dir(gitRepoPath, vbDirectory) = "" Then
        MsgBox "指定されたパスが存在しません: " & vbCrLf & vbCrLf & _
               gitRepoPath, vbCritical, "エラー"
        Exit Sub
    End If

    ' Gitリポジトリの確認
    If Not IsGitRepository(gitRepoPath) Then
        MsgBox "指定されたパスがGitリポジトリではありません: " & vbCrLf & vbCrLf & _
               gitRepoPath, vbCritical, "エラー"
        Exit Sub
    End If

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Debug.Print "========================================="
    Debug.Print "Git Log 可視化処理を開始します"
    Debug.Print "リポジトリ: " & gitRepoPath
    Debug.Print "取得件数: " & maxCommits & " 件"
    Debug.Print "========================================="

    ' Git Logを取得
    Debug.Print "コミット履歴を取得しています..."
    commitCount = 0

    On Error Resume Next
    commits = GetGitLog(gitRepoPath, maxCommits)
    If Err.Number <> 0 Then
        Debug.Print "GetGitLogでエラー発生: " & Err.Description
        Err.Clear
        GoTo CheckCommitCount
    End If

    Dim lowerBound As Long
    Dim upperBound As Long
    lowerBound = LBound(commits)
    upperBound = UBound(commits)

    If Err.Number <> 0 Then
        Err.Clear
        GoTo CheckCommitCount
    End If

    If Len(commits(lowerBound).Hash) > 0 Then
        commitCount = upperBound - lowerBound + 1
    End If
    On Error GoTo ErrorHandler

CheckCommitCount:
    If commitCount = 0 Then
        MsgBox "コミットが取得できませんでした。" & vbCrLf & _
               "リポジトリパスとGitのインストールを確認してください。", vbExclamation
        GoTo Cleanup
    End If

    Debug.Print "取得完了: " & commitCount & " 件"

    ' シートを準備
    Debug.Print "シートを準備しています..."
    PrepareSheets

    ' ダッシュボードシートを作成
    Debug.Print "ダッシュボードを作成しています..."
    CreateDashboardSheet commits, commitCount, gitRepoPath

    ' 履歴シートを作成
    Debug.Print "履歴シートを作成しています..."
    CreateHistorySheet commits, commitCount, gitRepoPath

    ' ブランチグラフシートを作成
    Debug.Print "ブランチグラフを作成しています..."
    CreateBranchGraphSheet commits, commitCount, gitRepoPath

    ' ダッシュボードシートをアクティブに
    ThisWorkbook.Sheets(SHEET_DASHBOARD).Select

    Debug.Print "========================================="
    Debug.Print "処理完了"
    Debug.Print "========================================="

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    If commitCount > 0 Then
        MsgBox "Git Log の可視化が完了しました。" & vbCrLf & vbCrLf & _
               "コミット数: " & commitCount & " 件", vbInformation, "処理完了"
    End If

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "エラーが発生しました: " & vbCrLf & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, vbCritical, "エラー"
End Sub

'==============================================================================
' Gitリポジトリかどうかを確認
'==============================================================================
Private Function IsGitRepository(ByVal repoPath As String) As Boolean
    Dim wsh As Object
    Dim exec As Object
    Dim command As String

    Set wsh = CreateObject("WScript.Shell")
    command = "cmd /c cd /d """ & repoPath & """ && " & GIT_COMMAND & " rev-parse --git-dir >nul 2>&1"
    Set exec = wsh.exec(command)

    Do While exec.Status = 0
        DoEvents
    Loop

    IsGitRepository = (exec.ExitCode = 0)
End Function

'==============================================================================
' Git Log を取得
'==============================================================================
Private Function GetGitLog(ByVal repoPath As String, ByVal maxCount As Long) As CommitInfo()
    Dim wsh As Object
    Dim fso As Object
    Dim command As String
    Dim output As String
    Dim lines() As String
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
Private Function ParseGitDate(ByVal dateStr As String) As Date
    On Error Resume Next
    ParseGitDate = CDate(Left(dateStr, 19))
    If Err.Number <> 0 Then
        ParseGitDate = Now
        Err.Clear
    End If
    On Error GoTo 0
End Function

'==============================================================================
' シートを準備
'==============================================================================
Private Sub PrepareSheets()
    Dim sheetNames As Variant
    Dim sheetName As Variant
    Dim ws As Worksheet
    Dim sheetExists As Boolean

    sheetNames = Array(SHEET_DASHBOARD, SHEET_HISTORY, SHEET_BRANCH_GRAPH)

    For Each sheetName In sheetNames
        sheetExists = False
        Set ws = Nothing

        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(CStr(sheetName))
        If Not ws Is Nothing Then
            sheetExists = True
        End If
        Err.Clear
        On Error GoTo 0

        If sheetExists Then
            ws.Cells.Clear
            On Error Resume Next
            ws.Cells.Interior.ColorIndex = xlNone
            On Error GoTo 0
        Else
            On Error Resume Next
            Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            If Not ws Is Nothing Then
                ws.Name = CStr(sheetName)
            End If
            On Error GoTo 0
        End If

        Set ws = Nothing
    Next sheetName
End Sub

'==============================================================================
' ダッシュボードシートを作成
'==============================================================================
Private Sub CreateDashboardSheet(ByRef commits() As CommitInfo, ByVal commitCount As Long, ByVal repoPath As String)
    Dim ws As Worksheet
    Dim i As Long
    Dim row As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_DASHBOARD)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "シート「" & SHEET_DASHBOARD & "」が見つかりません。" & vbCrLf & _
               "初期化を実行してから再度お試しください。", vbCritical, "エラー"
        Exit Sub
    End If

    ' 統計データの収集
    Dim authorDict As Object
    Set authorDict = CreateObject("Scripting.Dictionary")

    Dim minDate As Date
    Dim maxDate As Date
    Dim totalFiles As Long
    Dim totalInsertions As Long
    Dim totalDeletions As Long

    minDate = commits(0).CommitDate
    maxDate = commits(0).CommitDate
    totalFiles = 0
    totalInsertions = 0
    totalDeletions = 0

    For i = 0 To commitCount - 1
        ' 作者別カウント
        If authorDict.exists(commits(i).Author) Then
            authorDict(commits(i).Author) = authorDict(commits(i).Author) + 1
        Else
            authorDict.Add commits(i).Author, 1
        End If

        ' 日付範囲
        If commits(i).CommitDate < minDate Then minDate = commits(i).CommitDate
        If commits(i).CommitDate > maxDate Then maxDate = commits(i).CommitDate

        ' 変更量の合計
        totalFiles = totalFiles + commits(i).FilesChanged
        totalInsertions = totalInsertions + commits(i).Insertions
        totalDeletions = totalDeletions + commits(i).Deletions
    Next i

    With ws
        ' 全体の背景色を白に
        .Cells.Interior.Color = RGB(255, 255, 255)

        ' =================================================================
        ' タイトルエリア (行1-3)
        ' =================================================================
        .Range("B2:H2").Merge
        .Range("B2").Value = "Git Log ダッシュボード"
        With .Range("B2")
            .Font.Name = "Meiryo UI"
            .Font.Size = 20
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        .Range("B2:H3").Interior.Color = RGB(68, 114, 196)
        .Rows(2).RowHeight = 40
        .Rows(3).RowHeight = 5

        ' リポジトリ情報
        .Range("B4:H4").Merge
        .Range("B4").Value = "リポジトリ: " & repoPath
        With .Range("B4")
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
            .Font.Color = RGB(100, 100, 100)
        End With

        ' =================================================================
        ' サマリーセクション (行6-11)
        ' =================================================================
        .Range("B6:D6").Merge
        .Range("B6").Value = "サマリー"
        With .Range("B6")
            .Font.Name = "Meiryo UI"
            .Font.Size = 14
            .Font.Bold = True
            .Font.Color = RGB(68, 114, 196)
        End With
        With .Range("B6:D6").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(68, 114, 196)
            .Weight = xlMedium
        End With

        ' サマリー項目
        .Range("B8").Value = "総コミット数:"
        .Range("C8").Value = commitCount
        .Range("C8").Font.Bold = True
        .Range("C8").HorizontalAlignment = xlRight
        .Range("D8").Value = "件"

        .Range("B9").Value = "作者数:"
        .Range("C9").Value = authorDict.Count
        .Range("C9").Font.Bold = True
        .Range("C9").HorizontalAlignment = xlRight
        .Range("D9").Value = "人"

        .Range("B10").Value = "期間:"
        .Range("C10:D10").Merge
        .Range("C10").Value = Format(minDate, "yyyy/mm/dd") & " 〜 " & Format(maxDate, "yyyy/mm/dd")
        .Range("C10").Font.Bold = True
        .Range("C10").HorizontalAlignment = xlCenter

        .Range("B11").Value = "日数:"
        .Range("C11").Value = DateDiff("d", minDate, maxDate) + 1
        .Range("C11").Font.Bold = True
        .Range("C11").HorizontalAlignment = xlRight
        .Range("D11").Value = "日"

        ' サマリーエリアのスタイル
        .Range("B8:D11").Font.Name = "Meiryo UI"
        .Range("B8:D11").Font.Size = 11

        ' =================================================================
        ' 変更量セクション (行6-11, 右側)
        ' =================================================================
        .Range("F6:H6").Merge
        .Range("F6").Value = "変更量"
        With .Range("F6")
            .Font.Name = "Meiryo UI"
            .Font.Size = 14
            .Font.Bold = True
            .Font.Color = RGB(68, 114, 196)
        End With
        With .Range("F6:H6").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(68, 114, 196)
            .Weight = xlMedium
        End With

        .Range("F8").Value = "変更ファイル数:"
        .Range("G8").Value = totalFiles
        .Range("G8").Font.Bold = True
        .Range("G8").HorizontalAlignment = xlRight
        .Range("G8").NumberFormat = "#,##0"
        .Range("H8").Value = "ファイル"

        .Range("F9").Value = "追加行数:"
        .Range("G9").Value = totalInsertions
        .Range("G9").Font.Bold = True
        .Range("G9").Font.Color = RGB(0, 128, 0)
        .Range("G9").HorizontalAlignment = xlRight
        .Range("G9").NumberFormat = "#,##0"
        .Range("H9").Value = "行"

        .Range("F10").Value = "削除行数:"
        .Range("G10").Value = totalDeletions
        .Range("G10").Font.Bold = True
        .Range("G10").Font.Color = RGB(192, 0, 0)
        .Range("G10").HorizontalAlignment = xlRight
        .Range("G10").NumberFormat = "#,##0"
        .Range("H10").Value = "行"

        .Range("F11").Value = "純増行数:"
        .Range("G11").Value = totalInsertions - totalDeletions
        .Range("G11").Font.Bold = True
        If totalInsertions - totalDeletions >= 0 Then
            .Range("G11").Font.Color = RGB(0, 128, 0)
        Else
            .Range("G11").Font.Color = RGB(192, 0, 0)
        End If
        .Range("G11").HorizontalAlignment = xlRight
        .Range("G11").NumberFormat = "#,##0"
        .Range("H11").Value = "行"

        .Range("F8:H11").Font.Name = "Meiryo UI"
        .Range("F8:H11").Font.Size = 11

        ' =================================================================
        ' 作者別コミット数セクション (行13-)
        ' =================================================================
        .Range("B13:H13").Merge
        .Range("B13").Value = "作者別コミット数"
        With .Range("B13")
            .Font.Name = "Meiryo UI"
            .Font.Size = 14
            .Font.Bold = True
            .Font.Color = RGB(68, 114, 196)
        End With
        With .Range("B13:H13").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(68, 114, 196)
            .Weight = xlMedium
        End With

        ' ヘッダー
        .Range("B15").Value = "順位"
        .Range("C15").Value = "作者"
        .Range("D15").Value = "コミット数"
        .Range("E15").Value = "割合"
        .Range("F15:H15").Merge
        .Range("F15").Value = "グラフ"

        With .Range("B15:H15")
            .Font.Name = "Meiryo UI"
            .Font.Bold = True
            .Interior.Color = RGB(68, 114, 196)
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
        End With

        ' 作者データをソート（コミット数降順）
        Dim authors() As Variant
        Dim authorCounts() As Variant
        Dim authorCount As Long
        Dim keys As Variant
        Dim items As Variant

        authorCount = authorDict.Count
        ReDim authors(0 To authorCount - 1)
        ReDim authorCounts(0 To authorCount - 1)

        keys = authorDict.keys
        items = authorDict.items

        For i = 0 To authorCount - 1
            authors(i) = keys(i)
            authorCounts(i) = items(i)
        Next i

        ' バブルソート（降順）
        Dim j As Long
        Dim tempAuthor As Variant
        Dim tempCount As Variant

        For i = 0 To authorCount - 2
            For j = i + 1 To authorCount - 1
                If authorCounts(j) > authorCounts(i) Then
                    tempAuthor = authors(i)
                    tempCount = authorCounts(i)
                    authors(i) = authors(j)
                    authorCounts(i) = authorCounts(j)
                    authors(j) = tempAuthor
                    authorCounts(j) = tempCount
                End If
            Next j
        Next i

        ' データ行を出力（最大20人まで）
        Dim maxAuthors As Long
        maxAuthors = authorCount
        If maxAuthors > 20 Then maxAuthors = 20

        Dim maxCount As Long
        maxCount = authorCounts(0)

        For i = 0 To maxAuthors - 1
            row = 16 + i

            .Cells(row, 2).Value = i + 1
            .Cells(row, 2).HorizontalAlignment = xlCenter
            .Cells(row, 3).Value = authors(i)
            .Cells(row, 4).Value = authorCounts(i)
            .Cells(row, 4).HorizontalAlignment = xlRight
            .Cells(row, 5).Value = authorCounts(i) / commitCount
            .Cells(row, 5).NumberFormat = "0.0%"
            .Cells(row, 5).HorizontalAlignment = xlRight

            ' 簡易バーグラフ（セルの塗りつぶし）
            .Range(.Cells(row, 6), .Cells(row, 8)).Merge

            ' データバー的な表現
            Dim barWidth As Double
            barWidth = (authorCounts(i) / maxCount) * 100
            .Cells(row, 6).Value = String(Int(barWidth / 5), "█") & " " & authorCounts(i)
            .Cells(row, 6).Font.Color = RGB(68, 114, 196)
            .Cells(row, 6).Font.Name = "Consolas"
            .Cells(row, 6).Font.Size = 10

            ' 交互に色分け
            If i Mod 2 = 0 Then
                .Range(.Cells(row, 2), .Cells(row, 8)).Interior.Color = RGB(245, 245, 245)
            End If
        Next i

        ' 残りの作者がある場合
        If authorCount > 20 Then
            row = 16 + maxAuthors
            .Cells(row, 2).Value = "..."
            .Cells(row, 3).Value = "他 " & (authorCount - 20) & " 人"
            .Range(.Cells(row, 2), .Cells(row, 8)).Font.Color = RGB(128, 128, 128)
            .Range(.Cells(row, 2), .Cells(row, 8)).Font.Italic = True
        End If

        ' =================================================================
        ' 列幅調整
        ' =================================================================
        .Columns("A").ColumnWidth = 3
        .Columns("B").ColumnWidth = 16
        .Columns("C").ColumnWidth = 20
        .Columns("D").ColumnWidth = 12
        .Columns("E").ColumnWidth = 10
        .Columns("F").ColumnWidth = 10
        .Columns("G").ColumnWidth = 10
        .Columns("H").ColumnWidth = 10
        .Columns("I").ColumnWidth = 3

        .Range("A1").Select
    End With
End Sub

'==============================================================================
' 履歴シートを作成
'==============================================================================
Private Sub CreateHistorySheet(ByRef commits() As CommitInfo, ByVal commitCount As Long, ByVal repoPath As String)
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_HISTORY)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "シート「" & SHEET_HISTORY & "」が見つかりません。" & vbCrLf & _
               "初期化を実行してから再度お試しください。", vbCritical, "エラー"
        Exit Sub
    End If

    With ws
        ' タイトル
        .Range("A1").Value = "Git コミット履歴"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True

        ' リポジトリ情報
        .Range("A2").Value = "リポジトリ: " & repoPath & "  |  取得件数: " & commitCount & " 件"
        .Range("A2").Font.Size = 10
        .Range("A2").Font.Color = RGB(100, 100, 100)

        ' ヘッダー行
        .Range("A4").Value = "No"
        .Range("B4").Value = "ハッシュ"
        .Range("C4").Value = "作者"
        .Range("D4").Value = "日時"
        .Range("E4").Value = "コミットメッセージ"
        .Range("F4").Value = "ブランチ/タグ"
        .Range("G4").Value = "変更ファイル数"
        .Range("H4").Value = "追加行数"
        .Range("I4").Value = "削除行数"

        ' ヘッダー書式
        With .Range("A4:I4")
            .Font.Bold = True
            .Interior.Color = RGB(68, 114, 196)
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
        End With

        ' データ行
        Dim i As Long
        Dim row As Long

        For i = 0 To commitCount - 1
            row = i + 5

            .Cells(row, 1).Value = i + 1
            .Cells(row, 2).Value = commits(i).Hash
            .Cells(row, 3).Value = commits(i).Author
            .Cells(row, 4).Value = commits(i).CommitDate
            .Cells(row, 4).NumberFormat = "yyyy/mm/dd hh:mm"
            .Cells(row, 5).Value = commits(i).Subject
            .Cells(row, 6).Value = commits(i).RefNames
            .Cells(row, 7).Value = commits(i).FilesChanged
            .Cells(row, 8).Value = commits(i).Insertions
            .Cells(row, 9).Value = commits(i).Deletions

            ' 交互に色分け
            If i Mod 2 = 0 Then
                .Range(.Cells(row, 1), .Cells(row, 9)).Interior.Color = RGB(245, 245, 245)
            End If
        Next i

        ' 列幅調整
        .Columns("A").ColumnWidth = 5
        .Columns("B").ColumnWidth = 10
        .Columns("C").ColumnWidth = 15
        .Columns("D").ColumnWidth = 16
        .Columns("E").ColumnWidth = 60
        .Columns("F").ColumnWidth = 20
        .Columns("G").ColumnWidth = 12
        .Columns("H").ColumnWidth = 10
        .Columns("I").ColumnWidth = 10

        ' コミットメッセージ列の折り返し表示と上揃え
        .Columns("E").WrapText = True
        .Columns("E").VerticalAlignment = xlTop
        .Range(.Cells(5, 1), .Cells(commitCount + 4, 9)).VerticalAlignment = xlTop

        ' フィルターを設定
        .Range("A4:I4").AutoFilter
    End With

    ' ウィンドウ枠の固定（シートをアクティブにする必要がある）
    ws.Activate
    ws.Rows(5).Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select
End Sub

'==============================================================================
' ブランチグラフシートを作成
'==============================================================================
Private Sub CreateBranchGraphSheet(ByRef commits() As CommitInfo, ByVal commitCount As Long, ByVal repoPath As String)
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_BRANCH_GRAPH)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "シート「" & SHEET_BRANCH_GRAPH & "」が見つかりません。" & vbCrLf & _
               "初期化を実行してから再度お試しください。", vbCritical, "エラー"
        Exit Sub
    End If

    ' シートの図形をすべて削除
    Dim shp As Object
    For Each shp In ws.Shapes
        shp.Delete
    Next shp

    With ws
        ' タイトル
        .Range("A1").Value = "Git ブランチグラフ"
        .Range("A1").Font.Size = 18
        .Range("A1").Font.Bold = True

        .Range("A2").Value = "リポジトリ: " & repoPath
        .Range("A3").Value = "コミット数: " & commitCount

        ' グラフ描画エリアの設定
        Dim startRow As Long
        Dim startCol As Long
        Dim nodeSize As Double
        Dim vSpacing As Double
        Dim hSpacing As Double

        startRow = 5
        startCol = 2
        nodeSize = 12
        vSpacing = 25
        hSpacing = 60

        ' コミットハッシュとインデックスのマッピング
        Dim commitIndexMap As Object
        Set commitIndexMap = CreateObject("Scripting.Dictionary")

        Dim i As Long
        For i = 0 To commitCount - 1
            commitIndexMap.Add commits(i).Hash, i
        Next i

        ' ブランチレーンの割り当て
        Dim lanes() As String
        ReDim lanes(0 To 9)
        Dim laneCount As Long
        laneCount = 0

        Dim commitLanes() As Long
        ReDim commitLanes(0 To commitCount - 1)

        For i = 0 To commitCount - 1
            Dim lane As Long
            lane = -1

            If commits(i).ParentCount > 0 Then
                Dim parentHashes() As String
                parentHashes = Split(Trim(commits(i).ParentHashes), " ")

                Dim p As Long
                For p = 0 To UBound(parentHashes)
                    Dim parentHash As String
                    parentHash = Trim(parentHashes(p))

                    If Len(parentHash) > 0 Then
                        Dim j As Long
                        For j = i + 1 To commitCount - 1
                            If commits(j).Hash = parentHash Then
                                lane = commitLanes(j)
                                Exit For
                            End If
                        Next j

                        If lane >= 0 Then Exit For
                    End If
                Next p
            End If

            If lane < 0 Then
                lane = laneCount
                If laneCount < 10 Then laneCount = laneCount + 1
            End If

            commitLanes(i) = lane
        Next i

        ' コミットノードと接続線を描画
        For i = 0 To commitCount - 1
            Dim y As Double
            Dim x As Double

            y = .Cells(startRow + i, 1).Top
            x = .Cells(startRow, startCol + commitLanes(i)).Left

            Dim nodeColor As Long
            If commits(i).ParentCount = 0 Then
                nodeColor = RGB(255, 0, 0)
            ElseIf commits(i).ParentCount >= 2 Then
                nodeColor = RGB(0, 255, 0)
            Else
                nodeColor = RGB(0, 128, 255)
            End If

            Dim node As Object
            Set node = .Shapes.AddShape(msoShapeOval, x, y, nodeSize, nodeSize)
            node.Fill.ForeColor.RGB = nodeColor
            node.Line.ForeColor.RGB = RGB(0, 0, 0)
            node.Line.Weight = 1
            node.Name = "Node_" & commits(i).Hash

            Dim tooltip As Object
            Set tooltip = .Shapes.AddTextbox(msoTextOrientationHorizontal, x + nodeSize + 5, y - 3, 300, 15)
            tooltip.TextFrame.Characters.Text = commits(i).Hash & " " & commits(i).Subject
            tooltip.TextFrame.Characters.Font.Size = 8
            tooltip.Line.Visible = msoFalse
            tooltip.Fill.Visible = msoFalse

            If commits(i).ParentCount > 0 Then
                Dim parentHashes2() As String
                parentHashes2 = Split(Trim(commits(i).ParentHashes), " ")

                Dim k As Long
                For k = 0 To UBound(parentHashes2)
                    Dim parentHash2 As String
                    parentHash2 = Trim(parentHashes2(k))

                    If Len(parentHash2) > 0 And commitIndexMap.exists(parentHash2) Then
                        Dim parentIndex As Long
                        parentIndex = commitIndexMap(parentHash2)

                        Dim y2 As Double
                        Dim x2 As Double

                        y2 = .Cells(startRow + parentIndex, 1).Top
                        x2 = .Cells(startRow, startCol + commitLanes(parentIndex)).Left

                        Dim lineShape As Object
                        Set lineShape = .Shapes.AddLine(x + nodeSize / 2, y + nodeSize / 2, x2 + nodeSize / 2, y2 + nodeSize / 2)
                        lineShape.Line.ForeColor.RGB = RGB(128, 128, 128)
                        lineShape.Line.Weight = 1.5
                        lineShape.ZOrder msoSendToBack
                    End If
                Next k
            End If
        Next i

        .Columns("A").ColumnWidth = 3
        .Columns("B:K").ColumnWidth = 10
        .Rows(startRow & ":" & (startRow + commitCount)).RowHeight = 20
    End With
End Sub
