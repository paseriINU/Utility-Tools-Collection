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
Private Const SHEET_HISTORY As String = "履歴"
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

        .Range("B18").Value = "履歴"
        With .Range("B18")
            .Font.Name = "Meiryo UI"
            .Font.Size = 11
            .Font.Bold = True
            .Font.Color = RGB(68, 114, 196)
        End With
        .Range("C18:G18").Merge
        .Range("C18").Value = "コミット履歴の詳細一覧（ハッシュ、作者、日時、メッセージ、変更量等）"
        With .Range("C18")
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
        End With

        .Range("B19").Value = "ブランチグラフ"
        With .Range("B19")
            .Font.Name = "Meiryo UI"
            .Font.Size = 11
            .Font.Bold = True
            .Font.Color = RGB(68, 114, 196)
        End With
        .Range("C19:G19").Merge
        .Range("C19").Value = "ブランチ構造を視覚化（コミットノードと接続線）"
        With .Range("C19")
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
        End With

        ' =================================================================
        ' 色凡例セクション (行22-28)
        ' =================================================================
        .Range("B22:G22").Merge
        .Range("B22").Value = "ブランチグラフの色凡例"
        With .Range("B22")
            .Font.Name = "Meiryo UI"
            .Font.Size = 14
            .Font.Bold = True
            .Font.Color = RGB(68, 114, 196)
        End With
        With .Range("B22:G22").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(68, 114, 196)
            .Weight = xlMedium
        End With

        ' 初期コミット
        .Range("B24").Interior.Color = RGB(255, 0, 0)
        With .Range("B24").Borders
            .LineStyle = xlContinuous
            .Color = RGB(200, 200, 200)
        End With
        .Range("C24:E24").Merge
        .Range("C24").Value = "初期コミット（親コミットなし）"
        With .Range("C24")
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
        End With

        ' 通常コミット
        .Range("B25").Interior.Color = RGB(0, 128, 255)
        With .Range("B25").Borders
            .LineStyle = xlContinuous
            .Color = RGB(200, 200, 200)
        End With
        .Range("C25:E25").Merge
        .Range("C25").Value = "通常コミット（親コミット1つ）"
        With .Range("C25")
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
        End With

        ' マージコミット
        .Range("B26").Interior.Color = RGB(0, 255, 0)
        With .Range("B26").Borders
            .LineStyle = xlContinuous
            .Color = RGB(200, 200, 200)
        End With
        .Range("C26:E26").Merge
        .Range("C26").Value = "マージコミット（親コミット2つ以上）"
        With .Range("C26")
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

    ' 履歴シートを作成
    Debug.Print "履歴シートを作成しています..."
    CreateHistorySheet commits, commitCount, gitRepoPath

    ' ブランチグラフシートを作成
    Debug.Print "ブランチグラフを作成しています..."
    CreateBranchGraphSheet commits, commitCount, gitRepoPath

    ' 履歴シートをアクティブに
    Sheets(SHEET_HISTORY).Select

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

    Set wsh = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")

    tempFile = fso.GetSpecialFolder(2) & "\gitlog_" & fso.GetTempName & ".txt"

    command = "cmd /c chcp 65001 >nul && cd /d """ & repoPath & """ && " & _
              GIT_COMMAND & " log --all -n " & maxCount & _
              " --pretty=format:""%h|%H|%P|%an|%ae|%ai|%s|%d"" --numstat > """ & tempFile & """ 2>&1"

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
    lines = Split(output, vbLf)

    commitIndex = 0
    ReDim commits(0 To maxCount - 1)

    i = 0
    Do While i <= UBound(lines)
        Dim line As String
        line = Trim(lines(i))

        If InStr(line, "|") > 0 Then
            parts = Split(line, "|")

            If UBound(parts) >= 6 Then
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
                    .Subject = parts(6)
                    If UBound(parts) >= 7 Then
                        .RefNames = Trim(Replace(Replace(parts(7), "(", ""), ")", ""))
                    Else
                        .RefNames = ""
                    End If
                    .FilesChanged = 0
                    .Insertions = 0
                    .Deletions = 0
                End With

                i = i + 1
                Do While i <= UBound(lines)
                    line = Trim(lines(i))

                    If Len(line) = 0 Or InStr(line, "|") > 0 Then
                        Exit Do
                    End If

                    Dim statParts() As String
                    statParts = Split(line, vbTab)

                    If UBound(statParts) >= 2 Then
                        commits(commitIndex).FilesChanged = commits(commitIndex).FilesChanged + 1
                        If IsNumeric(statParts(0)) Then
                            commits(commitIndex).Insertions = commits(commitIndex).Insertions + CLng(statParts(0))
                        End If
                        If IsNumeric(statParts(1)) Then
                            commits(commitIndex).Deletions = commits(commitIndex).Deletions + CLng(statParts(1))
                        End If
                    End If

                    i = i + 1
                Loop

                commitIndex = commitIndex + 1

                If InStr(line, "|") > 0 Then
                    i = i - 1
                End If
            End If
        End If

        i = i + 1
    Loop

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

    sheetNames = Array(SHEET_HISTORY, SHEET_BRANCH_GRAPH)

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
' 履歴シートを作成（詳細情報付き）
'==============================================================================
Private Sub CreateHistorySheet(ByRef commits() As CommitInfo, ByVal commitCount As Long, ByVal repoPath As String)
    Dim ws As Worksheet
    Set ws = Sheets(SHEET_HISTORY)

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
        .Range("C4").Value = "フルハッシュ"
        .Range("D4").Value = "作者"
        .Range("E4").Value = "メール"
        .Range("F4").Value = "日時"
        .Range("G4").Value = "曜日"
        .Range("H4").Value = "種別"
        .Range("I4").Value = "コミットメッセージ"
        .Range("J4").Value = "ブランチ/タグ"
        .Range("K4").Value = "変更ファイル数"
        .Range("L4").Value = "追加行"
        .Range("M4").Value = "削除行"
        .Range("N4").Value = "変更量"
        .Range("O4").Value = "親コミット数"

        ' ヘッダー書式
        With .Range("A4:O4")
            .Font.Bold = True
            .Interior.Color = RGB(68, 114, 196)
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
        End With

        ' データ行
        Dim i As Long
        Dim row As Long
        Dim commitType As String
        Dim dayName As String

        For i = 0 To commitCount - 1
            row = i + 5

            ' コミット種別を判定
            If commits(i).ParentCount = 0 Then
                commitType = "初期"
            ElseIf commits(i).ParentCount >= 2 Then
                commitType = "マージ"
            Else
                commitType = "通常"
            End If

            ' 曜日を取得
            dayName = WeekdayName(Weekday(commits(i).CommitDate), True)

            .Cells(row, 1).Value = i + 1
            .Cells(row, 2).Value = commits(i).Hash
            .Cells(row, 3).Value = commits(i).FullHash
            .Cells(row, 4).Value = commits(i).Author
            .Cells(row, 5).Value = commits(i).AuthorEmail
            .Cells(row, 6).Value = commits(i).CommitDate
            .Cells(row, 6).NumberFormat = "yyyy/mm/dd hh:mm"
            .Cells(row, 7).Value = dayName
            .Cells(row, 8).Value = commitType
            .Cells(row, 9).Value = commits(i).Subject
            .Cells(row, 10).Value = commits(i).RefNames
            .Cells(row, 11).Value = commits(i).FilesChanged
            .Cells(row, 12).Value = commits(i).Insertions
            .Cells(row, 13).Value = commits(i).Deletions
            .Cells(row, 14).Value = commits(i).Insertions + commits(i).Deletions
            .Cells(row, 15).Value = commits(i).ParentCount

            ' コミット種別による色分け
            Select Case commitType
                Case "初期"
                    .Range(.Cells(row, 1), .Cells(row, 15)).Interior.Color = RGB(255, 230, 230)
                Case "マージ"
                    .Range(.Cells(row, 1), .Cells(row, 15)).Interior.Color = RGB(230, 255, 230)
                Case Else
                    If i Mod 2 = 0 Then
                        .Range(.Cells(row, 1), .Cells(row, 15)).Interior.Color = RGB(245, 245, 245)
                    End If
            End Select
        Next i

        ' 列幅調整
        .Columns("A").ColumnWidth = 5
        .Columns("B").ColumnWidth = 10
        .Columns("C").ColumnWidth = 42
        .Columns("D").ColumnWidth = 15
        .Columns("E").ColumnWidth = 25
        .Columns("F").ColumnWidth = 16
        .Columns("G").ColumnWidth = 6
        .Columns("H").ColumnWidth = 8
        .Columns("I").ColumnWidth = 50
        .Columns("J").ColumnWidth = 20
        .Columns("K").ColumnWidth = 12
        .Columns("L").ColumnWidth = 8
        .Columns("M").ColumnWidth = 8
        .Columns("N").ColumnWidth = 8
        .Columns("O").ColumnWidth = 10

        ' フィルターを設定
        .Range("A4:O4").AutoFilter

        ' ウィンドウ枠の固定
        .Rows(5).Select
        ActiveWindow.FreezePanes = True

        .Range("A1").Select
    End With
End Sub

'==============================================================================
' ブランチグラフシートを作成
'==============================================================================
Private Sub CreateBranchGraphSheet(ByRef commits() As CommitInfo, ByVal commitCount As Long, ByVal repoPath As String)
    Dim ws As Worksheet
    Set ws = Sheets(SHEET_BRANCH_GRAPH)

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
