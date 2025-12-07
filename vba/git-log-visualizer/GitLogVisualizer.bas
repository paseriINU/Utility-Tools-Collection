Attribute VB_Name = "GitLogVisualizer"
'==============================================================================
' Git Log 可視化ツール
'==============================================================================
' 概要:
'   Excelから実行し、Gitリポジトリのコミット履歴を取得して、
'   表形式・統計情報・グラフで視覚化するツールです。
'
' 機能:
'   - コミット履歴の一覧表示
'   - 作者別・日別・ブランチ別の統計
'   - グラフによる可視化
'   - ダッシュボードで概要表示
'
' 必要な環境:
'   - Microsoft Excel 2010以降
'   - Git がインストールされており、パスが通っていること
'
' 作成日: 2025-12-07
'==============================================================================

Option Explicit

'==============================================================================
' 設定: ここを編集してください
'==============================================================================
' 取得するコミット数（最近のN件）
Private Const COMMIT_COUNT As Long = 100

' Gitコマンドのパス（通常は "git" でOK。パスが通っていない場合はフルパス指定）
Private Const GIT_COMMAND As String = "git"

'==============================================================================
' デフォルトのGitリポジトリパスを取得
'==============================================================================
Private Function GetDefaultRepoPath() As String
    ' C:\Users\%USERNAME%\source\Git\project
    GetDefaultRepoPath = "C:\Users\" & Environ("USERNAME") & "\source\Git\project"
End Function

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
    FilesChanged As Long    ' 変更ファイル数
    Insertions As Long      ' 追加行数
    Deletions As Long       ' 削除行数
End Type

'==============================================================================
' メインプロシージャ: Git Log を可視化
'==============================================================================
Public Sub VisualizeGitLog()
    Dim commits() As CommitInfo
    Dim commitCount As Long
    Dim i As Long
    Dim gitRepoPath As String

    ' デフォルトのリポジトリパスを取得
    gitRepoPath = GetDefaultRepoPath()

    ' リポジトリの存在確認
    If Dir(gitRepoPath & "\.git", vbDirectory) = "" Then
        MsgBox "指定されたパスがGitリポジトリではありません: " & vbCrLf & vbCrLf & _
               gitRepoPath & vbCrLf & vbCrLf & _
               "パスを変更する場合は、GetDefaultRepoPath() 関数を編集してください。", vbCritical, "エラー"
        Exit Sub
    End If

    On Error GoTo ErrorHandler

    ' 処理開始メッセージ
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Debug.Print "========================================="
    Debug.Print "Git Log 可視化処理を開始します"
    Debug.Print "リポジトリ: " & gitRepoPath
    Debug.Print "取得件数: " & COMMIT_COUNT & " 件"
    Debug.Print "========================================="

    ' Git Logを取得
    Debug.Print "コミット履歴を取得しています..."
    commits = GetGitLog(gitRepoPath, COMMIT_COUNT)
    commitCount = UBound(commits) - LBound(commits) + 1

    If commitCount = 0 Then
        MsgBox "コミットが取得できませんでした。" & vbCrLf & _
               "リポジトリパスとGitのインストールを確認してください。", vbExclamation
        GoTo Cleanup
    End If

    Debug.Print "取得完了: " & commitCount & " 件"

    ' 既存のシートをクリア
    Debug.Print "シートを準備しています..."
    ClearAllSheets

    ' ダッシュボードシートを作成
    Debug.Print "ダッシュボードを作成しています..."
    CreateDashboard commits, commitCount, gitRepoPath

    ' コミット履歴シートを作成
    Debug.Print "コミット履歴シートを作成しています..."
    CreateCommitHistorySheet commits, commitCount

    ' 統計シートを作成
    Debug.Print "統計シートを作成しています..."
    CreateStatisticsSheet commits, commitCount

    ' グラフシートを作成
    Debug.Print "グラフシートを作成しています..."
    CreateChartsSheet commits, commitCount

    ' ダッシュボードシートをアクティブに
    Sheets("Dashboard").Select

    Debug.Print "========================================="
    Debug.Print "処理完了"
    Debug.Print "========================================="

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "Git Log の可視化が完了しました。" & vbCrLf & vbCrLf & _
           "コミット数: " & commitCount & " 件", vbInformation, "処理完了"

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "エラーが発生しました: " & vbCrLf & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, vbCritical, "エラー"
End Sub

'==============================================================================
' Git Log を取得
'==============================================================================
Private Function GetGitLog(ByVal repoPath As String, ByVal maxCount As Long) As CommitInfo()
    Dim wsh As Object
    Dim exec As Object
    Dim command As String
    Dim output As String
    Dim lines() As String
    Dim commits() As CommitInfo
    Dim i As Long
    Dim commitIndex As Long
    Dim parts() As String

    ' WScript.Shell を作成
    Set wsh = CreateObject("WScript.Shell")

    ' Git Log コマンド（全ブランチ、カスタムフォーマット）
    ' フォーマット: ハッシュ|フルハッシュ|作者|メール|日付|件名|ref名
    command = "cmd /c cd /d """ & repoPath & """ && " & _
              GIT_COMMAND & " log --all -n " & maxCount & _
              " --pretty=format:""%h|%H|%an|%ae|%ai|%s|%d"" --numstat"

    ' コマンド実行
    Set exec = wsh.exec(command)

    ' 出力を取得
    Do While exec.Status = 0
        DoEvents
    Loop
    output = exec.StdOut.ReadAll

    If Len(output) = 0 Then
        ReDim commits(0 To 0)
        GetGitLog = commits
        Exit Function
    End If

    ' 行に分割
    output = Replace(output, vbCrLf, vbLf)
    output = Replace(output, vbCr, vbLf)
    lines = Split(output, vbLf)

    ' コミット数をカウント（"|" を6個含む行がコミット情報）
    commitIndex = 0
    ReDim commits(0 To maxCount - 1)

    i = 0
    Do While i <= UBound(lines)
        Dim line As String
        line = Trim(lines(i))

        ' コミット情報行を判定（"|" を含む）
        If InStr(line, "|") > 0 Then
            parts = Split(line, "|")

            If UBound(parts) >= 5 Then
                ' コミット情報を格納
                With commits(commitIndex)
                    .Hash = parts(0)
                    .FullHash = parts(1)
                    .Author = parts(2)
                    .AuthorEmail = parts(3)
                    .CommitDate = ParseGitDate(parts(4))
                    .Subject = parts(5)
                    .RefNames = If(UBound(parts) >= 6, Trim(Replace(Replace(parts(6), "(", ""), ")", "")), "")
                    .FilesChanged = 0
                    .Insertions = 0
                    .Deletions = 0
                End With

                ' 次の行から numstat を解析
                i = i + 1
                Do While i <= UBound(lines)
                    line = Trim(lines(i))

                    ' 空行または次のコミットに到達したら終了
                    If Len(line) = 0 Or InStr(line, "|") > 0 Then
                        Exit Do
                    End If

                    ' numstat 行を解析（追加\t削除\tファイル名）
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

                ' 次のコミットがある場合は継続
                If InStr(line, "|") > 0 Then
                    i = i - 1
                End If
            End If
        End If

        i = i + 1
    Loop

    ' 配列のサイズを調整
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
    ' フォーマット例: 2025-12-07 12:34:56 +0900
    On Error Resume Next
    ParseGitDate = CDate(Left(dateStr, 19))
    If Err.Number <> 0 Then
        ParseGitDate = Now
        Err.Clear
    End If
    On Error GoTo 0
End Function

'==============================================================================
' すべてのシートをクリア（または作成）
'==============================================================================
Private Sub ClearAllSheets()
    Dim sheetNames As Variant
    Dim sheetName As Variant
    Dim ws As Worksheet

    sheetNames = Array("Dashboard", "CommitHistory", "Statistics", "Charts")

    ' シートが存在しない場合は作成、存在する場合はクリア
    For Each sheetName In sheetNames
        On Error Resume Next
        Set ws = Sheets(CStr(sheetName))

        If ws Is Nothing Then
            ' シートを新規作成
            Set ws = Sheets.Add(After:=Sheets(Sheets.Count))
            ws.Name = CStr(sheetName)
        Else
            ' 既存シートをクリア
            ws.Cells.Clear
            ws.Cells.Interior.ColorIndex = xlNone
        End If

        Set ws = Nothing
        On Error GoTo 0
    Next sheetName
End Sub

'==============================================================================
' ダッシュボードシートを作成
'==============================================================================
Private Sub CreateDashboard(ByRef commits() As CommitInfo, ByVal commitCount As Long, ByVal repoPath As String)
    Dim ws As Worksheet
    Set ws = Sheets("Dashboard")

    With ws
        ' タイトル
        .Range("A1").Value = "Git Log ダッシュボード"
        .Range("A1").Font.Size = 18
        .Range("A1").Font.Bold = True

        ' リポジトリ情報
        .Range("A3").Value = "リポジトリパス:"
        .Range("B3").Value = repoPath
        .Range("A4").Value = "取得コミット数:"
        .Range("B4").Value = commitCount
        .Range("A5").Value = "最新コミット:"
        If commitCount > 0 Then
            .Range("B5").Value = commits(0).CommitDate
        End If
        .Range("A6").Value = "最古コミット:"
        If commitCount > 0 Then
            .Range("B6").Value = commits(commitCount - 1).CommitDate
        End If

        ' 統計サマリー
        .Range("A8").Value = "統計サマリー"
        .Range("A8").Font.Size = 14
        .Range("A8").Font.Bold = True

        Dim authors As Object
        Set authors = CreateObject("Scripting.Dictionary")
        Dim totalInsertions As Long
        Dim totalDeletions As Long
        Dim i As Long

        For i = 0 To commitCount - 1
            If Not authors.exists(commits(i).Author) Then
                authors.Add commits(i).Author, 0
            End If
            authors(commits(i).Author) = authors(commits(i).Author) + 1
            totalInsertions = totalInsertions + commits(i).Insertions
            totalDeletions = totalDeletions + commits(i).Deletions
        Next i

        .Range("A10").Value = "作者数:"
        .Range("B10").Value = authors.Count
        .Range("A11").Value = "総追加行数:"
        .Range("B11").Value = totalInsertions
        .Range("A12").Value = "総削除行数:"
        .Range("B12").Value = totalDeletions

        ' 列幅調整
        .Columns("A:B").AutoFit
    End With
End Sub

'==============================================================================
' コミット履歴シートを作成
'==============================================================================
Private Sub CreateCommitHistorySheet(ByRef commits() As CommitInfo, ByVal commitCount As Long)
    Dim ws As Worksheet
    Set ws = Sheets("CommitHistory")

    With ws
        ' ヘッダー
        .Range("A1").Value = "No"
        .Range("B1").Value = "ハッシュ"
        .Range("C1").Value = "作者"
        .Range("D1").Value = "日時"
        .Range("E1").Value = "コミットメッセージ"
        .Range("F1").Value = "ブランチ/タグ"
        .Range("G1").Value = "変更ファイル数"
        .Range("H1").Value = "追加行数"
        .Range("I1").Value = "削除行数"

        ' ヘッダー書式
        With .Range("A1:I1")
            .Font.Bold = True
            .Interior.Color = RGB(68, 114, 196)
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
        End With

        ' データ行
        Dim i As Long
        For i = 0 To commitCount - 1
            Dim row As Long
            row = i + 2

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

            ' 交互に色を付ける
            If i Mod 2 = 0 Then
                .Range(.Cells(row, 1), .Cells(row, 9)).Interior.Color = RGB(242, 242, 242)
            End If
        Next i

        ' 列幅調整
        .Columns("A").ColumnWidth = 5
        .Columns("B").ColumnWidth = 10
        .Columns("C").ColumnWidth = 15
        .Columns("D").ColumnWidth = 16
        .Columns("E").ColumnWidth = 50
        .Columns("F").ColumnWidth = 20
        .Columns("G:I").ColumnWidth = 12

        ' フィルターを設定
        .Range("A1:I1").AutoFilter
    End With
End Sub

'==============================================================================
' 統計シートを作成
'==============================================================================
Private Sub CreateStatisticsSheet(ByRef commits() As CommitInfo, ByVal commitCount As Long)
    Dim ws As Worksheet
    Set ws = Sheets("Statistics")

    Dim authors As Object
    Dim dates As Object
    Set authors = CreateObject("Scripting.Dictionary")
    Set dates = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim author As Variant
    Dim dateKey As String

    ' 作者別・日別に集計
    For i = 0 To commitCount - 1
        ' 作者別
        author = commits(i).Author
        If Not authors.exists(author) Then
            authors.Add author, 0
        End If
        authors(author) = authors(author) + 1

        ' 日別
        dateKey = Format(commits(i).CommitDate, "yyyy-mm-dd")
        If Not dates.exists(dateKey) Then
            dates.Add dateKey, 0
        End If
        dates(dateKey) = dates(dateKey) + 1
    Next i

    With ws
        ' 作者別統計
        .Range("A1").Value = "作者別コミット数"
        .Range("A1").Font.Size = 14
        .Range("A1").Font.Bold = True

        .Range("A3").Value = "作者"
        .Range("B3").Value = "コミット数"
        With .Range("A3:B3")
            .Font.Bold = True
            .Interior.Color = RGB(68, 114, 196)
            .Font.Color = RGB(255, 255, 255)
        End With

        Dim row As Long
        row = 4
        For Each author In authors.Keys
            .Cells(row, 1).Value = author
            .Cells(row, 2).Value = authors(author)
            row = row + 1
        Next author

        ' 日別統計
        .Range("D1").Value = "日別コミット数"
        .Range("D1").Font.Size = 14
        .Range("D1").Font.Bold = True

        .Range("D3").Value = "日付"
        .Range("E3").Value = "コミット数"
        With .Range("D3:E3")
            .Font.Bold = True
            .Interior.Color = RGB(68, 114, 196)
            .Font.Color = RGB(255, 255, 255)
        End With

        ' 日付でソート（キーを配列に変換してソート）
        Dim dateKeys() As String
        ReDim dateKeys(0 To dates.Count - 1)
        i = 0
        For Each dateKey In dates.Keys
            dateKeys(i) = dateKey
            i = i + 1
        Next dateKey

        ' 簡易ソート（バブルソート）
        Dim j As Long
        Dim temp As String
        For i = 0 To UBound(dateKeys) - 1
            For j = i + 1 To UBound(dateKeys)
                If dateKeys(i) > dateKeys(j) Then
                    temp = dateKeys(i)
                    dateKeys(i) = dateKeys(j)
                    dateKeys(j) = temp
                End If
            Next j
        Next i

        row = 4
        For i = 0 To UBound(dateKeys)
            .Cells(row, 4).Value = dateKeys(i)
            .Cells(row, 5).Value = dates(dateKeys(i))
            row = row + 1
        Next i

        ' 列幅調整
        .Columns("A:B").AutoFit
        .Columns("D:E").AutoFit
    End With
End Sub

'==============================================================================
' グラフシートを作成
'==============================================================================
Private Sub CreateChartsSheet(ByRef commits() As CommitInfo, ByVal commitCount As Long)
    Dim ws As Worksheet
    Dim statsWs As Worksheet
    Dim chartObj As ChartObject

    Set ws = Sheets("Charts")
    Set statsWs = Sheets("Statistics")

    ' 作者別コミット数の棒グラフ
    Dim lastRow As Long
    lastRow = statsWs.Cells(statsWs.Rows.Count, 1).End(xlUp).row

    If lastRow >= 4 Then
        Set chartObj = ws.ChartObjects.Add(Left:=10, Top:=10, Width:=400, Height:=300)
        With chartObj.Chart
            .ChartType = xlColumnClustered
            .SetSourceData statsWs.Range("A3:B" & lastRow)
            .HasTitle = True
            .ChartTitle.Text = "作者別コミット数"
            .Axes(xlCategory).HasTitle = True
            .Axes(xlCategory).AxisTitle.Text = "作者"
            .Axes(xlValue).HasTitle = True
            .Axes(xlValue).AxisTitle.Text = "コミット数"
        End With
    End If

    ' 日別コミット数の折れ線グラフ
    lastRow = statsWs.Cells(statsWs.Rows.Count, 4).End(xlUp).row

    If lastRow >= 4 Then
        Set chartObj = ws.ChartObjects.Add(Left:=450, Top:=10, Width:=400, Height:=300)
        With chartObj.Chart
            .ChartType = xlLine
            .SetSourceData statsWs.Range("D3:E" & lastRow)
            .HasTitle = True
            .ChartTitle.Text = "日別コミット数"
            .Axes(xlCategory).HasTitle = True
            .Axes(xlCategory).AxisTitle.Text = "日付"
            .Axes(xlValue).HasTitle = True
            .Axes(xlValue).AxisTitle.Text = "コミット数"
        End With
    End If
End Sub

'==============================================================================
' テスト用プロシージャ
'==============================================================================
Public Sub TestVisualizeGitLog()
    VisualizeGitLog
End Sub
