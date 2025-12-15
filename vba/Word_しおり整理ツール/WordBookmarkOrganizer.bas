'==============================================================================
' Word文書のしおり（ブックマーク）整理ツール
'==============================================================================
' 概要:
'   Excelから実行し、Word文書のスタイル「表題1」「表題2」「表題3」に基づいて
'   アウトラインレベルを設定し、PDFエクスポート時に正しいしおりを生成します。
'
' 必要な参照設定:
'   - Microsoft Word XX.X Object Library
'
' 作成日: 2025-12-06
'==============================================================================

Option Explicit

'==============================================================================
' メインプロシージャ: Word文書のしおりを整理してPDF出力
'==============================================================================
Public Sub OrganizeWordBookmarks()
    Dim wordApp As Object           ' Word.Application
    Dim wordDoc As Object           ' Word.Document
    Dim para As Object              ' Word.Paragraph
    Dim filePath As String
    Dim outputWordPath As String
    Dim outputPdfPath As String
    Dim processedCount As Long
    Dim styleName As String
    Dim outlineLevel As Long
    Dim baseDir As String
    Dim inputDir As String
    Dim outputDir As String

    ' マクロのあるフォルダを基準にする
    baseDir = ThisWorkbook.Path
    If Right(baseDir, 1) <> "\" Then baseDir = baseDir & "\"

    inputDir = baseDir & "Input\"
    outputDir = baseDir & "Output\"

    ' Inputフォルダの存在確認と作成
    If Dir(inputDir, vbDirectory) = "" Then
        MkDir inputDir
        MsgBox "Inputフォルダを作成しました: " & vbCrLf & inputDir & vbCrLf & vbCrLf & _
               "このフォルダに処理したいWord文書を配置してください。", vbInformation
        Exit Sub
    End If

    ' Outputフォルダの存在確認と作成
    If Dir(outputDir, vbDirectory) = "" Then
        MkDir outputDir
    End If

    ' Inputフォルダから処理対象のWord文書を選択
    filePath = SelectWordFileFromInput(inputDir)
    If filePath = "" Then
        Exit Sub
    End If

    ' 出力ファイルパスを設定（Outputフォルダ）
    Dim fileName As String
    fileName = Mid(filePath, InStrRev(filePath, "\") + 1)
    outputWordPath = outputDir & fileName
    outputPdfPath = outputDir & Left(fileName, InStrRev(fileName, ".") - 1) & ".pdf"

    On Error GoTo ErrorHandler

    ' Wordアプリケーションを起動
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False  ' バックグラウンドで実行

    ' Word文書を開く
    Set wordDoc = wordApp.Documents.Open(filePath)

    ' 処理開始メッセージ
    Debug.Print "========================================="
    Debug.Print "Word文書のしおり整理を開始します"
    Debug.Print "対象ファイル: " & filePath
    Debug.Print "========================================="

    processedCount = 0

    ' すべての段落をループ
    For Each para In wordDoc.Paragraphs
        ' 段落が存在し、スタイルが設定されている場合のみ処理
        If Not para.Range Is Nothing Then
            ' スタイル名を取得
            On Error Resume Next
            styleName = para.Style.NameLocal
            If Err.Number <> 0 Then
                styleName = ""
                Err.Clear
            End If
            On Error GoTo ErrorHandler

            ' スタイル名に応じてアウトラインレベルを設定
            outlineLevel = 0

            Select Case styleName
                Case "表題1"
                    outlineLevel = 1
                Case "表題2"
                    outlineLevel = 2
                Case "表題3"
                    outlineLevel = 3
            End Select

            ' アウトラインレベルを設定
            If outlineLevel > 0 Then
                para.OutlineLevel = outlineLevel
                processedCount = processedCount + 1
                Debug.Print "[" & outlineLevel & "] " & Left(para.Range.Text, 50)
            End If
        End If
    Next para

    ' Outputフォルダに名前を付けて保存
    wordDoc.SaveAs2 outputWordPath

    ' PDFとしてエクスポート
    Debug.Print "========================================="
    Debug.Print "PDFをエクスポートしています..."
    wordDoc.ExportAsFixedFormat _
        OutputFileName:=outputPdfPath, _
        ExportFormat:=17, _
        OpenAfterExport:=False, _
        OptimizeFor:=0, _
        CreateBookmarks:=1

    Debug.Print "Word文書を出力しました: " & outputWordPath
    Debug.Print "PDFを出力しました: " & outputPdfPath
    Debug.Print "========================================="
    Debug.Print "処理完了: " & processedCount & " 個の見出しを処理しました"
    Debug.Print "========================================="

    ' Word文書を閉じる
    wordDoc.Close SaveChanges:=False
    wordApp.Quit

    Set wordDoc = Nothing
    Set wordApp = Nothing

    ' 完了メッセージ
    MsgBox "しおりの整理とPDF出力が完了しました。" & vbCrLf & vbCrLf & _
           "処理件数: " & processedCount & " 個" & vbCrLf & _
           "Word出力先: " & outputWordPath & vbCrLf & _
           "PDF出力先: " & outputPdfPath, vbInformation, "処理完了"

    Exit Sub

ErrorHandler:
    ' エラー処理
    MsgBox "エラーが発生しました: " & vbCrLf & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, vbCritical, "エラー"

    ' Wordオブジェクトのクリーンアップ
    On Error Resume Next
    If Not wordDoc Is Nothing Then
        wordDoc.Close SaveChanges:=False
    End If
    If Not wordApp Is Nothing Then
        wordApp.Quit
    End If
    Set wordDoc = Nothing
    Set wordApp = Nothing
    On Error GoTo 0
End Sub

'==============================================================================
' Inputフォルダから処理対象のWord文書を選択
'==============================================================================
Private Function SelectWordFileFromInput(ByVal inputDir As String) As String
    Dim fileList() As String
    Dim fileCount As Long
    Dim fileName As String
    Dim i As Long
    Dim selectedIndex As Long
    Dim msg As String

    ' 配列の初期化
    fileCount = 0
    ReDim fileList(0 To 99)  ' 最大100ファイルまで対応

    ' .docxファイルを検索
    fileName = Dir(inputDir & "*.docx")
    Do While fileName <> ""
        fileList(fileCount) = fileName
        fileCount = fileCount + 1
        fileName = Dir()
    Loop

    ' .docファイルを検索
    fileName = Dir(inputDir & "*.doc")
    Do While fileName <> ""
        ' .docxと重複しないようにチェック
        Dim isDuplicate As Boolean
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

    ' ファイルが見つからない場合
    If fileCount = 0 Then
        MsgBox "Inputフォルダ内にWord文書(.docx/.doc)が見つかりませんでした。" & vbCrLf & vbCrLf & _
               "フォルダ: " & inputDir, vbExclamation, "ファイルなし"
        SelectWordFileFromInput = ""
        Exit Function
    End If

    ' ファイルが1つだけの場合は自動選択
    If fileCount = 1 Then
        SelectWordFileFromInput = inputDir & fileList(0)
        Exit Function
    End If

    ' 複数ファイルがある場合は選択させる
    msg = "処理するWord文書を選択してください:" & vbCrLf & vbCrLf
    For i = 0 To fileCount - 1
        msg = msg & (i + 1) & ". " & fileList(i) & vbCrLf
    Next i
    msg = msg & vbCrLf & "番号を入力してください (1-" & fileCount & "):"

    Dim userInput As String
    userInput = InputBox(msg, "ファイル選択", "1")

    ' キャンセルされた場合
    If userInput = "" Then
        SelectWordFileFromInput = ""
        Exit Function
    End If

    ' 入力値の検証
    If Not IsNumeric(userInput) Then
        MsgBox "無効な入力です。数値を入力してください。", vbExclamation
        SelectWordFileFromInput = ""
        Exit Function
    End If

    selectedIndex = CLng(userInput) - 1
    If selectedIndex < 0 Or selectedIndex >= fileCount Then
        MsgBox "無効な番号です。1から" & fileCount & "の間で入力してください。", vbExclamation
        SelectWordFileFromInput = ""
        Exit Function
    End If

    ' 選択されたファイルのフルパスを返す
    SelectWordFileFromInput = inputDir & fileList(selectedIndex)
End Function

'==============================================================================
' メインシート初期化
'==============================================================================
Public Sub InitializeWordBookmarkOrganizer()
    Dim ws As Worksheet
    Dim mainSheetName As String

    mainSheetName = "WordBookmarkOrganizer"

    On Error Resume Next
    Application.DisplayAlerts = False

    ' 既存のメインシートがあれば削除
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = mainSheetName Then
            ws.Delete
            Exit For
        End If
    Next ws

    Application.DisplayAlerts = True
    On Error GoTo 0

    ' 新しいシートを作成
    Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
    ws.Name = mainSheetName

    ' シートを初期化
    FormatBookmarkMainSheet ws

    MsgBox "WordBookmarkOrganizerシートを初期化しました。", vbInformation, "初期化完了"
End Sub

'==============================================================================
' メインシートのフォーマット
'==============================================================================
Private Sub FormatBookmarkMainSheet(ByRef ws As Worksheet)
    Dim btn As Button
    Dim baseDir As String

    baseDir = ThisWorkbook.Path
    If Right(baseDir, 1) <> "\" Then baseDir = baseDir & "\"

    Application.ScreenUpdating = False

    With ws
        ' 全体の背景色を白に
        .Cells.Interior.Color = RGB(255, 255, 255)

        ' =================================================================
        ' タイトルエリア (行1-3)
        ' =================================================================
        .Range("B2:H2").Merge
        .Range("B2").Value = "Word しおり整理ツール"
        With .Range("B2")
            .Font.Name = "Meiryo UI"
            .Font.Size = 20
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        .Range("B2:H3").Interior.Color = RGB(91, 155, 213)
        .Rows(2).RowHeight = 40
        .Rows(3).RowHeight = 5

        ' =================================================================
        ' 説明エリア (行5-7)
        ' =================================================================
        .Range("B5:H5").Merge
        .Range("B5").Value = "Word文書のスタイル（表題1〜3）に基づいてアウトラインレベルを設定し、"
        With .Range("B5")
            .Font.Name = "Meiryo UI"
            .Font.Size = 11
            .Font.Color = RGB(64, 64, 64)
        End With

        .Range("B6:H6").Merge
        .Range("B6").Value = "PDFエクスポート時に正しいしおり（ブックマーク）を生成します。"
        With .Range("B6")
            .Font.Name = "Meiryo UI"
            .Font.Size = 11
            .Font.Color = RGB(64, 64, 64)
        End With

        ' =================================================================
        ' フォルダ設定セクション (行9-15)
        ' =================================================================
        .Range("B9:H9").Merge
        .Range("B9").Value = "フォルダ設定"
        With .Range("B9")
            .Font.Name = "Meiryo UI"
            .Font.Size = 14
            .Font.Bold = True
            .Font.Color = RGB(91, 155, 213)
        End With

        ' セクション下線
        With .Range("B9:H9").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(91, 155, 213)
            .Weight = xlMedium
        End With

        ' 入力フォルダ
        .Range("B11").Value = "入力フォルダ:"
        With .Range("B11")
            .Font.Name = "Meiryo UI"
            .Font.Size = 11
            .Font.Bold = True
        End With

        .Range("C11:G11").Merge
        .Range("C11").Value = baseDir & "Input\"
        With .Range("C11:G11")
            .Interior.Color = RGB(242, 242, 242)
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
            .HorizontalAlignment = xlLeft
        End With
        With .Range("C11:G11").Borders
            .LineStyle = xlContinuous
            .Color = RGB(200, 200, 200)
        End With

        ' 出力フォルダ
        .Range("B13").Value = "出力フォルダ:"
        With .Range("B13")
            .Font.Name = "Meiryo UI"
            .Font.Size = 11
            .Font.Bold = True
        End With

        .Range("C13:G13").Merge
        .Range("C13").Value = baseDir & "Output\"
        With .Range("C13:G13")
            .Interior.Color = RGB(242, 242, 242)
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
            .HorizontalAlignment = xlLeft
        End With
        With .Range("C13:G13").Borders
            .LineStyle = xlContinuous
            .Color = RGB(200, 200, 200)
        End With

        ' =================================================================
        ' ボタンエリア (行16-18)
        ' =================================================================
        .Rows(16).RowHeight = 10

        ' 整理実行ボタン
        Set btn = .Buttons.Add(.Range("C17").Left, .Range("C17").Top, 180, 35)
        With btn
            .Name = "btnOrganizeBookmarks"
            .Caption = "しおりを整理してPDF出力"
            .OnAction = "OrganizeWordBookmarks"
            .Font.Name = "Meiryo UI"
            .Font.Size = 11
            .Font.Bold = True
        End With

        ' =================================================================
        ' 対応スタイルセクション (行21-28)
        ' =================================================================
        .Range("B21:H21").Merge
        .Range("B21").Value = "対応スタイル（アウトラインレベル）"
        With .Range("B21")
            .Font.Name = "Meiryo UI"
            .Font.Size = 14
            .Font.Bold = True
            .Font.Color = RGB(91, 155, 213)
        End With

        ' セクション下線
        With .Range("B21:H21").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(91, 155, 213)
            .Weight = xlMedium
        End With

        ' 表題1
        .Range("B23").Value = "表題1"
        With .Range("B23")
            .Font.Name = "Meiryo UI"
            .Font.Size = 11
            .Font.Bold = True
            .Font.Color = RGB(91, 155, 213)
        End With
        .Range("C23").Value = "→"
        With .Range("C23")
            .Font.Name = "Meiryo UI"
            .Font.Size = 11
            .HorizontalAlignment = xlCenter
        End With
        .Range("D23:E23").Merge
        .Range("D23").Value = "レベル1（大見出し）"
        With .Range("D23")
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
        End With

        ' 表題2
        .Range("B24").Value = "表題2"
        With .Range("B24")
            .Font.Name = "Meiryo UI"
            .Font.Size = 11
            .Font.Bold = True
            .Font.Color = RGB(91, 155, 213)
        End With
        .Range("C24").Value = "→"
        With .Range("C24")
            .Font.Name = "Meiryo UI"
            .Font.Size = 11
            .HorizontalAlignment = xlCenter
        End With
        .Range("D24:E24").Merge
        .Range("D24").Value = "レベル2（中見出し）"
        With .Range("D24")
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
        End With

        ' 表題3
        .Range("B25").Value = "表題3"
        With .Range("B25")
            .Font.Name = "Meiryo UI"
            .Font.Size = 11
            .Font.Bold = True
            .Font.Color = RGB(91, 155, 213)
        End With
        .Range("C25").Value = "→"
        With .Range("C25")
            .Font.Name = "Meiryo UI"
            .Font.Size = 11
            .HorizontalAlignment = xlCenter
        End With
        .Range("D25:E25").Merge
        .Range("D25").Value = "レベル3（小見出し）"
        With .Range("D25")
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
        End With

        ' =================================================================
        ' 使い方セクション (行28-36)
        ' =================================================================
        .Range("B28:H28").Merge
        .Range("B28").Value = "使い方"
        With .Range("B28")
            .Font.Name = "Meiryo UI"
            .Font.Size = 14
            .Font.Bold = True
            .Font.Color = RGB(91, 155, 213)
        End With

        ' セクション下線
        With .Range("B28:H28").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(91, 155, 213)
            .Weight = xlMedium
        End With

        .Range("B30").Value = "1."
        .Range("C30").Value = "処理したいWord文書をInputフォルダに配置"
        .Range("B31").Value = "2."
        .Range("C31").Value = "「しおりを整理してPDF出力」ボタンをクリック"
        .Range("B32").Value = "3."
        .Range("C32").Value = "複数ファイルがある場合は番号で選択"
        .Range("B33").Value = "4."
        .Range("C33").Value = "処理完了後、OutputフォルダにWord文書とPDFが出力されます"

        .Range("B30:B33").Font.Name = "Meiryo UI"
        .Range("B30:B33").Font.Size = 10
        .Range("B30:B33").Font.Bold = True
        .Range("B30:B33").Font.Color = RGB(91, 155, 213)
        .Range("C30:C33").Font.Name = "Meiryo UI"
        .Range("C30:C33").Font.Size = 10

        ' =================================================================
        ' 出力ファイルセクション (行36-42)
        ' =================================================================
        .Range("B36:H36").Merge
        .Range("B36").Value = "出力ファイル"
        With .Range("B36")
            .Font.Name = "Meiryo UI"
            .Font.Size = 14
            .Font.Bold = True
            .Font.Color = RGB(91, 155, 213)
        End With

        ' セクション下線
        With .Range("B36:H36").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(91, 155, 213)
            .Weight = xlMedium
        End With

        .Range("B38").Value = "・"
        .Range("C38:G38").Merge
        .Range("C38").Value = "Word文書（.docx）- アウトラインレベルが設定された文書"
        With .Range("C38")
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
        End With

        .Range("B39").Value = "・"
        .Range("C39:G39").Merge
        .Range("C39").Value = "PDFファイル（.pdf）- しおり付きPDF"
        With .Range("C39")
            .Font.Name = "Meiryo UI"
            .Font.Size = 10
        End With

        .Range("B38:B39").Font.Name = "Meiryo UI"
        .Range("B38:B39").Font.Size = 10
        .Range("B38:B39").Font.Bold = True

        ' =================================================================
        ' 必要環境セクション (行42-48)
        ' =================================================================
        .Range("B42:H42").Merge
        .Range("B42").Value = "必要な環境"
        With .Range("B42")
            .Font.Name = "Meiryo UI"
            .Font.Size = 14
            .Font.Bold = True
            .Font.Color = RGB(91, 155, 213)
        End With

        ' セクション下線
        With .Range("B42:H42").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(91, 155, 213)
            .Weight = xlMedium
        End With

        .Range("B44").Value = "・"
        .Range("C44").Value = "Microsoft Excel 2010以降"
        .Range("B45").Value = "・"
        .Range("C45").Value = "Microsoft Word 2010以降"

        .Range("B44:B45").Font.Name = "Meiryo UI"
        .Range("B44:B45").Font.Size = 10
        .Range("B44:B45").Font.Bold = True
        .Range("C44:C45").Font.Name = "Meiryo UI"
        .Range("C44:C45").Font.Size = 10

        ' =================================================================
        ' 対応ファイル形式セクション (行48-52)
        ' =================================================================
        .Range("B48:H48").Merge
        .Range("B48").Value = "対応ファイル形式"
        With .Range("B48")
            .Font.Name = "Meiryo UI"
            .Font.Size = 14
            .Font.Bold = True
            .Font.Color = RGB(91, 155, 213)
        End With

        ' セクション下線
        With .Range("B48:H48").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(91, 155, 213)
            .Weight = xlMedium
        End With

        .Range("B50").Value = "入力:"
        .Range("C50").Value = ".docx, .doc"
        .Range("B50").Font.Name = "Meiryo UI"
        .Range("B50").Font.Size = 10
        .Range("B50").Font.Bold = True
        .Range("C50").Font.Name = "Meiryo UI"
        .Range("C50").Font.Size = 10

        ' =================================================================
        ' 列幅・行高の調整
        ' =================================================================
        .Columns("A").ColumnWidth = 3
        .Columns("B").ColumnWidth = 16
        .Columns("C").ColumnWidth = 15
        .Columns("D").ColumnWidth = 12
        .Columns("E").ColumnWidth = 15
        .Columns("F").ColumnWidth = 12
        .Columns("G").ColumnWidth = 15
        .Columns("H").ColumnWidth = 12
        .Columns("I").ColumnWidth = 3

        ' セルA1を選択
        .Range("A1").Select
    End With

    Application.ScreenUpdating = True
End Sub

'==============================================================================
' テスト用: イミディエイトウィンドウに情報を出力
'==============================================================================
Public Sub TestOrganizeWordBookmarks()
    ' イミディエイトウィンドウを開いた状態でこのマクロを実行してください
    OrganizeWordBookmarks
End Sub
