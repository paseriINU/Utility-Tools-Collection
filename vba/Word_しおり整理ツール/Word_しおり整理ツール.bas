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
' 注意:
'   - 初期化とシートフォーマット機能は Word_しおり整理ツール_Setup.bas に分離されています
'
' 作成日: 2025-12-06
' 更新日: 2025-12-17 - セットアップモジュールを分離
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
' テスト用: イミディエイトウィンドウに情報を出力
'==============================================================================
Public Sub TestOrganizeWordBookmarks()
    ' イミディエイトウィンドウを開いた状態でこのマクロを実行してください
    OrganizeWordBookmarks
End Sub
