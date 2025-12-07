Attribute VB_Name = "WordBookmarkOrganizer"
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
    Dim pdfPath As String
    Dim fileDialog As FileDialog
    Dim processedCount As Long
    Dim styleName As String
    Dim outlineLevel As Long

    ' ファイル選択ダイアログを表示
    Set fileDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fileDialog
        .Title = "しおりを整理するWord文書を選択してください"
        .Filters.Clear
        .Filters.Add "Word文書", "*.docx;*.doc"
        .AllowMultiSelect = False

        If .Show = -1 Then
            filePath = .SelectedItems(1)
        Else
            MsgBox "ファイルが選択されませんでした。処理を中止します。", vbExclamation
            Exit Sub
        End If
    End With

    ' Word文書が存在するか確認
    If Dir(filePath) = "" Then
        MsgBox "指定されたファイルが見つかりません: " & vbCrLf & filePath, vbCritical
        Exit Sub
    End If

    ' PDF出力パスを設定（同じフォルダ）
    pdfPath = Left(filePath, InStrRev(filePath, ".") - 1) & ".pdf"

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

    ' 変更を保存
    wordDoc.Save

    ' PDFとしてエクスポート
    Debug.Print "========================================="
    Debug.Print "PDFをエクスポートしています..."
    wordDoc.ExportAsFixedFormat _
        OutputFileName:=pdfPath, _
        ExportFormat:=17, _
        OpenAfterExport:=False, _
        OptimizeFor:=0, _
        CreateBookmarks:=1

    Debug.Print "PDFを出力しました: " & pdfPath
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
           "出力先: " & pdfPath, vbInformation, "処理完了"

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
' テスト用: イミディエイトウィンドウに情報を出力
'==============================================================================
Public Sub TestOrganizeWordBookmarks()
    ' イミディエイトウィンドウを開いた状態でこのマクロを実行してください
    OrganizeWordBookmarks
End Sub
