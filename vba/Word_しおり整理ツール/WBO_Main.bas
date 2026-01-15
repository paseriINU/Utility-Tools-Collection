Attribute VB_Name = "WBO_Main"
Option Explicit

' ============================================================================
' Word しおり整理ツール - メインモジュール
' エントリーポイント。各モジュールを呼び出すオーケストレーター
' ============================================================================

' ============================================================================
' メインプロシージャ: Word文書のしおりを整理してPDF出力
' ============================================================================
Public Sub OrganizeWordBookmarks()
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim wsSettings As Worksheet
    Dim filePath As String
    Dim outputWordPath As String
    Dim outputPdfPath As String
    Dim inputDir As String
    Dim outputDir As String
    Dim processedCount As Long
    Dim doPdfOutput As Boolean
    Dim hasSections As Boolean
    Dim isHyohyoDocument As Boolean
    Dim missingStyles As String

    ' 動的スタイル設定
    Dim styleSettings() As StyleSetting
    Dim styleCount As Long

    On Error GoTo ErrorHandler

    ' === 1. 設定シート取得 ===
    Set wsSettings = GetSettingsSheet()
    If wsSettings Is Nothing Then Exit Sub

    ' === 2. フォルダパス取得と検証 ===
    inputDir = GetInputFolder(wsSettings)
    outputDir = GetOutputFolder(wsSettings)

    If Not ValidateFolders(inputDir, outputDir) Then Exit Sub

    ' === 3. 設定読み込み ===
    If Not LoadSettings(wsSettings, styleSettings, styleCount, doPdfOutput) Then
        MsgBox "設定の読み込みに失敗しました。" & vbCrLf & _
               "設定シートを確認してください。", vbExclamation
        Exit Sub
    End If

    ' === 4. ファイル選択 ===
    filePath = SelectWordFile(inputDir)
    If filePath = "" Then Exit Sub

    ' 出力ファイルパスを設定
    Dim fileName As String
    fileName = Mid(filePath, InStrRev(filePath, "\") + 1)
    outputWordPath = outputDir & fileName
    outputPdfPath = outputDir & Left(fileName, InStrRev(fileName, ".") - 1) & ".pdf"

    ' === 5. Word起動・文書オープン ===
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    Set wordDoc = wordApp.Documents.Open(filePath)

    ' === 6. 文書判定 ===
    hasSections = HasSectionsInDoc(wordDoc)
    isHyohyoDocument = HasHyohyoOnPage1(wordDoc)

    ' === 7. スタイル検証 ===
    missingStyles = ValidateStyles(wordDoc, styleSettings, styleCount, hasSections)
    If missingStyles <> "" Then
        MsgBox "エラー: 以下のスタイルがWord文書に存在しません。" & vbCrLf & vbCrLf & _
               missingStyles & vbCrLf & _
               "処理を中止します。", vbCritical, "スタイルエラー"
        GoTo CleanUp
    End If

    ' === 8. 処理開始ログ ===
    Debug.Print "========================================="
    Debug.Print "Word文書のしおり整理を開始します"
    Debug.Print "対象ファイル: " & filePath
    Debug.Print "節構造: " & IIf(hasSections, "あり", "なし")
    Debug.Print "帳票文書: " & IIf(isHyohyoDocument, "あり", "なし")
    Debug.Print "スタイル設定数: " & styleCount
    Debug.Print "========================================="

    ' === 9. 文書処理 ===
    processedCount = ProcessDocument(wordDoc, styleSettings, styleCount, hasSections, isHyohyoDocument)

    ' === 10. ヘッダーフィールド更新 ===
    UpdateHeaderFields wordDoc, styleSettings, styleCount

    ' === 11. 保存・出力 ===
    SaveAndExport wordDoc, outputWordPath, outputPdfPath, doPdfOutput

    Debug.Print "========================================="
    Debug.Print "処理完了: " & processedCount & " 個の見出しを処理しました"
    Debug.Print "========================================="

    ' === 12. 完了メッセージ ===
    Dim msg As String
    msg = "しおりの整理が完了しました。" & vbCrLf & vbCrLf & _
          "処理件数: " & processedCount & " 個" & vbCrLf & _
          "Word出力先: " & outputWordPath

    If doPdfOutput Then
        msg = msg & vbCrLf & "PDF出力先: " & outputPdfPath
    End If

    MsgBox msg, vbInformation, "処理完了"

CleanUp:
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
    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & vbCrLf & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, vbCritical, "エラー"
    Resume CleanUp
End Sub

' ============================================================================
' テスト用: イミディエイトウィンドウに情報を出力
' ============================================================================
Public Sub TestOrganizeWordBookmarks()
    OrganizeWordBookmarks
End Sub
