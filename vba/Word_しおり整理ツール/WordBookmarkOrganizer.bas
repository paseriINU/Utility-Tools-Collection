Attribute VB_Name = "WordBookmarkOrganizer"
Option Explicit

' ============================================================================
' Word しおり整理ツール - メインモジュール
' パターン検出、スタイル適用、連番チェック、PDF出力を行う
' ============================================================================

' === パターン設定構造体 ===
Private Type PatternConfig
    Level As Long
    Description As String
    RegexPattern As String
    StyleName As String
End Type

' === 連番トラッカー構造体 ===
Private Type SequenceTracker
    Level1Expected As Long
    Level2Expected As Long
    Level3Expected As Long
    Level4Expected As Long
    Warnings As Collection
End Type

' ============================================================================
' メインプロシージャ: Word文書のしおりを整理してPDF出力
' ============================================================================
Public Sub OrganizeWordBookmarks()
    Dim wordApp As Object           ' Word.Application
    Dim wordDoc As Object           ' Word.Document
    Dim para As Object              ' Word.Paragraph
    Dim filePath As String
    Dim outputWordPath As String
    Dim outputPdfPath As String
    Dim processedCount As Long
    Dim baseDir As String
    Dim inputDir As String
    Dim outputDir As String

    ' 設定
    Dim configs(1 To 4) As PatternConfig
    Dim exception1Style As String
    Dim exception2Style As String
    Dim doSequenceCheck As Boolean
    Dim doPdfOutput As Boolean

    ' 連番トラッカー
    Dim tracker As SequenceTracker
    Set tracker.Warnings = New Collection
    tracker.Level1Expected = 1
    tracker.Level2Expected = 1
    tracker.Level3Expected = 1
    tracker.Level4Expected = 1

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

    ' Excelシートから設定を読み込み
    If Not LoadSettings(configs, exception1Style, exception2Style, doSequenceCheck, doPdfOutput) Then
        MsgBox "設定の読み込みに失敗しました。" & vbCrLf & _
               "シートを初期化してください。", vbExclamation
        Exit Sub
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

    ' レベル1-4のスタイル名リスト（例外処理用）
    Dim levelStyles As Collection
    Set levelStyles = New Collection
    Dim i As Long
    For i = 1 To 4
        If configs(i).StyleName <> "" Then
            levelStyles.Add configs(i).StyleName
        End If
    Next i

    ' すべての段落をループ
    For Each para In wordDoc.Paragraphs
        processedCount = processedCount + ProcessParagraph(para, configs, exception1Style, exception2Style, _
                                                           levelStyles, wordDoc, doSequenceCheck, tracker)
    Next para

    ' 図形（Shape）内のテキストも処理
    Dim shp As Object
    Dim shapePara As Object
    For Each shp In wordDoc.Shapes
        On Error Resume Next
        If shp.TextFrame.HasText Then
            For Each shapePara In shp.TextFrame.TextRange.Paragraphs
                processedCount = processedCount + ProcessParagraph(shapePara, configs, exception1Style, exception2Style, _
                                                                   levelStyles, wordDoc, doSequenceCheck, tracker)
            Next shapePara
        End If
        Err.Clear
        On Error GoTo ErrorHandler
    Next shp

    ' Outputフォルダに名前を付けて保存
    wordDoc.SaveAs2 outputWordPath

    ' PDFとしてエクスポート
    If doPdfOutput Then
        Debug.Print "========================================="
        Debug.Print "PDFをエクスポートしています..."
        wordDoc.ExportAsFixedFormat _
            OutputFileName:=outputPdfPath, _
            ExportFormat:=17, _
            OpenAfterExport:=False, _
            OptimizeFor:=0, _
            CreateBookmarks:=1

        Debug.Print "PDFを出力しました: " & outputPdfPath
    End If

    Debug.Print "Word文書を出力しました: " & outputWordPath
    Debug.Print "========================================="
    Debug.Print "処理完了: " & processedCount & " 個の見出しを処理しました"
    Debug.Print "========================================="

    ' Word文書を閉じる
    wordDoc.Close SaveChanges:=False
    wordApp.Quit

    Set wordDoc = Nothing
    Set wordApp = Nothing

    ' 完了メッセージの作成
    Dim msg As String
    msg = "しおりの整理が完了しました。" & vbCrLf & vbCrLf & _
          "処理件数: " & processedCount & " 個" & vbCrLf & _
          "Word出力先: " & outputWordPath

    If doPdfOutput Then
        msg = msg & vbCrLf & "PDF出力先: " & outputPdfPath
    End If

    ' 連番チェック警告があれば追加
    If doSequenceCheck And tracker.Warnings.Count > 0 Then
        msg = msg & vbCrLf & vbCrLf & "■ 連番チェック警告:" & vbCrLf
        Dim warning As Variant
        For Each warning In tracker.Warnings
            msg = msg & "  - " & warning & vbCrLf
        Next
    End If

    MsgBox msg, vbInformation, "処理完了"

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

' ============================================================================
' 段落処理（共通関数）
' 戻り値: 処理された場合は1、それ以外は0
' ============================================================================
Private Function ProcessParagraph(ByRef para As Object, _
                                  ByRef configs() As PatternConfig, _
                                  ByVal exception1Style As String, _
                                  ByVal exception2Style As String, _
                                  ByRef levelStyles As Collection, _
                                  ByRef wordDoc As Object, _
                                  ByVal doSequenceCheck As Boolean, _
                                  ByRef tracker As SequenceTracker) As Long
    On Error GoTo ErrorHandler

    If para Is Nothing Then
        ProcessParagraph = 0
        Exit Function
    End If

    If para.Range Is Nothing Then
        ProcessParagraph = 0
        Exit Function
    End If

    Dim paraText As String
    Dim currentStyle As String
    Dim currentOutline As Long
    Dim detectedLevel As Long
    Dim targetStyle As String
    Dim matchedNumber As Long

    ' 段落テキストを取得（改行を除去）
    paraText = Trim(Replace(para.Range.Text, vbCr, ""))
    paraText = Replace(paraText, Chr(13), "")

    ' 空の段落はスキップ
    If paraText = "" Then
        ProcessParagraph = 0
        Exit Function
    End If

    ' 現在のスタイル名を取得
    On Error Resume Next
    currentStyle = para.Style.NameLocal
    If Err.Number <> 0 Then
        currentStyle = ""
        Err.Clear
    End If
    On Error GoTo ErrorHandler

    ' 現在のアウトラインレベルを取得
    On Error Resume Next
    currentOutline = para.OutlineLevel
    If Err.Number <> 0 Then
        currentOutline = 10  ' 本文レベル
        Err.Clear
    End If
    On Error GoTo ErrorHandler

    ' パターン検出（レベル4→3→2→1の順で評価）
    detectedLevel = 0
    targetStyle = ""
    matchedNumber = 0

    ' レベル4: X-X,X
    If detectedLevel = 0 And configs(4).RegexPattern <> "" Then
        If MatchPattern(paraText, configs(4).RegexPattern) Then
            detectedLevel = 4
            targetStyle = configs(4).StyleName
            matchedNumber = ExtractLastNumber(paraText, configs(4).RegexPattern)
        End If
    End If

    ' レベル3: X-X
    If detectedLevel = 0 And configs(3).RegexPattern <> "" Then
        If MatchPattern(paraText, configs(3).RegexPattern) Then
            detectedLevel = 3
            targetStyle = configs(3).StyleName
            matchedNumber = ExtractSecondNumber(paraText, configs(3).RegexPattern)
        End If
    End If

    ' レベル2: 第X章
    If detectedLevel = 0 And configs(2).RegexPattern <> "" Then
        If MatchPattern(paraText, configs(2).RegexPattern) Then
            detectedLevel = 2
            targetStyle = configs(2).StyleName
            matchedNumber = ExtractFirstNumber(paraText, configs(2).RegexPattern)
        End If
    End If

    ' レベル1: 第X部
    If detectedLevel = 0 And configs(1).RegexPattern <> "" Then
        If MatchPattern(paraText, configs(1).RegexPattern) Then
            detectedLevel = 1
            targetStyle = configs(1).StyleName
            matchedNumber = ExtractFirstNumber(paraText, configs(1).RegexPattern)
        End If
    End If

    ' 例外1: パターン外だが既にレベル1-4のスタイルが適用されている
    If detectedLevel = 0 And exception1Style <> "" Then
        If IsStyleInList(currentStyle, levelStyles) Then
            detectedLevel = -1  ' 例外1
            targetStyle = exception1Style
        End If
    End If

    ' 例外2: アウトライン設定済み（段落またはスタイル）
    If detectedLevel = 0 And exception2Style <> "" Then
        ' 段落のOutlineLevelが1-9の場合、またはスタイルにアウトラインが定義されている場合
        If (currentOutline >= 1 And currentOutline <= 9) Then
            detectedLevel = -2  ' 例外2
            targetStyle = exception2Style
        ElseIf HasOutlineDefinedInStyle(para, wordDoc) Then
            detectedLevel = -2
            targetStyle = exception2Style
        End If
    End If

    ' スタイル適用
    If detectedLevel <> 0 And targetStyle <> "" Then
        ApplyStyle para, targetStyle
        Debug.Print "[レベル" & detectedLevel & "] " & Left(paraText, 50)

        ' 連番チェック（レベル1-4のみ）
        If doSequenceCheck And detectedLevel > 0 Then
            TrackSequence tracker, detectedLevel, matchedNumber, paraText
        End If

        ProcessParagraph = 1
        Exit Function
    End If

    ProcessParagraph = 0
    Exit Function

ErrorHandler:
    ProcessParagraph = 0
End Function

' ============================================================================
' Excelシートから設定を読み込み
' ============================================================================
Private Function LoadSettings(ByRef configs() As PatternConfig, _
                              ByRef exception1Style As String, _
                              ByRef exception2Style As String, _
                              ByRef doSequenceCheck As Boolean, _
                              ByRef doPdfOutput As Boolean) As Boolean
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_MAIN)

    ' レベル1-4の設定を読み込み
    configs(1).Level = 1
    configs(1).Description = "第X部"
    configs(1).RegexPattern = CStr(ws.Cells(ROW_PATTERN_LEVEL1, COL_REGEX).Value)
    configs(1).StyleName = CStr(ws.Cells(ROW_PATTERN_LEVEL1, COL_STYLE_NAME).Value)

    configs(2).Level = 2
    configs(2).Description = "第X章"
    configs(2).RegexPattern = CStr(ws.Cells(ROW_PATTERN_LEVEL2, COL_REGEX).Value)
    configs(2).StyleName = CStr(ws.Cells(ROW_PATTERN_LEVEL2, COL_STYLE_NAME).Value)

    configs(3).Level = 3
    configs(3).Description = "X-X"
    configs(3).RegexPattern = CStr(ws.Cells(ROW_PATTERN_LEVEL3, COL_REGEX).Value)
    configs(3).StyleName = CStr(ws.Cells(ROW_PATTERN_LEVEL3, COL_STYLE_NAME).Value)

    configs(4).Level = 4
    configs(4).Description = "X-X,X"
    configs(4).RegexPattern = CStr(ws.Cells(ROW_PATTERN_LEVEL4, COL_REGEX).Value)
    configs(4).StyleName = CStr(ws.Cells(ROW_PATTERN_LEVEL4, COL_STYLE_NAME).Value)

    ' 例外スタイルを読み込み
    exception1Style = CStr(ws.Cells(ROW_PATTERN_EXCEPTION1, COL_STYLE_NAME).Value)
    exception2Style = CStr(ws.Cells(ROW_PATTERN_EXCEPTION2, COL_STYLE_NAME).Value)

    ' オプション設定を読み込み
    doSequenceCheck = (CStr(ws.Cells(ROW_OPTION_SEQUENCE_CHECK, COL_OPTION_VALUE).Value) = "はい")
    doPdfOutput = (CStr(ws.Cells(ROW_OPTION_PDF_OUTPUT, COL_OPTION_VALUE).Value) = "はい")

    LoadSettings = True
    Exit Function

ErrorHandler:
    LoadSettings = False
End Function

' ============================================================================
' 正規表現パターンマッチ
' ============================================================================
Private Function MatchPattern(ByVal text As String, ByVal pattern As String) As Boolean
    On Error GoTo ErrorHandler

    If text = "" Or pattern = "" Then
        MatchPattern = False
        Exit Function
    End If

    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")

    With regex
        .Global = False
        .IgnoreCase = False
        .MultiLine = False
        .pattern = pattern
    End With

    MatchPattern = regex.Test(text)

    Set regex = Nothing
    Exit Function

ErrorHandler:
    MatchPattern = False
End Function

' ============================================================================
' 数値抽出（最初の数値）- 全角対応
' ============================================================================
Private Function ExtractFirstNumber(ByVal text As String, ByVal pattern As String) As Long
    On Error GoTo ErrorHandler

    ' 全角数字を半角に変換
    text = ConvertToHalfWidth(text)

    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = "[0-9]+"
    regex.Global = False

    Dim matches As Object
    Set matches = regex.Execute(text)

    If matches.Count > 0 Then
        ExtractFirstNumber = CLng(matches(0).Value)
    Else
        ExtractFirstNumber = 0
    End If

    Set regex = Nothing
    Exit Function

ErrorHandler:
    ExtractFirstNumber = 0
End Function

' ============================================================================
' 数値抽出（2番目の数値）- X-X パターン用、全角対応
' ============================================================================
Private Function ExtractSecondNumber(ByVal text As String, ByVal pattern As String) As Long
    On Error GoTo ErrorHandler

    ' 全角数字を半角に変換
    text = ConvertToHalfWidth(text)

    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = "[0-9]+"
    regex.Global = True

    Dim matches As Object
    Set matches = regex.Execute(text)

    If matches.Count >= 2 Then
        ExtractSecondNumber = CLng(matches(1).Value)
    ElseIf matches.Count = 1 Then
        ExtractSecondNumber = CLng(matches(0).Value)
    Else
        ExtractSecondNumber = 0
    End If

    Set regex = Nothing
    Exit Function

ErrorHandler:
    ExtractSecondNumber = 0
End Function

' ============================================================================
' 数値抽出（最後の数値）- X-X,X パターン用、全角対応
' ============================================================================
Private Function ExtractLastNumber(ByVal text As String, ByVal pattern As String) As Long
    On Error GoTo ErrorHandler

    ' 全角数字を半角に変換
    text = ConvertToHalfWidth(text)

    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = "[0-9]+"
    regex.Global = True

    Dim matches As Object
    Set matches = regex.Execute(text)

    If matches.Count > 0 Then
        ExtractLastNumber = CLng(matches(matches.Count - 1).Value)
    Else
        ExtractLastNumber = 0
    End If

    Set regex = Nothing
    Exit Function

ErrorHandler:
    ExtractLastNumber = 0
End Function

' ============================================================================
' 全角を半角に変換（数字のみ）
' ============================================================================
Private Function ConvertToHalfWidth(ByVal text As String) As String
    Dim i As Long
    Dim char As String
    Dim result As String

    result = ""
    For i = 1 To Len(text)
        char = Mid(text, i, 1)
        Select Case AscW(char)
            Case &HFF10 To &HFF19  ' ０-９
                result = result & Chr(AscW(char) - &HFF10 + Asc("0"))
            Case Else
                result = result & char
        End Select
    Next i

    ConvertToHalfWidth = result
End Function

' ============================================================================
' スタイルがリストに含まれるかチェック
' ============================================================================
Private Function IsStyleInList(ByVal styleName As String, ByRef styleList As Collection) As Boolean
    On Error GoTo ErrorHandler

    Dim item As Variant
    For Each item In styleList
        If styleName = CStr(item) Then
            IsStyleInList = True
            Exit Function
        End If
    Next

    IsStyleInList = False
    Exit Function

ErrorHandler:
    IsStyleInList = False
End Function

' ============================================================================
' スタイルにアウトラインが定義されているかチェック
' ============================================================================
Private Function HasOutlineDefinedInStyle(ByRef para As Object, ByRef wordDoc As Object) As Boolean
    On Error GoTo ErrorHandler

    Dim styleName As String
    styleName = para.Style.NameLocal

    Dim style As Object
    Set style = wordDoc.Styles(styleName)

    ' スタイルのOutlineLevelが本文(10)以外なら、アウトライン定義あり
    If style.ParagraphFormat.OutlineLevel >= 1 And _
       style.ParagraphFormat.OutlineLevel <= 9 Then
        HasOutlineDefinedInStyle = True
    Else
        HasOutlineDefinedInStyle = False
    End If

    Exit Function

ErrorHandler:
    HasOutlineDefinedInStyle = False
End Function

' ============================================================================
' スタイル適用
' ============================================================================
Private Sub ApplyStyle(ByRef para As Object, ByVal styleName As String)
    On Error Resume Next

    para.Style = styleName

    If Err.Number <> 0 Then
        Debug.Print "[警告] スタイル '" & styleName & "' が見つかりません: " & _
                    Left(para.Range.Text, 50)
        Err.Clear
    End If

    On Error GoTo 0
End Sub

' ============================================================================
' 連番トラッキング
' ============================================================================
Private Sub TrackSequence(ByRef tracker As SequenceTracker, _
                          ByVal level As Long, _
                          ByVal actualNum As Long, _
                          ByVal paraText As String)
    Dim expectedNum As Long
    Dim warningMsg As String

    Select Case level
        Case 1
            expectedNum = tracker.Level1Expected
            If actualNum <> expectedNum Then
                warningMsg = "レベル1: 期待=" & expectedNum & ", 実際=" & actualNum & _
                             " (" & Left(paraText, 30) & "...)"
                tracker.Warnings.Add warningMsg
            End If
            tracker.Level1Expected = actualNum + 1
            ' 下位レベルをリセット
            tracker.Level2Expected = 1
            tracker.Level3Expected = 1
            tracker.Level4Expected = 1

        Case 2
            expectedNum = tracker.Level2Expected
            If actualNum <> expectedNum Then
                warningMsg = "レベル2: 期待=" & expectedNum & ", 実際=" & actualNum & _
                             " (" & Left(paraText, 30) & "...)"
                tracker.Warnings.Add warningMsg
            End If
            tracker.Level2Expected = actualNum + 1
            ' 下位レベルをリセット
            tracker.Level3Expected = 1
            tracker.Level4Expected = 1

        Case 3
            expectedNum = tracker.Level3Expected
            If actualNum <> expectedNum Then
                warningMsg = "レベル3: 期待=" & expectedNum & ", 実際=" & actualNum & _
                             " (" & Left(paraText, 30) & "...)"
                tracker.Warnings.Add warningMsg
            End If
            tracker.Level3Expected = actualNum + 1
            ' 下位レベルをリセット
            tracker.Level4Expected = 1

        Case 4
            expectedNum = tracker.Level4Expected
            If actualNum <> expectedNum Then
                warningMsg = "レベル4: 期待=" & expectedNum & ", 実際=" & actualNum & _
                             " (" & Left(paraText, 30) & "...)"
                tracker.Warnings.Add warningMsg
            End If
            tracker.Level4Expected = actualNum + 1
    End Select
End Sub

' ============================================================================
' Inputフォルダから処理対象のWord文書を選択
' ============================================================================
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

' ============================================================================
' テスト用: イミディエイトウィンドウに情報を出力
' ============================================================================
Public Sub TestOrganizeWordBookmarks()
    ' イミディエイトウィンドウを開いた状態でこのマクロを実行してください
    OrganizeWordBookmarks
End Sub
