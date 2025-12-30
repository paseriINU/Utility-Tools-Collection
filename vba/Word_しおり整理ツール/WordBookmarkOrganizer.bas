Option Explicit

' ============================================================================
' Word しおり整理ツール - メインモジュール
' ヘッダー参照方式: セクションのヘッダーからパターンを抽出し、
' そのセクション内の本文で該当テキストを検索してスタイルを適用
' ============================================================================

' === スタイル設定構造体 ===
Private Type StyleConfig
    Level1Style As String  ' 第X部用（パターンマッチ）
    Level2Style As String  ' 第X章用（ヘッダー参照/フォールバック:パターンマッチ）
    Level3Style As String  ' 第X節用（ヘッダー参照、自動判定）
    Level4Style As String  ' X-X用（ヘッダー参照）
    Level5Style As String  ' X-X,X用（ヘッダー参照）
    Exception1Style As String
    Exception2Style As String
End Type

' === デフォルトスタイル名定数（ヘッダーのSTYLEREF更新用） ===
Private Const DEFAULT_LEVEL1_STYLE As String = "表題1"
Private Const DEFAULT_LEVEL2_STYLE As String = "表題2"
Private Const DEFAULT_LEVEL3_STYLE As String = "表題3"
Private Const DEFAULT_LEVEL4_STYLE As String = "表題4"
Private Const DEFAULT_LEVEL5_STYLE As String = "表題5"

' === ヘッダーパターン構造体 ===
Private Type HeaderPatterns
    Level2Text As String   ' 第X章
    Level3Text As String   ' 第X節
    Level4Text As String   ' X-X
    Level5Text As String   ' X-X,X
End Type

' ============================================================================
' メインプロシージャ: Word文書のしおりを整理してPDF出力
' ============================================================================
Public Sub OrganizeWordBookmarks()
    Dim wordApp As Object           ' Word.Application
    Dim wordDoc As Object           ' Word.Document
    Dim filePath As String
    Dim outputWordPath As String
    Dim outputPdfPath As String
    Dim processedCount As Long
    Dim baseDir As String
    Dim inputDir As String
    Dim outputDir As String

    ' 設定
    Dim styles As StyleConfig
    Dim doPdfOutput As Boolean

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
    If Not LoadSettings(styles, doPdfOutput) Then
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

    ' 文書内に「第X節」がヘッダーに存在するか判定
    Dim hasSections As Boolean
    hasSections = HasSectionsInDocument(wordDoc)

    ' スタイルの存在確認（節がない場合はLevel 5のチェックをスキップ）
    Dim missingStyles As String
    missingStyles = ValidateStyles(wordDoc, styles, hasSections)
    If missingStyles <> "" Then
        MsgBox "エラー: 以下のスタイルがWord文書に存在しません。" & vbCrLf & vbCrLf & _
               missingStyles & vbCrLf & _
               "処理を中止します。" & vbCrLf & vbCrLf & _
               "Word文書にスタイルを作成するか、" & vbCrLf & _
               "Excelシートのスタイル名を修正してください。", _
               vbCritical, "スタイルエラー"
        wordDoc.Close SaveChanges:=False
        wordApp.Quit
        Set wordDoc = Nothing
        Set wordApp = Nothing
        Exit Sub
    End If

    ' 処理開始メッセージ
    Debug.Print "========================================="
    Debug.Print "Word文書のしおり整理を開始します"
    Debug.Print "対象ファイル: " & filePath
    Debug.Print "節構造: " & IIf(hasSections, "あり（5レベル）", "なし（4レベル）")
    Debug.Print "========================================="

    processedCount = 0

    ' セクションごとに処理
    Dim sect As Object
    Dim sectIndex As Long
    Dim prevPatterns As HeaderPatterns
    Dim currPatterns As HeaderPatterns

    ' 初期化
    prevPatterns.Level2Text = ""
    prevPatterns.Level3Text = ""
    prevPatterns.Level4Text = ""
    prevPatterns.Level5Text = ""

    sectIndex = 0
    For Each sect In wordDoc.Sections
        sectIndex = sectIndex + 1
        Debug.Print "-----------------------------------------"
        Debug.Print "セクション " & sectIndex & " を処理中..."

        ' このセクションのヘッダーからパターンを抽出
        currPatterns = ExtractHeaderPatterns(sect)

        Debug.Print "  ヘッダーパターン: Level2=" & currPatterns.Level2Text & _
                    ", Level3=" & currPatterns.Level3Text & _
                    ", Level4=" & currPatterns.Level4Text & _
                    ", Level5=" & currPatterns.Level5Text

        ' 前セクションとの差分を計算（新しく追加されたパターンのみ）
        Dim searchLevel2 As String
        Dim searchLevel3 As String
        Dim searchLevel4 As String
        Dim searchLevel5 As String

        searchLevel2 = ""
        searchLevel3 = ""
        searchLevel4 = ""
        searchLevel5 = ""

        If currPatterns.Level2Text <> "" And currPatterns.Level2Text <> prevPatterns.Level2Text Then
            searchLevel2 = currPatterns.Level2Text
        End If
        If currPatterns.Level3Text <> "" And currPatterns.Level3Text <> prevPatterns.Level3Text Then
            searchLevel3 = currPatterns.Level3Text
        End If
        If currPatterns.Level4Text <> "" And currPatterns.Level4Text <> prevPatterns.Level4Text Then
            searchLevel4 = currPatterns.Level4Text
        End If
        If currPatterns.Level5Text <> "" And currPatterns.Level5Text <> prevPatterns.Level5Text Then
            searchLevel5 = currPatterns.Level5Text
        End If

        Debug.Print "  検索対象: Level2=" & searchLevel2 & _
                    ", Level3=" & searchLevel3 & _
                    ", Level4=" & searchLevel4 & _
                    ", Level5=" & searchLevel5

        ' セクション内の段落を処理
        Dim para As Object
        For Each para In sect.Range.Paragraphs
            processedCount = processedCount + ProcessParagraphByHeader(para, styles, _
                searchLevel2, searchLevel3, searchLevel4, searchLevel5, hasSections, wordDoc)
        Next para

        ' セクション内の図形も処理
        Dim shp As Object
        Dim shapePara As Object
        On Error Resume Next
        For Each shp In sect.Range.ShapeRange
            If shp.TextFrame.HasText Then
                For Each shapePara In shp.TextFrame.TextRange.Paragraphs
                    processedCount = processedCount + ProcessParagraphByHeader(shapePara, styles, _
                        searchLevel2, searchLevel3, searchLevel4, searchLevel5, hasSections, wordDoc)
                Next shapePara
            End If
        Next shp
        Err.Clear
        On Error GoTo ErrorHandler

        ' 現在のパターンを保存
        prevPatterns = currPatterns
    Next sect

    ' ヘッダー内のフィールドを更新（STYLEREFのスタイル名も変更）
    UpdateHeaderFields wordDoc, styles

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
' セクションのヘッダーからパターンを抽出
' ============================================================================
Private Function ExtractHeaderPatterns(ByRef sect As Object) As HeaderPatterns
    Dim result As HeaderPatterns
    Dim headerText As String

    result.Level2Text = ""
    result.Level3Text = ""
    result.Level4Text = ""
    result.Level5Text = ""

    On Error Resume Next

    ' プライマリヘッダーからテキストを取得
    ' wdHeaderFooterPrimary = 1
    headerText = sect.Headers(1).Range.Text
    headerText = Trim(Replace(headerText, vbCr, " "))
    headerText = Replace(headerText, Chr(13), " ")

    If Err.Number <> 0 Then
        Err.Clear
        ExtractHeaderPatterns = result
        Exit Function
    End If

    On Error GoTo 0

    Debug.Print "  ヘッダーテキスト: " & Left(headerText, 100)

    ' ヘッダーが空の場合はスキップ
    If Trim(headerText) = "" Then
        ExtractHeaderPatterns = result
        Exit Function
    End If

    ' パターンを抽出（正規表現で）
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True

    ' 第X章を抽出
    regex.Pattern = "第[0-9０-９]+章"
    Dim matches As Object
    Set matches = regex.Execute(headerText)
    If matches.Count > 0 Then
        result.Level2Text = ConvertToHalfWidth(matches(matches.Count - 1).Value)
    End If

    ' 第X節を抽出
    regex.Pattern = "第[0-9０-９]+節"
    Set matches = regex.Execute(headerText)
    If matches.Count > 0 Then
        result.Level3Text = ConvertToHalfWidth(matches(matches.Count - 1).Value)
    End If

    ' X-X,X を先に抽出（X-Xより具体的）
    regex.Pattern = "[0-9０-９]+[-－ー][0-9０-９]+[,，.．][0-9０-９]+"
    Set matches = regex.Execute(headerText)
    If matches.Count > 0 Then
        result.Level5Text = ConvertToHalfWidth(matches(matches.Count - 1).Value)
    End If

    ' X-X を抽出（X-X,Xを除外）
    regex.Pattern = "[0-9０-９]+[-－ー][0-9０-９]+(?![,，.．0-9０-９])"
    Set matches = regex.Execute(headerText)
    If matches.Count > 0 Then
        result.Level4Text = ConvertToHalfWidth(matches(matches.Count - 1).Value)
    End If

    Set regex = Nothing
    ExtractHeaderPatterns = result
End Function

' ============================================================================
' 段落処理（ヘッダー参照方式）
' 戻り値: 処理された場合は1、それ以外は0
' hasSections: 文書内に「第X節」が存在する場合True
'   True の場合: Level3=第X節, Level4=X-X, Level5=X-X,X
'   False の場合: Level3=X-X, Level4=X-X,X, Level5=未使用
' ============================================================================
Private Function ProcessParagraphByHeader(ByRef para As Object, _
                                          ByRef styles As StyleConfig, _
                                          ByVal searchLevel2 As String, _
                                          ByVal searchLevel3 As String, _
                                          ByVal searchLevel4 As String, _
                                          ByVal searchLevel5 As String, _
                                          ByVal hasSections As Boolean, _
                                          ByRef wordDoc As Object) As Long
    On Error GoTo ErrorHandler

    If para Is Nothing Then
        ProcessParagraphByHeader = 0
        Exit Function
    End If

    If para.Range Is Nothing Then
        ProcessParagraphByHeader = 0
        Exit Function
    End If

    Dim paraText As String
    Dim paraTextHalf As String
    Dim detectedLevel As Long
    Dim targetStyle As String

    ' 段落テキストを取得（改行を除去）
    paraText = Trim(Replace(para.Range.Text, vbCr, ""))
    paraText = Replace(paraText, Chr(13), "")

    ' 半角変換版も用意（比較用）
    paraTextHalf = ConvertToHalfWidth(paraText)

    ' 空の段落はスキップ
    If paraText = "" Then
        ProcessParagraphByHeader = 0
        Exit Function
    End If

    detectedLevel = 0
    targetStyle = ""

    If hasSections Then
        ' ========================================
        ' 節ありの場合（5レベル構造）
        ' Level3=第X節, Level4=X-X, Level5=X-X,X
        ' ========================================

        ' レベル5: X-X,X（ヘッダーから抽出したテキストで検索）
        If detectedLevel = 0 And searchLevel5 <> "" And styles.Level5Style <> "" Then
            If InStr(paraTextHalf, searchLevel5) > 0 Then
                detectedLevel = 5
                targetStyle = styles.Level5Style
            End If
        End If

        ' レベル4: X-X（ヘッダーから抽出したテキストで検索）
        If detectedLevel = 0 And searchLevel4 <> "" And styles.Level4Style <> "" Then
            If InStr(paraTextHalf, searchLevel4) > 0 Then
                detectedLevel = 4
                targetStyle = styles.Level4Style
            End If
        End If

        ' レベル3: 第X節（ヘッダー参照のみ - パターンマッチなし）
        If detectedLevel = 0 And searchLevel3 <> "" And styles.Level3Style <> "" Then
            If InStr(paraTextHalf, searchLevel3) > 0 Then
                detectedLevel = 3
                targetStyle = styles.Level3Style
            End If
        End If
    Else
        ' ========================================
        ' 節なしの場合（4レベル構造）
        ' Level3=X-X, Level4=X-X,X, Level5=未使用
        ' ========================================

        ' レベル4: X-X,X（ヘッダーから抽出したテキストで検索）
        ' ※節なしの場合、searchLevel5にX-X,Xが入っている
        If detectedLevel = 0 And searchLevel5 <> "" And styles.Level4Style <> "" Then
            If InStr(paraTextHalf, searchLevel5) > 0 Then
                detectedLevel = 4
                targetStyle = styles.Level4Style
            End If
        End If

        ' レベル3: X-X（ヘッダーから抽出したテキストで検索）
        ' ※節なしの場合、searchLevel4にX-Xが入っている
        If detectedLevel = 0 And searchLevel4 <> "" And styles.Level3Style <> "" Then
            If InStr(paraTextHalf, searchLevel4) > 0 Then
                detectedLevel = 3
                targetStyle = styles.Level3Style
            End If
        End If
    End If

    ' レベル2: 第X章（ヘッダー参照またはパターンマッチ）
    If detectedLevel = 0 And styles.Level2Style <> "" Then
        If searchLevel2 <> "" Then
            ' ヘッダーから抽出したテキストで検索
            If InStr(paraTextHalf, searchLevel2) > 0 Then
                detectedLevel = 2
                targetStyle = styles.Level2Style
            End If
        Else
            ' ヘッダーが空の場合はパターンマッチでフォールバック
            If MatchPattern(paraText, "第[0-9０-９]+章") Then
                detectedLevel = 2
                targetStyle = styles.Level2Style
            End If
        End If
    End If

    ' レベル1: 第X部（パターンマッチ - ヘッダー情報なし）
    If detectedLevel = 0 And styles.Level1Style <> "" Then
        If MatchPattern(paraText, "第[0-9０-９]+部") Then
            detectedLevel = 1
            targetStyle = styles.Level1Style
        End If
    End If

    ' 例外1: パターン外だが既にレベル1-5のスタイルが適用されている
    If detectedLevel = 0 And styles.Exception1Style <> "" Then
        Dim currentStyle As String
        On Error Resume Next
        currentStyle = para.Style.NameLocal
        If Err.Number <> 0 Then
            currentStyle = ""
            Err.Clear
        End If
        On Error GoTo ErrorHandler

        If IsLevelStyle(currentStyle, styles) Then
            detectedLevel = -1
            targetStyle = styles.Exception1Style
        End If
    End If

    ' 例外2: アウトライン設定済み（段落またはスタイル）
    If detectedLevel = 0 And styles.Exception2Style <> "" Then
        Dim currentOutline As Long
        On Error Resume Next
        currentOutline = para.OutlineLevel
        If Err.Number <> 0 Then
            currentOutline = 10
            Err.Clear
        End If
        On Error GoTo ErrorHandler

        If (currentOutline >= 1 And currentOutline <= 9) Then
            detectedLevel = -2
            targetStyle = styles.Exception2Style
        ElseIf HasOutlineDefinedInStyle(para, wordDoc) Then
            detectedLevel = -2
            targetStyle = styles.Exception2Style
        End If
    End If

    ' スタイル適用
    If detectedLevel <> 0 And targetStyle <> "" Then
        ApplyStyle para, targetStyle
        Debug.Print "  [レベル" & detectedLevel & "] " & Left(paraText, 50)
        ProcessParagraphByHeader = 1
        Exit Function
    End If

    ProcessParagraphByHeader = 0
    Exit Function

ErrorHandler:
    ProcessParagraphByHeader = 0
End Function

' ============================================================================
' スタイルがレベル1-5のいずれかかチェック
' ============================================================================
Private Function IsLevelStyle(ByVal styleName As String, ByRef styles As StyleConfig) As Boolean
    If styleName = "" Then
        IsLevelStyle = False
        Exit Function
    End If

    If styleName = styles.Level1Style Or _
       styleName = styles.Level2Style Or _
       styleName = styles.Level3Style Or _
       styleName = styles.Level4Style Or _
       styleName = styles.Level5Style Then
        IsLevelStyle = True
    Else
        IsLevelStyle = False
    End If
End Function

' ============================================================================
' Excelシートから設定を読み込み
' ============================================================================
Private Function LoadSettings(ByRef styles As StyleConfig, _
                              ByRef doPdfOutput As Boolean) As Boolean
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_MAIN)

    ' スタイル名を読み込み
    styles.Level1Style = CStr(ws.Cells(ROW_PATTERN_LEVEL1, COL_STYLE_NAME).Value)
    styles.Level2Style = CStr(ws.Cells(ROW_PATTERN_LEVEL2, COL_STYLE_NAME).Value)
    styles.Level3Style = CStr(ws.Cells(ROW_PATTERN_LEVEL3, COL_STYLE_NAME).Value)
    styles.Level4Style = CStr(ws.Cells(ROW_PATTERN_LEVEL4, COL_STYLE_NAME).Value)
    styles.Level5Style = CStr(ws.Cells(ROW_PATTERN_LEVEL5, COL_STYLE_NAME).Value)
    styles.Exception1Style = CStr(ws.Cells(ROW_PATTERN_EXCEPTION1, COL_STYLE_NAME).Value)
    styles.Exception2Style = CStr(ws.Cells(ROW_PATTERN_EXCEPTION2, COL_STYLE_NAME).Value)

    ' オプション設定を読み込み
    doPdfOutput = (CStr(ws.Cells(ROW_OPTION_PDF_OUTPUT, COL_OPTION_VALUE).Value) = "はい")

    LoadSettings = True
    Exit Function

ErrorHandler:
    LoadSettings = False
End Function

' ============================================================================
' 文書内のヘッダーに「第X節」が存在するかチェック
' ============================================================================
Private Function HasSectionsInDocument(ByRef wordDoc As Object) As Boolean
    On Error Resume Next

    Dim sect As Object
    Dim headerText As String
    Dim regex As Object

    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.Pattern = "第[0-9０-９]+節"

    For Each sect In wordDoc.Sections
        ' プライマリヘッダーからテキストを取得
        headerText = sect.Headers(1).Range.Text
        If Err.Number = 0 Then
            If regex.Test(headerText) Then
                Set regex = Nothing
                HasSectionsInDocument = True
                Exit Function
            End If
        End If
        Err.Clear
    Next sect

    Set regex = Nothing
    HasSectionsInDocument = False
End Function

' ============================================================================
' スタイルの存在確認
' 存在しないスタイルがあれば、そのスタイル名を改行区切りで返す
' すべて存在すれば空文字列を返す
' hasSections: 節構造がある場合はTrue（Level5まで検証）
' ============================================================================
Private Function ValidateStyles(ByRef wordDoc As Object, ByRef styles As StyleConfig, _
                                ByVal hasSections As Boolean) As String
    Dim missingStyles As String
    missingStyles = ""

    ' 各スタイルの存在確認
    If styles.Level1Style <> "" Then
        If Not StyleExists(wordDoc, styles.Level1Style) Then
            missingStyles = missingStyles & "  - " & styles.Level1Style & " (レベル1: 第X部)" & vbCrLf
        End If
    End If

    If styles.Level2Style <> "" Then
        If Not StyleExists(wordDoc, styles.Level2Style) Then
            missingStyles = missingStyles & "  - " & styles.Level2Style & " (レベル2: 第X章)" & vbCrLf
        End If
    End If

    If styles.Level3Style <> "" Then
        If Not StyleExists(wordDoc, styles.Level3Style) Then
            If hasSections Then
                missingStyles = missingStyles & "  - " & styles.Level3Style & " (レベル3: 第X節)" & vbCrLf
            Else
                missingStyles = missingStyles & "  - " & styles.Level3Style & " (レベル3: X-X)" & vbCrLf
            End If
        End If
    End If

    If styles.Level4Style <> "" Then
        If Not StyleExists(wordDoc, styles.Level4Style) Then
            If hasSections Then
                missingStyles = missingStyles & "  - " & styles.Level4Style & " (レベル4: X-X)" & vbCrLf
            Else
                missingStyles = missingStyles & "  - " & styles.Level4Style & " (レベル4: X-X,X)" & vbCrLf
            End If
        End If
    End If

    ' Level5は節がある場合のみチェック
    If hasSections And styles.Level5Style <> "" Then
        If Not StyleExists(wordDoc, styles.Level5Style) Then
            missingStyles = missingStyles & "  - " & styles.Level5Style & " (レベル5: X-X,X)" & vbCrLf
        End If
    End If

    If styles.Exception1Style <> "" Then
        If Not StyleExists(wordDoc, styles.Exception1Style) Then
            missingStyles = missingStyles & "  - " & styles.Exception1Style & " (例外1)" & vbCrLf
        End If
    End If

    If styles.Exception2Style <> "" Then
        If Not StyleExists(wordDoc, styles.Exception2Style) Then
            missingStyles = missingStyles & "  - " & styles.Exception2Style & " (例外2)" & vbCrLf
        End If
    End If

    ValidateStyles = missingStyles
End Function

' ============================================================================
' 指定したスタイルがWord文書に存在するかチェック
' ============================================================================
Private Function StyleExists(ByRef wordDoc As Object, ByVal styleName As String) As Boolean
    On Error Resume Next

    Dim testStyle As Object
    Set testStyle = wordDoc.styles(styleName)

    If Err.Number <> 0 Then
        StyleExists = False
        Err.Clear
    Else
        StyleExists = True
    End If

    On Error GoTo 0
End Function

' ============================================================================
' 正規表現パターンマッチ
' ============================================================================
Private Function MatchPattern(ByVal text As String, ByVal Pattern As String) As Boolean
    On Error GoTo ErrorHandler

    If text = "" Or Pattern = "" Then
        MatchPattern = False
        Exit Function
    End If

    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")

    With regex
        .Global = False
        .IgnoreCase = False
        .MultiLine = False
        .Pattern = Pattern
    End With

    MatchPattern = regex.Test(text)

    Set regex = Nothing
    Exit Function

ErrorHandler:
    MatchPattern = False
End Function

' ============================================================================
' 全角を半角に変換（数字、ハイフン、カンマ、ピリオド）
' ============================================================================
Private Function ConvertToHalfWidth(ByVal text As String) As String
    Dim i As Long
    Dim char As String
    Dim result As String
    Dim charCode As Long

    result = ""
    For i = 1 To Len(text)
        char = Mid(text, i, 1)
        charCode = AscW(char)

        Select Case charCode
            Case &HFF10 To &HFF19  ' ０-９ → 0-9
                result = result & Chr(charCode - &HFF10 + Asc("0"))
            Case &HFF0D, &H2212, &H30FC  ' －、−、ー → -
                result = result & "-"
            Case &HFF0C  ' ， → ,
                result = result & ","
            Case &HFF0E  ' ． → .
                result = result & "."
            Case Else
                result = result & char
        End Select
    Next i

    ConvertToHalfWidth = result
End Function

' ============================================================================
' スタイルにアウトラインが定義されているかチェック
' ============================================================================
Private Function HasOutlineDefinedInStyle(ByRef para As Object, ByRef wordDoc As Object) As Boolean
    On Error GoTo ErrorHandler

    Dim styleName As String
    styleName = para.Style.NameLocal

    Dim style As Object
    Set style = wordDoc.styles(styleName)

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

    para.style = styleName

    If Err.Number <> 0 Then
        Debug.Print "  [警告] スタイル '" & styleName & "' が見つかりません: " & _
                    Left(para.Range.text, 50)
        Err.Clear
    End If

    On Error GoTo 0
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
' ヘッダー内のフィールドを更新し、STYLEREFのスタイル名も変更
' ============================================================================
Private Sub UpdateHeaderFields(ByRef wordDoc As Object, ByRef styles As StyleConfig)
    On Error Resume Next

    Dim sect As Object
    Dim hdr As Object
    Dim fld As Object
    Dim fieldCode As String
    Dim newFieldCode As String

    Debug.Print "========================================="
    Debug.Print "ヘッダーのフィールドを更新しています..."

    For Each sect In wordDoc.Sections
        ' wdHeaderFooterPrimary = 1
        For Each hdr In sect.Headers
            ' ヘッダー内のフィールドを更新
            For Each fld In hdr.Range.Fields
                fieldCode = fld.Code.Text

                ' STYLEREFフィールドの場合、スタイル名を更新
                If InStr(fieldCode, "STYLEREF") > 0 Then
                    newFieldCode = fieldCode

                    ' デフォルトスタイル名を新しいスタイル名に置換
                    If styles.Level1Style <> DEFAULT_LEVEL1_STYLE And _
                       styles.Level1Style <> "" Then
                        newFieldCode = Replace(newFieldCode, DEFAULT_LEVEL1_STYLE, styles.Level1Style)
                    End If
                    If styles.Level2Style <> DEFAULT_LEVEL2_STYLE And _
                       styles.Level2Style <> "" Then
                        newFieldCode = Replace(newFieldCode, DEFAULT_LEVEL2_STYLE, styles.Level2Style)
                    End If
                    If styles.Level3Style <> DEFAULT_LEVEL3_STYLE And _
                       styles.Level3Style <> "" Then
                        newFieldCode = Replace(newFieldCode, DEFAULT_LEVEL3_STYLE, styles.Level3Style)
                    End If
                    If styles.Level4Style <> DEFAULT_LEVEL4_STYLE And _
                       styles.Level4Style <> "" Then
                        newFieldCode = Replace(newFieldCode, DEFAULT_LEVEL4_STYLE, styles.Level4Style)
                    End If
                    If styles.Level5Style <> DEFAULT_LEVEL5_STYLE And _
                       styles.Level5Style <> "" Then
                        newFieldCode = Replace(newFieldCode, DEFAULT_LEVEL5_STYLE, styles.Level5Style)
                    End If

                    ' フィールドコードが変更された場合、更新
                    If newFieldCode <> fieldCode Then
                        fld.Code.Text = newFieldCode
                        Debug.Print "  STYLEREF更新: " & Trim(fieldCode) & " → " & Trim(newFieldCode)
                    End If
                End If

                ' フィールドを更新
                fld.Update
            Next fld
        Next hdr

        ' フッターも更新
        ' wdHeaderFooterPrimary = 1
        For Each hdr In sect.Footers
            For Each fld In hdr.Range.Fields
                fld.Update
            Next fld
        Next hdr
    Next sect

    ' 文書全体のフィールドも更新
    wordDoc.Fields.Update

    Debug.Print "ヘッダーのフィールド更新完了"

    On Error GoTo 0
End Sub

' ============================================================================
' テスト用: イミディエイトウィンドウに情報を出力
' ============================================================================
Public Sub TestOrganizeWordBookmarks()
    ' イミディエイトウィンドウを開いた状態でこのマクロを実行してください
    OrganizeWordBookmarks
End Sub
