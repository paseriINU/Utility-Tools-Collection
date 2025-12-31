Option Explicit

' ============================================================================
' Word しおり整理ツール - メインモジュール
' パターンマッチ方式: 段落テキストを正規表現でパターンマッチし、
' 該当するスタイルを適用。「参照」を含む段落はスキップ。
' ============================================================================

' === スタイル設定構造体 ===
Private Type StyleConfig
    Level1Style As String  ' 第X部用（パターンマッチ、ヘッダー空白時のみ）
    Level2Style As String  ' 第X章用（パターンマッチ、ヘッダー非空白時のみ）
    Level3Style As String  ' 第X節/X-X用（パターンマッチ、自動判定）
    Level4Style As String  ' X-X/X-X,X用（パターンマッチ）
    Level5Style As String  ' X-X,X用（パターンマッチ、節あり時のみ）
    Exception1Style As String
    Exception2Style As String
End Type

' === デフォルトスタイル名定数（ヘッダーのSTYLEREF更新用） ===
Private Const DEFAULT_LEVEL1_STYLE As String = "表題1"
Private Const DEFAULT_LEVEL2_STYLE As String = "表題2"
Private Const DEFAULT_LEVEL3_STYLE As String = "表題3"
Private Const DEFAULT_LEVEL4_STYLE As String = "表題4"
Private Const DEFAULT_LEVEL5_STYLE As String = "表題5"

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
    Dim inputDir As String
    Dim outputDir As String

    ' 設定
    Dim styles As StyleConfig
    Dim doPdfOutput As Boolean

    ' Excelシートからフォルダパスを読み取る
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_MAIN)

    inputDir = CStr(ws.Cells(ROW_INPUT_FOLDER, 3).Value)  ' C10セル
    outputDir = CStr(ws.Cells(ROW_OUTPUT_FOLDER, 3).Value)  ' C12セル

    ' 末尾に\を追加（なければ）
    If Right(inputDir, 1) <> "\" Then inputDir = inputDir & "\"
    If Right(outputDir, 1) <> "\" Then outputDir = outputDir & "\"

    ' Inputフォルダの存在確認（存在しない場合はエラー）
    If Dir(inputDir, vbDirectory) = "" Then
        MsgBox "エラー: 入力フォルダが存在しません。" & vbCrLf & vbCrLf & _
               inputDir & vbCrLf & vbCrLf & _
               "フォルダ設定を確認してください。", vbCritical, "フォルダエラー"
        Exit Sub
    End If

    ' Outputフォルダの存在確認（存在しない場合はエラー）
    If Dir(outputDir, vbDirectory) = "" Then
        MsgBox "エラー: 出力フォルダが存在しません。" & vbCrLf & vbCrLf & _
               outputDir & vbCrLf & vbCrLf & _
               "フォルダ設定を確認してください。", vbCritical, "フォルダエラー"
        Exit Sub
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

    ' 1ページ目に「帳票」があるか判定
    Dim isHyohyoDocument As Boolean
    isHyohyoDocument = HasHyohyoOnFirstPage(wordDoc)

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
    Debug.Print "帳票文書: " & IIf(isHyohyoDocument, "あり（帳票番号パターン有効）", "なし")
    Debug.Print "========================================="

    processedCount = 0

    ' セクションごとに処理
    Dim sect As Object
    Dim sectIndex As Long

    sectIndex = 0
    For Each sect In wordDoc.Sections
        sectIndex = sectIndex + 1
        Debug.Print "-----------------------------------------"
        Debug.Print "セクション " & sectIndex & " を処理中..."

        ' ヘッダーが空白かどうか判定（ヘッダーテキスト自体が空かどうか）
        Dim headerIsEmpty As Boolean
        headerIsEmpty = IsHeaderEmpty(sect)
        Debug.Print "  ヘッダー空白: " & headerIsEmpty

        ' セクション内の段落を処理（パターンマッチ方式）
        Dim para As Object
        For Each para In sect.Range.Paragraphs
            processedCount = processedCount + ProcessParagraph(para, styles, hasSections, headerIsEmpty, isHyohyoDocument, wordDoc)
        Next para

        ' セクション内の図形も処理
        Dim shp As Object
        Dim shapePara As Object
        On Error Resume Next
        For Each shp In sect.Range.ShapeRange
            If shp.TextFrame.HasText Then
                For Each shapePara In shp.TextFrame.TextRange.Paragraphs
                    processedCount = processedCount + ProcessParagraph(shapePara, styles, hasSections, headerIsEmpty, isHyohyoDocument, wordDoc)
                Next shapePara
            End If
        Next shp
        Err.Clear
        On Error GoTo ErrorHandler
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
' セクションのヘッダーが空白かどうかをチェック
' ============================================================================
Private Function IsHeaderEmpty(ByRef sect As Object) As Boolean
    On Error Resume Next

    Dim headerText As String
    headerText = sect.Headers(1).Range.Text
    headerText = Trim(Replace(headerText, vbCr, ""))
    headerText = Replace(headerText, Chr(13), "")

    If Err.Number <> 0 Then
        Err.Clear
        IsHeaderEmpty = True
        Exit Function
    End If

    On Error GoTo 0

    IsHeaderEmpty = (Trim(headerText) = "")
End Function

' ============================================================================
' 段落処理（パターンマッチ方式）
' 戻り値: 処理された場合は1、それ以外は0
' hasSections: 文書内に「第X節」が存在する場合True
'   True の場合: Level3=第X節, Level4=X-X, Level5=X-X,X
'   False の場合: Level3=X-X, Level4=X-X,X, Level5=未使用
' headerIsEmpty: ヘッダーが空白の場合True
' isHyohyoDocument: 1ページ目に「帳票」がある場合True
'   True の場合: (X123)/(XX12)パターンにLevel5スタイルを適用
' ============================================================================
Private Function ProcessParagraph(ByRef para As Object, _
                                  ByRef styles As StyleConfig, _
                                  ByVal hasSections As Boolean, _
                                  ByVal headerIsEmpty As Boolean, _
                                  ByVal isHyohyoDocument As Boolean, _
                                  ByRef wordDoc As Object) As Long
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
    Dim paraTextHalf As String
    Dim detectedLevel As Long
    Dim targetStyle As String

    ' 段落テキストを取得（改行・改ページ等の制御文字を除去）
    paraText = Trim(Replace(para.Range.Text, vbCr, ""))
    paraText = Replace(paraText, Chr(13), "")   ' キャリッジリターン
    paraText = Replace(paraText, Chr(12), "")   ' 改ページ（フォームフィード）
    paraText = Replace(paraText, Chr(11), "")   ' 段落内改行（Shift+Enter）
    paraText = Replace(paraText, Chr(7), "")    ' セル終端マーカー（表内）
    paraText = Trim(paraText)                   ' 再度トリム

    ' 半角変換版も用意（比較用）
    paraTextHalf = ConvertToHalfWidth(paraText)

    ' 空の段落はスキップ
    If paraText = "" Then
        ProcessParagraph = 0
        Exit Function
    End If

    ' 「参照」を含む段落はスキップ
    If InStr(paraText, "参照") > 0 Then
        ProcessParagraph = 0
        Exit Function
    End If

    ' 「・」で始まる段落はスキップ（目次形式: 「・　第1章」など）
    If Left(paraText, 1) = "・" Then
        ProcessParagraph = 0
        Exit Function
    End If

    ' ハイパーリンクを含む段落はスキップ（目次など）
    On Error Resume Next
    If para.Range.Hyperlinks.Count > 0 Then
        ProcessParagraph = 0
        Exit Function
    End If
    If Err.Number <> 0 Then
        Err.Clear
    End If
    On Error GoTo ErrorHandler

    ' 表内の段落はスキップ
    On Error Resume Next
    If para.Range.Information(12) = True Then  ' 12 = wdWithInTable
        ProcessParagraph = 0
        Exit Function
    End If
    If Err.Number <> 0 Then
        Err.Clear
    End If
    On Error GoTo ErrorHandler

    detectedLevel = 0
    targetStyle = ""

    ' ========================================
    ' 最初に「第X部」をチェック（段落先頭かつヘッダー空白時のみ）
    ' 「第X部」で始まる段落は他のレベルの処理をスキップ
    ' ========================================
    If styles.Level1Style <> "" And headerIsEmpty Then
        If MatchPattern(paraText, "^第[0-9０-９]+部") Then
            detectedLevel = 1
            targetStyle = styles.Level1Style
            GoTo ApplyStyleAndExit
        End If
    End If

    If hasSections Then
        ' ========================================
        ' 節ありの場合（5レベル構造）
        ' Level3=第X節, Level4=X-X, Level5=X-X,X
        ' ========================================

        ' レベル5: X-X,X（パターンマッチ）
        If detectedLevel = 0 And styles.Level5Style <> "" Then
            If MatchPattern(paraTextHalf, "^[0-9]+-[0-9]+[,\.][0-9]+") Then
                detectedLevel = 5
                targetStyle = styles.Level5Style
            End If
        End If

        ' レベル4: X-X（パターンマッチ、X-X,Xを除外）
        If detectedLevel = 0 And styles.Level4Style <> "" Then
            If MatchPattern(paraTextHalf, "^[0-9]+-[0-9]+(?![,\.0-9])") Then
                detectedLevel = 4
                targetStyle = styles.Level4Style
            End If
        End If

        ' レベル3: 第X節（パターンマッチ）
        If detectedLevel = 0 And styles.Level3Style <> "" Then
            If MatchPattern(paraText, "^第[0-9０-９]+節") Then
                detectedLevel = 3
                targetStyle = styles.Level3Style
            End If
        End If
    Else
        ' ========================================
        ' 節なしの場合（4レベル構造）
        ' Level3=X-X, Level4=X-X,X, Level5=未使用
        ' ========================================

        ' レベル4: X-X,X（パターンマッチ）
        If detectedLevel = 0 And styles.Level4Style <> "" Then
            If MatchPattern(paraTextHalf, "^[0-9]+-[0-9]+[,\.][0-9]+") Then
                detectedLevel = 4
                targetStyle = styles.Level4Style
            End If
        End If

        ' レベル3: X-X（パターンマッチ、X-X,Xを除外）
        If detectedLevel = 0 And styles.Level3Style <> "" Then
            If MatchPattern(paraTextHalf, "^[0-9]+-[0-9]+(?![,\.0-9])") Then
                detectedLevel = 3
                targetStyle = styles.Level3Style
            End If
        End If
    End If

    ' レベル2: 第X章（パターンマッチ）
    If detectedLevel = 0 And styles.Level2Style <> "" Then
        If MatchPattern(paraText, "^第[0-9０-９]+章") Then
            detectedLevel = 2
            targetStyle = styles.Level2Style
        End If
    End If

    ' 帳票文書の場合: (X123)/(XX12)パターンにLevel5スタイルを適用
    ' (X123): 英字1文字 + 数字3文字
    ' (XX12): 英字2文字 + 数字2文字
    ' ※全角・半角両対応（paraTextHalfで半角統一済み）
    If detectedLevel = 0 And isHyohyoDocument And styles.Level5Style <> "" Then
        ' 半角変換後のテキストでパターンマッチ
        If MatchPattern(paraTextHalf, "\([A-Za-z][0-9]{3}\)") Or _
           MatchPattern(paraTextHalf, "\([A-Za-z]{2}[0-9]{2}\)") Then
            detectedLevel = 5
            targetStyle = styles.Level5Style
            Debug.Print "  [帳票番号検出] " & Left(paraText, 50)
        End If
    End If

    ' 特定テキストの例外処理（スタイルはレベル3、アウトラインはレベル1）
    If detectedLevel = 0 And styles.Level3Style <> "" Then
        If paraText = "本書の記述について" Or paraText = "修正履歴" Then
            ApplyStyle para, styles.Level3Style
            ' アウトラインレベルを1に設定（しおりの階層用）
            On Error Resume Next
            para.Range.ParagraphFormat.OutlineLevel = 1  ' wdOutlineLevel1
            On Error GoTo ErrorHandler
            Debug.Print "  [特定テキスト→アウトライン1] " & paraText
            ProcessParagraph = 1
            Exit Function
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

ApplyStyleAndExit:
    ' スタイル適用
    If detectedLevel <> 0 And targetStyle <> "" Then
        ApplyStyle para, targetStyle
        Debug.Print "  [レベル" & detectedLevel & "] " & Left(paraText, 50)
        ProcessParagraph = 1
        Exit Function
    End If

    ProcessParagraph = 0
    Exit Function

ErrorHandler:
    ProcessParagraph = 0
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
' 1ページ目に「帳票」の文字があるかチェック
' ============================================================================
Private Function HasHyohyoOnFirstPage(ByRef wordDoc As Object) As Boolean
    On Error Resume Next

    Dim searchRange As Object

    ' 文書全体を検索対象に
    Set searchRange = wordDoc.Content

    ' 「帳票」を検索
    searchRange.Find.ClearFormatting
    If searchRange.Find.Execute(FindText:="帳票") Then
        ' 見つかった位置が1ページ目かどうかを確認
        ' wdActiveEndPageNumber = 3
        If searchRange.Information(3) = 1 Then
            HasHyohyoOnFirstPage = True
            Exit Function
        End If
    End If

    HasHyohyoOnFirstPage = False
    On Error GoTo 0
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
' 全角を半角に変換（数字、英字、ハイフン、カンマ、ピリオド、括弧）
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
            Case &HFF21 To &HFF3A  ' Ａ-Ｚ → A-Z
                result = result & Chr(charCode - &HFF21 + Asc("A"))
            Case &HFF41 To &HFF5A  ' ａ-ｚ → a-z
                result = result & Chr(charCode - &HFF41 + Asc("a"))
            Case &HFF0D, &H2212, &H30FC  ' －、−、ー → -
                result = result & "-"
            Case &HFF0C  ' ， → ,
                result = result & ","
            Case &HFF0E  ' ． → .
                result = result & "."
            Case &HFF08  ' （ → (
                result = result & "("
            Case &HFF09  ' ） → )
                result = result & ")"
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
