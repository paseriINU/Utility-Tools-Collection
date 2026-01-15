Option Explicit

' ============================================================================
' Word しおり整理ツール - メインモジュール
' パターンマッチ方式: 段落テキストを正規表現でパターンマッチし、
' 該当するスタイルを適用。「参照」を含む段落はスキップ。
' ============================================================================

' === スタイル設定構造体（動的配列対応） ===
Private Type StyleSetting
    Category As String      ' 種別: パターン, 帳票, 特定, 例外
    Level As String         ' レベル: 1, 2, 3, 3-節, 4, 4-節, 5-節 など
    Pattern As String       ' パターン（正規表現）またはテキスト
    StyleName As String     ' 適用スタイル名
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
    Dim doPdfOutput As Boolean

    ' 動的スタイル設定
    Dim styleSettings() As StyleSetting
    Dim styleCount As Long

    ' 設定シートからフォルダパスを読み取る
    Dim wsSettings As Worksheet
    On Error Resume Next
    Set wsSettings = ThisWorkbook.Worksheets(SHEET_SETTINGS)
    If wsSettings Is Nothing Then
        MsgBox "エラー: 「設定」シートが見つかりません。" & vbCrLf & _
               "初期化マクロを実行してください。", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    inputDir = CStr(wsSettings.Cells(SETTINGS_ROW_INPUT_FOLDER, SETTINGS_COL_VALUE).Value)
    outputDir = CStr(wsSettings.Cells(SETTINGS_ROW_OUTPUT_FOLDER, SETTINGS_COL_VALUE).Value)

    ' 末尾に\を追加（なければ）
    If Right(inputDir, 1) <> "\" Then inputDir = inputDir & "\"
    If Right(outputDir, 1) <> "\" Then outputDir = outputDir & "\"

    ' Inputフォルダの存在確認
    If Dir(inputDir, vbDirectory) = "" Then
        MsgBox "エラー: 入力フォルダが存在しません。" & vbCrLf & vbCrLf & _
               inputDir & vbCrLf & vbCrLf & _
               "設定シートのフォルダ設定を確認してください。", vbCritical, "フォルダエラー"
        Exit Sub
    End If

    ' Outputフォルダの存在確認
    If Dir(outputDir, vbDirectory) = "" Then
        MsgBox "エラー: 出力フォルダが存在しません。" & vbCrLf & vbCrLf & _
               outputDir & vbCrLf & vbCrLf & _
               "設定シートのフォルダ設定を確認してください。", vbCritical, "フォルダエラー"
        Exit Sub
    End If

    ' 設定シートから設定を動的に読み込み
    If Not LoadDynamicSettings(wsSettings, styleSettings, styleCount, doPdfOutput) Then
        MsgBox "設定の読み込みに失敗しました。" & vbCrLf & _
               "設定シートを確認してください。", vbExclamation
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
    wordApp.Visible = False

    ' Word文書を開く
    Set wordDoc = wordApp.Documents.Open(filePath)

    ' 文書内に「第X節」がヘッダーに存在するか判定
    Dim hasSections As Boolean
    hasSections = HasSectionsInDocument(wordDoc)

    ' 1ページ目に「帳票」があるか判定
    Dim isHyohyoDocument As Boolean
    isHyohyoDocument = HasHyohyoOnFirstPage(wordDoc)

    ' スタイルの存在確認
    Dim missingStyles As String
    missingStyles = ValidateDynamicStyles(wordDoc, styleSettings, styleCount, hasSections)
    If missingStyles <> "" Then
        MsgBox "エラー: 以下のスタイルがWord文書に存在しません。" & vbCrLf & vbCrLf & _
               missingStyles & vbCrLf & _
               "処理を中止します。", vbCritical, "スタイルエラー"
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
    Debug.Print "節構造: " & IIf(hasSections, "あり", "なし")
    Debug.Print "帳票文書: " & IIf(isHyohyoDocument, "あり", "なし")
    Debug.Print "スタイル設定数: " & styleCount
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

        ' ヘッダーが空白かどうか判定
        Dim headerIsEmpty As Boolean
        headerIsEmpty = IsHeaderEmpty(sect)
        Debug.Print "  ヘッダー空白: " & headerIsEmpty

        ' セクション内の段落を処理
        Dim para As Object
        For Each para In sect.Range.Paragraphs
            processedCount = processedCount + ProcessParagraphDynamic(para, styleSettings, styleCount, hasSections, headerIsEmpty, isHyohyoDocument, wordDoc)
        Next para

        ' セクション内の図形も処理
        Dim shp As Object
        Dim shapePara As Object
        On Error Resume Next
        For Each shp In sect.Range.ShapeRange
            If shp.TextFrame.HasText Then
                For Each shapePara In shp.TextFrame.TextRange.Paragraphs
                    processedCount = processedCount + ProcessParagraphDynamic(shapePara, styleSettings, styleCount, hasSections, headerIsEmpty, isHyohyoDocument, wordDoc)
                Next shapePara
            End If
        Next shp
        Err.Clear
        On Error GoTo ErrorHandler
    Next sect

    ' ヘッダー内のフィールドを更新
    UpdateHeaderFieldsDynamic wordDoc, styleSettings, styleCount

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

    ' 完了メッセージ
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
    MsgBox "エラーが発生しました: " & vbCrLf & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, vbCritical, "エラー"

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
' 設定シートから動的に設定を読み込み
' ============================================================================
Private Function LoadDynamicSettings(ByRef wsSettings As Worksheet, _
                                     ByRef styleSettings() As StyleSetting, _
                                     ByRef styleCount As Long, _
                                     ByRef doPdfOutput As Boolean) As Boolean
    On Error GoTo ErrorHandler

    Dim row As Long
    Dim category As String
    Dim maxRows As Long

    ' 最大行数を設定（空行が続いたら終了）
    maxRows = 100
    styleCount = 0
    ReDim styleSettings(0 To maxRows - 1)

    ' スタイル設定を読み込み
    row = SETTINGS_ROW_STYLE_START
    Dim emptyRowCount As Long
    emptyRowCount = 0

    Do While row < SETTINGS_ROW_STYLE_START + maxRows
        category = Trim(CStr(wsSettings.Cells(row, SETTINGS_COL_LABEL).Value))

        ' 空行が3行連続したら終了
        If category = "" Then
            emptyRowCount = emptyRowCount + 1
            If emptyRowCount >= 3 Then Exit Do
            row = row + 1
            GoTo NextRow
        End If

        emptyRowCount = 0

        ' 有効な設定行を追加
        styleSettings(styleCount).Category = category
        styleSettings(styleCount).Level = Trim(CStr(wsSettings.Cells(row, SETTINGS_COL_VALUE).Value))
        styleSettings(styleCount).Pattern = Trim(CStr(wsSettings.Cells(row, SETTINGS_COL_PATTERN).Value))
        styleSettings(styleCount).StyleName = Trim(CStr(wsSettings.Cells(row, SETTINGS_COL_STYLE).Value))

        styleCount = styleCount + 1
        row = row + 1
NextRow:
    Loop

    ' 配列をリサイズ
    If styleCount > 0 Then
        ReDim Preserve styleSettings(0 To styleCount - 1)
    End If

    ' オプション設定を読み込み
    doPdfOutput = (CStr(wsSettings.Cells(SETTINGS_ROW_PDF_OUTPUT, SETTINGS_COL_VALUE).Value) = "はい")

    LoadDynamicSettings = True
    Exit Function

ErrorHandler:
    LoadDynamicSettings = False
End Function

' ============================================================================
' 段落処理（動的設定対応）
' ============================================================================
Private Function ProcessParagraphDynamic(ByRef para As Object, _
                                         ByRef styleSettings() As StyleSetting, _
                                         ByVal styleCount As Long, _
                                         ByVal hasSections As Boolean, _
                                         ByVal headerIsEmpty As Boolean, _
                                         ByVal isHyohyoDocument As Boolean, _
                                         ByRef wordDoc As Object) As Long
    On Error GoTo ErrorHandler

    If para Is Nothing Then
        ProcessParagraphDynamic = 0
        Exit Function
    End If

    If para.Range Is Nothing Then
        ProcessParagraphDynamic = 0
        Exit Function
    End If

    Dim paraText As String
    Dim paraTextHalf As String

    ' 段落テキストを取得
    paraText = Trim(Replace(para.Range.Text, vbCr, ""))
    paraText = Replace(paraText, Chr(13), "")
    paraText = Replace(paraText, Chr(12), "")
    paraText = Replace(paraText, Chr(11), "")
    paraText = Replace(paraText, Chr(7), "")
    paraText = Trim(paraText)

    ' 半角変換版
    paraTextHalf = ConvertToHalfWidth(paraText)

    ' 空の段落はスキップ
    If paraText = "" Then
        ProcessParagraphDynamic = 0
        Exit Function
    End If

    ' 「参照」を含む段落はスキップ
    If InStr(paraText, "参照") > 0 Then
        ProcessParagraphDynamic = 0
        Exit Function
    End If

    ' 「・」で始まる段落はスキップ
    If Left(paraText, 1) = "・" Then
        ProcessParagraphDynamic = 0
        Exit Function
    End If

    ' ハイパーリンクを含む段落はスキップ
    On Error Resume Next
    If para.Range.Hyperlinks.Count > 0 Then
        ProcessParagraphDynamic = 0
        Exit Function
    End If
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo ErrorHandler

    ' 表内の段落はスキップ
    On Error Resume Next
    If para.Range.Information(12) = True Then
        ProcessParagraphDynamic = 0
        Exit Function
    End If
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo ErrorHandler

    ' 動的設定をループして処理
    Dim i As Long
    Dim setting As StyleSetting
    Dim matched As Boolean
    Dim targetStyle As String
    Dim detectedLevel As String

    matched = False

    For i = 0 To styleCount - 1
        setting = styleSettings(i)

        ' スタイル名が空なら次へ
        If setting.StyleName = "" Then GoTo NextSetting

        Select Case setting.Category
            Case "パターン"
                ' レベルに「-節」が含まれる場合は節構造あり時のみ適用
                If InStr(setting.Level, "-節") > 0 Then
                    If Not hasSections Then GoTo NextSetting
                ElseIf setting.Level = "1" Then
                    ' レベル1はヘッダー空白時のみ
                    If Not headerIsEmpty Then GoTo NextSetting
                Else
                    ' 「X-節」でないレベル設定は節構造なし時のみ適用
                    ' ただしレベル1,2は除く
                    If setting.Level <> "1" And setting.Level <> "2" Then
                        ' 同じレベルで「-節」付きがあればスキップ（節構造あり時）
                        If hasSections Then
                            If HasSectionVariant(styleSettings, styleCount, setting.Level) Then
                                GoTo NextSetting
                            End If
                        End If
                    End If
                End If

                ' パターンマッチ（半角変換済みテキストでも試す）
                If setting.Pattern <> "" Then
                    If MatchPattern(paraText, setting.Pattern) Or _
                       MatchPattern(paraTextHalf, setting.Pattern) Then
                        matched = True
                        targetStyle = setting.StyleName
                        detectedLevel = setting.Level
                        Exit For
                    End If
                End If

            Case "帳票"
                ' 帳票文書の場合のみ適用
                If isHyohyoDocument And setting.Pattern <> "" Then
                    If MatchPattern(paraTextHalf, setting.Pattern) Then
                        matched = True
                        targetStyle = setting.StyleName
                        detectedLevel = "帳票"
                        Debug.Print "  [帳票番号検出] " & Left(paraText, 50)
                        Exit For
                    End If
                End If

            Case "特定"
                ' 完全一致
                If setting.Pattern <> "" Then
                    If paraText = setting.Pattern Then
                        ApplyStyle para, setting.StyleName
                        ' アウトラインレベルを設定
                        Dim outlineLevel As Long
                        If IsNumeric(setting.Level) Then
                            outlineLevel = CLng(setting.Level)
                            If outlineLevel >= 1 And outlineLevel <= 9 Then
                                On Error Resume Next
                                para.Range.ParagraphFormat.OutlineLevel = outlineLevel
                                On Error GoTo ErrorHandler
                            End If
                        End If
                        Debug.Print "  [特定テキスト→アウトライン" & setting.Level & "] " & paraText
                        ProcessParagraphDynamic = 1
                        Exit Function
                    End If
                End If

            Case "例外"
                ' 例外1: パターン外だが見出しスタイル適用済み
                If setting.Level = "1" Then
                    Dim currentStyle As String
                    On Error Resume Next
                    currentStyle = para.Style.NameLocal
                    If Err.Number <> 0 Then
                        currentStyle = ""
                        Err.Clear
                    End If
                    On Error GoTo ErrorHandler

                    If IsLevelStyleDynamic(currentStyle, styleSettings, styleCount) Then
                        matched = True
                        targetStyle = setting.StyleName
                        detectedLevel = "例外1"
                        Exit For
                    End If
                End If

                ' 例外2: アウトライン設定済み
                If setting.Level = "2" Then
                    Dim currentOutline As Long
                    On Error Resume Next
                    currentOutline = para.OutlineLevel
                    If Err.Number <> 0 Then
                        currentOutline = 10
                        Err.Clear
                    End If
                    On Error GoTo ErrorHandler

                    If (currentOutline >= 1 And currentOutline <= 9) Then
                        matched = True
                        targetStyle = setting.StyleName
                        detectedLevel = "例外2"
                        Exit For
                    ElseIf HasOutlineDefinedInStyle(para, wordDoc) Then
                        matched = True
                        targetStyle = setting.StyleName
                        detectedLevel = "例外2"
                        Exit For
                    End If
                End If
        End Select
NextSetting:
    Next i

    ' スタイル適用
    If matched And targetStyle <> "" Then
        ApplyStyle para, targetStyle
        Debug.Print "  [" & detectedLevel & "] " & Left(paraText, 50)
        ProcessParagraphDynamic = 1
        Exit Function
    End If

    ProcessParagraphDynamic = 0
    Exit Function

ErrorHandler:
    ProcessParagraphDynamic = 0
End Function

' ============================================================================
' 同じレベルで「-節」バリアントがあるかチェック
' ============================================================================
Private Function HasSectionVariant(ByRef styleSettings() As StyleSetting, _
                                   ByVal styleCount As Long, _
                                   ByVal level As String) As Boolean
    Dim i As Long
    For i = 0 To styleCount - 1
        If styleSettings(i).Level = level & "-節" Then
            HasSectionVariant = True
            Exit Function
        End If
    Next i
    HasSectionVariant = False
End Function

' ============================================================================
' スタイルがレベル系スタイルかチェック（動的版）
' ============================================================================
Private Function IsLevelStyleDynamic(ByVal styleName As String, _
                                     ByRef styleSettings() As StyleSetting, _
                                     ByVal styleCount As Long) As Boolean
    If styleName = "" Then
        IsLevelStyleDynamic = False
        Exit Function
    End If

    Dim i As Long
    For i = 0 To styleCount - 1
        If styleSettings(i).Category = "パターン" And styleSettings(i).StyleName = styleName Then
            IsLevelStyleDynamic = True
            Exit Function
        End If
    Next i

    IsLevelStyleDynamic = False
End Function

' ============================================================================
' スタイルの存在確認（動的版）
' ============================================================================
Private Function ValidateDynamicStyles(ByRef wordDoc As Object, _
                                       ByRef styleSettings() As StyleSetting, _
                                       ByVal styleCount As Long, _
                                       ByVal hasSections As Boolean) As String
    Dim missingStyles As String
    Dim checkedStyles As Collection
    Dim i As Long
    Dim setting As StyleSetting

    missingStyles = ""
    Set checkedStyles = New Collection

    For i = 0 To styleCount - 1
        setting = styleSettings(i)

        ' 空のスタイル名はスキップ
        If setting.StyleName = "" Then GoTo NextStyle

        ' 節構造に応じてスキップ
        If InStr(setting.Level, "-節") > 0 And Not hasSections Then GoTo NextStyle
        If setting.Category = "パターン" And setting.Level <> "1" And setting.Level <> "2" Then
            If Not InStr(setting.Level, "-節") > 0 And hasSections Then
                If HasSectionVariant(styleSettings, styleCount, setting.Level) Then
                    GoTo NextStyle
                End If
            End If
        End If

        ' 既にチェック済みならスキップ
        On Error Resume Next
        checkedStyles.Add setting.StyleName, setting.StyleName
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextStyle
        End If
        On Error GoTo 0

        ' スタイルの存在確認
        If Not StyleExists(wordDoc, setting.StyleName) Then
            missingStyles = missingStyles & "  - " & setting.StyleName & _
                           " (" & setting.Category & ": " & setting.Pattern & ")" & vbCrLf
        End If
NextStyle:
    Next i

    ValidateDynamicStyles = missingStyles
End Function

' ============================================================================
' ヘッダー内のフィールドを更新（動的版）
' ============================================================================
Private Sub UpdateHeaderFieldsDynamic(ByRef wordDoc As Object, _
                                      ByRef styleSettings() As StyleSetting, _
                                      ByVal styleCount As Long)
    On Error Resume Next

    Dim sect As Object
    Dim hdr As Object
    Dim fld As Object
    Dim fieldCode As String
    Dim newFieldCode As String
    Dim i As Long

    Debug.Print "========================================="
    Debug.Print "ヘッダーのフィールドを更新しています..."

    For Each sect In wordDoc.Sections
        For Each hdr In sect.Headers
            For Each fld In hdr.Range.Fields
                fieldCode = fld.Code.Text

                ' STYLEREFフィールドの場合、スタイル名を更新
                If InStr(fieldCode, "STYLEREF") > 0 Then
                    newFieldCode = fieldCode

                    ' デフォルトスタイル名を新しいスタイル名に置換
                    For i = 0 To styleCount - 1
                        If styleSettings(i).Category = "パターン" Then
                            Select Case styleSettings(i).Level
                                Case "1"
                                    If styleSettings(i).StyleName <> DEFAULT_LEVEL1_STYLE Then
                                        newFieldCode = Replace(newFieldCode, DEFAULT_LEVEL1_STYLE, styleSettings(i).StyleName)
                                    End If
                                Case "2"
                                    If styleSettings(i).StyleName <> DEFAULT_LEVEL2_STYLE Then
                                        newFieldCode = Replace(newFieldCode, DEFAULT_LEVEL2_STYLE, styleSettings(i).StyleName)
                                    End If
                                Case "3", "3-節"
                                    If styleSettings(i).StyleName <> DEFAULT_LEVEL3_STYLE Then
                                        newFieldCode = Replace(newFieldCode, DEFAULT_LEVEL3_STYLE, styleSettings(i).StyleName)
                                    End If
                                Case "4", "4-節"
                                    If styleSettings(i).StyleName <> DEFAULT_LEVEL4_STYLE Then
                                        newFieldCode = Replace(newFieldCode, DEFAULT_LEVEL4_STYLE, styleSettings(i).StyleName)
                                    End If
                                Case "5", "5-節"
                                    If styleSettings(i).StyleName <> DEFAULT_LEVEL5_STYLE Then
                                        newFieldCode = Replace(newFieldCode, DEFAULT_LEVEL5_STYLE, styleSettings(i).StyleName)
                                    End If
                            End Select
                        End If
                    Next i

                    If newFieldCode <> fieldCode Then
                        fld.Code.Text = newFieldCode
                        Debug.Print "  STYLEREF更新: " & Trim(fieldCode) & " → " & Trim(newFieldCode)
                    End If
                End If

                fld.Update
            Next fld
        Next hdr

        For Each hdr In sect.Footers
            For Each fld In hdr.Range.Fields
                fld.Update
            Next fld
        Next hdr
    Next sect

    wordDoc.Fields.Update

    Debug.Print "ヘッダーのフィールド更新完了"

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
' 1ページ目に「帳票」の文字があるかチェック
' ============================================================================
Private Function HasHyohyoOnFirstPage(ByRef wordDoc As Object) As Boolean
    On Error Resume Next

    Dim searchRange As Object
    Dim shp As Object

    Set searchRange = wordDoc.Content
    searchRange.Find.ClearFormatting
    If searchRange.Find.Execute(FindText:="帳票") Then
        If searchRange.Information(3) = 1 Then
            HasHyohyoOnFirstPage = True
            Exit Function
        End If
    End If

    For Each shp In wordDoc.Shapes
        If shp.TextFrame.HasText Then
            If InStr(shp.TextFrame.TextRange.Text, "帳票") > 0 Then
                If shp.Anchor.Information(3) = 1 Then
                    HasHyohyoOnFirstPage = True
                    Exit Function
                End If
            End If
        End If
    Next shp

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
Private Function MatchPattern(ByVal Text As String, ByVal Pattern As String) As Boolean
    On Error GoTo ErrorHandler

    If Text = "" Or Pattern = "" Then
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

    MatchPattern = regex.Test(Text)

    Set regex = Nothing
    Exit Function

ErrorHandler:
    MatchPattern = False
End Function

' ============================================================================
' 全角を半角に変換
' ============================================================================
Private Function ConvertToHalfWidth(ByVal Text As String) As String
    Dim i As Long
    Dim char As String
    Dim result As String
    Dim charCode As Long

    result = ""
    For i = 1 To Len(Text)
        char = Mid(Text, i, 1)
        charCode = AscW(char)

        Select Case charCode
            Case &HFF10 To &HFF19
                result = result & Chr(charCode - &HFF10 + Asc("0"))
            Case &HFF21 To &HFF3A
                result = result & Chr(charCode - &HFF21 + Asc("A"))
            Case &HFF41 To &HFF5A
                result = result & Chr(charCode - &HFF41 + Asc("a"))
            Case &HFF0D, &H2212, &H30FC
                result = result & "-"
            Case &HFF0C
                result = result & ","
            Case &HFF0E
                result = result & "."
            Case &HFF08
                result = result & "("
            Case &HFF09
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
                    Left(para.Range.Text, 50)
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

    fileCount = 0
    ReDim fileList(0 To 99)

    fileName = Dir(inputDir & "*.docx")
    Do While fileName <> ""
        fileList(fileCount) = fileName
        fileCount = fileCount + 1
        fileName = Dir()
    Loop

    fileName = Dir(inputDir & "*.doc")
    Do While fileName <> ""
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

    If fileCount = 0 Then
        MsgBox "Inputフォルダ内にWord文書(.docx/.doc)が見つかりませんでした。" & vbCrLf & vbCrLf & _
               "フォルダ: " & inputDir, vbExclamation, "ファイルなし"
        SelectWordFileFromInput = ""
        Exit Function
    End If

    If fileCount = 1 Then
        SelectWordFileFromInput = inputDir & fileList(0)
        Exit Function
    End If

    msg = "処理するWord文書を選択してください:" & vbCrLf & vbCrLf
    For i = 0 To fileCount - 1
        msg = msg & (i + 1) & ". " & fileList(i) & vbCrLf
    Next i
    msg = msg & vbCrLf & "番号を入力してください (1-" & fileCount & "):"

    Dim userInput As String
    userInput = InputBox(msg, "ファイル選択", "1")

    If userInput = "" Then
        SelectWordFileFromInput = ""
        Exit Function
    End If

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

    SelectWordFileFromInput = inputDir & fileList(selectedIndex)
End Function

' ============================================================================
' テスト用: イミディエイトウィンドウに情報を出力
' ============================================================================
Public Sub TestOrganizeWordBookmarks()
    OrganizeWordBookmarks
End Sub
