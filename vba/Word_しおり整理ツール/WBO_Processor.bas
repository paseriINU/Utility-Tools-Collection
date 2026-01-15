Attribute VB_Name = "WBO_Processor"
Option Explicit

' ============================================================================
' Word しおり整理ツール - 段落処理モジュール
' 文書処理、段落マッチング、スタイル適用ロジックを提供
' ============================================================================

' ============================================================================
' 文書全体を処理
' ============================================================================
Public Function ProcessDocument(ByRef wordDoc As Object, _
                                ByRef styleSettings() As StyleSetting, _
                                ByVal styleCount As Long, _
                                ByVal hasSections As Boolean, _
                                ByVal isHyohyoDocument As Boolean) As Long
    Dim processedCount As Long
    Dim sect As Object
    Dim sectIndex As Long
    Dim para As Object
    Dim shp As Object
    Dim shapePara As Object
    Dim headerIsEmpty As Boolean

    processedCount = 0
    sectIndex = 0

    For Each sect In wordDoc.Sections
        sectIndex = sectIndex + 1
        Debug.Print "-----------------------------------------"
        Debug.Print "セクション " & sectIndex & " を処理中..."

        headerIsEmpty = IsHeaderEmpty(sect)
        Debug.Print "  ヘッダー空白: " & headerIsEmpty

        ' セクション内の段落を処理
        For Each para In sect.Range.Paragraphs
            processedCount = processedCount + ProcessParagraph(para, styleSettings, styleCount, hasSections, headerIsEmpty, isHyohyoDocument, wordDoc)
        Next para

        ' セクション内の図形も処理
        On Error Resume Next
        For Each shp In sect.Range.ShapeRange
            If shp.TextFrame.HasText Then
                For Each shapePara In shp.TextFrame.TextRange.Paragraphs
                    processedCount = processedCount + ProcessParagraph(shapePara, styleSettings, styleCount, hasSections, headerIsEmpty, isHyohyoDocument, wordDoc)
                Next shapePara
            End If
        Next shp
        Err.Clear
        On Error GoTo 0
    Next sect

    ProcessDocument = processedCount
End Function

' ============================================================================
' 段落処理
' ============================================================================
Public Function ProcessParagraph(ByRef para As Object, _
                                 ByRef styleSettings() As StyleSetting, _
                                 ByVal styleCount As Long, _
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

    paraText = CleanParagraphText(para.Range.Text)
    paraTextHalf = ToHalfWidth(paraText)

    ' スキップ条件チェック
    If Not ShouldProcessParagraph(para, paraText) Then
        ProcessParagraph = 0
        Exit Function
    End If

    ' パターンマッチングループ
    Dim i As Long
    Dim setting As StyleSetting
    Dim matched As Boolean
    Dim targetStyle As String
    Dim detectedLevel As String

    matched = False

    For i = 0 To styleCount - 1
        setting = styleSettings(i)

        If setting.StyleName = "" Then GoTo NextSetting

        Select Case setting.Category
            Case "パターン"
                If MatchPatternCategory(para, paraText, paraTextHalf, setting, hasSections, headerIsEmpty, styleSettings, styleCount, matched, targetStyle, detectedLevel) Then
                    Exit For
                End If

            Case "帳票"
                If MatchHyohyoCategory(paraTextHalf, setting, isHyohyoDocument, paraText, matched, targetStyle, detectedLevel) Then
                    Exit For
                End If

            Case "特定"
                If MatchSpecificCategory(para, paraText, setting, matched, targetStyle, detectedLevel) Then
                    ProcessParagraph = 1
                    Exit Function
                End If

            Case "例外"
                If MatchExceptionCategory(para, setting, styleSettings, styleCount, wordDoc, matched, targetStyle, detectedLevel) Then
                    Exit For
                End If
        End Select
NextSetting:
    Next i

    ' スタイル適用
    If matched And targetStyle <> "" Then
        ApplyStyle para, targetStyle
        Debug.Print "  [" & detectedLevel & "] " & Left(paraText, 50)
        ProcessParagraph = 1
        Exit Function
    End If

    ProcessParagraph = 0
    Exit Function

ErrorHandler:
    ProcessParagraph = 0
End Function

' ============================================================================
' 段落テキストをクリーンアップ
' ============================================================================
Private Function CleanParagraphText(ByVal Text As String) As String
    Dim result As String
    result = Trim(Replace(Text, vbCr, ""))
    result = Replace(result, Chr(13), "")
    result = Replace(result, Chr(12), "")
    result = Replace(result, Chr(11), "")
    result = Replace(result, Chr(7), "")
    CleanParagraphText = Trim(result)
End Function

' ============================================================================
' 段落を処理すべきかチェック
' ============================================================================
Private Function ShouldProcessParagraph(ByRef para As Object, ByVal paraText As String) As Boolean
    ShouldProcessParagraph = False

    ' 空の段落はスキップ
    If paraText = "" Then Exit Function

    ' 「参照」を含む段落はスキップ
    If InStr(paraText, "参照") > 0 Then Exit Function

    ' 「・」で始まる段落はスキップ
    If Left(paraText, 1) = "・" Then Exit Function

    ' ハイパーリンクを含む段落はスキップ
    On Error Resume Next
    If para.Range.Hyperlinks.Count > 0 Then Exit Function
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0

    ' 表内の段落はスキップ
    On Error Resume Next
    If para.Range.Information(12) = True Then Exit Function
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0

    ShouldProcessParagraph = True
End Function

' ============================================================================
' パターンカテゴリのマッチング
' ============================================================================
Private Function MatchPatternCategory(ByRef para As Object, _
                                      ByVal paraText As String, _
                                      ByVal paraTextHalf As String, _
                                      ByRef setting As StyleSetting, _
                                      ByVal hasSections As Boolean, _
                                      ByVal headerIsEmpty As Boolean, _
                                      ByRef styleSettings() As StyleSetting, _
                                      ByVal styleCount As Long, _
                                      ByRef matched As Boolean, _
                                      ByRef targetStyle As String, _
                                      ByRef detectedLevel As String) As Boolean
    MatchPatternCategory = False

    ' レベルに「-節」が含まれる場合は節構造あり時のみ適用
    If InStr(setting.Level, "-節") > 0 Then
        If Not hasSections Then Exit Function
    ElseIf setting.Level = "1" Then
        ' レベル1はヘッダー空白時のみ
        If Not headerIsEmpty Then Exit Function
    Else
        ' 同じレベルで「-節」付きがあればスキップ（節構造あり時）
        If setting.Level <> "1" And setting.Level <> "2" Then
            If hasSections Then
                If HasSectionVariant(styleSettings, styleCount, setting.Level) Then
                    Exit Function
                End If
            End If
        End If
    End If

    ' パターンマッチ
    If setting.Pattern <> "" Then
        If RegexMatch(paraText, setting.Pattern) Or _
           RegexMatch(paraTextHalf, setting.Pattern) Then
            matched = True
            targetStyle = setting.StyleName
            detectedLevel = setting.Level
            MatchPatternCategory = True
        End If
    End If
End Function

' ============================================================================
' 帳票カテゴリのマッチング
' ============================================================================
Private Function MatchHyohyoCategory(ByVal paraTextHalf As String, _
                                     ByRef setting As StyleSetting, _
                                     ByVal isHyohyoDocument As Boolean, _
                                     ByVal paraText As String, _
                                     ByRef matched As Boolean, _
                                     ByRef targetStyle As String, _
                                     ByRef detectedLevel As String) As Boolean
    MatchHyohyoCategory = False

    If isHyohyoDocument And setting.Pattern <> "" Then
        If RegexMatch(paraTextHalf, setting.Pattern) Then
            matched = True
            targetStyle = setting.StyleName
            detectedLevel = "帳票"
            Debug.Print "  [帳票番号検出] " & Left(paraText, 50)
            MatchHyohyoCategory = True
        End If
    End If
End Function

' ============================================================================
' 特定カテゴリのマッチング
' ============================================================================
Private Function MatchSpecificCategory(ByRef para As Object, _
                                       ByVal paraText As String, _
                                       ByRef setting As StyleSetting, _
                                       ByRef matched As Boolean, _
                                       ByRef targetStyle As String, _
                                       ByRef detectedLevel As String) As Boolean
    MatchSpecificCategory = False

    If setting.Pattern <> "" Then
        If paraText = setting.Pattern Then
            ApplyStyle para, setting.StyleName

            ' アウトラインレベルを設定
            If IsNumeric(setting.Level) Then
                Dim outlineLevel As Long
                outlineLevel = CLng(setting.Level)
                If outlineLevel >= 1 And outlineLevel <= 9 Then
                    On Error Resume Next
                    para.Range.ParagraphFormat.OutlineLevel = outlineLevel
                    On Error GoTo 0
                End If
            End If

            Debug.Print "  [特定→アウトライン" & setting.Level & "] " & paraText
            MatchSpecificCategory = True
        End If
    End If
End Function

' ============================================================================
' 例外カテゴリのマッチング
' ============================================================================
Private Function MatchExceptionCategory(ByRef para As Object, _
                                        ByRef setting As StyleSetting, _
                                        ByRef styleSettings() As StyleSetting, _
                                        ByVal styleCount As Long, _
                                        ByRef wordDoc As Object, _
                                        ByRef matched As Boolean, _
                                        ByRef targetStyle As String, _
                                        ByRef detectedLevel As String) As Boolean
    MatchExceptionCategory = False

    ' 例外1: パターン外だが見出しスタイル適用済み
    If setting.Level = "1" Then
        Dim currentStyle As String
        On Error Resume Next
        currentStyle = para.style.NameLocal
        If Err.Number <> 0 Then
            currentStyle = ""
            Err.Clear
        End If
        On Error GoTo 0

        If IsLevelStyle(currentStyle, styleSettings, styleCount) Then
            matched = True
            targetStyle = setting.StyleName
            detectedLevel = "例外1"
            MatchExceptionCategory = True
            Exit Function
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
        On Error GoTo 0

        If (currentOutline >= 1 And currentOutline <= 9) Then
            matched = True
            targetStyle = setting.StyleName
            detectedLevel = "例外2"
            MatchExceptionCategory = True
        ElseIf HasOutlineInStyle(para, wordDoc) Then
            matched = True
            targetStyle = setting.StyleName
            detectedLevel = "例外2"
            MatchExceptionCategory = True
        End If
    End If
End Function

' ============================================================================
' 全角を半角に変換
' ============================================================================
Private Function ToHalfWidth(ByVal Text As String) As String
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

    ToHalfWidth = result
End Function

' ============================================================================
' 正規表現マッチ
' ============================================================================
Private Function RegexMatch(ByVal Text As String, ByVal Pattern As String) As Boolean
    On Error GoTo ErrorHandler

    If Text = "" Or Pattern = "" Then
        RegexMatch = False
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

    RegexMatch = regex.Test(Text)

    Set regex = Nothing
    Exit Function

ErrorHandler:
    RegexMatch = False
End Function
