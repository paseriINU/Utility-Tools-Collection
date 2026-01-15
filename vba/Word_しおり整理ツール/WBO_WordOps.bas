Attribute VB_Name = "WBO_WordOps"
Option Explicit

' ============================================================================
' Word しおり整理ツール - Word操作モジュール
' Word文書操作、スタイル適用、文書判定機能を提供
' ============================================================================

' ============================================================================
' スタイルの存在確認（動的版）
' ============================================================================
Public Function ValidateStyles(ByRef wordDoc As Object, _
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

        If setting.StyleName = "" Then GoTo NextStyle
        If InStr(setting.Level, "-節") > 0 And Not hasSections Then GoTo NextStyle
        If setting.Category = "パターン" And setting.Level <> "1" And setting.Level <> "2" Then
            If Not InStr(setting.Level, "-節") > 0 And hasSections Then
                If HasSectionVariant(styleSettings, styleCount, setting.Level) Then
                    GoTo NextStyle
                End If
            End If
        End If

        On Error Resume Next
        checkedStyles.Add setting.StyleName, setting.StyleName
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextStyle
        End If
        On Error GoTo 0

        If Not StyleExists(wordDoc, setting.StyleName) Then
            missingStyles = missingStyles & "  - " & setting.StyleName & _
                           " (" & setting.Category & ": " & setting.Pattern & ")" & vbCrLf
        End If
NextStyle:
    Next i

    ValidateStyles = missingStyles
End Function

' ============================================================================
' 同じレベルで「-節」バリアントがあるかチェック
' ============================================================================
Public Function HasSectionVariant(ByRef styleSettings() As StyleSetting, _
                                  ByVal styleCount As Long, _
                                  ByVal Level As String) As Boolean
    Dim i As Long
    For i = 0 To styleCount - 1
        If styleSettings(i).Level = Level & "-節" Then
            HasSectionVariant = True
            Exit Function
        End If
    Next i
    HasSectionVariant = False
End Function

' ============================================================================
' スタイルがレベル系スタイルかチェック（動的版）
' ============================================================================
Public Function IsLevelStyle(ByVal styleName As String, _
                             ByRef styleSettings() As StyleSetting, _
                             ByVal styleCount As Long) As Boolean
    If styleName = "" Then
        IsLevelStyle = False
        Exit Function
    End If

    Dim i As Long
    For i = 0 To styleCount - 1
        If styleSettings(i).Category = "パターン" And styleSettings(i).StyleName = styleName Then
            IsLevelStyle = True
            Exit Function
        End If
    Next i

    IsLevelStyle = False
End Function

' ============================================================================
' 指定したスタイルがWord文書に存在するかチェック
' ============================================================================
Public Function StyleExists(ByRef wordDoc As Object, ByVal styleName As String) As Boolean
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
' スタイル適用
' ============================================================================
Public Sub ApplyStyle(ByRef para As Object, ByVal styleName As String)
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
' スタイルにアウトラインが定義されているかチェック
' ============================================================================
Public Function HasOutlineInStyle(ByRef para As Object, ByRef wordDoc As Object) As Boolean
    On Error GoTo ErrorHandler

    Dim styleName As String
    styleName = para.style.NameLocal

    Dim style As Object
    Set style = wordDoc.styles(styleName)

    If style.ParagraphFormat.OutlineLevel >= 1 And _
       style.ParagraphFormat.OutlineLevel <= 9 Then
        HasOutlineInStyle = True
    Else
        HasOutlineInStyle = False
    End If

    Exit Function

ErrorHandler:
    HasOutlineInStyle = False
End Function

' ============================================================================
' 文書内のヘッダーに「第X節」が存在するかチェック
' ============================================================================
Public Function HasSectionsInDoc(ByRef wordDoc As Object) As Boolean
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
                HasSectionsInDoc = True
                Exit Function
            End If
        End If
        Err.Clear
    Next sect

    Set regex = Nothing
    HasSectionsInDoc = False
End Function

' ============================================================================
' 1ページ目に「帳票」の文字があるかチェック
' ============================================================================
Public Function HasHyohyoOnPage1(ByRef wordDoc As Object) As Boolean
    On Error Resume Next

    Dim searchRange As Object
    Dim shp As Object

    Set searchRange = wordDoc.Content
    searchRange.Find.ClearFormatting
    If searchRange.Find.Execute(FindText:="帳票") Then
        If searchRange.Information(3) = 1 Then
            HasHyohyoOnPage1 = True
            Exit Function
        End If
    End If

    For Each shp In wordDoc.Shapes
        If shp.TextFrame.HasText Then
            If InStr(shp.TextFrame.TextRange.Text, "帳票") > 0 Then
                If shp.Anchor.Information(3) = 1 Then
                    HasHyohyoOnPage1 = True
                    Exit Function
                End If
            End If
        End If
    Next shp

    HasHyohyoOnPage1 = False
    On Error GoTo 0
End Function

' ============================================================================
' セクションのヘッダーが空白かどうかをチェック
' ============================================================================
Public Function IsHeaderEmpty(ByRef sect As Object) As Boolean
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
' ヘッダー内のフィールドを更新
' ============================================================================
Public Sub UpdateHeaderFields(ByRef wordDoc As Object, _
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

                If InStr(fieldCode, "STYLEREF") > 0 Then
                    newFieldCode = fieldCode

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
' Word文書を保存してPDFエクスポート
' ============================================================================
Public Sub SaveAndExport(ByRef wordDoc As Object, _
                         ByVal outputWordPath As String, _
                         ByVal outputPdfPath As String, _
                         ByVal doPdfOutput As Boolean)
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
End Sub
