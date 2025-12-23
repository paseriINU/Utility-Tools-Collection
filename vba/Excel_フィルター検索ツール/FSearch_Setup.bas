Option Explicit

' ========================================
' フィルター検索ツール - 初期化モジュール
' フォームとコントロールを自動作成
' ========================================

' フォーム名（短縮版）
Private Const FORM_NAME As String = "FSearchForm"

' ========================================
' 公開プロシージャ
' ========================================

' フォームを自動作成して初期化
Public Sub InitializeFilterSearchTool()
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False

    ' 既存フォームがあれば削除
    DeleteExistingForm

    ' フォームを作成
    CreateFilterSearchForm

    Application.ScreenUpdating = True

    MsgBox "フィルター検索ツールの初期化が完了しました。" & vbCrLf & vbCrLf & _
           "使い方:" & vbCrLf & _
           "  ShowFilterSearchForm マクロを実行してください。", vbInformation, "初期化完了"

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True

    If Err.Number = 1004 Then
        MsgBox "VBAプロジェクトへのアクセスが許可されていません。" & vbCrLf & vbCrLf & _
               "以下の設定を有効にしてください:" & vbCrLf & _
               "  ファイル → オプション → トラストセンター → トラストセンターの設定" & vbCrLf & _
               "  → マクロの設定 → 「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」", _
               vbCritical, "エラー"
    Else
        MsgBox "エラーが発生しました: " & Err.Description, vbCritical, "エラー"
    End If
End Sub

' ========================================
' 内部プロシージャ
' ========================================

' 既存のフォームを削除
Private Sub DeleteExistingForm()
    Dim vbComp As Object

    On Error Resume Next
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        If vbComp.Name = FORM_NAME Then
            ThisWorkbook.VBProject.VBComponents.Remove vbComp
            Exit For
        End If
    Next vbComp
    On Error GoTo 0
End Sub

' フォームを作成
Private Sub CreateFilterSearchForm()
    Dim vbComp As Object
    Dim frm As Object
    Dim ctrl As Object
    Dim yPos As Single
    Dim labelWidth As Single
    Dim textWidth As Single
    Dim btnWidth As Single
    Dim margin As Single
    Dim rowHeight As Single

    ' フォームを追加
    Set vbComp = ThisWorkbook.VBProject.VBComponents.Add(3) ' vbext_ct_MSForm = 3
    vbComp.Name = FORM_NAME

    Set frm = vbComp.Designer

    ' フォームのプロパティを設定
    With frm
        .Caption = "フィルター検索"
        .Width = 320
        .Height = 280
        .StartUpPosition = 1 ' CenterOwner
    End With

    ' サイズ設定
    margin = 12
    labelWidth = 80
    textWidth = 200
    btnWidth = 70
    rowHeight = 24
    yPos = margin

    ' ----------------------------------------
    ' タイトルラベル
    ' ----------------------------------------
    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblTitle")
    With ctrl
        .Left = margin
        .Top = yPos
        .Width = 280
        .Height = 18
        .Caption = "検索キーワード（OR条件）"
        .Font.Bold = True
    End With
    yPos = yPos + 24

    ' ----------------------------------------
    ' キーワード入力欄 (5つ)
    ' ----------------------------------------
    Dim i As Integer
    For i = 1 To 5
        ' ラベル
        Set ctrl = frm.Controls.Add("Forms.Label.1", "lblWord" & i)
        With ctrl
            .Left = margin
            .Top = yPos + 2
            .Width = labelWidth
            .Height = 16
            .Caption = "キーワード" & i & ":"
        End With

        ' テキストボックス
        Set ctrl = frm.Controls.Add("Forms.TextBox.1", "txtWord" & i)
        With ctrl
            .Left = margin + labelWidth
            .Top = yPos
            .Width = textWidth
            .Height = 20
        End With

        yPos = yPos + rowHeight
    Next i

    yPos = yPos + 10

    ' ----------------------------------------
    ' ボタン
    ' ----------------------------------------
    ' 検索ボタン
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnSearch")
    With ctrl
        .Left = margin
        .Top = yPos
        .Width = btnWidth
        .Height = 26
        .Caption = "検索"
    End With

    ' クリアボタン
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnClear")
    With ctrl
        .Left = margin + btnWidth + 10
        .Top = yPos
        .Width = btnWidth
        .Height = 26
        .Caption = "クリア"
    End With

    ' 閉じるボタン
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnClose")
    With ctrl
        .Left = margin + (btnWidth + 10) * 2
        .Top = yPos
        .Width = btnWidth
        .Height = 26
        .Caption = "閉じる"
    End With

    ' ----------------------------------------
    ' フォームにコードを追加
    ' ----------------------------------------
    AddFormCode vbComp

End Sub

' フォームにコードを追加
Private Sub AddFormCode(vbComp As Object)
    Dim code As String

    code = "Option Explicit" & vbCrLf & vbCrLf

    ' 検索ボタンクリック
    code = code & "Private Sub btnSearch_Click()" & vbCrLf
    code = code & "    Dim keywords() As String" & vbCrLf
    code = code & "    Dim count As Long" & vbCrLf
    code = code & "    Dim i As Long" & vbCrLf
    code = code & "    " & vbCrLf
    code = code & "    ReDim keywords(1 To 5)" & vbCrLf
    code = code & "    count = 0" & vbCrLf
    code = code & "    " & vbCrLf
    code = code & "    If Trim(txtWord1.Value) <> """" Then count = count + 1: keywords(count) = Trim(txtWord1.Value)" & vbCrLf
    code = code & "    If Trim(txtWord2.Value) <> """" Then count = count + 1: keywords(count) = Trim(txtWord2.Value)" & vbCrLf
    code = code & "    If Trim(txtWord3.Value) <> """" Then count = count + 1: keywords(count) = Trim(txtWord3.Value)" & vbCrLf
    code = code & "    If Trim(txtWord4.Value) <> """" Then count = count + 1: keywords(count) = Trim(txtWord4.Value)" & vbCrLf
    code = code & "    If Trim(txtWord5.Value) <> """" Then count = count + 1: keywords(count) = Trim(txtWord5.Value)" & vbCrLf
    code = code & "    " & vbCrLf
    code = code & "    If count = 0 Then" & vbCrLf
    code = code & "        MsgBox ""キーワードを1つ以上入力してください。"", vbExclamation, ""入力エラー""" & vbCrLf
    code = code & "        txtWord1.SetFocus" & vbCrLf
    code = code & "        Exit Sub" & vbCrLf
    code = code & "    End If" & vbCrLf
    code = code & "    " & vbCrLf
    code = code & "    ReDim Preserve keywords(1 To count)" & vbCrLf
    code = code & "    Call FSearch.ApplyOrFilter(keywords)" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf

    ' クリアボタンクリック
    code = code & "Private Sub btnClear_Click()" & vbCrLf
    code = code & "    Call FSearch.ClearFilter" & vbCrLf
    code = code & "    txtWord1.Value = """": txtWord2.Value = """": txtWord3.Value = """"" & vbCrLf
    code = code & "    txtWord4.Value = """": txtWord5.Value = """"" & vbCrLf
    code = code & "    txtWord1.SetFocus" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf

    ' 閉じるボタンクリック
    code = code & "Private Sub btnClose_Click()" & vbCrLf
    code = code & "    Unload Me" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf

    ' Enterキー対応
    Dim j As Integer
    For j = 1 To 5
        code = code & "Private Sub txtWord" & j & "_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)" & vbCrLf
        code = code & "    If KeyCode = 13 Then btnSearch_Click" & vbCrLf
        code = code & "End Sub" & vbCrLf & vbCrLf
    Next j

    ' コードを追加
    vbComp.CodeModule.AddFromString code
End Sub
