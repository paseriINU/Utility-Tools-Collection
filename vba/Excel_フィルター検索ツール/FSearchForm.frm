Option Explicit

' ========================================
' フィルター検索フォーム
' ========================================

Private Sub UserForm_Initialize()
    ' フォーム初期化
    Me.Caption = "フィルター検索"

    ' ラベルの初期化
    lblTitle.Caption = "検索キーワード（OR条件）"
    lblWord1.Caption = "キーワード1:"
    lblWord2.Caption = "キーワード2:"
    lblWord3.Caption = "キーワード3:"
    lblWord4.Caption = "キーワード4:"
    lblWord5.Caption = "キーワード5:"

    ' テキストボックスをクリア
    txtWord1.Value = ""
    txtWord2.Value = ""
    txtWord3.Value = ""
    txtWord4.Value = ""
    txtWord5.Value = ""

    ' ボタンのキャプション
    btnSearch.Caption = "検索"
    btnClear.Caption = "クリア"
    btnClose.Caption = "閉じる"

    txtWord1.SetFocus
End Sub

Private Sub btnSearch_Click()
    ' 検索実行
    Dim keywords() As String
    Dim count As Long

    ' 入力されたキーワードを収集
    ReDim keywords(1 To 5)
    count = 0

    If Trim(txtWord1.Value) <> "" Then count = count + 1: keywords(count) = Trim(txtWord1.Value)
    If Trim(txtWord2.Value) <> "" Then count = count + 1: keywords(count) = Trim(txtWord2.Value)
    If Trim(txtWord3.Value) <> "" Then count = count + 1: keywords(count) = Trim(txtWord3.Value)
    If Trim(txtWord4.Value) <> "" Then count = count + 1: keywords(count) = Trim(txtWord4.Value)
    If Trim(txtWord5.Value) <> "" Then count = count + 1: keywords(count) = Trim(txtWord5.Value)

    ' キーワードが1つもない場合
    If count = 0 Then
        MsgBox "キーワードを1つ以上入力してください。", vbExclamation, "入力エラー"
        txtWord1.SetFocus
        Exit Sub
    End If

    ' 配列をリサイズ
    ReDim Preserve keywords(1 To count)

    ' フォームを非表示
    Me.Hide

    ' フィルター実行
    Call FSearch.ApplyOrFilter(keywords)

    ' フォームを閉じる
    Unload Me
End Sub

Private Sub btnClear_Click()
    ' キーワードのみクリア（フィルターはそのまま）
    txtWord1.Value = ""
    txtWord2.Value = ""
    txtWord3.Value = ""
    txtWord4.Value = ""
    txtWord5.Value = ""
    txtWord1.SetFocus
End Sub

Private Sub btnClose_Click()
    ' フォームを閉じる
    Unload Me
End Sub

Private Sub txtWord1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then btnSearch_Click
End Sub

Private Sub txtWord2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then btnSearch_Click
End Sub

Private Sub txtWord3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then btnSearch_Click
End Sub

Private Sub txtWord4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then btnSearch_Click
End Sub

Private Sub txtWord5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then btnSearch_Click
End Sub
