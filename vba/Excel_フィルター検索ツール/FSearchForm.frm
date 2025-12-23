VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FSearchForm
   Caption         =   "フィルター検索"
   ClientHeight    =   4500
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4680
   OleObjectBlob   =   "FSearchForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "FSearchForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
End Sub

Private Sub btnSearch_Click()
    ' 検索実行
    Dim keywords() As String
    Dim count As Long
    Dim i As Long

    ' 入力されたキーワードを収集
    ReDim keywords(1 To 5)
    count = 0

    If Trim(txtWord1.Value) <> "" Then
        count = count + 1
        keywords(count) = Trim(txtWord1.Value)
    End If
    If Trim(txtWord2.Value) <> "" Then
        count = count + 1
        keywords(count) = Trim(txtWord2.Value)
    End If
    If Trim(txtWord3.Value) <> "" Then
        count = count + 1
        keywords(count) = Trim(txtWord3.Value)
    End If
    If Trim(txtWord4.Value) <> "" Then
        count = count + 1
        keywords(count) = Trim(txtWord4.Value)
    End If
    If Trim(txtWord5.Value) <> "" Then
        count = count + 1
        keywords(count) = Trim(txtWord5.Value)
    End If

    ' キーワードが1つもない場合
    If count = 0 Then
        MsgBox "キーワードを1つ以上入力してください。", vbExclamation, "入力エラー"
        txtWord1.SetFocus
        Exit Sub
    End If

    ' 配列をリサイズ
    ReDim Preserve keywords(1 To count)

    ' フィルター実行
    Call FilterSearch.ApplyOrFilter(keywords)

End Sub

Private Sub btnClear_Click()
    ' フィルタークリア
    Call FilterSearch.ClearFilter

    ' テキストボックスもクリア
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
    If KeyCode = 13 Then ' Enterキー
        btnSearch_Click
    End If
End Sub

Private Sub txtWord2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        btnSearch_Click
    End If
End Sub

Private Sub txtWord3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        btnSearch_Click
    End If
End Sub

Private Sub txtWord4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        btnSearch_Click
    End If
End Sub

Private Sub txtWord5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        btnSearch_Click
    End If
End Sub
