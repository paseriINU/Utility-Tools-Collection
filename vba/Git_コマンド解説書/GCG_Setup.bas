Attribute VB_Name = "GCG_Setup"
Option Explicit

'==============================================================================
' Git コマンド解説書 生成ツール - セットアップモジュール
' 初期化とシート作成機能を提供
' ※このモジュールは初期化後に削除可能
'==============================================================================

'==============================================================================
' メイン: Git解説書を生成
'==============================================================================
Public Sub CreateGitCommandGuide()
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False

    ' シートを作成
    CreateSheet "Git基礎知識"
    CreateSheet "基本コマンド"
    CreateSheet "ブランチ操作"
    CreateSheet "リモート操作"
    CreateSheet "履歴・差分確認"
    CreateSheet "取り消し・修正"
    CreateSheet "実践シナリオ"
    CreateSheet "トラブル対処"

    ' 各シートの内容を作成
    FormatBasicsSheet
    FormatBasicCommandsSheet
    FormatBranchSheet
    FormatRemoteSheet
    FormatHistorySheet
    FormatUndoSheet
    FormatScenarioSheet
    FormatTroubleshootSheet

    ' 最初のシートをアクティブに
    Sheets("Git基礎知識").Activate

    Application.ScreenUpdating = True

    MsgBox "Git コマンド解説書を作成しました。" & vbCrLf & vbCrLf & _
           "【シート構成】" & vbCrLf & _
           "1. Git基礎知識 - Gitの概念を理解" & vbCrLf & _
           "2. 基本コマンド - よく使うコマンド" & vbCrLf & _
           "3. ブランチ操作 - ブランチの使い方" & vbCrLf & _
           "4. リモート操作 - GitHubとの連携" & vbCrLf & _
           "5. 履歴・差分確認 - 変更の確認方法" & vbCrLf & _
           "6. 取り消し・修正 - 間違いの修正方法" & vbCrLf & _
           "7. 実践シナリオ - 実際の作業フロー" & vbCrLf & _
           "8. トラブル対処 - よくある問題と解決法", _
           vbInformation, "作成完了"

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub

'==============================================================================
' シート作成ヘルパー
'==============================================================================
Private Sub CreateSheet(ByVal sheetName As String)
    Dim ws As Worksheet
    Dim exists As Boolean

    exists = False
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    If Not ws Is Nothing Then exists = True
    Err.Clear
    On Error GoTo 0

    If exists Then
        ws.Cells.Clear
        ws.Cells.Interior.ColorIndex = xlNone
    Else
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = sheetName
    End If
End Sub

