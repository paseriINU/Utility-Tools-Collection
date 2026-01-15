Attribute VB_Name = "GLV_Config"
Option Explicit

'==============================================================================
' Git Log 可視化ツール - 設定モジュール
' 定数、型定義を提供
' ※このモジュールはSetupモジュールを削除しても動作するよう設計
'==============================================================================

' ============================================================================
' シート名定数
' ============================================================================
Public Const SHEET_MAIN As String = "メイン"
Public Const SHEET_DASHBOARD As String = "ダッシュボード"
Public Const SHEET_HISTORY As String = "コミット履歴"
Public Const SHEET_BRANCH_GRAPH As String = "ブランチグラフ"

' ============================================================================
' メインシートのセル位置
' ============================================================================
Public Const CELL_REPO_PATH As String = "D8"
Public Const CELL_COMMIT_COUNT As String = "D10"

' ============================================================================
' Gitコマンドのパス
' ============================================================================
' 通常は "git" でOK。パスが通っていない場合はフルパス指定
Public Const GIT_COMMAND As String = "git"

' ============================================================================
' データ構造
' ============================================================================
Public Type CommitInfo
    Hash As String          ' コミットハッシュ（短縮）
    FullHash As String      ' コミットハッシュ（フル）
    Author As String        ' 作者名
    AuthorEmail As String   ' 作者メール
    CommitDate As Date      ' コミット日時
    Subject As String       ' コミットメッセージ（件名）
    RefNames As String      ' ブランチ・タグ名
    ParentHashes As String  ' 親コミットハッシュ（スペース区切り）
    ParentCount As Long     ' 親コミット数（0=初期, 1=通常, 2+=マージ）
    FilesChanged As Long    ' 変更ファイル数
    Insertions As Long      ' 追加行数
    Deletions As Long       ' 削除行数
End Type

