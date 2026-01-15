Attribute VB_Name = "FDE_Config"
Option Explicit

'==============================================================================
' Git 差分ファイル抽出ツール（VBA版） - 設定モジュール
' 定数、型定義を提供
' ※バッチ版「Git_差分ファイル抽出ツール.bat」のVBA版
'==============================================================================

' ============================================================================
' シート名定数
' ============================================================================
Public Const SHEET_MAIN As String = "メイン"
Public Const SHEET_RESULT As String = "比較結果"

' ============================================================================
' メインシートのセル位置
' ============================================================================
Public Const CELL_REPO_PATH As String = "D8"
Public Const CELL_OUTPUT_FOLDER As String = "D10"
Public Const CELL_BASE_REF As String = "D14"
Public Const CELL_TARGET_REF As String = "D16"
Public Const CELL_COMPARE_MODE As String = "D12"

' ============================================================================
' 比較結果シートの列定義
' ============================================================================
Public Const COL_DIFF_MARK As Long = 1      ' A列: 抽出対象マーク
Public Const COL_RELATIVE_PATH As Long = 2  ' B列: 相対パス
Public Const COL_FILE_NAME As Long = 3      ' C列: ファイル名
Public Const COL_STATUS As Long = 4         ' D列: 状態（A/M/D）
Public Const COL_BASE_EXISTS As Long = 5    ' E列: 修正前存在
Public Const COL_BASE_SIZE As Long = 6      ' F列: 修正前サイズ
Public Const COL_TARGET_EXISTS As Long = 7  ' G列: 修正後存在
Public Const COL_TARGET_SIZE As Long = 8    ' H列: 修正後サイズ
Public Const COL_CHANGE_TYPE As Long = 9    ' I列: 変更種別

' ============================================================================
' Git差分ステータス定数
' ============================================================================
Public Const STATUS_ADDED As String = "A"      ' 追加
Public Const STATUS_MODIFIED As String = "M"   ' 変更
Public Const STATUS_DELETED As String = "D"    ' 削除
Public Const STATUS_RENAMED As String = "R"    ' 名前変更
Public Const STATUS_COPIED As String = "C"     ' コピー

' ============================================================================
' 変更種別の日本語表示
' ============================================================================
Public Const CHANGE_ADDED As String = "新規"
Public Const CHANGE_MODIFIED As String = "変更"
Public Const CHANGE_DELETED As String = "削除"
Public Const CHANGE_RENAMED As String = "名前変更"

' ============================================================================
' 比較モード定数
' ============================================================================
Public Const MODE_BRANCH As String = "ブランチ間"
Public Const MODE_COMMIT As String = "コミット間"

' ============================================================================
' Gitコマンドのパス
' ============================================================================
Public Const GIT_COMMAND As String = "git"

' ============================================================================
' 差分ファイル情報を格納する型
' ============================================================================
Public Type DiffFileInfo
    RelativePath As String      ' 相対パス
    FileName As String          ' ファイル名
    Status As String            ' 状態（A/M/D/R/C）
    ChangeType As String        ' 変更種別（日本語）
    BaseExists As Boolean       ' 修正前に存在
    BaseSize As Double          ' 修正前サイズ
    TargetExists As Boolean     ' 修正後に存在
    TargetSize As Double        ' 修正後サイズ
End Type

' ============================================================================
' ブランチ情報を格納する型
' ============================================================================
Public Type BranchInfo
    Name As String              ' ブランチ名
    IsCurrent As Boolean        ' 現在のブランチか
    LastCommit As String        ' 最新コミットハッシュ
    LastCommitDate As String    ' 最新コミット日時
End Type

' ============================================================================
' コミット情報を格納する型
' ============================================================================
Public Type CommitInfo
    Hash As String              ' コミットハッシュ（短縮）
    FullHash As String          ' コミットハッシュ（フル）
    CommitDate As String        ' コミット日時
    Subject As String           ' コミットメッセージ
    Author As String            ' 作者
End Type

