Attribute VB_Name = "JRM_Config"
Option Explicit

'==============================================================================
' JP1 REST ジョブ管理ツール - 設定モジュール
' 定数、設定取得機能を提供
' ※このモジュールはSetupモジュールを削除しても動作するよう設計
'==============================================================================

' ============================================================================
' シート名定数
' ============================================================================
Public Const SHEET_SETTINGS As String = "設定"
Public Const SHEET_TREE As String = "ツリー表示"
Public Const SHEET_LOG As String = "実行ログ"

' ============================================================================
' 設定セル位置（設定シート）
' ============================================================================
Public Const ROW_WEB_CONSOLE_HOST As Long = 7
Public Const ROW_WEB_CONSOLE_PORT As Long = 8
Public Const ROW_USE_HTTPS As Long = 9
Public Const ROW_MANAGER_HOST As Long = 10
Public Const ROW_SCHEDULER_SERVICE As Long = 11
Public Const ROW_JP1_USER As Long = 12
Public Const ROW_JP1_PASSWORD As Long = 13
Public Const ROW_ROOT_PATH As Long = 14
Public Const ROW_WAIT_COMPLETION As Long = 16
Public Const ROW_POLLING_INTERVAL As Long = 17
Public Const ROW_TIMEOUT As Long = 18
Public Const ROW_DEBUG_MODE As Long = 20
Public Const COL_SETTING_VALUE As Long = 3

' ============================================================================
' ツリー表示シートの列位置
' ============================================================================
Public Const COL_EXPAND As Long = 1          ' 展開/折りたたみ（>[v]）
Public Const COL_UNIT_NAME As Long = 2       ' インデント付きユニット名
Public Const COL_UNIT_PATH As Long = 3       ' ユニットパス（フルパス）
Public Const COL_UNIT_TYPE As Long = 4       ' ユニット種別
Public Const COL_STATUS As Long = 5          ' 状態
Public Const COL_LAST_RESULT As Long = 6     ' 最終実行結果
Public Const COL_EXEC_ID As Long = 7         ' execID
Public Const COL_START_TIME As Long = 8      ' 開始時刻
Public Const COL_END_TIME As Long = 9        ' 終了時刻
Public Const COL_SELECT As Long = 10         ' 選択チェック
Public Const ROW_TREE_HEADER As Long = 4
Public Const ROW_TREE_DATA_START As Long = 5

' ============================================================================
' ログシートの行位置
' ============================================================================
Public Const ROW_LOG_HEADER As Long = 4
Public Const ROW_LOG_DATA_START As Long = 5

' ============================================================================
' ユニット種別の日本語表示名
' ============================================================================
Public Const TYPE_GROUP As String = "グループ"
Public Const TYPE_ROOTNET As String = "ルートジョブネット"
Public Const TYPE_NET As String = "ネストジョブネット"
Public Const TYPE_JOB As String = "ジョブ"
Public Const TYPE_UNKNOWN As String = "その他"

' ============================================================================
' 展開アイコン・チェックボックス
' ============================================================================
Public Const ICON_COLLAPSED As String = ">"  ' 折りたたみ状態
Public Const ICON_EXPANDED As String = "v"   ' 展開状態
Public Const CHECK_ON As String = "v"
Public Const CHECK_OFF As String = ""

' ============================================================================
' モジュールレベル変数
' ============================================================================
Public g_DebugMode As Boolean

' ============================================================================
' 設定取得
' ============================================================================
Public Function GetConfig() As Object
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_SETTINGS)

    Dim config As Object
    Set config = CreateObject("Scripting.Dictionary")

    config("WebConsoleHost") = Trim(CStr(ws.Cells(ROW_WEB_CONSOLE_HOST, COL_SETTING_VALUE).Value))
    config("WebConsolePort") = Trim(CStr(ws.Cells(ROW_WEB_CONSOLE_PORT, COL_SETTING_VALUE).Value))
    config("UseHttps") = Trim(CStr(ws.Cells(ROW_USE_HTTPS, COL_SETTING_VALUE).Value))
    config("ManagerHost") = Trim(CStr(ws.Cells(ROW_MANAGER_HOST, COL_SETTING_VALUE).Value))
    config("SchedulerService") = Trim(CStr(ws.Cells(ROW_SCHEDULER_SERVICE, COL_SETTING_VALUE).Value))
    config("JP1User") = Trim(CStr(ws.Cells(ROW_JP1_USER, COL_SETTING_VALUE).Value))
    config("JP1Password") = Trim(CStr(ws.Cells(ROW_JP1_PASSWORD, COL_SETTING_VALUE).Value))
    config("RootPath") = Trim(CStr(ws.Cells(ROW_ROOT_PATH, COL_SETTING_VALUE).Value))
    config("WaitCompletion") = Trim(CStr(ws.Cells(ROW_WAIT_COMPLETION, COL_SETTING_VALUE).Value))
    config("PollingInterval") = ws.Cells(ROW_POLLING_INTERVAL, COL_SETTING_VALUE).Value
    config("Timeout") = ws.Cells(ROW_TIMEOUT, COL_SETTING_VALUE).Value
    config("DebugMode") = Trim(CStr(ws.Cells(ROW_DEBUG_MODE, COL_SETTING_VALUE).Value))

    ' デバッグモード設定をモジュール変数に保存
    g_DebugMode = (config("DebugMode") = "はい")

    ' 必須項目チェック
    If config("WebConsoleHost") = "" Or config("SchedulerService") = "" Or config("JP1User") = "" Then
        MsgBox "接続設定が不完全です。" & vbCrLf & _
               "設定シートで必須項目を入力してください。", vbExclamation
        Set GetConfig = Nothing
        Exit Function
    End If

    Set GetConfig = config
    Exit Function

ErrorHandler:
    MsgBox "設定の取得に失敗しました。" & vbCrLf & Err.Description, vbCritical
    Set GetConfig = Nothing
End Function

' ============================================================================
' ユーティリティ関数
' ============================================================================

' インデントレベル取得
Public Function GetIndentLevel(unitPath As String) As Long
    If unitPath = "" Or unitPath = "/" Then
        GetIndentLevel = 0
        Exit Function
    End If

    ' パスの "/" の数 - 1 がインデントレベル
    Dim slashCount As Long
    slashCount = Len(unitPath) - Len(Replace(unitPath, "/", ""))

    GetIndentLevel = slashCount - 1
    If GetIndentLevel < 0 Then GetIndentLevel = 0
End Function

' ユニットタイプの日本語表示名を取得
Public Function GetTypeDisplayName(unitType As String) As String
    Select Case UCase(unitType)
        Case "GROUP"
            GetTypeDisplayName = TYPE_GROUP
        Case "ROOTNET"
            GetTypeDisplayName = TYPE_ROOTNET
        Case "NET"
            GetTypeDisplayName = TYPE_NET
        Case Else
            If InStr(UCase(unitType), "JOB") > 0 Then
                GetTypeDisplayName = TYPE_JOB
            Else
                GetTypeDisplayName = TYPE_UNKNOWN
            End If
    End Select
End Function

' PowerShell文字列エスケープ
Public Function EscapePSString(s As String) As String
    EscapePSString = Replace(Replace(s, "'", "''"), "`", "``")
End Function

' 値抽出
Public Function ExtractValue(text As String, prefix As String) As String
    Dim startPos As Long
    startPos = InStr(text, prefix)

    If startPos = 0 Then
        ExtractValue = ""
        Exit Function
    End If

    startPos = startPos + Len(prefix)

    Dim endPos As Long
    endPos = InStr(startPos, text, vbCrLf)

    If endPos = 0 Then
        endPos = InStr(startPos, text, vbLf)
    End If

    If endPos = 0 Then
        endPos = Len(text) + 1
    End If

    ExtractValue = Trim(Mid(text, startPos, endPos - startPos))
End Function

' 終了状態かどうかをチェック
Public Function IsTerminalStatus(status As String) As Boolean
    Select Case UCase(status)
        Case "NORMAL", "WARNING", "ABNORMAL", "KILLED", "BYPASS", "NOTRUN", "END", _
             "NEST_END_NORMAL", "NEST_END_WARNING", "NEST_END_ABNORMAL"
            IsTerminalStatus = True
        Case Else
            IsTerminalStatus = False
    End Select
End Function

