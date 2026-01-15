Attribute VB_Name = "JM_Config"
Option Explicit

'==============================================================================
' JP1 ジョブ管理ツール - 設定モジュール
' 定数、設定取得機能、共通ユーティリティを提供
' ※このモジュールはSetupモジュールを削除しても動作するよう設計
'==============================================================================

' デバッグモード（Trueにすると[DEBUG-XX]ログが出力されます）
Public Const DEBUG_MODE As Boolean = False

' ============================================================================
' シート名定数
' ============================================================================
Public Const SHEET_SETTINGS As String = "設定"
Public Const SHEET_JOBLIST As String = "ジョブ一覧"
Public Const SHEET_LOG As String = "実行ログ"

' ============================================================================
' 設定セル位置（設定シート）
' ※ボタンが上部（3-4行目）にあるため、設定は6行目以降に配置
' ============================================================================
Public Const ROW_EXEC_MODE As Long = 7
Public Const ROW_JP1_SERVER As Long = 9
Public Const ROW_REMOTE_USER As Long = 10
Public Const ROW_REMOTE_PASSWORD As Long = 11
Public Const ROW_JP1_USER As Long = 12
Public Const ROW_JP1_PASSWORD As Long = 13
Public Const ROW_SCHEDULER_SERVICE As Long = 14
Public Const ROW_ROOT_PATH As Long = 15
Public Const ROW_WAIT_COMPLETION As Long = 17
Public Const ROW_TIMEOUT As Long = 18
Public Const ROW_POLLING_INTERVAL As Long = 19
Public Const COL_SETTING_VALUE As Long = 3

' ============================================================================
' ジョブ一覧シートの列位置
' ============================================================================
Public Const COL_SELECT As Long = 1         ' 選択（チェック列）
Public Const COL_ORDER As Long = 2          ' 順序（自動採番）
Public Const COL_UNIT_TYPE As Long = 3      ' 種別（グループ/ジョブネット/ジョブ）
Public Const COL_JOBNET_PATH As Long = 4
Public Const COL_JOBNET_NAME As Long = 5
Public Const COL_COMMENT As Long = 6
Public Const COL_SCRIPT As Long = 7         ' スクリプトファイル名 (sc) ※非表示列
Public Const COL_PARAMETER As Long = 8      ' パラメーター (prm) ※非表示列
Public Const COL_WORK_PATH As Long = 9      ' ワークパス (wkp) ※非表示列
Public Const COL_HOLD As Long = 10
Public Const COL_LAST_STATUS As Long = 11
Public Const COL_LAST_EXEC_TIME As Long = 12
Public Const COL_LAST_END_TIME As Long = 13
Public Const COL_LAST_MESSAGE As Long = 14   ' ログファイルパス（最終列）
Public Const ROW_JOBLIST_HEADER As Long = 4
Public Const ROW_JOBLIST_DATA_START As Long = 5

' ============================================================================
' モジュールレベル変数
' ============================================================================
Public g_AdminChecked As Boolean
Public g_IsAdmin As Boolean
Public g_LogFilePath As String

' ============================================================================
' 設定取得
' ============================================================================
Public Function GetConfig() As Object
    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_SETTINGS)

    Dim config As Object
    Set config = CreateObject("Scripting.Dictionary")

    ' 実行モード（ローカル/リモート）
    config("ExecMode") = CStr(ws.Cells(ROW_EXEC_MODE, COL_SETTING_VALUE).Value)

    config("JP1Server") = CStr(ws.Cells(ROW_JP1_SERVER, COL_SETTING_VALUE).Value)
    config("RemoteUser") = CStr(ws.Cells(ROW_REMOTE_USER, COL_SETTING_VALUE).Value)
    config("RemotePassword") = CStr(ws.Cells(ROW_REMOTE_PASSWORD, COL_SETTING_VALUE).Value)
    config("JP1User") = CStr(ws.Cells(ROW_JP1_USER, COL_SETTING_VALUE).Value)
    config("JP1Password") = CStr(ws.Cells(ROW_JP1_PASSWORD, COL_SETTING_VALUE).Value)
    config("SchedulerService") = CStr(ws.Cells(ROW_SCHEDULER_SERVICE, COL_SETTING_VALUE).Value)
    config("RootPath") = CStr(ws.Cells(ROW_ROOT_PATH, COL_SETTING_VALUE).Value)
    config("WaitCompletion") = CStr(ws.Cells(ROW_WAIT_COMPLETION, COL_SETTING_VALUE).Value)
    config("Timeout") = CLng(ws.Cells(ROW_TIMEOUT, COL_SETTING_VALUE).Value)
    config("PollingInterval") = CLng(ws.Cells(ROW_POLLING_INTERVAL, COL_SETTING_VALUE).Value)

    ' 必須項目チェック（ローカルモードとリモートモードで異なる）
    If config("ExecMode") = "ローカル" Then
        ' ローカルモード: JP1ユーザーのみ必須
        If config("JP1User") = "" Then
            MsgBox "JP1ユーザーを入力してください。", vbExclamation
            Set GetConfig = Nothing
            Exit Function
        End If
    Else
        ' リモートモード: 接続情報が必須
        If config("JP1Server") = "" Or config("RemoteUser") = "" Or config("JP1User") = "" Then
            MsgBox "接続設定が不完全です。設定シートで設定を入力してください。", vbExclamation
            Set GetConfig = Nothing
            Exit Function
        End If
    End If

    Set GetConfig = config
End Function

' ============================================================================
' PowerShell文字列エスケープ
' ============================================================================
Public Function EscapePSString(str As String) As String
    ' PowerShell文字列内のシングルクォートをエスケープ
    EscapePSString = Replace(str, "'", "''")
End Function

' ============================================================================
' 管理者権限チェック
' ============================================================================
Public Function IsRunningAsAdmin() As Boolean
    ' キャッシュを利用
    If g_AdminChecked Then
        IsRunningAsAdmin = g_IsAdmin
        Exit Function
    End If

    ' PowerShellで管理者権限をチェック
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")

    Dim cmd As String
    cmd = "powershell -NoProfile -Command ""$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent()); if ($principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { exit 0 } else { exit 1 }"""

    Dim exitCode As Long
    exitCode = shell.Run(cmd, 0, True)

    g_IsAdmin = (exitCode = 0)
    g_AdminChecked = True

    IsRunningAsAdmin = g_IsAdmin
End Function

Public Function EnsureAdminForRemoteMode(config As Object) As Boolean
    ' ローカルモードなら管理者不要
    If config("ExecMode") = "ローカル" Then
        EnsureAdminForRemoteMode = True
        Exit Function
    End If

    ' 既に管理者なら問題なし
    If IsRunningAsAdmin() Then
        EnsureAdminForRemoteMode = True
        Exit Function
    End If

    ' 管理者でない場合、ユーザーに選択させる
    Dim response As VbMsgBoxResult
    response = MsgBox( _
        "リモート実行モードでは、WinRM設定の変更に管理者権限が必要です。" & vbCrLf & vbCrLf & _
        "現在、管理者権限で実行されていません。" & vbCrLf & vbCrLf & _
        "[はい] 管理者としてExcelを再起動して実行" & vbCrLf & _
        "[いいえ] このまま続行（WinRMが既に設定済みの場合）" & vbCrLf & _
        "[キャンセル] 処理を中止", _
        vbYesNoCancel + vbExclamation, "管理者権限が必要")

    Select Case response
        Case vbYes
            ' 管理者権限でExcelを再起動
            RestartAsAdmin
            EnsureAdminForRemoteMode = False

        Case vbNo
            ' そのまま続行
            EnsureAdminForRemoteMode = True

        Case vbCancel
            ' 処理を中止
            EnsureAdminForRemoteMode = False
    End Select
End Function

Public Sub RestartAsAdmin()
    ' 現在のブックを保存
    If ThisWorkbook.Saved = False Then
        Dim saveResponse As VbMsgBoxResult
        saveResponse = MsgBox("ブックを保存しますか？", vbYesNoCancel + vbQuestion, "保存確認")
        If saveResponse = vbYes Then
            ThisWorkbook.Save
        ElseIf saveResponse = vbCancel Then
            Exit Sub
        End If
    End If

    ' 管理者権限でExcelを再起動
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")

    Dim excelPath As String
    excelPath = Application.Path & "\EXCEL.EXE"

    Dim workbookPath As String
    workbookPath = ThisWorkbook.FullName

    ' PowerShellでStart-Process -Verb RunAsを実行
    Dim cmd As String
    cmd = "powershell -NoProfile -Command ""Start-Process -FilePath '" & Replace(excelPath, "'", "''") & "' -ArgumentList '""" & Replace(workbookPath, "'", "''") & """' -Verb RunAs"""

    shell.Run cmd, 0, False

    ' このブックのみを閉じる（他のExcelブックは維持）
    ThisWorkbook.Close SaveChanges:=False
End Sub

' ============================================================================
' ログファイル関連
' ============================================================================
Public Function CreateLogFile() As String
    ' ログファイルのパスを生成して初期ヘッダーを書き込む
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' ログフォルダ（Excelブックと同じフォルダにLogsサブフォルダ）
    Dim logFolder As String
    logFolder = ThisWorkbook.Path & "\Logs"

    ' Logsフォルダが存在しない場合は作成
    If Not fso.FolderExists(logFolder) Then
        fso.CreateFolder logFolder
    End If

    ' ログファイル名（JP1_実行ログ_yyyyMMdd_HHmmss.txt）
    Dim logFileName As String
    logFileName = "JP1_実行ログ_" & Format(Now, "yyyyMMdd_HHmmss") & ".txt"

    Dim logFilePath As String
    logFilePath = logFolder & "\" & logFileName

    ' ADODB.Streamを使用してUTF-8（BOMなし）でヘッダーを書き込む
    Dim logContent As String
    logContent = "================================================================================" & vbCrLf
    logContent = logContent & "JP1 ジョブ管理ツール - 実行ログ" & vbCrLf
    logContent = logContent & "================================================================================" & vbCrLf
    logContent = logContent & "開始日時: " & Format(Now, "yyyy/mm/dd HH:mm:ss") & vbCrLf
    logContent = logContent & "実行モード: " & Worksheets(SHEET_SETTINGS).Cells(ROW_EXEC_MODE, COL_SETTING_VALUE).Value & vbCrLf
    logContent = logContent & "================================================================================" & vbCrLf
    logContent = logContent & "" & vbCrLf

    Dim utfStream As Object
    Set utfStream = CreateObject("ADODB.Stream")
    utfStream.Type = 2 ' adTypeText
    utfStream.Charset = "UTF-8"
    utfStream.Open
    utfStream.WriteText logContent

    ' BOMをスキップしてバイナリで保存
    utfStream.Position = 0
    utfStream.Type = 1 ' adTypeBinary
    utfStream.Position = 3 ' BOM（3バイト）をスキップ

    Dim binStream As Object
    Set binStream = CreateObject("ADODB.Stream")
    binStream.Type = 1 ' adTypeBinary
    binStream.Open
    utfStream.CopyTo binStream
    binStream.SaveToFile logFilePath, 2 ' adSaveCreateOverWrite

    binStream.Close
    utfStream.Close
    Set binStream = Nothing
    Set utfStream = Nothing

    CreateLogFile = logFilePath
End Function

Public Function GetLogFilePath() As String
    ' 現在のログファイルパスを返す
    GetLogFilePath = g_LogFilePath
End Function

