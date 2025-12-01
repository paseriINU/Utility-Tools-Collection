' ====================================================================
' PowerShellスクリプト（Invoke-RemoteBatch.ps1）をバックグラウンドで実行
' ウィンドウを表示せず、完了時のみ通知
' タスクスケジューラや自動実行に最適
' ====================================================================
Option Explicit

Dim objShell, fso, scriptPath
Dim computerName, userName, batchPath, outputLog, useSSL
Dim command, result

Set objShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' ====================================================================
' 設定項目（ここを編集してください）
' ====================================================================
computerName = "192.168.1.100"           ' リモートサーバのIPアドレス
userName = "Administrator"               ' ユーザー名
batchPath = "C:\Scripts\target_script.bat"  ' 実行するバッチファイル
useSSL = False                           ' HTTPS使用（True/False）

' 出力ログファイル（このVBScriptと同じフォルダ）
outputLog = fso.GetParentFolderName(WScript.ScriptFullName) & "\remote_output.log"

' ====================================================================
' メイン処理
' ====================================================================

' PowerShellスクリプトのパス
scriptPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\Invoke-RemoteBatch.ps1"

' スクリプトファイルの存在確認
If Not fso.FileExists(scriptPath) Then
    MsgBox "PowerShellスクリプトが見つかりません：" & vbCrLf & scriptPath, vbCritical, "エラー"
    WScript.Quit
End If

' コマンドを構築
command = "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & scriptPath & """ " & _
          "-ComputerName """ & computerName & """ " & _
          "-UserName """ & userName & """ " & _
          "-BatchPath """ & batchPath & """ " & _
          "-OutputLog """ & outputLog & """"

If useSSL Then
    command = command & " -UseSSL"
End If

' 実行（ウィンドウ非表示: 0, 待機: True）
result = objShell.Run(command, 0, True)

' 結果を通知（完了時のみ）
If result = 0 Then
    MsgBox "リモートバッチ実行が完了しました。" & vbCrLf & vbCrLf & _
           "リモートサーバ: " & computerName & vbCrLf & _
           "ログファイル: " & outputLog, vbInformation, "実行完了"
Else
    MsgBox "リモートバッチ実行に失敗しました。" & vbCrLf & vbCrLf & _
           "終了コード: " & result & vbCrLf & _
           "ログファイル: " & outputLog, vbCritical, "実行失敗"
End If

' クリーンアップ
Set fso = Nothing
Set objShell = Nothing
