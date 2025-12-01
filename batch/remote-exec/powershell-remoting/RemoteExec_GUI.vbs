' ====================================================================
' PowerShellスクリプト（Invoke-RemoteBatch.ps1）をGUIで実行
' ダブルクリックで実行可能
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

' 確認ダイアログ
result = MsgBox("以下の設定でリモートバッチを実行します：" & vbCrLf & vbCrLf & _
                "リモートサーバ: " & computerName & vbCrLf & _
                "実行ユーザー: " & userName & vbCrLf & _
                "実行ファイル: " & batchPath & vbCrLf & _
                "出力ログ: " & outputLog & vbCrLf & _
                "HTTPS使用: " & IIf(useSSL, "はい", "いいえ") & vbCrLf & vbCrLf & _
                "実行しますか？", vbQuestion + vbYesNo, "リモートバッチ実行")

If result = vbNo Then
    WScript.Quit
End If

' コマンドを構築
command = "powershell.exe -ExecutionPolicy Bypass -WindowStyle Normal -File """ & scriptPath & """ " & _
          "-ComputerName """ & computerName & """ " & _
          "-UserName """ & userName & """ " & _
          "-BatchPath """ & batchPath & """ " & _
          "-OutputLog """ & outputLog & """"

If useSSL Then
    command = command & " -UseSSL"
End If

' 実行（ウィンドウ表示: 1, 待機: True）
result = objShell.Run(command, 1, True)

' 結果を表示
If result = 0 Then
    MsgBox "リモートバッチ実行が正常に完了しました。" & vbCrLf & vbCrLf & _
           "終了コード: " & result & vbCrLf & _
           "ログファイル: " & outputLog, vbInformation, "実行完了"
Else
    MsgBox "リモートバッチ実行に失敗しました。" & vbCrLf & vbCrLf & _
           "終了コード: " & result & vbCrLf & _
           "ログファイルを確認してください: " & outputLog, vbCritical, "実行失敗"
End If

' クリーンアップ
Set fso = Nothing
Set objShell = Nothing

' IIf関数（VBScript用）
Function IIf(condition, trueValue, falseValue)
    If condition Then
        IIf = trueValue
    Else
        IIf = falseValue
    End If
End Function
