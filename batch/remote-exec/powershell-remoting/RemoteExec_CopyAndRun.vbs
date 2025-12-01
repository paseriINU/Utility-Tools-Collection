' ====================================================================
' サーバ上のPowerShellスクリプトを一時コピーして実行
' より安全な方法（ネットワーク切断に強い）
' ====================================================================
Option Explicit

Dim objShell, fso
Dim networkScriptPath, localTempPath
Dim computerName, userName, batchPath, outputLog, useSSL
Dim command, result

Set objShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' ====================================================================
' 設定項目（ここを編集してください）
' ====================================================================

' サーバ上のPowerShellスクリプト（ネットワークパス）
networkScriptPath = "\\192.168.1.100\Share\Scripts\Invoke-RemoteBatch.ps1"

' リモート実行の設定
computerName = "192.168.1.100"           ' リモートサーバのIPアドレス
userName = "Administrator"               ' ユーザー名
batchPath = "C:\Scripts\target_script.bat"  ' 実行するバッチファイル
useSSL = False                           ' HTTPS使用（True/False）

' 出力ログファイル（ローカル）
outputLog = "C:\Logs\remote_output.log"

' ====================================================================
' メイン処理
' ====================================================================

' 一時フォルダのパス
localTempPath = objShell.ExpandEnvironmentStrings("%TEMP%") & "\Invoke-RemoteBatch.ps1"

' ネットワークパスの存在確認
If Not fso.FileExists(networkScriptPath) Then
    MsgBox "PowerShellスクリプトが見つかりません：" & vbCrLf & vbCrLf & _
           networkScriptPath & vbCrLf & vbCrLf & _
           "ネットワークパスにアクセスできるか確認してください。", vbCritical, "エラー"
    WScript.Quit
End If

' 確認ダイアログ
result = MsgBox("以下の設定でリモートバッチを実行します：" & vbCrLf & vbCrLf & _
                "スクリプト: " & networkScriptPath & vbCrLf & _
                "リモートサーバ: " & computerName & vbCrLf & _
                "実行ユーザー: " & userName & vbCrLf & _
                "実行ファイル: " & batchPath & vbCrLf & vbCrLf & _
                "実行しますか？", vbQuestion + vbYesNo, "リモートバッチ実行")

If result = vbNo Then
    WScript.Quit
End If

' スクリプトをローカル一時フォルダにコピー
On Error Resume Next
fso.CopyFile networkScriptPath, localTempPath, True
If Err.Number <> 0 Then
    MsgBox "スクリプトのコピーに失敗しました：" & vbCrLf & vbCrLf & _
           "エラー: " & Err.Description, vbCritical, "エラー"
    WScript.Quit
End If
On Error Goto 0

' コマンドを構築（ローカルの一時ファイルを使用）
command = "powershell.exe -ExecutionPolicy Bypass -File """ & localTempPath & """ " & _
          "-ComputerName """ & computerName & """ " & _
          "-UserName """ & userName & """ " & _
          "-BatchPath """ & batchPath & """ " & _
          "-OutputLog """ & outputLog & """"

If useSSL Then
    command = command & " -UseSSL"
End If

' 実行
result = objShell.Run(command, 1, True)

' 一時ファイルを削除
On Error Resume Next
fso.DeleteFile localTempPath, True
On Error Goto 0

' 結果を表示
If result = 0 Then
    MsgBox "リモートバッチ実行が正常に完了しました。" & vbCrLf & vbCrLf & _
           "ログファイル: " & outputLog, vbInformation, "実行完了"
Else
    MsgBox "リモートバッチ実行に失敗しました。" & vbCrLf & vbCrLf & _
           "終了コード: " & result, vbCritical, "実行失敗"
End If

' クリーンアップ
Set fso = Nothing
Set objShell = Nothing
