' ====================================================================
' ネットワークパス上のPowerShellスクリプトを実行
' VBSをローカルに、.ps1をサーバ上に配置する場合に使用
' ====================================================================
Option Explicit

Dim objShell, fso
Dim networkScriptPath, computerName, userName, batchPath, outputLog, useSSL
Dim localTempPath, command, result

Set objShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' ====================================================================
' 設定項目（ここを編集してください）
' ====================================================================

' ネットワークパス上のPowerShellスクリプト
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

' コマンドを構築（ネットワークパスを直接指定）
command = "powershell.exe -ExecutionPolicy Bypass -File """ & networkScriptPath & """ " & _
          "-ComputerName """ & computerName & """ " & _
          "-UserName """ & userName & """ " & _
          "-BatchPath """ & batchPath & """ " & _
          "-OutputLog """ & outputLog & """"

If useSSL Then
    command = command & " -UseSSL"
End If

' 実行
result = objShell.Run(command, 1, True)

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
