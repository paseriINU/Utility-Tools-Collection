Attribute VB_Name = "JRM_Api"
Option Explicit

'==============================================================================
' JP1 REST ジョブ管理ツール - API モジュール
' REST API呼び出し、PowerShell実行、レスポンスパースを提供
'==============================================================================

'==============================================================================
' REST API: ユニット一覧取得
'==============================================================================
Public Function GetUnitList(config As Object, location As String) As Collection
    On Error GoTo ErrorHandler

    Dim psScript As String
    psScript = BuildStatusesAPIScript(config, location, "NO")

    Dim result As String
    result = ExecutePowerShell(psScript)

    ' APIエラーチェック
    If InStr(result, "API_ERROR:") > 0 Then
        Dim errorCode As String
        Dim errorMsg As String
        errorCode = ExtractValue(result, "API_ERROR:")
        errorMsg = ExtractValue(result, "ERROR_MESSAGE:")

        Dim userMsg As String
        userMsg = "REST API接続エラー" & vbCrLf & vbCrLf
        userMsg = userMsg & "エラーコード: " & errorCode & vbCrLf
        userMsg = userMsg & "エラー内容: " & errorMsg & vbCrLf & vbCrLf

        Select Case errorCode
            Case "401"
                userMsg = userMsg & "原因: JP1ユーザー名またはパスワードが間違っています"
            Case "403"
                userMsg = userMsg & "原因: JP1ユーザーに参照権限がありません"
            Case "404"
                userMsg = userMsg & "原因: 指定したパスまたはManagerホスト名が存在しません"
            Case "412"
                userMsg = userMsg & "原因: Web Consoleサーバに接続できません"
            Case "500"
                userMsg = userMsg & "原因: サーバ側でエラーが発生しました"
            Case Else
                userMsg = userMsg & "接続設定を確認してください"
        End Select

        MsgBox userMsg, vbCritical, "接続エラー"
        Set GetUnitList = Nothing
        Exit Function
    End If

    ' PowerShell実行エラーチェック
    If InStr(result, "PS_ERROR:") > 0 Or Len(result) = 0 Then
        Dim psError As String
        psError = "PowerShell実行エラー" & vbCrLf & vbCrLf

        If Len(result) = 0 Then
            psError = psError & "応答がありませんでした。" & vbCrLf
            psError = psError & "Web Consoleサーバへの接続を確認してください。"
        Else
            psError = psError & Left(result, 500)
        End If

        MsgBox psError, vbCritical, "実行エラー"
        Set GetUnitList = Nothing
        Exit Function
    End If

    ' デバッグモード: API応答をMsgBoxで表示
    If g_DebugMode Then
        Dim debugMsg As String
        debugMsg = "=== API応答デバッグ ===" & vbCrLf & vbCrLf
        debugMsg = debugMsg & "応答長: " & Len(result) & " 文字" & vbCrLf & vbCrLf
        debugMsg = debugMsg & "応答内容（先頭1000文字）:" & vbCrLf
        debugMsg = debugMsg & Left(result, 1000)
        MsgBox debugMsg, vbInformation, "デバッグモード"
    End If

    ' 結果をパース
    Set GetUnitList = ParseStatusesResponse(result)
    Exit Function

ErrorHandler:
    MsgBox "GetUnitListエラー: " & Err.Description, vbCritical, "VBAエラー"
    Set GetUnitList = Nothing
End Function

'==============================================================================
' REST API: 即時実行登録
'==============================================================================
Public Function ExecuteImmediateExec(config As Object, unitPath As String) As Object
    On Error GoTo ErrorHandler

    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result("Success") = False

    Dim psScript As String
    psScript = BuildImmediateExecAPIScript(config, unitPath)

    Dim output As String
    output = ExecutePowerShell(psScript)

    ' 結果をパース
    If InStr(output, "EXEC_ID:") > 0 Then
        Dim execID As String
        execID = ExtractValue(output, "EXEC_ID:")
        result("Success") = True
        result("ExecID") = Trim(execID)
    Else
        result("Success") = False
        result("ErrorMessage") = "execIDの取得に失敗しました"

        If InStr(output, "API_ERROR:") > 0 Then
            result("ErrorMessage") = ExtractValue(output, "ERROR_MESSAGE:")
        End If
    End If

    Set ExecuteImmediateExec = result
    Exit Function

ErrorHandler:
    Set result = CreateObject("Scripting.Dictionary")
    result("Success") = False
    result("ErrorMessage") = Err.Description
    Set ExecuteImmediateExec = result
End Function

'==============================================================================
' REST API: 実行状態ポーリング
'==============================================================================
Public Function PollExecutionStatus(config As Object, unitPath As String, execID As String) As Object
    On Error GoTo ErrorHandler

    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result("Success") = False

    Dim pollingInterval As Long
    pollingInterval = CLng(config("PollingInterval"))
    If pollingInterval < 1 Then pollingInterval = 5

    Dim timeout As Long
    timeout = CLng(config("Timeout"))

    Dim startTime As Date
    startTime = Now

    Do
        ' API呼び出し
        Dim psScript As String
        psScript = BuildStatusesAPIScriptWithExecID(config, unitPath, execID)

        Dim output As String
        output = ExecutePowerShell(psScript)

        ' 状態を確認
        Dim status As String
        status = ExtractValue(output, "STATUS:")

        result("Status") = status
        result("StartTime") = ExtractValue(output, "START_TIME:")
        result("EndTime") = ExtractValue(output, "END_TIME:")

        ' 終了状態かチェック
        If IsTerminalStatus(status) Then
            result("Success") = True
            Exit Do
        End If

        ' タイムアウトチェック
        If timeout > 0 Then
            If DateDiff("s", startTime, Now) > timeout Then
                result("Success") = False
                result("ErrorMessage") = "タイムアウト"
                Exit Do
            End If
        End If

        ' 待機
        Application.Wait Now + TimeSerial(0, 0, pollingInterval)
        DoEvents
    Loop

    Set PollExecutionStatus = result
    Exit Function

ErrorHandler:
    Set result = CreateObject("Scripting.Dictionary")
    result("Success") = False
    result("ErrorMessage") = Err.Description
    Set PollExecutionStatus = result
End Function

'==============================================================================
' REST API: 実行結果詳細取得
'==============================================================================
Public Function GetExecResultDetails(config As Object, unitPath As String, execID As String) As Object
    On Error GoTo ErrorHandler

    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result("Success") = False

    Dim psScript As String
    psScript = BuildExecResultDetailsAPIScript(config, unitPath, execID)

    Dim output As String
    output = ExecutePowerShell(psScript)

    If InStr(output, "RESULT_DETAILS_START") > 0 Then
        Dim startPos As Long
        Dim endPos As Long
        startPos = InStr(output, "RESULT_DETAILS_START") + Len("RESULT_DETAILS_START") + 2
        endPos = InStr(output, "RESULT_DETAILS_END")

        If endPos > startPos Then
            result("Details") = Mid(output, startPos, endPos - startPos)
            result("Success") = True
        End If
    Else
        result("Success") = False
        result("ErrorMessage") = "ログの取得に失敗しました"
    End If

    Set GetExecResultDetails = result
    Exit Function

ErrorHandler:
    Set result = CreateObject("Scripting.Dictionary")
    result("Success") = False
    result("ErrorMessage") = Err.Description
    Set GetExecResultDetails = result
End Function

'==============================================================================
' PowerShellスクリプト生成: statuses API
'==============================================================================
Public Function BuildStatusesAPIScript(config As Object, location As String, searchLowerUnits As String) As String
    Dim script As String

    script = BuildAPIHeader(config)

    ' URLエンコード
    script = script & "$location = '" & EscapePSString(location) & "'" & vbCrLf
    script = script & "$encodedLocation = [System.Uri]::EscapeDataString($location)" & vbCrLf
    script = script & vbCrLf

    ' URL構築
    script = script & "$url = ""${baseUrl}/objects/statuses?mode=search""" & vbCrLf
    script = script & "$url += ""&manager=${managerHost}""" & vbCrLf
    script = script & "$url += ""&serviceName=${schedulerService}""" & vbCrLf
    script = script & "$url += ""&location=${encodedLocation}""" & vbCrLf
    script = script & "$url += ""&searchLowerUnits=" & searchLowerUnits & """" & vbCrLf
    script = script & "$url += ""&searchTarget=DEFINITION""" & vbCrLf
    script = script & vbCrLf

    ' API呼び出し
    script = script & "try {" & vbCrLf
    script = script & "  $response = Invoke-WebRequest -Uri $url -Method GET -Headers $headers -TimeoutSec 30 -UseBasicParsing" & vbCrLf
    script = script & "  $responseBytes = $response.RawContentStream.ToArray()" & vbCrLf
    script = script & "  $responseText = [System.Text.Encoding]::UTF8.GetString($responseBytes)" & vbCrLf
    script = script & "  $json = $responseText | ConvertFrom-Json" & vbCrLf
    script = script & vbCrLf
    script = script & "  Write-Output 'JSON_START'" & vbCrLf
    script = script & "  Write-Output ($json | ConvertTo-Json -Depth 10 -Compress)" & vbCrLf
    script = script & "  Write-Output 'JSON_END'" & vbCrLf
    script = script & "} catch {" & vbCrLf
    script = script & "  Write-Output ""API_ERROR:$($_.Exception.Response.StatusCode.Value__)""" & vbCrLf
    script = script & "  Write-Output ""ERROR_MESSAGE:$($_.Exception.Message)""" & vbCrLf
    script = script & "}" & vbCrLf

    BuildStatusesAPIScript = script
End Function

'==============================================================================
' PowerShellスクリプト生成: statuses API (execID指定)
'==============================================================================
Public Function BuildStatusesAPIScriptWithExecID(config As Object, unitPath As String, execID As String) As String
    Dim script As String

    script = BuildAPIHeader(config)

    ' パスを分解
    script = script & "$unitPath = '" & EscapePSString(unitPath) & "'" & vbCrLf
    script = script & "$lastSlash = $unitPath.LastIndexOf('/')" & vbCrLf
    script = script & "$parentPath = $unitPath.Substring(0, $lastSlash)" & vbCrLf
    script = script & "if (-not $parentPath) { $parentPath = '/' }" & vbCrLf
    script = script & "$unitName = $unitPath.Substring($lastSlash + 1)" & vbCrLf
    script = script & "$encodedParent = [System.Uri]::EscapeDataString($parentPath)" & vbCrLf
    script = script & "$encodedName = [System.Uri]::EscapeDataString($unitName)" & vbCrLf
    script = script & vbCrLf

    ' URL構築
    script = script & "$url = ""${baseUrl}/objects/statuses?mode=search""" & vbCrLf
    script = script & "$url += ""&manager=${managerHost}""" & vbCrLf
    script = script & "$url += ""&serviceName=${schedulerService}""" & vbCrLf
    script = script & "$url += ""&location=${encodedParent}""" & vbCrLf
    script = script & "$url += ""&searchLowerUnits=NO""" & vbCrLf
    script = script & "$url += ""&searchTarget=DEFINITION_AND_STATUS""" & vbCrLf
    script = script & "$url += ""&unitName=${encodedName}""" & vbCrLf
    script = script & "$url += ""&unitNameMatchMethods=EQ""" & vbCrLf
    script = script & "$url += ""&generation=EXECID""" & vbCrLf
    script = script & "$url += ""&execID=" & execID & """" & vbCrLf
    script = script & vbCrLf

    ' API呼び出し
    script = script & "try {" & vbCrLf
    script = script & "  $response = Invoke-WebRequest -Uri $url -Method GET -Headers $headers -TimeoutSec 30 -UseBasicParsing" & vbCrLf
    script = script & "  $responseBytes = $response.RawContentStream.ToArray()" & vbCrLf
    script = script & "  $responseText = [System.Text.Encoding]::UTF8.GetString($responseBytes)" & vbCrLf
    script = script & "  $json = $responseText | ConvertFrom-Json" & vbCrLf
    script = script & vbCrLf
    script = script & "  if ($json.statuses -and $json.statuses.Count -gt 0) {" & vbCrLf
    script = script & "    $unit = $json.statuses[0]" & vbCrLf
    script = script & "    $status = if ($unit.unitStatus) { $unit.unitStatus.status } else { 'N/A' }" & vbCrLf
    script = script & "    $startTime = if ($unit.unitStatus) { $unit.unitStatus.startTime } else { '' }" & vbCrLf
    script = script & "    $endTime = if ($unit.unitStatus) { $unit.unitStatus.endTime } else { '' }" & vbCrLf
    script = script & "    Write-Output ""STATUS:$status""" & vbCrLf
    script = script & "    Write-Output ""START_TIME:$startTime""" & vbCrLf
    script = script & "    Write-Output ""END_TIME:$endTime""" & vbCrLf
    script = script & "  } else {" & vbCrLf
    script = script & "    Write-Output 'STATUS:NOT_FOUND'" & vbCrLf
    script = script & "  }" & vbCrLf
    script = script & "} catch {" & vbCrLf
    script = script & "  Write-Output ""API_ERROR:$($_.Exception.Response.StatusCode.Value__)""" & vbCrLf
    script = script & "  Write-Output ""ERROR_MESSAGE:$($_.Exception.Message)""" & vbCrLf
    script = script & "}" & vbCrLf

    BuildStatusesAPIScriptWithExecID = script
End Function

'==============================================================================
' PowerShellスクリプト生成: 即時実行登録 API
'==============================================================================
Public Function BuildImmediateExecAPIScript(config As Object, unitPath As String) As String
    Dim script As String

    script = BuildAPIHeader(config)

    ' URLエンコード
    script = script & "$unitPath = '" & EscapePSString(unitPath) & "'" & vbCrLf
    script = script & "$encodedPath = [System.Uri]::EscapeDataString($unitPath)" & vbCrLf
    script = script & vbCrLf

    ' URL構築
    script = script & "$url = ""${baseUrl}/objects/definitions/${encodedPath}/actions/registerImmediateExec/invoke""" & vbCrLf
    script = script & "$url += ""?manager=${managerHost}&serviceName=${schedulerService}""" & vbCrLf
    script = script & vbCrLf

    ' API呼び出し（POST）
    script = script & "try {" & vbCrLf
    script = script & "  $response = Invoke-WebRequest -Uri $url -Method POST -Headers $headers -TimeoutSec 30 -UseBasicParsing" & vbCrLf
    script = script & "  $responseBytes = $response.RawContentStream.ToArray()" & vbCrLf
    script = script & "  $responseText = [System.Text.Encoding]::UTF8.GetString($responseBytes)" & vbCrLf
    script = script & "  $json = $responseText | ConvertFrom-Json" & vbCrLf
    script = script & "  Write-Output ""EXEC_ID:$($json.execID)""" & vbCrLf
    script = script & "} catch {" & vbCrLf
    script = script & "  Write-Output ""API_ERROR:$($_.Exception.Response.StatusCode.Value__)""" & vbCrLf
    script = script & "  Write-Output ""ERROR_MESSAGE:$($_.Exception.Message)""" & vbCrLf
    script = script & "}" & vbCrLf

    BuildImmediateExecAPIScript = script
End Function

'==============================================================================
' PowerShellスクリプト生成: 実行結果詳細取得 API
'==============================================================================
Public Function BuildExecResultDetailsAPIScript(config As Object, unitPath As String, execID As String) As String
    Dim script As String

    script = BuildAPIHeader(config)

    ' URLエンコード
    script = script & "$unitPath = '" & EscapePSString(unitPath) & "'" & vbCrLf
    script = script & "$execID = '" & execID & "'" & vbCrLf
    script = script & "$encodedPath = [System.Uri]::EscapeDataString($unitPath)" & vbCrLf
    script = script & vbCrLf

    ' URL構築
    script = script & "$url = ""${baseUrl}/objects/statuses/${encodedPath}:${execID}/actions/execResultDetails/invoke""" & vbCrLf
    script = script & "$url += ""?manager=${managerHost}&serviceName=${schedulerService}""" & vbCrLf
    script = script & vbCrLf

    ' API呼び出し
    script = script & "try {" & vbCrLf
    script = script & "  $response = Invoke-WebRequest -Uri $url -Method GET -Headers $headers -TimeoutSec 60 -UseBasicParsing" & vbCrLf
    script = script & "  $responseBytes = $response.RawContentStream.ToArray()" & vbCrLf
    script = script & "  $responseText = [System.Text.Encoding]::UTF8.GetString($responseBytes)" & vbCrLf
    script = script & "  $json = $responseText | ConvertFrom-Json" & vbCrLf
    script = script & "  Write-Output 'RESULT_DETAILS_START'" & vbCrLf
    script = script & "  Write-Output $json.execResultDetails" & vbCrLf
    script = script & "  Write-Output 'RESULT_DETAILS_END'" & vbCrLf
    script = script & "  Write-Output ""ALL:$($json.all)""" & vbCrLf
    script = script & "} catch {" & vbCrLf
    script = script & "  Write-Output ""API_ERROR:$($_.Exception.Response.StatusCode.Value__)""" & vbCrLf
    script = script & "  Write-Output ""ERROR_MESSAGE:$($_.Exception.Message)""" & vbCrLf
    script = script & "}" & vbCrLf

    BuildExecResultDetailsAPIScript = script
End Function

'==============================================================================
' PowerShellスクリプト共通ヘッダー
'==============================================================================
Public Function BuildAPIHeader(config As Object) As String
    Dim script As String

    ' UTF-8エンコーディング
    script = "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8" & vbCrLf
    script = script & "$OutputEncoding = [System.Text.Encoding]::UTF8" & vbCrLf
    script = script & vbCrLf

    ' 接続設定
    Dim protocol As String
    If config("UseHttps") = "はい" Then
        protocol = "https"
    Else
        protocol = "http"
    End If

    script = script & "$protocol = '" & protocol & "'" & vbCrLf
    script = script & "$webConsoleHost = '" & config("WebConsoleHost") & "'" & vbCrLf
    script = script & "$webConsolePort = '" & config("WebConsolePort") & "'" & vbCrLf
    script = script & "$managerHost = '" & config("ManagerHost") & "'" & vbCrLf
    script = script & "$schedulerService = '" & config("SchedulerService") & "'" & vbCrLf
    script = script & "$baseUrl = ""${protocol}://${webConsoleHost}:${webConsolePort}/ajs/api/v1""" & vbCrLf
    script = script & vbCrLf

    ' 認証ヘッダー
    script = script & "$authString = '" & config("JP1User") & ":" & config("JP1Password") & "'" & vbCrLf
    script = script & "$authBytes = [System.Text.Encoding]::UTF8.GetBytes($authString)" & vbCrLf
    script = script & "$authBase64 = [System.Convert]::ToBase64String($authBytes)" & vbCrLf
    script = script & "$headers = @{ 'Accept-Language' = 'ja'; 'X-AJS-Authorization' = $authBase64 }" & vbCrLf
    script = script & vbCrLf

    ' HTTPS設定
    If config("UseHttps") = "はい" Then
        script = script & "Add-Type @""" & vbCrLf
        script = script & "using System.Net;" & vbCrLf
        script = script & "using System.Security.Cryptography.X509Certificates;" & vbCrLf
        script = script & "public class TrustAllCertsPolicy : ICertificatePolicy {" & vbCrLf
        script = script & "    public bool CheckValidationResult(ServicePoint sp, X509Certificate cert, WebRequest req, int problem) { return true; }" & vbCrLf
        script = script & "}" & vbCrLf
        script = script & """@" & vbCrLf
        script = script & "[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy" & vbCrLf
        script = script & "[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12" & vbCrLf
        script = script & vbCrLf
    End If

    BuildAPIHeader = script
End Function

'==============================================================================
' PowerShell実行
'==============================================================================
Public Function ExecutePowerShell(script As String) As String
    On Error GoTo ErrorHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim tempFolder As String
    tempFolder = fso.GetSpecialFolder(2) ' Temp folder

    Dim timestamp As String
    timestamp = Format(Now, "yyyymmddhhnnss") & "_" & Int(Rnd * 10000)

    Dim scriptPath As String
    scriptPath = tempFolder & "\jp1rest_" & timestamp & ".ps1"

    Dim outputPath As String
    outputPath = tempFolder & "\jp1rest_output_" & timestamp & ".txt"

    ' デバッグモード: ファイルパスを表示
    If g_DebugMode Then
        MsgBox "デバッグファイル:" & vbCrLf & vbCrLf & _
               "スクリプト: " & scriptPath & vbCrLf & _
               "出力: " & outputPath, vbInformation, "デバッグモード"
    End If

    ' ADODB.Streamを使用してUTF-8（BOMなし）で保存
    Dim utfStream As Object
    Set utfStream = CreateObject("ADODB.Stream")
    utfStream.Type = 2 ' adTypeText
    utfStream.Charset = "UTF-8"
    utfStream.Open
    utfStream.WriteText script

    ' BOMを除去してUTF-8で保存
    Dim binStream As Object
    Set binStream = CreateObject("ADODB.Stream")
    binStream.Type = 1 ' adTypeBinary
    binStream.Open

    utfStream.Position = 3 ' Skip UTF-8 BOM (3 bytes)
    utfStream.CopyTo binStream
    utfStream.Close

    binStream.SaveToFile scriptPath, 2
    binStream.Close
    Set utfStream = Nothing
    Set binStream = Nothing

    ' PowerShell実行コマンド（UTF-8出力を明示的に設定）
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")

    Dim cmd As String
    cmd = "powershell -NoProfile -ExecutionPolicy Bypass -Command ""& {" & _
          "$OutputEncoding = [System.Text.Encoding]::UTF8; " & _
          "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; " & _
          "& '" & scriptPath & "'" & _
          "}"" > """ & outputPath & """ 2>&1"

    ' デバッグモード: 1 = 表示, 0 = 非表示
    Dim windowStyle As Long
    If g_DebugMode Then
        windowStyle = 1  ' 表示
    Else
        windowStyle = 0  ' 非表示
    End If

    shell.Run cmd, windowStyle, True

    ' 結果ファイルを読み込む
    Dim output As String
    output = ""

    If fso.FileExists(outputPath) Then
        ' ファイルサイズ確認
        Dim fileSize As Long
        fileSize = fso.GetFile(outputPath).Size

        If g_DebugMode Then
            Debug.Print "Output file size: " & fileSize & " bytes"
        End If

        If fileSize > 0 Then
            ' UTF-8として読み込み
            Set utfStream = CreateObject("ADODB.Stream")
            utfStream.Type = 2
            utfStream.Charset = "UTF-8"
            utfStream.Open
            utfStream.LoadFromFile outputPath

            If Not utfStream.EOS Then
                output = utfStream.ReadText
            End If

            utfStream.Close
            Set utfStream = Nothing
        End If

        ' デバッグモードでない場合は出力ファイル削除
        If Not g_DebugMode Then
            On Error Resume Next
            fso.DeleteFile outputPath
            On Error GoTo 0
        Else
            Debug.Print "API Output File: " & outputPath
        End If
    Else
        If g_DebugMode Then
            MsgBox "出力ファイルが作成されませんでした: " & outputPath, vbExclamation, "デバッグモード"
        End If
    End If

    ' デバッグモードでない場合はスクリプトファイル削除
    If Not g_DebugMode Then
        On Error Resume Next
        fso.DeleteFile scriptPath
        On Error GoTo 0
    Else
        Debug.Print "Script File: " & scriptPath
        Debug.Print "API Response Length: " & Len(output)
        If Len(output) > 0 Then
            Debug.Print "API Response (first 500 chars): " & Left(output, 500)
        End If
    End If

    ExecutePowerShell = output
    Exit Function

ErrorHandler:
    If g_DebugMode Then
        MsgBox "ExecutePowerShellエラー: " & Err.Description, vbCritical, "デバッグモード"
    End If
    ExecutePowerShell = "PS_ERROR:" & Err.Description
End Function

'==============================================================================
' statuses APIレスポンスパース
'==============================================================================
Public Function ParseStatusesResponse(response As String) As Collection
    On Error GoTo ErrorHandler

    Set ParseStatusesResponse = New Collection

    ' JSON部分を抽出
    If InStr(response, "JSON_START") = 0 Then
        Exit Function
    End If

    Dim startPos As Long
    Dim endPos As Long
    startPos = InStr(response, "JSON_START") + Len("JSON_START") + 2
    endPos = InStr(response, "JSON_END")

    If endPos <= startPos Then
        Exit Function
    End If

    Dim jsonStr As String
    jsonStr = Trim(Mid(response, startPos, endPos - startPos))

    ' 簡易JSONパース（statusesの各要素を抽出）
    Dim units As Collection
    Set units = New Collection

    ' statuses配列を探す
    Dim statusesStart As Long
    statusesStart = InStr(jsonStr, """statuses"":")

    If statusesStart = 0 Then
        Set ParseStatusesResponse = units
        Exit Function
    End If

    ' 各ユニットを抽出（簡易パース）
    Dim pos As Long
    pos = statusesStart

    Do
        ' "definition"を探す
        Dim defStart As Long
        defStart = InStr(pos, jsonStr, """definition"":")

        If defStart = 0 Then Exit Do

        ' unitNameを抽出
        Dim unitNameStart As Long
        unitNameStart = InStr(defStart, jsonStr, """unitName"":""")

        If unitNameStart = 0 Then Exit Do

        unitNameStart = unitNameStart + Len("""unitName"":""")
        Dim unitNameEnd As Long
        unitNameEnd = InStr(unitNameStart, jsonStr, """")

        Dim unitName As String
        unitName = Mid(jsonStr, unitNameStart, unitNameEnd - unitNameStart)

        ' simpleUnitNameを抽出
        Dim simpleNameStart As Long
        simpleNameStart = InStr(defStart, jsonStr, """simpleUnitName"":""")

        Dim simpleName As String
        If simpleNameStart > 0 And simpleNameStart < defStart + 500 Then
            simpleNameStart = simpleNameStart + Len("""simpleUnitName"":""")
            Dim simpleNameEnd As Long
            simpleNameEnd = InStr(simpleNameStart, jsonStr, """")
            simpleName = Mid(jsonStr, simpleNameStart, simpleNameEnd - simpleNameStart)
        Else
            simpleName = unitName
        End If

        ' unitTypeを抽出
        Dim unitTypeStart As Long
        unitTypeStart = InStr(defStart, jsonStr, """unitType"":""")

        Dim unitType As String
        If unitTypeStart > 0 And unitTypeStart < defStart + 500 Then
            unitTypeStart = unitTypeStart + Len("""unitType"":""")
            Dim unitTypeEnd As Long
            unitTypeEnd = InStr(unitTypeStart, jsonStr, """")
            unitType = Mid(jsonStr, unitTypeStart, unitTypeEnd - unitTypeStart)
        Else
            unitType = "UNKNOWN"
        End If

        ' unitStatusを探す
        Dim statusStart As Long
        statusStart = InStr(defStart, jsonStr, """unitStatus"":")

        Dim execID As String
        Dim status As String
        Dim unitStartTime As String
        Dim unitEndTime As String

        execID = ""
        status = ""
        unitStartTime = ""
        unitEndTime = ""

        If statusStart > 0 And statusStart < defStart + 2000 Then
            ' execIDを抽出
            Dim execIDStart As Long
            execIDStart = InStr(statusStart, jsonStr, """execID"":""")
            If execIDStart > 0 And execIDStart < statusStart + 500 Then
                execIDStart = execIDStart + Len("""execID"":""")
                Dim execIDEnd As Long
                execIDEnd = InStr(execIDStart, jsonStr, """")
                execID = Mid(jsonStr, execIDStart, execIDEnd - execIDStart)
            End If

            ' statusを抽出
            Dim statusValStart As Long
            statusValStart = InStr(statusStart, jsonStr, """status"":""")
            If statusValStart > 0 And statusValStart < statusStart + 500 Then
                statusValStart = statusValStart + Len("""status"":""")
                Dim statusValEnd As Long
                statusValEnd = InStr(statusValStart, jsonStr, """")
                status = Mid(jsonStr, statusValStart, statusValEnd - statusValStart)
            End If

            ' startTimeを抽出
            Dim startTimeStart As Long
            startTimeStart = InStr(statusStart, jsonStr, """startTime"":""")
            If startTimeStart > 0 And startTimeStart < statusStart + 800 Then
                startTimeStart = startTimeStart + Len("""startTime"":""")
                Dim startTimeEnd As Long
                startTimeEnd = InStr(startTimeStart, jsonStr, """")
                unitStartTime = Mid(jsonStr, startTimeStart, startTimeEnd - startTimeStart)
            End If

            ' endTimeを抽出
            Dim endTimeStart As Long
            endTimeStart = InStr(statusStart, jsonStr, """endTime"":""")
            If endTimeStart > 0 And endTimeStart < statusStart + 1000 Then
                endTimeStart = endTimeStart + Len("""endTime"":""")
                Dim endTimeEnd As Long
                endTimeEnd = InStr(endTimeStart, jsonStr, """")
                unitEndTime = Mid(jsonStr, endTimeStart, endTimeEnd - endTimeStart)
            End If
        End If

        ' ユニット情報を作成
        Dim unitInfo As Object
        Set unitInfo = CreateObject("Scripting.Dictionary")
        unitInfo("Path") = unitName
        unitInfo("Name") = simpleName
        unitInfo("Type") = unitType
        unitInfo("ExecID") = execID
        unitInfo("Status") = status
        unitInfo("StartTime") = unitStartTime
        unitInfo("EndTime") = unitEndTime

        units.Add unitInfo

        pos = unitNameEnd + 1
    Loop

    Set ParseStatusesResponse = units
    Exit Function

ErrorHandler:
    Set ParseStatusesResponse = New Collection
End Function

