Attribute VB_Name = "JM_Executor"
Option Explicit

'==============================================================================
' JP1 ジョブ管理ツール - 実行モジュール
' PowerShellスクリプト生成、ジョブ実行、結果更新機能を提供
'==============================================================================

' ============================================================================
' 単一ジョブの実行
' ============================================================================
Public Function ExecuteSingleJob(ByVal config As Object, ByVal jobnetPath As String, ByVal isHold As Boolean, ByVal logFilePath As String) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result("Status") = ""
    result("StartTime") = ""
    result("EndTime") = ""
    result("Message") = ""

    Dim waitCompletion As Boolean
    waitCompletion = (config("WaitCompletion") = "はい")

    Dim psScript As String
    psScript = BuildExecuteJobScript(config, jobnetPath, waitCompletion, isHold, logFilePath)

    Dim output As String
    output = ExecutePowerShell(psScript)

    ' 結果をパース
    Dim lines() As String
    lines = Split(output, vbCrLf)

    Dim line As String
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        line = lines(i)
        If InStr(line, "RESULT_STATUS:") > 0 Then
            result("Status") = Trim(Replace(line, "RESULT_STATUS:", ""))
        ElseIf InStr(line, "RESULT_START:") > 0 Then
            result("StartTime") = Trim(Replace(line, "RESULT_START:", ""))
        ElseIf InStr(line, "RESULT_END:") > 0 Then
            result("EndTime") = Trim(Replace(line, "RESULT_END:", ""))
        ElseIf InStr(line, "RESULT_MESSAGE:") > 0 Then
            result("Message") = Trim(Replace(line, "RESULT_MESSAGE:", ""))
        ElseIf InStr(line, "RESULT_LOGPATH:") > 0 Then
            result("LogPath") = Trim(Replace(line, "RESULT_LOGPATH:", ""))
        ElseIf InStr(line, "RESULT_DETAIL:") > 0 Then
            result("Detail") = Trim(Replace(line, "RESULT_DETAIL:", ""))
        ElseIf InStr(line, "ERROR:") > 0 Then
            result("Status") = "エラー"
            result("Message") = line
        End If
    Next i

    If result("Status") = "" Then
        result("Status") = "不明"
        result("Message") = output
    End If

    Set ExecuteSingleJob = result
End Function

' ============================================================================
' ジョブ実行用PowerShellスクリプト生成
' ============================================================================
Public Function BuildExecuteJobScript(ByVal config As Object, ByVal jobnetPath As String, ByVal waitCompletion As Boolean, ByVal isHold As Boolean, ByVal logFilePath As String) As String
    Dim script As String
    Dim isRemote As Boolean
    isRemote = (config("ExecMode") <> "ローカル")

    script = "$ErrorActionPreference = 'Stop'" & vbCrLf
    script = script & vbCrLf

    ' UTF-8エンコーディング設定
    script = script & "# UTF-8エンコーディング設定" & vbCrLf
    script = script & "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8" & vbCrLf
    script = script & "$OutputEncoding = [System.Text.Encoding]::UTF8" & vbCrLf
    script = script & "chcp 65001 | Out-Null" & vbCrLf
    script = script & vbCrLf

    ' ログ出力関数を定義
    script = script & "# デバッグモードフラグ" & vbCrLf
    script = script & "$debugMode = $" & IIf(DEBUG_MODE, "true", "false") & vbCrLf
    script = script & vbCrLf
    script = script & "# ログ出力関数" & vbCrLf
    script = script & "$logFile = '" & Replace(logFilePath, "'", "''") & "'" & vbCrLf
    script = script & "function Write-Log {" & vbCrLf
    script = script & "  param([string]$Message)" & vbCrLf
    script = script & "  if ($Message -match '^\[DEBUG-' -and -not $debugMode) { return }" & vbCrLf
    script = script & "  $timestamp = Get-Date -Format 'yyyy/MM/dd HH:mm:ss'" & vbCrLf
    script = script & "  $logLine = ""[$timestamp] $Message""" & vbCrLf
    script = script & "  Write-Host $logLine" & vbCrLf
    script = script & "  Add-Content -Path $logFile -Value $logLine -Encoding UTF8" & vbCrLf
    script = script & "}" & vbCrLf
    script = script & vbCrLf

    ' 実行モードフラグ
    script = script & "# 実行モード設定" & vbCrLf
    script = script & "$isRemote = $" & IIf(isRemote, "true", "false") & vbCrLf
    script = script & "$session = $null" & vbCrLf
    script = script & vbCrLf

    ' JP1コマンド実行関数（ローカル/リモート共通）
    script = script & "# JP1コマンド実行関数（ローカル/リモート共通）" & vbCrLf
    script = script & "function Invoke-JP1Command {" & vbCrLf
    script = script & "  param([string]$CommandName, [string[]]$Arguments)" & vbCrLf
    script = script & "  $searchPaths = @('C:\Program Files\HITACHI\JP1AJS3\bin','C:\Program Files (x86)\HITACHI\JP1AJS3\bin','C:\Program Files\Hitachi\JP1AJS2\bin','C:\Program Files (x86)\Hitachi\JP1AJS2\bin')" & vbCrLf
    script = script & "  if ($isRemote) {" & vbCrLf
    script = script & "    Invoke-Command -Session $session -ScriptBlock {" & vbCrLf
    script = script & "      param($cmdName, $cmdArgs, $paths)" & vbCrLf
    script = script & "      $cmdPath = $null" & vbCrLf
    script = script & "      foreach ($p in $paths) { if (Test-Path ""$p\$cmdName"") { $cmdPath = ""$p\$cmdName""; break } }" & vbCrLf
    script = script & "      if (-not $cmdPath) { return @{ ExitCode = 1; Output = ""$cmdName not found"" } }" & vbCrLf
    script = script & "      $output = & $cmdPath $cmdArgs 2>&1" & vbCrLf
    script = script & "      @{ ExitCode = $LASTEXITCODE; Output = $output }" & vbCrLf
    script = script & "    } -ArgumentList $CommandName, $Arguments, $searchPaths" & vbCrLf
    script = script & "  } else {" & vbCrLf
    script = script & "    $cmdPath = $null" & vbCrLf
    script = script & "    foreach ($p in $searchPaths) { if (Test-Path ""$p\$CommandName"") { $cmdPath = ""$p\$CommandName""; break } }" & vbCrLf
    script = script & "    if (-not $cmdPath) { return @{ ExitCode = 1; Output = ""$CommandName not found"" } }" & vbCrLf
    script = script & "    $output = & $cmdPath $Arguments 2>&1" & vbCrLf
    script = script & "    @{ ExitCode = $LASTEXITCODE; Output = $output }" & vbCrLf
    script = script & "  }" & vbCrLf
    script = script & "}" & vbCrLf
    script = script & vbCrLf

    ' ファイル読み取り関数（ローカル/リモート共通）
    script = script & "# ファイル読み取り関数（ローカル/リモート共通）" & vbCrLf
    script = script & "function Read-FileContent {" & vbCrLf
    script = script & "  param([string]$FilePath)" & vbCrLf
    script = script & "  if ($isRemote) {" & vbCrLf
    script = script & "    Invoke-Command -Session $session -ScriptBlock {" & vbCrLf
    script = script & "      param($path)" & vbCrLf
    script = script & "      if (Test-Path $path) { Get-Content $path -Encoding Default -ErrorAction SilentlyContinue }" & vbCrLf
    script = script & "    } -ArgumentList $FilePath" & vbCrLf
    script = script & "  } else {" & vbCrLf
    script = script & "    if (Test-Path $FilePath) { Get-Content $FilePath -Encoding Default -ErrorAction SilentlyContinue }" & vbCrLf
    script = script & "  }" & vbCrLf
    script = script & "}" & vbCrLf
    script = script & vbCrLf

    ' リモートモードの場合: WinRM設定変数
    If isRemote Then
        script = script & "# WinRM設定変数" & vbCrLf
        script = script & "$originalTrustedHosts = $null" & vbCrLf
        script = script & "$winrmConfigChanged = $false" & vbCrLf
        script = script & "$winrmServiceWasStarted = $false" & vbCrLf
        script = script & "$securePass = ConvertTo-SecureString '" & EscapePSString(config("RemotePassword")) & "' -AsPlainText -Force" & vbCrLf
        script = script & "$cred = New-Object System.Management.Automation.PSCredential('" & config("RemoteUser") & "', $securePass)" & vbCrLf
        script = script & vbCrLf
    End If

    script = script & "try {" & vbCrLf
    script = script & "  Write-Log '--------------------------------------------------------------------------------'" & vbCrLf
    script = script & "  Write-Log 'ジョブネット: " & jobnetPath & "'" & vbCrLf

    If isRemote Then
        script = script & "  Write-Log '接続先: " & config("JP1Server") & " (リモートモード)'" & vbCrLf
    End If

    script = script & "  Write-Log '--------------------------------------------------------------------------------'" & vbCrLf
    script = script & vbCrLf

    ' リモートモードの場合: WinRMセットアップ
    If isRemote Then
        script = script & "  # WinRMサービス起動確認" & vbCrLf
        script = script & "  $winrmService = Get-Service -Name WinRM -ErrorAction SilentlyContinue" & vbCrLf
        script = script & "  if ($winrmService.Status -ne 'Running') {" & vbCrLf
        script = script & "    Write-Log '[準備] WinRMサービスを起動'" & vbCrLf
        script = script & "    Start-Service -Name WinRM -ErrorAction Stop" & vbCrLf
        script = script & "    $winrmServiceWasStarted = $true" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & "  $originalTrustedHosts = (Get-Item WSMan:\localhost\Client\TrustedHosts -ErrorAction SilentlyContinue).Value" & vbCrLf
        script = script & "  if ($originalTrustedHosts -notmatch '" & config("JP1Server") & "') {" & vbCrLf
        script = script & "    Write-Log '[準備] TrustedHostsに接続先を追加'" & vbCrLf
        script = script & "    if ($originalTrustedHosts) {" & vbCrLf
        script = script & "      Set-Item WSMan:\localhost\Client\TrustedHosts -Value ""$originalTrustedHosts," & config("JP1Server") & """ -Force" & vbCrLf
        script = script & "    } else {" & vbCrLf
        script = script & "      Set-Item WSMan:\localhost\Client\TrustedHosts -Value '" & config("JP1Server") & "' -Force" & vbCrLf
        script = script & "    }" & vbCrLf
        script = script & "    $winrmConfigChanged = $true" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & "  Write-Log '[接続] リモートセッション作成中...'" & vbCrLf
        script = script & "  $session = New-PSSession -ComputerName '" & config("JP1Server") & "' -Credential $cred -ErrorAction Stop" & vbCrLf
        script = script & "  Write-Log '[接続] セッション確立完了'" & vbCrLf
        script = script & vbCrLf
    End If

    ' 保留解除処理（共通）
    If isHold Then
        script = script & "  # 保留解除" & vbCrLf
        script = script & "  Write-Log '[実行] ajsplan -r - 保留解除'" & vbCrLf
        script = script & "  $releaseResult = Invoke-JP1Command 'ajsplan.exe' @('-F', '" & config("SchedulerService") & "', '-r', '" & jobnetPath & "')" & vbCrLf
        script = script & "  Write-Log ""結果: $($releaseResult.Output -join ' ')""" & vbCrLf
        script = script & "  if ($releaseResult.ExitCode -ne 0) {" & vbCrLf
        script = script & "    Write-Log '[ERROR] 保留解除失敗'" & vbCrLf
        script = script & "    Write-Output ""RESULT_STATUS:保留解除失敗""" & vbCrLf
        script = script & "    Write-Output ""RESULT_MESSAGE:$($releaseResult.Output -join ' ')""" & vbCrLf
        script = script & "    Write-Output ""RESULT_LOGPATH:$logFile""" & vbCrLf
        script = script & "    exit" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & "  Write-Log '[成功] 保留解除完了'" & vbCrLf
        script = script & vbCrLf
    End If

    ' ajsentry実行前に現在の最新実行IDを取得（比較用）
    script = script & "  # ajsentry実行前の実行IDを取得（比較用）" & vbCrLf
    script = script & "  $beforeIdResult = Invoke-JP1Command 'ajsshow.exe' @('-F', '" & config("SchedulerService") & "', '-g', '1', '-i', '%##', '" & jobnetPath & "')" & vbCrLf
    script = script & "  $beforeExecId = ($beforeIdResult.Output -join '').Trim()" & vbCrLf
    script = script & vbCrLf

    ' ajsentry実行（共通）-n: 即時実行, -w: 完了待ち
    script = script & "  # ajsentry実行（即時実行・完了待ち）" & vbCrLf
    script = script & "  Write-Log '[実行] ajsentry - ジョブ起動（-wオプションで完了待ち）'" & vbCrLf
    script = script & "  $entryResult = Invoke-JP1Command 'ajsentry.exe' @('-F', '" & config("SchedulerService") & "', '-n', '-w', '" & jobnetPath & "')" & vbCrLf
    script = script & "  $entryOutput = if ($entryResult.Output) { $entryResult.Output -join ' ' } else { '' }" & vbCrLf
    script = script & "  $entryExitCode = $entryResult.ExitCode" & vbCrLf
    script = script & "  if ($entryOutput) { Write-Log ""結果: $entryOutput"" } else { Write-Log ""結果: 正常終了 (ExitCode=$entryExitCode)"" }" & vbCrLf
    script = script & vbCrLf
    script = script & "  # ajsentryの実行結果をチェック" & vbCrLf
    script = script & "  if ($entryExitCode -ne 0 -or $entryOutput -match 'KAVS\d+-E') {" & vbCrLf
    script = script & "    $errMsg = if ($entryOutput) { $entryOutput } else { ""ExitCode=$entryExitCode"" }" & vbCrLf
    script = script & "    Write-Log ""[ERROR] ajsentryエラー: $errMsg""" & vbCrLf
    script = script & "    Write-Output ""RESULT_STATUS:実行エラー""" & vbCrLf
    script = script & "    Write-Output ""RESULT_MESSAGE:$errMsg""" & vbCrLf
    script = script & "    Write-Output ""RESULT_LOGPATH:$logFile""" & vbCrLf
    script = script & "    exit" & vbCrLf
    script = script & "  }" & vbCrLf
    script = script & vbCrLf
    script = script & "  # 実行IDを取得（ajsentry後の最新世代）" & vbCrLf
    script = script & "  $execIdResult = Invoke-JP1Command 'ajsshow.exe' @('-F', '" & config("SchedulerService") & "', '-g', '1', '-i', '%##', '" & jobnetPath & "')" & vbCrLf
    script = script & "  $execId = ($execIdResult.Output -join '').Trim()" & vbCrLf
    script = script & "  # 実行IDが空または不正な場合のチェック" & vbCrLf
    script = script & "  if (-not $execId -or $execId -match 'KAVS\d+-E') {" & vbCrLf
    script = script & "    Write-Log '[ERROR] 実行IDの取得に失敗しました'" & vbCrLf
    script = script & "    Write-Output ""RESULT_STATUS:実行ID取得失敗""" & vbCrLf
    script = script & "    Write-Output ""RESULT_MESSAGE:$($execIdResult.Output -join ' ')""" & vbCrLf
    script = script & "    Write-Output ""RESULT_LOGPATH:$logFile""" & vbCrLf
    script = script & "    exit" & vbCrLf
    script = script & "  }" & vbCrLf
    script = script & "  Write-Log ""実行ID: $execId""" & vbCrLf
    script = script & vbCrLf
    script = script & "  # 実行IDが変わったことを確認（今回の実行であることを保証）" & vbCrLf
    script = script & "  if ($execId -eq $beforeExecId) {" & vbCrLf
    script = script & "    Write-Log '[ERROR] 実行IDが変化していません。ジョブが実行されませんでした。'" & vbCrLf
    script = script & "    Write-Output ""RESULT_STATUS:実行ID未変化""" & vbCrLf
    script = script & "    Write-Output ""RESULT_MESSAGE:実行IDが変化していません（前回: $beforeExecId）""" & vbCrLf
    script = script & "    Write-Output ""RESULT_LOGPATH:$logFile""" & vbCrLf
    script = script & "    exit" & vbCrLf
    script = script & "  }" & vbCrLf
    script = script & vbCrLf

    If waitCompletion Then
        ' 完了待ちモードの結果取得・エラー詳細処理
        script = script & BuildWaitCompletionScript(config, jobnetPath)
    Else
        script = script & "  Write-Log '[完了] 起動成功（完了待ちなし）'" & vbCrLf
        script = script & "  Write-Output ""RESULT_STATUS:起動成功""" & vbCrLf
        script = script & "  Write-Output ""RESULT_LOGPATH:$logFile""" & vbCrLf
        script = script & "  $msgOutput = if ($entryOutput) { $entryOutput } else { ""起動完了 (ExitCode=$entryExitCode)"" }" & vbCrLf
        script = script & "  Write-Output ""RESULT_MESSAGE:$msgOutput""" & vbCrLf
    End If

    ' リモートモードの場合: セッション終了
    If isRemote Then
        script = script & "  Write-Log '[クリーンアップ] セッション終了'" & vbCrLf
        script = script & "  Remove-PSSession $session" & vbCrLf
    End If

    script = script & "} catch {" & vbCrLf
    script = script & "  Write-Log ""[EXCEPTION] $($_.Exception.Message)""" & vbCrLf
    script = script & "  Write-Output ""ERROR: $($_.Exception.Message)""" & vbCrLf
    script = script & "}" & vbCrLf

    ' リモートモードの場合: WinRM設定復元
    If isRemote Then
        script = script & "finally {" & vbCrLf
        script = script & "  Write-Log '[クリーンアップ] WinRM設定を復元中...'" & vbCrLf
        script = script & "  if ($winrmConfigChanged) {" & vbCrLf
        script = script & "    if ($originalTrustedHosts) {" & vbCrLf
        script = script & "      Set-Item WSMan:\localhost\Client\TrustedHosts -Value $originalTrustedHosts -Force -ErrorAction SilentlyContinue" & vbCrLf
        script = script & "    } else {" & vbCrLf
        script = script & "      Clear-Item WSMan:\localhost\Client\TrustedHosts -Force -ErrorAction SilentlyContinue" & vbCrLf
        script = script & "    }" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & "  if ($winrmServiceWasStarted) {" & vbCrLf
        script = script & "    Stop-Service -Name WinRM -Force -ErrorAction SilentlyContinue" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & "  Write-Log '[完了] 処理終了'" & vbCrLf
        script = script & "}" & vbCrLf
    End If

    BuildExecuteJobScript = script
End Function

' ============================================================================
' 完了待ちモードのスクリプト部分
' ============================================================================
Private Function BuildWaitCompletionScript(ByVal config As Object, ByVal jobnetPath As String) As String
    Dim script As String

    ' ajsentry -w終了後、ajsshowで1回だけ結果を取得
    script = "  # ajsentry終了後、ajsshowで1回だけ結果を取得" & vbCrLf
    script = script & "  $statusResult = Invoke-JP1Command 'ajsshow.exe' @('-F', '" & config("SchedulerService") & "', '-g', '1', '-i', '%CC', '" & jobnetPath & "')" & vbCrLf
    script = script & "  $jobStatus = ($statusResult.Output -join ' ').Trim()" & vbCrLf
    script = script & "  Write-Log ""ジョブネット状態: $jobStatus""" & vbCrLf
    script = script & vbCrLf
    script = script & "  # ajsshowコマンド自体のエラーチェック" & vbCrLf
    script = script & "  if ($statusResult.ExitCode -ne 0 -or $jobStatus -match 'KAVS\d+-E') {" & vbCrLf
    script = script & "    Write-Log ""[ERROR] ajsshowコマンドエラー: $jobStatus""" & vbCrLf
    script = script & "    Write-Output ""RESULT_STATUS:コマンドエラー""" & vbCrLf
    script = script & "    Write-Output ""RESULT_MESSAGE:$jobStatus""" & vbCrLf
    script = script & "    Write-Output ""RESULT_LOGPATH:$logFile""" & vbCrLf
    script = script & "    exit" & vbCrLf
    script = script & "  }" & vbCrLf
    script = script & vbCrLf
    script = script & "  # 状態判定" & vbCrLf
    script = script & "  if ($jobStatus -match '正常終了') {" & vbCrLf
    script = script & "    Write-Log '[完了] 正常終了'" & vbCrLf
    script = script & "    Write-Output ""RESULT_STATUS:正常終了""" & vbCrLf
    script = script & "    Write-Output ""RESULT_LOGPATH:$logFile""" & vbCrLf
    script = script & "  } elseif ($jobStatus -match '警告検出終了|警告終了') {" & vbCrLf
    script = script & "    Write-Log '[完了] 警告検出終了'" & vbCrLf
    script = script & "    Write-Output ""RESULT_STATUS:警告検出終了""" & vbCrLf
    script = script & "  } elseif ($jobStatus -match '異常終了|異常検出終了|強制終了|中断') {" & vbCrLf
    script = script & "    Write-Log '[完了] 異常終了'" & vbCrLf
    script = script & "    Write-Output ""RESULT_STATUS:異常終了""" & vbCrLf
    script = script & "  } else {" & vbCrLf
    script = script & "    Write-Log ""[完了] 状態: $jobStatus""" & vbCrLf
    script = script & "    Write-Output ""RESULT_STATUS:$jobStatus""" & vbCrLf
    script = script & "  }" & vbCrLf
    script = script & vbCrLf

    ' 詳細情報取得
    script = script & "  # 詳細情報取得" & vbCrLf
    script = script & "  $detailStatusResult = Invoke-JP1Command 'ajsshow.exe' @('-F', '" & config("SchedulerService") & "', '-g', '1', '-i', '%JJ %CC %SS %EE', '" & jobnetPath & "')" & vbCrLf
    script = script & "  $lastStatusStr = $detailStatusResult.Output -join ' '" & vbCrLf
    script = script & "  Write-Log ""詳細ステータス: $lastStatusStr""" & vbCrLf
    script = script & "  if ($detailStatusResult.ExitCode -ne 0 -or $lastStatusStr -match 'KAVS\d+-E') {" & vbCrLf
    script = script & "    Write-Log ""[ERROR] 詳細情報取得エラー: $lastStatusStr""" & vbCrLf
    script = script & "  }" & vbCrLf
    script = script & vbCrLf

    ' 時間抽出
    script = script & "  # 時間抽出" & vbCrLf
    script = script & "  $timePattern = '\d{4}/\d{2}/\d{2}\s+\d{2}:\d{2}'" & vbCrLf
    script = script & "  $allTimes = [regex]::Matches($lastStatusStr, $timePattern)" & vbCrLf
    script = script & "  $startTimeStr = ''" & vbCrLf
    script = script & "  $endTimeStr = ''" & vbCrLf
    script = script & "  if ($allTimes.Count -ge 1) {" & vbCrLf
    script = script & "    $startTimeStr = $allTimes[0].Value" & vbCrLf
    script = script & "    Write-Output ""RESULT_START:$startTimeStr""" & vbCrLf
    script = script & "    Write-Log ""開始時刻: $startTimeStr""" & vbCrLf
    script = script & "  }" & vbCrLf
    script = script & "  if ($allTimes.Count -ge 2) {" & vbCrLf
    script = script & "    $endTimeStr = $allTimes[1].Value" & vbCrLf
    script = script & "    Write-Output ""RESULT_END:$endTimeStr""" & vbCrLf
    script = script & "    Write-Log ""終了時刻: $endTimeStr""" & vbCrLf
    script = script & "  }" & vbCrLf
    script = script & "  if ($startTimeStr -and $endTimeStr) {" & vbCrLf
    script = script & "    try {" & vbCrLf
    script = script & "      $startDt = [datetime]::ParseExact($startTimeStr, 'yyyy/MM/dd HH:mm', $null)" & vbCrLf
    script = script & "      $endDt = [datetime]::ParseExact($endTimeStr, 'yyyy/MM/dd HH:mm', $null)" & vbCrLf
    script = script & "      $duration = $endDt - $startDt" & vbCrLf
    script = script & "      $durationStr = '{0:D2}:{1:D2}:{2:D2}' -f [int]$duration.TotalHours, $duration.Minutes, $duration.Seconds" & vbCrLf
    script = script & "      Write-Log ""実行時間: $durationStr""" & vbCrLf
    script = script & "    } catch { }" & vbCrLf
    script = script & "  }" & vbCrLf
    script = script & "  $cleanMsg = $lastStatusStr -replace 'KAVS\d+-[IEW][^\r\n]*', '' -replace '\s+', ' '" & vbCrLf
    script = script & "  Write-Output ""RESULT_MESSAGE:$cleanMsg""" & vbCrLf
    script = script & vbCrLf

    ' ジョブネット内のジョブ一覧を表示
    script = script & "  # ジョブネット内のジョブ状態一覧を取得" & vbCrLf
    script = script & "  Write-Log ''" & vbCrLf
    script = script & "  Write-Log '【ジョブ実行結果一覧】'" & vbCrLf
    script = script & "  $jobListResult = Invoke-JP1Command 'ajsshow.exe' @('-F', '" & config("SchedulerService") & "', '-g', '1', '-R', '-f', '%JJ %TT %CC %RR', '" & jobnetPath & "')" & vbCrLf
    script = script & "  Write-Log ('  ' + '-' * 78)" & vbCrLf
    script = script & "  Write-Log ('  {0,-40} {1,-10} {2,-12} {3}' -f 'ジョブ名', 'タイプ', '状態', '戻り値')" & vbCrLf
    script = script & "  Write-Log ('  ' + '-' * 78)" & vbCrLf
    script = script & "  foreach ($jobLine in $jobListResult.Output) {" & vbCrLf
    script = script & "    if ($jobLine -match '^(/[^\s]+)\s+(\S+)\s+(\S+)\s*(.*)$') {" & vbCrLf
    script = script & "      $jName = $matches[1]" & vbCrLf
    script = script & "      $jType = $matches[2]" & vbCrLf
    script = script & "      $jStatus = $matches[3]" & vbCrLf
    script = script & "      $jReturn = $matches[4].Trim()" & vbCrLf
    script = script & "      if ($jName.Length -gt 40) { $jName = '...' + $jName.Substring($jName.Length - 37) }" & vbCrLf
    script = script & "      $statusMark = switch -Regex ($jStatus) {" & vbCrLf
    script = script & "        '正常終了' { '[OK]' }" & vbCrLf
    script = script & "        '異常終了|異常検出' { '[NG]' }" & vbCrLf
    script = script & "        '警告' { '[!]' }" & vbCrLf
    script = script & "        '未実行|未起動' { '[-]' }" & vbCrLf
    script = script & "        default { '' }" & vbCrLf
    script = script & "      }" & vbCrLf
    script = script & "      Write-Log ('  {0,-40} {1,-10} {2} {3,-10} {4}' -f $jName, $jType, $statusMark, $jStatus, $jReturn)" & vbCrLf
    script = script & "    }" & vbCrLf
    script = script & "  }" & vbCrLf
    script = script & "  Write-Log ('  ' + '-' * 78)" & vbCrLf
    script = script & "  Write-Log ''" & vbCrLf
    script = script & vbCrLf

    ' エラー詳細取得
    script = script & "  # エラー詳細取得" & vbCrLf
    script = script & "  if ($jobStatus -match '警告検出終了|警告終了|異常終了|異常検出終了|強制終了|中断') {" & vbCrLf
    script = script & "    Write-Log '[詳細取得] 異常終了したジョブを検索中...'" & vbCrLf
    script = script & "    $failedJobsResult = Invoke-JP1Command 'ajsshow.exe' @('-F', '" & config("SchedulerService") & "', '-R', '-f', '%J %T %C %R', '" & jobnetPath & "')" & vbCrLf
    script = script & "    $failedJobsStr = $failedJobsResult.Output -join ""`n""" & vbCrLf
    script = script & "    Write-Log ""[DEBUG] ajsshow -R -f 結果: $failedJobsStr""" & vbCrLf
    script = script & vbCrLf
    script = script & "    $failedJobPath = ''" & vbCrLf
    script = script & "    $nonZeroReturnJobPath = ''" & vbCrLf
    script = script & "    foreach ($line in $failedJobsResult.Output) {" & vbCrLf
    script = script & "      if ($line -match '^(/[^\s]+)\s+(\w*job|\w*jb)\s+(異常終了|警告終了|警告検出終了|Abnormal|Warning|ended abnormally|ended with warning)') {" & vbCrLf
    script = script & "        $failedJobPath = $matches[1]" & vbCrLf
    script = script & "        Write-Log ""[DEBUG] 異常終了ジョブ検出: $failedJobPath""" & vbCrLf
    script = script & "        break" & vbCrLf
    script = script & "      }" & vbCrLf
    script = script & "      if (-not $nonZeroReturnJobPath -and $line -match '^(/[^\s]+)\s+(\w*job|\w*jb)\s+\S+\s+([1-9]\d*|-\d+)') {" & vbCrLf
    script = script & "        $nonZeroReturnJobPath = $matches[1]" & vbCrLf
    script = script & "      }" & vbCrLf
    script = script & "    }" & vbCrLf
    script = script & "    if (-not $failedJobPath -and $nonZeroReturnJobPath) { $failedJobPath = $nonZeroReturnJobPath }" & vbCrLf
    script = script & vbCrLf
    script = script & "    if ($failedJobPath) {" & vbCrLf
    script = script & "      Write-Log ""[DEBUG] failedJobPath: $failedJobPath""" & vbCrLf
    script = script & "      $detailResult = Invoke-JP1Command 'ajsshow.exe' @('-F', '" & config("SchedulerService") & "', '-g', '1', '-i', '%## %ll %rr', $failedJobPath)" & vbCrLf
    script = script & "      $detailStr = $detailResult.Output -join ""`n""" & vbCrLf
    script = script & "      Write-Log ""[DEBUG] 詳細結果: $detailStr""" & vbCrLf
    script = script & vbCrLf
    script = script & "      $stderrFile = ''" & vbCrLf
    script = script & "      if ($detailStr -match '[A-Za-z]:[^\r\n]+\.err') { $stderrFile = $matches[0] }" & vbCrLf
    script = script & "      if ($stderrFile) {" & vbCrLf
    script = script & "        Write-Output ""RESULT_LOGPATH:$logFile""" & vbCrLf
    script = script & "        Write-Log ""[DEBUG] 標準エラーファイル: $stderrFile""" & vbCrLf
    script = script & "        $logContent = Read-FileContent $stderrFile" & vbCrLf
    script = script & "        if ($logContent) {" & vbCrLf
    script = script & "          Write-Log '[詳細] 標準エラーログ:'" & vbCrLf
    script = script & "          foreach ($line in $logContent) { Write-Log ""  $line"" }" & vbCrLf
    script = script & "        } else {" & vbCrLf
    script = script & "          Write-Log '標準エラーログを取得できませんでした'" & vbCrLf
    script = script & "        }" & vbCrLf
    script = script & "      }" & vbCrLf
    script = script & "    } else {" & vbCrLf
    script = script & "      Write-Log '異常終了したジョブが見つかりませんでした'" & vbCrLf
    script = script & "    }" & vbCrLf
    script = script & "  }" & vbCrLf

    BuildWaitCompletionScript = script
End Function

' ============================================================================
' ジョブ一覧取得用PowerShellスクリプト生成
' ============================================================================
Public Function BuildGetJobListScript(config As Object) As String
    Dim script As String

    script = "$ErrorActionPreference = 'Stop'" & vbCrLf
    script = script & vbCrLf

    ' UTF-8エンコーディング設定（日本語パス対応）
    script = script & "# UTF-8エンコーディング設定" & vbCrLf
    script = script & "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8" & vbCrLf
    script = script & "$OutputEncoding = [System.Text.Encoding]::UTF8" & vbCrLf
    script = script & "chcp 65001 | Out-Null" & vbCrLf
    script = script & vbCrLf

    ' ログ出力関数を定義（コンソール表示用）
    script = script & "function Write-Log {" & vbCrLf
    script = script & "  param([string]$Message)" & vbCrLf
    script = script & "  $timestamp = Get-Date -Format 'yyyy/MM/dd HH:mm:ss'" & vbCrLf
    script = script & "  Write-Host ""[$timestamp] $Message""" & vbCrLf
    script = script & "}" & vbCrLf
    script = script & vbCrLf

    ' ローカルモードとリモートモードで処理を分岐
    If config("ExecMode") = "ローカル" Then
        script = script & BuildLocalJobListScript(config)
    Else
        script = script & BuildRemoteJobListScript(config)
    End If

    BuildGetJobListScript = script
End Function

Private Function BuildLocalJobListScript(config As Object) As String
    Dim script As String

    script = "try {" & vbCrLf
    script = script & "  Write-Log '--------------------------------------------------------------------------------'" & vbCrLf
    script = script & "  Write-Log 'ジョブ一覧取得開始'" & vbCrLf
    script = script & "  Write-Log ""対象パス: " & config("RootPath") & """" & vbCrLf
    script = script & "  Write-Log '--------------------------------------------------------------------------------'" & vbCrLf
    script = script & vbCrLf
    script = script & "  # JP1コマンドパスの検出" & vbCrLf
    script = script & "  $ajsprintPath = $null" & vbCrLf
    script = script & "  $searchPaths = @(" & vbCrLf
    script = script & "    'C:\Program Files\HITACHI\JP1AJS3\bin\ajsprint.exe'," & vbCrLf
    script = script & "    'C:\Program Files (x86)\HITACHI\JP1AJS3\bin\ajsprint.exe'," & vbCrLf
    script = script & "    'C:\Program Files\Hitachi\JP1AJS2\bin\ajsprint.exe'," & vbCrLf
    script = script & "    'C:\Program Files (x86)\Hitachi\JP1AJS2\bin\ajsprint.exe'" & vbCrLf
    script = script & "  )" & vbCrLf
    script = script & "  foreach ($path in $searchPaths) {" & vbCrLf
    script = script & "    if (Test-Path $path) { $ajsprintPath = $path; break }" & vbCrLf
    script = script & "  }" & vbCrLf
    script = script & "  if (-not $ajsprintPath) {" & vbCrLf
    script = script & "    Write-Log '[ERROR] JP1コマンド(ajsprint.exe)が見つかりません'" & vbCrLf
    script = script & "    Write-Output ""ERROR: JP1コマンド(ajsprint.exe)が見つかりません。JP1/AJS3 Managerがインストールされているか確認してください。""" & vbCrLf
    script = script & "    exit 1" & vbCrLf
    script = script & "  }" & vbCrLf
    script = script & "  Write-Log ""JP1コマンドパス: $ajsprintPath""" & vbCrLf
    script = script & vbCrLf
    script = script & "  # ローカルでajsprintを実行" & vbCrLf
    script = script & "  Write-Log '[実行] ajsprint - ジョブ一覧取得'" & vbCrLf
    script = script & "  Write-Log ""コマンド: ajsprint.exe -F " & config("SchedulerService") & " " & config("RootPath") & " -R""" & vbCrLf
    script = script & "  $result = & $ajsprintPath -F " & config("SchedulerService") & " '" & config("RootPath") & "' -R 2>&1" & vbCrLf
    script = script & "  Write-Log '[成功] ジョブ一覧取得完了'" & vbCrLf
    script = script & "  $result | ForEach-Object { Write-Output $_ }" & vbCrLf
    script = script & "} catch {" & vbCrLf
    script = script & "  Write-Log ""[ERROR] $($_.Exception.Message)""" & vbCrLf
    script = script & "  Write-Output ""ERROR: $($_.Exception.Message)""" & vbCrLf
    script = script & "}" & vbCrLf

    BuildLocalJobListScript = script
End Function

Private Function BuildRemoteJobListScript(config As Object) As String
    Dim script As String

    script = "Write-Log '--------------------------------------------------------------------------------'" & vbCrLf
    script = script & "Write-Log 'ジョブ一覧取得開始（リモート）'" & vbCrLf
    script = script & "Write-Log ""対象パス: " & config("RootPath") & """" & vbCrLf
    script = script & "Write-Log ""接続先: " & config("JP1Server") & """" & vbCrLf
    script = script & "Write-Log '--------------------------------------------------------------------------------'" & vbCrLf
    script = script & vbCrLf
    script = script & "$securePass = ConvertTo-SecureString '" & EscapePSString(config("RemotePassword")) & "' -AsPlainText -Force" & vbCrLf
    script = script & "$cred = New-Object System.Management.Automation.PSCredential('" & config("RemoteUser") & "', $securePass)" & vbCrLf
    script = script & vbCrLf
    script = script & "$originalTrustedHosts = $null" & vbCrLf
    script = script & "$winrmConfigChanged = $false" & vbCrLf
    script = script & "$winrmServiceWasStarted = $false" & vbCrLf
    script = script & vbCrLf
    script = script & "try {" & vbCrLf
    script = script & "  Write-Log '[実行] WinRMサービス確認'" & vbCrLf
    script = script & "  $winrmService = Get-Service -Name WinRM -ErrorAction SilentlyContinue" & vbCrLf
    script = script & "  if ($winrmService.Status -ne 'Running') {" & vbCrLf
    script = script & "    Start-Service -Name WinRM -ErrorAction Stop" & vbCrLf
    script = script & "    $winrmServiceWasStarted = $true" & vbCrLf
    script = script & "  }" & vbCrLf
    script = script & "  $originalTrustedHosts = (Get-Item WSMan:\localhost\Client\TrustedHosts -ErrorAction SilentlyContinue).Value" & vbCrLf
    script = script & "  if ($originalTrustedHosts -notmatch '" & config("JP1Server") & "') {" & vbCrLf
    script = script & "    if ($originalTrustedHosts) {" & vbCrLf
    script = script & "      Set-Item WSMan:\localhost\Client\TrustedHosts -Value ""$originalTrustedHosts," & config("JP1Server") & """ -Force" & vbCrLf
    script = script & "    } else {" & vbCrLf
    script = script & "      Set-Item WSMan:\localhost\Client\TrustedHosts -Value '" & config("JP1Server") & "' -Force" & vbCrLf
    script = script & "    }" & vbCrLf
    script = script & "    $winrmConfigChanged = $true" & vbCrLf
    script = script & "  }" & vbCrLf
    script = script & "  Write-Log '[実行] リモートセッション作成'" & vbCrLf
    script = script & "  $session = New-PSSession -ComputerName '" & config("JP1Server") & "' -Credential $cred -ErrorAction Stop" & vbCrLf
    script = script & "  Write-Log '[成功] リモートセッション作成完了'" & vbCrLf
    script = script & vbCrLf
    script = script & "  Write-Log '[実行] ajsprint - ジョブ一覧取得'" & vbCrLf
    script = script & "  Write-Log ""コマンド: ajsprint.exe -F " & config("SchedulerService") & " " & config("RootPath") & " -R""" & vbCrLf
    script = script & "  $result = Invoke-Command -Session $session -ScriptBlock {" & vbCrLf
    script = script & "    param($schedulerService, $rootPath)" & vbCrLf
    script = script & "    if ([string]::IsNullOrWhiteSpace($rootPath)) { Write-Output 'ERROR: rootPath is empty'; return }" & vbCrLf
    script = script & "    $ajsprintPath = $null" & vbCrLf
    script = script & "    $searchPaths = @('C:\Program Files\HITACHI\JP1AJS3\bin\ajsprint.exe','C:\Program Files (x86)\HITACHI\JP1AJS3\bin\ajsprint.exe','C:\Program Files\Hitachi\JP1AJS2\bin\ajsprint.exe','C:\Program Files (x86)\Hitachi\JP1AJS2\bin\ajsprint.exe')" & vbCrLf
    script = script & "    foreach ($p in $searchPaths) { if (Test-Path $p) { $ajsprintPath = $p; break } }" & vbCrLf
    script = script & "    if (-not $ajsprintPath) { Write-Output 'ERROR: ajsprint.exe not found'; return }" & vbCrLf
    script = script & "    $output = & $ajsprintPath '-F' $schedulerService $rootPath '-R' 2>&1" & vbCrLf
    script = script & "    $output | Where-Object { $_ -notmatch '^KAVS\d+-I' }" & vbCrLf
    script = script & "  } -ArgumentList '" & config("SchedulerService") & "', '" & config("RootPath") & "'" & vbCrLf
    script = script & "  Write-Log '[成功] ジョブ一覧取得完了'" & vbCrLf
    script = script & "  Remove-PSSession $session" & vbCrLf
    script = script & "  $result | ForEach-Object { Write-Output $_ }" & vbCrLf
    script = script & "} catch {" & vbCrLf
    script = script & "  Write-Log ""[ERROR] $($_.Exception.Message)""" & vbCrLf
    script = script & "  Write-Output ""ERROR: $($_.Exception.Message)""" & vbCrLf
    script = script & "} finally {" & vbCrLf
    script = script & "  if ($winrmConfigChanged) {" & vbCrLf
    script = script & "    if ($originalTrustedHosts) {" & vbCrLf
    script = script & "      Set-Item WSMan:\localhost\Client\TrustedHosts -Value $originalTrustedHosts -Force -ErrorAction SilentlyContinue" & vbCrLf
    script = script & "    } else {" & vbCrLf
    script = script & "      Clear-Item WSMan:\localhost\Client\TrustedHosts -Force -ErrorAction SilentlyContinue" & vbCrLf
    script = script & "    }" & vbCrLf
    script = script & "  }" & vbCrLf
    script = script & "  if ($winrmServiceWasStarted) {" & vbCrLf
    script = script & "    Stop-Service -Name WinRM -Force -ErrorAction SilentlyContinue" & vbCrLf
    script = script & "  }" & vbCrLf
    script = script & "}" & vbCrLf

    BuildRemoteJobListScript = script
End Function

' ============================================================================
' グループ一覧取得用PowerShellスクリプト生成
' ============================================================================
Public Function BuildGetGroupListScript(config As Object) As String
    Dim script As String

    script = "$ErrorActionPreference = 'Stop'" & vbCrLf
    script = script & vbCrLf

    ' UTF-8エンコーディング設定（日本語パス対応）
    script = script & "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8" & vbCrLf
    script = script & "$OutputEncoding = [System.Text.Encoding]::UTF8" & vbCrLf
    script = script & "chcp 65001 | Out-Null" & vbCrLf
    script = script & vbCrLf

    ' ログ出力関数を定義（コンソール表示用）
    script = script & "function Write-Log {" & vbCrLf
    script = script & "  param([string]$Message)" & vbCrLf
    script = script & "  $timestamp = Get-Date -Format 'yyyy/MM/dd HH:mm:ss'" & vbCrLf
    script = script & "  Write-Host ""[$timestamp] $Message""" & vbCrLf
    script = script & "}" & vbCrLf
    script = script & vbCrLf

    ' ローカルモードとリモートモードで処理を分岐
    If config("ExecMode") = "ローカル" Then
        script = script & BuildLocalGroupListScript(config)
    Else
        script = script & BuildRemoteGroupListScript(config)
    End If

    BuildGetGroupListScript = script
End Function

Private Function BuildLocalGroupListScript(config As Object) As String
    Dim script As String

    script = "try {" & vbCrLf
    script = script & "  Write-Log '--------------------------------------------------------------------------------'" & vbCrLf
    script = script & "  Write-Log 'グループ一覧取得開始'" & vbCrLf
    script = script & "  Write-Log '--------------------------------------------------------------------------------'" & vbCrLf
    script = script & vbCrLf
    script = script & "  $ajsprintPath = $null" & vbCrLf
    script = script & "  $searchPaths = @(" & vbCrLf
    script = script & "    'C:\Program Files\HITACHI\JP1AJS3\bin\ajsprint.exe'," & vbCrLf
    script = script & "    'C:\Program Files (x86)\HITACHI\JP1AJS3\bin\ajsprint.exe'," & vbCrLf
    script = script & "    'C:\Program Files\Hitachi\JP1AJS2\bin\ajsprint.exe'," & vbCrLf
    script = script & "    'C:\Program Files (x86)\Hitachi\JP1AJS2\bin\ajsprint.exe'" & vbCrLf
    script = script & "  )" & vbCrLf
    script = script & "  foreach ($path in $searchPaths) {" & vbCrLf
    script = script & "    if (Test-Path $path) { $ajsprintPath = $path; break }" & vbCrLf
    script = script & "  }" & vbCrLf
    script = script & "  if (-not $ajsprintPath) {" & vbCrLf
    script = script & "    Write-Log '[ERROR] JP1コマンド(ajsprint.exe)が見つかりません'" & vbCrLf
    script = script & "    Write-Output 'ERROR: JP1コマンド(ajsprint.exe)が見つかりません。'" & vbCrLf
    script = script & "    exit 1" & vbCrLf
    script = script & "  }" & vbCrLf
    script = script & "  Write-Log ""JP1コマンドパス: $ajsprintPath""" & vbCrLf
    script = script & vbCrLf
    script = script & "  Write-Log '[実行] ajsprint - グループ一覧取得'" & vbCrLf
    script = script & "  Write-Log ""コマンド: ajsprint.exe -F " & config("SchedulerService") & " /* -R""" & vbCrLf
    script = script & "  $result = & $ajsprintPath -F " & config("SchedulerService") & " '/*' -R 2>&1" & vbCrLf
    script = script & "  Write-Log '[成功] グループ一覧取得完了'" & vbCrLf
    script = script & "  $result | ForEach-Object { Write-Output $_ }" & vbCrLf
    script = script & "} catch {" & vbCrLf
    script = script & "  Write-Log ""[ERROR] $($_.Exception.Message)""" & vbCrLf
    script = script & "  Write-Output ""ERROR: $($_.Exception.Message)""" & vbCrLf
    script = script & "}" & vbCrLf

    BuildLocalGroupListScript = script
End Function

Private Function BuildRemoteGroupListScript(config As Object) As String
    Dim script As String

    script = "Write-Log '--------------------------------------------------------------------------------'" & vbCrLf
    script = script & "Write-Log 'グループ一覧取得開始（リモート）'" & vbCrLf
    script = script & "Write-Log ""接続先: " & config("JP1Server") & """" & vbCrLf
    script = script & "Write-Log '--------------------------------------------------------------------------------'" & vbCrLf
    script = script & vbCrLf
    script = script & "$securePass = ConvertTo-SecureString '" & EscapePSString(config("RemotePassword")) & "' -AsPlainText -Force" & vbCrLf
    script = script & "$cred = New-Object System.Management.Automation.PSCredential('" & config("RemoteUser") & "', $securePass)" & vbCrLf
    script = script & vbCrLf
    script = script & "$originalTrustedHosts = $null" & vbCrLf
    script = script & "$winrmConfigChanged = $false" & vbCrLf
    script = script & "$winrmServiceWasStarted = $false" & vbCrLf
    script = script & vbCrLf
    script = script & "try {" & vbCrLf
    script = script & "  Write-Log '[実行] WinRMサービス確認'" & vbCrLf
    script = script & "  $winrmService = Get-Service -Name WinRM -ErrorAction SilentlyContinue" & vbCrLf
    script = script & "  if ($winrmService.Status -ne 'Running') {" & vbCrLf
    script = script & "    Start-Service -Name WinRM -ErrorAction Stop" & vbCrLf
    script = script & "    $winrmServiceWasStarted = $true" & vbCrLf
    script = script & "  }" & vbCrLf
    script = script & "  $originalTrustedHosts = (Get-Item WSMan:\localhost\Client\TrustedHosts -ErrorAction SilentlyContinue).Value" & vbCrLf
    script = script & "  if ($originalTrustedHosts -notmatch '" & config("JP1Server") & "') {" & vbCrLf
    script = script & "    if ($originalTrustedHosts) {" & vbCrLf
    script = script & "      Set-Item WSMan:\localhost\Client\TrustedHosts -Value ""$originalTrustedHosts," & config("JP1Server") & """ -Force" & vbCrLf
    script = script & "    } else {" & vbCrLf
    script = script & "      Set-Item WSMan:\localhost\Client\TrustedHosts -Value '" & config("JP1Server") & "' -Force" & vbCrLf
    script = script & "    }" & vbCrLf
    script = script & "    $winrmConfigChanged = $true" & vbCrLf
    script = script & "  }" & vbCrLf
    script = script & "  Write-Log '[実行] リモートセッション作成'" & vbCrLf
    script = script & "  $session = New-PSSession -ComputerName '" & config("JP1Server") & "' -Credential $cred -ErrorAction Stop" & vbCrLf
    script = script & "  Write-Log '[成功] リモートセッション作成完了'" & vbCrLf
    script = script & vbCrLf
    script = script & "  Write-Log '[実行] ajsprint - グループ一覧取得'" & vbCrLf
    script = script & "  Write-Log ""コマンド: ajsprint.exe -F " & config("SchedulerService") & " /* -R""" & vbCrLf
    script = script & "  $result = Invoke-Command -Session $session -ScriptBlock {" & vbCrLf
    script = script & "    param($schedulerService)" & vbCrLf
    script = script & "    $ajsprintPath = $null" & vbCrLf
    script = script & "    $searchPaths = @('C:\Program Files\HITACHI\JP1AJS3\bin\ajsprint.exe','C:\Program Files (x86)\HITACHI\JP1AJS3\bin\ajsprint.exe','C:\Program Files\Hitachi\JP1AJS2\bin\ajsprint.exe','C:\Program Files (x86)\Hitachi\JP1AJS2\bin\ajsprint.exe')" & vbCrLf
    script = script & "    foreach ($p in $searchPaths) { if (Test-Path $p) { $ajsprintPath = $p; break } }" & vbCrLf
    script = script & "    if (-not $ajsprintPath) { Write-Output 'ERROR: ajsprint.exe not found'; return }" & vbCrLf
    script = script & "    $output = & $ajsprintPath '-F' $schedulerService '/*' '-R' 2>&1" & vbCrLf
    script = script & "    $output | Where-Object { $_ -notmatch '^KAVS\d+-I' }" & vbCrLf
    script = script & "  } -ArgumentList '" & config("SchedulerService") & "'" & vbCrLf
    script = script & "  Write-Log '[成功] グループ一覧取得完了'" & vbCrLf
    script = script & "  Remove-PSSession $session" & vbCrLf
    script = script & "  $result | ForEach-Object { Write-Output $_ }" & vbCrLf
    script = script & "} catch {" & vbCrLf
    script = script & "  Write-Log ""[ERROR] $($_.Exception.Message)""" & vbCrLf
    script = script & "  Write-Output ""ERROR: $($_.Exception.Message)""" & vbCrLf
    script = script & "} finally {" & vbCrLf
    script = script & "  if ($winrmConfigChanged) {" & vbCrLf
    script = script & "    if ($originalTrustedHosts) {" & vbCrLf
    script = script & "      Set-Item WSMan:\localhost\Client\TrustedHosts -Value $originalTrustedHosts -Force -ErrorAction SilentlyContinue" & vbCrLf
    script = script & "    } else {" & vbCrLf
    script = script & "      Clear-Item WSMan:\localhost\Client\TrustedHosts -Force -ErrorAction SilentlyContinue" & vbCrLf
    script = script & "    }" & vbCrLf
    script = script & "  }" & vbCrLf
    script = script & "  if ($winrmServiceWasStarted) {" & vbCrLf
    script = script & "    Stop-Service -Name WinRM -Force -ErrorAction SilentlyContinue" & vbCrLf
    script = script & "  }" & vbCrLf
    script = script & "}" & vbCrLf

    BuildRemoteGroupListScript = script
End Function

' ============================================================================
' PowerShell実行
' ============================================================================
Public Function ExecutePowerShell(script As String) As String
    ' 一時ファイルにスクリプトを保存（UTF-8 BOMなしで保存）
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim tempFolder As String
    tempFolder = fso.GetSpecialFolder(2) ' Temp folder

    Dim timestamp As String
    timestamp = Format(Now, "yyyymmddhhnnss") & "_" & Int(Rnd * 10000)

    Dim scriptPath As String
    scriptPath = tempFolder & "\jp1_temp_" & timestamp & ".ps1"

    Dim outputPath As String
    outputPath = tempFolder & "\jp1_output_" & timestamp & ".txt"

    ' スクリプトをラップして結果をファイルに出力
    Dim wrappedScript As String
    wrappedScript = script & vbCrLf
    wrappedScript = wrappedScript & "# 出力完了マーカー" & vbCrLf

    ' ADODB.Streamを使用してUTF-8（BOM付き）で保存
    Dim utfStream As Object
    Set utfStream = CreateObject("ADODB.Stream")
    utfStream.Type = 2 ' adTypeText
    utfStream.Charset = "UTF-8"
    utfStream.Open
    utfStream.WriteText wrappedScript
    utfStream.SaveToFile scriptPath, 2 ' adSaveCreateOverWrite（BOM付きで保存）
    utfStream.Close
    Set utfStream = Nothing

    ' PowerShell実行（リアルタイム表示・結果をファイルに出力）
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")

    Dim cmd As String
    cmd = "powershell -NoProfile -ExecutionPolicy Bypass -Command ""& {" & _
          "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; " & _
          "Start-Transcript -Path '" & outputPath & "' -Force | Out-Null; " & _
          "try { & '" & scriptPath & "' } finally { Stop-Transcript | Out-Null }" & _
          "}"""

    ' 1 = vbNormalFocus（通常表示）、True で完了まで待機
    shell.Run cmd, 1, True

    ' 結果ファイルを読み込む
    Dim output As String
    output = ""

    If fso.FileExists(outputPath) Then
        ' UTF-8で読み込み
        Set utfStream = CreateObject("ADODB.Stream")
        utfStream.Type = 2 ' adTypeText
        utfStream.Charset = "UTF-8"
        utfStream.Open
        utfStream.LoadFromFile outputPath

        If Not utfStream.EOS Then
            output = utfStream.ReadText
        End If

        utfStream.Close
        Set utfStream = Nothing

        ' 出力ファイル削除
        On Error Resume Next
        fso.DeleteFile outputPath
        On Error GoTo 0
    End If

    ' スクリプトファイル削除
    On Error Resume Next
    fso.DeleteFile scriptPath
    On Error GoTo 0

    ExecutePowerShell = output
End Function

' ============================================================================
' ジョブ一覧シートの実行結果更新
' ============================================================================
Public Sub UpdateJobListStatus(ByVal row As Long, ByVal result As Object)
    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_JOBLIST)

    ws.Cells(row, COL_LAST_STATUS).Value = result("Status")
    ws.Cells(row, COL_LAST_EXEC_TIME).Value = result("StartTime")
    ws.Cells(row, COL_LAST_END_TIME).Value = result("EndTime")

    ' ログパスを記録（N列にハイパーリンク設定）
    If result("LogPath") <> "" Then
        Dim logPath As String
        logPath = result("LogPath")
        ws.Cells(row, COL_LAST_MESSAGE).Value = logPath
        On Error Resume Next
        ws.Hyperlinks.Add Anchor:=ws.Cells(row, COL_LAST_MESSAGE), _
                          Address:=logPath, _
                          TextToDisplay:=logPath
        On Error GoTo 0
    End If

    ' 保留解除された場合（成功時）、保留列をクリアしてハイライトを解除
    If result("Status") = "正常終了" Or result("Status") = "起動成功" Then
        If ws.Cells(row, COL_HOLD).Value = "保留中" Then
            ws.Cells(row, COL_HOLD).Value = ""
            ws.Cells(row, COL_HOLD).Font.Bold = False
            ws.Cells(row, COL_HOLD).Font.Color = RGB(0, 0, 0)
            ws.Range(ws.Cells(row, COL_SELECT), ws.Cells(row, COL_LAST_MESSAGE)).Interior.ColorIndex = xlNone
        End If
    End If

    ' 色付け
    If result("Status") = "正常終了" Then
        ws.Cells(row, COL_LAST_STATUS).Interior.Color = RGB(198, 239, 206)  ' 緑（正常）
    ElseIf result("Status") = "起動成功" Then
        ws.Cells(row, COL_LAST_STATUS).Interior.Color = RGB(255, 235, 156)  ' 黄（起動のみ）
    ElseIf result("Status") = "警告検出終了" Then
        ws.Cells(row, COL_LAST_STATUS).Interior.Color = RGB(255, 192, 0)    ' オレンジ（警告）
    Else
        ws.Cells(row, COL_LAST_STATUS).Interior.Color = RGB(255, 199, 206)  ' 赤（異常）
    End If
End Sub

