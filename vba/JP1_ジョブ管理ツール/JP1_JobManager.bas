Option Explicit

'==============================================================================
' JP1 ジョブ管理ツール - メインモジュール
'   - ジョブネット一覧取得（ajsprint経由）
'   - ジョブ実行処理
'
' 注意: 初期化処理は JP1_JobManager_Setup.bas にあります
'       定数はSetupモジュールで Public として定義されています
'==============================================================================

' 管理者権限状態を保持
Private g_AdminChecked As Boolean
Private g_IsAdmin As Boolean

' 現在の実行セッションのログファイルパス
Private g_LogFilePath As String

'==============================================================================
' ジョブ一覧取得
'==============================================================================
Public Sub GetJobList()
    On Error GoTo ErrorHandler

    Dim config As Object
    Set config = GetConfig()

    If config Is Nothing Then Exit Sub

    ' リモートモードの場合、管理者権限をチェック
    If Not EnsureAdminForRemoteMode(config) Then Exit Sub

    ' パスワード入力（リモートモードの場合のみリモートパスワードが必要）
    If config("ExecMode") <> "ローカル" Then
        If config("RemotePassword") = "" Then
            config("RemotePassword") = InputBox("リモートサーバのパスワードを入力してください:", "パスワード入力")
            If config("RemotePassword") = "" Then
                MsgBox "パスワードが入力されませんでした。", vbExclamation
                Exit Sub
            End If
        End If
    End If

    If config("JP1Password") = "" Then
        config("JP1Password") = InputBox("JP1パスワードを入力してください:", "パスワード入力")
        If config("JP1Password") = "" Then
            MsgBox "パスワードが入力されませんでした。", vbExclamation
            Exit Sub
        End If
    End If

    Application.ScreenUpdating = False
    Application.StatusBar = "ジョブ一覧を取得中..."

    ' PowerShellスクリプト生成・実行
    Dim psScript As String
    psScript = BuildGetJobListScript(config)

    Dim result As String
    result = ExecutePowerShell(psScript)

    ' 結果をパース（戻り値で成功/失敗を判定）
    ' コマンド実行時は -F オプションでスケジューラサービスを指定するため
    ' パスにはスケジューラサービス名を含めない（ルートパス + "/" + ユニット名）
    Dim parseSuccess As Boolean
    parseSuccess = ParseJobListResult(result, config("RootPath"))

    Application.StatusBar = False
    Application.ScreenUpdating = True

    ' エラーの場合は完了メッセージを表示しない
    If Not parseSuccess Then
        Exit Sub
    End If

    MsgBox "ジョブ一覧の取得が完了しました。" & vbCrLf & _
           "ジョブ一覧シートを確認してください。", vbInformation

    Worksheets(SHEET_JOBLIST).Activate
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました。" & vbCrLf & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description & vbCrLf & _
           "発生場所: GetJobList", vbCritical, "VBAエラー"
End Sub

Private Function BuildGetJobListScript(config As Object) As String
    Dim script As String

    script = "$ErrorActionPreference = 'Stop'" & vbCrLf
    script = script & vbCrLf

    ' UTF-8エンコーディング設定（日本語パス対応）
    script = script & "# UTF-8エンコーディング設定" & vbCrLf
    script = script & "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8" & vbCrLf
    script = script & "$OutputEncoding = [System.Text.Encoding]::UTF8" & vbCrLf
    script = script & "chcp 65001 | Out-Null" & vbCrLf
    script = script & vbCrLf

    ' ローカルモードとリモートモードで処理を分岐
    If config("ExecMode") = "ローカル" Then
        ' ローカル実行モード（WinRM不使用）
        script = script & "try {" & vbCrLf
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
        script = script & "    Write-Output ""ERROR: JP1コマンド(ajsprint.exe)が見つかりません。JP1/AJS3 Managerがインストールされているか確認してください。""" & vbCrLf
        script = script & "    exit 1" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & vbCrLf
        script = script & "  # ローカルでajsprintを実行" & vbCrLf
        script = script & "  $result = & $ajsprintPath -F " & config("SchedulerService") & " '" & config("RootPath") & "' -R 2>&1" & vbCrLf
        script = script & "  $result | ForEach-Object { Write-Output $_ }" & vbCrLf
        script = script & "} catch {" & vbCrLf
        script = script & "  Write-Output ""ERROR: $($_.Exception.Message)""" & vbCrLf
        script = script & "}" & vbCrLf
    Else
        ' リモート実行モード（WinRM使用）
        ' 認証情報
        script = script & "$securePass = ConvertTo-SecureString '" & EscapePSString(config("RemotePassword")) & "' -AsPlainText -Force" & vbCrLf
        script = script & "$cred = New-Object System.Management.Automation.PSCredential('" & config("RemoteUser") & "', $securePass)" & vbCrLf
        script = script & vbCrLf

        ' WinRM設定の保存と自動設定
        script = script & "$originalTrustedHosts = $null" & vbCrLf
        script = script & "$winrmConfigChanged = $false" & vbCrLf
        script = script & "$winrmServiceWasStarted = $false" & vbCrLf
        script = script & vbCrLf

        script = script & "try {" & vbCrLf
        script = script & "  # WinRMサービスの起動確認（TrustedHosts取得前に起動が必要）" & vbCrLf
        script = script & "  $winrmService = Get-Service -Name WinRM -ErrorAction SilentlyContinue" & vbCrLf
        script = script & "  if ($winrmService.Status -ne 'Running') {" & vbCrLf
        script = script & "    Start-Service -Name WinRM -ErrorAction Stop" & vbCrLf
        script = script & "    $winrmServiceWasStarted = $true" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & vbCrLf
        script = script & "  # 現在のTrustedHostsを取得（WinRMサービス起動後に取得）" & vbCrLf
        script = script & "  $originalTrustedHosts = (Get-Item WSMan:\localhost\Client\TrustedHosts -ErrorAction SilentlyContinue).Value" & vbCrLf
        script = script & vbCrLf
        script = script & "  # TrustedHostsに接続先を追加（必要な場合のみ）" & vbCrLf
        script = script & "  if ($originalTrustedHosts -notmatch '" & config("JP1Server") & "') {" & vbCrLf
        script = script & "    if ($originalTrustedHosts) {" & vbCrLf
        script = script & "      Set-Item WSMan:\localhost\Client\TrustedHosts -Value ""$originalTrustedHosts," & config("JP1Server") & """ -Force" & vbCrLf
        script = script & "    } else {" & vbCrLf
        script = script & "      Set-Item WSMan:\localhost\Client\TrustedHosts -Value '" & config("JP1Server") & "' -Force" & vbCrLf
        script = script & "    }" & vbCrLf
        script = script & "    $winrmConfigChanged = $true" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & vbCrLf

        ' リモート実行
        script = script & "  $session = New-PSSession -ComputerName '" & config("JP1Server") & "' -Credential $cred -ErrorAction Stop" & vbCrLf
        script = script & vbCrLf
        script = script & "  $result = Invoke-Command -Session $session -ScriptBlock {" & vbCrLf
        script = script & "    param($schedulerService, $rootPath)" & vbCrLf
        script = script & "    if ([string]::IsNullOrWhiteSpace($rootPath)) { Write-Output 'ERROR: rootPath is empty'; return }" & vbCrLf
        script = script & "    $ajsprintPath = $null" & vbCrLf
        script = script & "    $searchPaths = @('C:\Program Files\HITACHI\JP1AJS3\bin\ajsprint.exe','C:\Program Files (x86)\HITACHI\JP1AJS3\bin\ajsprint.exe','C:\Program Files\Hitachi\JP1AJS2\bin\ajsprint.exe','C:\Program Files (x86)\Hitachi\JP1AJS2\bin\ajsprint.exe')" & vbCrLf
        script = script & "    foreach ($p in $searchPaths) { if (Test-Path $p) { $ajsprintPath = $p; break } }" & vbCrLf
        script = script & "    if (-not $ajsprintPath) { Write-Output 'ERROR: ajsprint.exe not found'; return }" & vbCrLf
        script = script & "    $output = & $ajsprintPath '-F' $schedulerService $rootPath '-R' 2>&1" & vbCrLf
        script = script & "    # KAVS情報メッセージ（-I）を除外、unit=行のみ出力" & vbCrLf
        script = script & "    $output | Where-Object { $_ -notmatch '^KAVS\d+-I' }" & vbCrLf
        script = script & "  } -ArgumentList '" & config("SchedulerService") & "', '" & config("RootPath") & "'" & vbCrLf
        script = script & vbCrLf
        script = script & "  Remove-PSSession $session" & vbCrLf
        script = script & vbCrLf
        script = script & "  $result | ForEach-Object { Write-Output $_ }" & vbCrLf
        script = script & "} catch {" & vbCrLf
        script = script & "  Write-Output ""ERROR: $($_.Exception.Message)""" & vbCrLf
        script = script & "} finally {" & vbCrLf
        script = script & "  # WinRM設定の復元" & vbCrLf
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
    End If

    BuildGetJobListScript = script
End Function

Private Function ParseJobListResult(result As String, rootPath As String) As Boolean
    ' 戻り値: True=成功, False=エラー
    ' JP1 ajsprint出力形式（ネスト対応）:
    '   unit=ユニット名,,admin,グループ;    ← 2番目のフィールドは空
    '   {
    '       ty=n;
    '       cm="コメント";
    '       unit=子ユニット名,,admin,グループ;  ← ネストされたユニット
    '       {
    '           ty=n;
    '           ...
    '       }
    '   }
    ' フルパス = ルートパス + "/" + ユニット名
    ' ※コマンド実行時は -F オプションでスケジューラサービスを指定するため
    '   パスにはスケジューラサービス名を含めない
    ParseJobListResult = False

    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_JOBLIST)

    ' 既存データをクリア
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_JOBNET_PATH).End(xlUp).Row
    If lastRow >= ROW_JOBLIST_DATA_START Then
        ws.Range(ws.Cells(ROW_JOBLIST_DATA_START, COL_ORDER), ws.Cells(lastRow, COL_LAST_MESSAGE)).Clear
    End If

    ' 結果をパース
    Dim lines() As String
    lines = Split(result, vbCrLf)

    Dim row As Long
    row = ROW_JOBLIST_DATA_START

    Dim i As Long

    ' ルートパスの末尾のスラッシュを正規化
    Dim basePath As String
    basePath = rootPath
    If Right(basePath, 1) = "/" Then
        basePath = Left(basePath, Len(basePath) - 1)
    End If

    ' ネスト対応のためスタック構造を使用
    ' 配列でスタックをシミュレート（最大ネスト深度10）
    Const MAX_DEPTH As Long = 10
    Dim unitStack(1 To MAX_DEPTH) As String   ' unit=...ヘッダーのスタック
    Dim blockStack(1 To MAX_DEPTH) As String  ' ブロック内容のスタック
    Dim pathStack(1 To MAX_DEPTH) As String   ' フルパスのスタック
    Dim stackDepth As Long                     ' 現在のスタック深度

    stackDepth = 0

    For i = LBound(lines) To UBound(lines)
        Dim line As String
        ' Trim()はスペースのみ除去するため、TAB文字も明示的に除去
        line = Trim(Replace(lines(i), vbTab, ""))

        ' 空行はスキップ
        If line = "" Then GoTo NextLine

        ' エラーチェック
        If InStr(line, "ERROR:") > 0 Then
            MsgBox "エラーが発生しました:" & vbCrLf & line, vbExclamation
            Exit Function
        End If

        ' unit= で始まる行（ヘッダー）- ブロック内外問わず検出
        If InStr(line, "unit=") > 0 Then
            ' 次の行が{かどうかを先読み
            Dim nextIdx As Long
            nextIdx = i + 1
            If nextIdx <= UBound(lines) Then
                Dim nextLine As String
                nextLine = Trim(Replace(lines(nextIdx), vbTab, ""))
                If Left(nextLine, 1) = "{" Then
                    ' 新しいユニット定義開始 - スタックにプッシュ
                    stackDepth = stackDepth + 1
                    If stackDepth <= MAX_DEPTH Then
                        unitStack(stackDepth) = line
                        blockStack(stackDepth) = ""

                        ' ユニット名を取得（unit=の最初のフィールド）
                        Dim unitName As String
                        unitName = ExtractUnitName(line)

                        ' フルパスを構築
                        If stackDepth = 1 Then
                            ' ルートレベル: ajsprintで指定したパスのユニット自体が
                            ' 最初に出力されるため、basePathをそのまま使用
                            pathStack(stackDepth) = basePath
                        Else
                            ' ネストレベル: 親のパス + "/" + ユニット名
                            pathStack(stackDepth) = pathStack(stackDepth - 1) & "/" & unitName
                        End If
                    End If
                End If
            End If
            GoTo NextLine
        End If

        ' ブロック開始 {
        If Left(line, 1) = "{" Then
            ' {の後に内容がある場合
            If Len(line) > 1 And stackDepth > 0 Then
                blockStack(stackDepth) = blockStack(stackDepth) & " " & Mid(line, 2)
            End If
            GoTo NextLine
        End If

        ' ブロック終了 }
        If Right(line, 1) = "}" Or line = "}" Then
            ' }の前に内容がある場合
            If Len(line) > 1 And stackDepth > 0 Then
                blockStack(stackDepth) = blockStack(stackDepth) & " " & Left(line, Len(line) - 1)
            End If

            ' スタックからポップして処理
            If stackDepth > 0 Then
                Dim currentHeader As String
                Dim currentBlock As String
                Dim currentFullPath As String
                currentHeader = unitStack(stackDepth)
                currentBlock = blockStack(stackDepth)
                currentFullPath = pathStack(stackDepth)

                ' ユニットタイプを抽出（ty=xxx; から xxx を取得）
                Dim unitType As String
                Dim unitTypeDisplay As String
                unitType = ExtractUnitType(currentBlock)
                unitTypeDisplay = GetUnitTypeDisplayName(unitType)

                ' ty=が存在する場合に一覧に追加
                If unitType <> "" And currentFullPath <> "" Then
                    ws.Cells(row, COL_ORDER).Value = ""
                    ' 種別を設定
                    ws.Cells(row, COL_UNIT_TYPE).Value = unitTypeDisplay
                    ws.Cells(row, COL_UNIT_TYPE).HorizontalAlignment = xlCenter
                    ' フルパス（ルートからのパス）を設定
                    ws.Cells(row, COL_JOBNET_PATH).Value = currentFullPath
                    ' ユニット名を設定（unit=の最初のフィールド）
                    ws.Cells(row, COL_JOBNET_NAME).Value = ExtractUnitName(currentHeader)
                    ws.Cells(row, COL_COMMENT).Value = ExtractCommentFromBlock(currentBlock)
                    ' スクリプトファイル名 (sc=)
                    ws.Cells(row, COL_SCRIPT).Value = ExtractAttributeFromBlock(currentBlock, "sc")
                    ' パラメーター (prm=)
                    ws.Cells(row, COL_PARAMETER).Value = ExtractAttributeFromBlock(currentBlock, "prm")
                    ' ワークパス (wkp=)
                    ws.Cells(row, COL_WORK_PATH).Value = ExtractAttributeFromBlock(currentBlock, "wkp")

                    ' 保留状態を解析
                    Dim isHold As Boolean
                    isHold = (InStr(currentBlock, "hd=h") > 0) Or (InStr(currentBlock, "hd=H") > 0)

                    If isHold Then
                        ws.Cells(row, COL_HOLD).Value = "保留中"
                        ws.Cells(row, COL_HOLD).HorizontalAlignment = xlCenter
                        ws.Range(ws.Cells(row, COL_ORDER), ws.Cells(row, COL_LAST_MESSAGE)).Interior.Color = RGB(255, 235, 156)
                        ws.Cells(row, COL_HOLD).Font.Bold = True
                        ws.Cells(row, COL_HOLD).Font.Color = RGB(156, 87, 0)
                    Else
                        ws.Cells(row, COL_HOLD).Value = ""
                    End If

                    ' 順序列の書式
                    With ws.Cells(row, COL_ORDER)
                        .HorizontalAlignment = xlCenter
                    End With

                    ' 罫線
                    ws.Range(ws.Cells(row, COL_ORDER), ws.Cells(row, COL_LAST_MESSAGE)).Borders.LineStyle = xlContinuous

                    row = row + 1
                End If

                ' スタックをクリアしてポップ
                unitStack(stackDepth) = ""
                blockStack(stackDepth) = ""
                pathStack(stackDepth) = ""
                stackDepth = stackDepth - 1
            End If
            GoTo NextLine
        End If

        ' ブロック内のコンテンツを収集（ただし、次のunit=行の前まで）
        If stackDepth > 0 Then
            ' 次の行がunit=かチェック
            Dim isNextUnit As Boolean
            isNextUnit = False
            If i + 1 <= UBound(lines) Then
                Dim checkLine As String
                checkLine = Trim(Replace(lines(i + 1), vbTab, ""))
                If InStr(checkLine, "unit=") > 0 Then
                    isNextUnit = True
                End If
            End If

            ' unit=の直前でなければブロック内容に追加
            If Not isNextUnit Then
                blockStack(stackDepth) = blockStack(stackDepth) & " " & line
            End If
        End If

NextLine:
    Next i

    ' データがない場合
    If row = ROW_JOBLIST_DATA_START Then
        MsgBox "ユニットが見つかりませんでした。" & vbCrLf & _
               "取得パスを確認してください。", vbExclamation
        Exit Function
    End If

    ' 成功
    ParseJobListResult = True
End Function

Private Function ExtractJobNameFromHeader(header As String) As String
    ' unit=パス,名前,admin,グループ; から名前を抽出
    Dim parts() As String
    Dim unitPart As String
    Dim startPos As Long

    startPos = InStr(header, "unit=")
    If startPos > 0 Then
        unitPart = Mid(header, startPos + 5)
        ' セミコロンを除去
        If Right(unitPart, 1) = ";" Then
            unitPart = Left(unitPart, Len(unitPart) - 1)
        End If
        ' カンマで分割
        parts = Split(unitPart, ",")
        If UBound(parts) >= 1 Then
            ExtractJobNameFromHeader = parts(1)
        End If
    End If
End Function

Private Function ExtractCommentFromBlock(blockContent As String) As String
    ' cm="コメント"; からコメントを抽出
    Dim startPos As Long
    Dim endPos As Long

    startPos = InStr(blockContent, "cm=""")
    If startPos > 0 Then
        startPos = startPos + 4
        endPos = InStr(startPos, blockContent, """")
        If endPos > startPos Then
            ExtractCommentFromBlock = Mid(blockContent, startPos, endPos - startPos)
        End If
    End If
End Function

Private Function ExtractAttributeFromBlock(blockContent As String, attrName As String) As String
    ' 指定された属性名の値を抽出
    ' 形式1: attr="value"; (ダブルクォート囲み)
    ' 形式2: attr=value; (クォートなし)
    Dim startPos As Long
    Dim endPos As Long
    Dim searchStr As String

    ExtractAttributeFromBlock = ""

    ' ダブルクォート形式を先にチェック: attr="value"
    searchStr = attrName & "="""
    startPos = InStr(blockContent, searchStr)
    If startPos > 0 Then
        startPos = startPos + Len(searchStr)
        endPos = InStr(startPos, blockContent, """")
        If endPos > startPos Then
            ExtractAttributeFromBlock = Mid(blockContent, startPos, endPos - startPos)
            Exit Function
        End If
    End If

    ' クォートなし形式: attr=value;
    searchStr = attrName & "="
    startPos = InStr(blockContent, searchStr)
    If startPos > 0 Then
        startPos = startPos + Len(searchStr)
        ' セミコロンまたはスペースまでを取得
        endPos = InStr(startPos, blockContent, ";")
        Dim endPosSpace As Long
        endPosSpace = InStr(startPos, blockContent, " ")

        If endPos > startPos Then
            If endPosSpace > startPos And endPosSpace < endPos Then
                endPos = endPosSpace
            End If
            ExtractAttributeFromBlock = Mid(blockContent, startPos, endPos - startPos)
        End If
    End If
End Function

Private Function ExtractUnitPath(line As String) As String
    ' unit=/path/to/jobnet から /path/to/jobnet を抽出
    ' 注: JP1のajsprintでは最初のフィールドがユニット名
    Dim startPos As Long
    Dim endPos As Long

    startPos = InStr(line, "unit=")
    If startPos > 0 Then
        startPos = startPos + 5
        endPos = InStr(startPos, line, ",")
        If endPos > startPos Then
            ExtractUnitPath = Mid(line, startPos, endPos - startPos)
        End If
    End If
End Function

Private Function ExtractUnitName(line As String) As String
    ' unit=ユニット名,,admin,group; からユニット名を抽出
    ' 最初のフィールド（カンマまで）を返す
    ExtractUnitName = ExtractUnitPath(line)
End Function

Private Function ExtractUnitType(blockContent As String) As String
    ' ty=xxx; から xxx を抽出
    ' 例: ty=n; → n, ty=pj; → pj, ty=jdj; → jdj
    Dim startPos As Long
    Dim endPos As Long
    Dim tyValue As String

    ExtractUnitType = ""

    startPos = InStr(blockContent, "ty=")
    If startPos > 0 Then
        startPos = startPos + 3
        ' セミコロンまたはスペースまでを取得
        endPos = InStr(startPos, blockContent, ";")
        Dim endPosSpace As Long
        endPosSpace = InStr(startPos, blockContent, " ")

        If endPos > startPos Then
            If endPosSpace > startPos And endPosSpace < endPos Then
                endPos = endPosSpace
            End If
            ExtractUnitType = Mid(blockContent, startPos, endPos - startPos)
        End If
    End If
End Function

Private Function GetUnitTypeDisplayName(unitType As String) As String
    ' ユニットタイプコードを日本語表示名に変換
    ' JP1/AJS3の全ユニット種別に対応
    Select Case LCase(unitType)
        ' グループ系
        Case "g"
            GetUnitTypeDisplayName = "グループ"
        Case "mg"
            GetUnitTypeDisplayName = "マネージャーグループ"

        ' ジョブネット系
        Case "n"
            GetUnitTypeDisplayName = "ジョブネット"
        Case "rn"
            GetUnitTypeDisplayName = "リカバリーネット"
        Case "rm"
            GetUnitTypeDisplayName = "リモートネット"
        Case "mn"
            GetUnitTypeDisplayName = "マネージャーネット"

        ' 標準ジョブ系
        Case "j"
            GetUnitTypeDisplayName = "ジョブ"
        Case "rj"
            GetUnitTypeDisplayName = "リカバリージョブ"
        Case "pj"
            GetUnitTypeDisplayName = "判定ジョブ"
        Case "rp"
            GetUnitTypeDisplayName = "リカバリー判定"
        Case "qj"
            GetUnitTypeDisplayName = "キュージョブ"
        Case "rq"
            GetUnitTypeDisplayName = "リカバリーキュー"

        ' 判定変数系
        Case "jdj"
            GetUnitTypeDisplayName = "判定変数参照"
        Case "rjdj"
            GetUnitTypeDisplayName = "リカバリー判定変数"
        Case "orj"
            GetUnitTypeDisplayName = "OR分岐"
        Case "rorj"
            GetUnitTypeDisplayName = "リカバリーOR分岐"

        ' イベント監視系
        Case "evwj"
            GetUnitTypeDisplayName = "イベント監視"
        Case "revwj"
            GetUnitTypeDisplayName = "リカバリーイベント"
        Case "flwj"
            GetUnitTypeDisplayName = "ファイル監視"
        Case "rflwj"
            GetUnitTypeDisplayName = "リカバリーファイル監視"
        Case "mlwj"
            GetUnitTypeDisplayName = "メール受信監視"
        Case "rmlwj"
            GetUnitTypeDisplayName = "リカバリーメール受信"
        Case "mqwj"
            GetUnitTypeDisplayName = "MQ受信監視"
        Case "rmqwj"
            GetUnitTypeDisplayName = "リカバリーMQ受信"
        Case "mswj"
            GetUnitTypeDisplayName = "MSMQメッセージ受信監視"
        Case "rmswj"
            GetUnitTypeDisplayName = "リカバリーMSMQ受信"
        Case "lfwj"
            GetUnitTypeDisplayName = "ログファイル監視"
        Case "rlfwj"
            GetUnitTypeDisplayName = "リカバリーログ監視"
        Case "ntwj"
            GetUnitTypeDisplayName = "Windows NT イベントログ監視"
        Case "rntwj"
            GetUnitTypeDisplayName = "リカバリー NT イベントログ"
        Case "tmwj"
            GetUnitTypeDisplayName = "実行間隔制御"
        Case "rtmwj"
            GetUnitTypeDisplayName = "リカバリー実行間隔"

        ' 送信系
        Case "evsj"
            GetUnitTypeDisplayName = "イベント送信"
        Case "revsj"
            GetUnitTypeDisplayName = "リカバリーイベント送信"
        Case "mlsj"
            GetUnitTypeDisplayName = "メール送信"
        Case "rmlsj"
            GetUnitTypeDisplayName = "リカバリーメール送信"
        Case "mqsj"
            GetUnitTypeDisplayName = "MQ送信"
        Case "rmqsj"
            GetUnitTypeDisplayName = "リカバリーMQ送信"
        Case "mssj"
            GetUnitTypeDisplayName = "MSMQメッセージ送信"
        Case "rmssj"
            GetUnitTypeDisplayName = "リカバリーMSMQ送信"
        Case "cmsj"
            GetUnitTypeDisplayName = "JP1イベント送信"
        Case "rcmsj"
            GetUnitTypeDisplayName = "リカバリーJP1送信"

        ' PowerShell系
        Case "pwlj"
            GetUnitTypeDisplayName = "ローカルPowerShell"
        Case "rpwlj"
            GetUnitTypeDisplayName = "リカバリーローカルPS"
        Case "pwrj"
            GetUnitTypeDisplayName = "リモートPowerShell"
        Case "rpwrj"
            GetUnitTypeDisplayName = "リカバリーリモートPS"

        ' カスタム系
        Case "cj"
            GetUnitTypeDisplayName = "カスタムジョブ"
        Case "rcj"
            GetUnitTypeDisplayName = "リカバリーカスタム"
        Case "cpj"
            GetUnitTypeDisplayName = "カスタムPCジョブ"
        Case "rcpj"
            GetUnitTypeDisplayName = "リカバリーカスタムPC"

        ' 外部連携系
        Case "fxj"
            GetUnitTypeDisplayName = "ファイル転送"
        Case "rfxj"
            GetUnitTypeDisplayName = "リカバリーファイル転送"
        Case "htpj"
            GetUnitTypeDisplayName = "HTTP接続"
        Case "rhtpj"
            GetUnitTypeDisplayName = "リカバリーHTTP"

        ' その他
        Case "nc"
            GetUnitTypeDisplayName = "ネットコネクタ"
        Case "hln"
            GetUnitTypeDisplayName = "リンク"
        Case "rc"
            GetUnitTypeDisplayName = "リリースコネクタ"
        Case "rr"
            GetUnitTypeDisplayName = "ルートジョブネット起動条件"

        ' 未知のタイプはそのまま表示
        Case Else
            If unitType <> "" Then
                GetUnitTypeDisplayName = unitType
            Else
                GetUnitTypeDisplayName = ""
            End If
    End Select
End Function

Private Function ExtractJobName(line As String) As String
    ' unit=/path/to/jobnet,ジョブ名,ty=n から ジョブ名 を抽出
    ' ajsprintの出力形式: unit=/path,name,ty=type,cm="comment";
    Dim startPos As Long
    Dim endPos As Long
    Dim fields() As String
    Dim unitPart As String

    ' unit= の後ろを取得
    startPos = InStr(line, "unit=")
    If startPos > 0 Then
        unitPart = Mid(line, startPos + 5)
        ' セミコロンまでを取得
        endPos = InStr(unitPart, ";")
        If endPos > 0 Then
            unitPart = Left(unitPart, endPos - 1)
        End If

        ' カンマで分割
        fields = Split(unitPart, ",")

        ' 2番目のフィールドがジョブ名（ty=で始まらない場合）
        If UBound(fields) >= 1 Then
            If InStr(fields(1), "ty=") = 0 And InStr(fields(1), "cm=") = 0 Then
                ExtractJobName = Trim(fields(1))
                Exit Function
            End If
        End If

        ' 2番目がty=の場合はパスの最後の部分を使用
        If UBound(fields) >= 0 Then
            ExtractJobName = GetLastPathComponent(fields(0))
        End If
    End If
End Function

Private Function ExtractComment(line As String) As String
    ' cm="comment" からコメントを抽出
    Dim startPos As Long
    Dim endPos As Long

    startPos = InStr(line, "cm=""")
    If startPos > 0 Then
        startPos = startPos + 4
        endPos = InStr(startPos, line, """")
        If endPos > startPos Then
            ExtractComment = Mid(line, startPos, endPos - startPos)
        End If
    End If
End Function

Private Function ExtractHoldStatus(line As String) As Boolean
    ' hd=y（保留）を検出
    ' JP1のajsprint出力で hd=y はホールド(保留)を示す
    ExtractHoldStatus = (InStr(line, ",hd=y") > 0 Or InStr(line, " hd=y") > 0)
End Function

Private Function GetLastPathComponent(path As String) As String
    Dim parts() As String
    parts = Split(path, "/")
    If UBound(parts) >= 0 Then
        GetLastPathComponent = parts(UBound(parts))
    End If
End Function

'==============================================================================
' 選択ジョブ実行
'==============================================================================
Public Sub ExecuteCheckedJobs()
    On Error GoTo ErrorHandler

    Dim config As Object
    Set config = GetConfig()

    If config Is Nothing Then Exit Sub

    ' リモートモードの場合、管理者権限をチェック
    If Not EnsureAdminForRemoteMode(config) Then Exit Sub

    ' パスワード入力（リモートモードの場合のみリモートパスワードが必要）
    If config("ExecMode") <> "ローカル" Then
        If config("RemotePassword") = "" Then
            config("RemotePassword") = InputBox("リモートサーバのパスワードを入力してください:", "パスワード入力")
            If config("RemotePassword") = "" Then
                MsgBox "パスワードが入力されませんでした。", vbExclamation
                Exit Sub
            End If
        End If
    End If

    If config("JP1Password") = "" Then
        config("JP1Password") = InputBox("JP1パスワードを入力してください:", "パスワード入力")
        If config("JP1Password") = "" Then
            MsgBox "パスワードが入力されませんでした。", vbExclamation
            Exit Sub
        End If
    End If

    ' 順序が指定されたジョブを取得
    Dim jobs As Collection
    Set jobs = GetOrderedJobs()

    If jobs.Count = 0 Then
        MsgBox "実行するジョブが選択されていません。" & vbCrLf & _
               "ジョブ一覧シートの「順序」列に数字（1, 2, 3...）を入力してください。", vbExclamation
        Exit Sub
    End If

    ' 保留中のジョブ数をカウント
    Dim holdCount As Long
    holdCount = 0
    Dim j As Variant
    For Each j In jobs
        If j("IsHold") Then holdCount = holdCount + 1
    Next j

    ' 確認
    Dim msg As String
    msg = "以下の " & jobs.Count & " 件のジョブを実行します：" & vbCrLf & vbCrLf
    Dim cnt As Long
    cnt = 0
    For Each j In jobs
        cnt = cnt + 1
        If cnt <= 5 Then
            Dim holdMark As String
            If j("IsHold") Then
                holdMark = " [保留中]"
            Else
                holdMark = ""
            End If
            msg = msg & cnt & ". " & j("Path") & holdMark & vbCrLf
        ElseIf cnt = 6 Then
            msg = msg & "..." & vbCrLf
        End If
    Next j

    If holdCount > 0 Then
        msg = msg & vbCrLf & "※ 保留中のジョブが " & holdCount & " 件あります。自動で保留解除してから実行します。" & vbCrLf
    End If
    msg = msg & vbCrLf & "実行しますか？"

    If MsgBox(msg, vbYesNo + vbQuestion, "実行確認") = vbNo Then Exit Sub

    ' ログファイルの初期化
    g_LogFilePath = CreateLogFile()

    ' 実行
    Application.ScreenUpdating = False

    Dim wsLog As Worksheet
    Set wsLog = Worksheets(SHEET_LOG)
    Dim logRow As Long
    logRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1
    If logRow < 4 Then logRow = 4

    Dim success As Boolean
    success = True

    For Each j In jobs
        Application.StatusBar = "実行中: " & j("Path")

        Dim execResult As Object
        Set execResult = ExecuteSingleJob(config, j("Path"), j("IsHold"), g_LogFilePath)

        ' 結果をログに記録
        wsLog.Cells(logRow, 1).Value = Now
        wsLog.Cells(logRow, 2).Value = j("Path")
        wsLog.Cells(logRow, 3).Value = execResult("Status")
        wsLog.Cells(logRow, 4).Value = execResult("StartTime")
        wsLog.Cells(logRow, 5).Value = execResult("EndTime")
        wsLog.Cells(logRow, 6).Value = execResult("Message")

        ' 色付け
        If execResult("Status") = "正常終了" Then
            wsLog.Cells(logRow, 3).Interior.Color = RGB(198, 239, 206)
        Else
            wsLog.Cells(logRow, 3).Interior.Color = RGB(255, 199, 206)
        End If

        wsLog.Range(wsLog.Cells(logRow, 1), wsLog.Cells(logRow, 6)).Borders.LineStyle = xlContinuous

        ' ジョブ一覧シートも更新
        UpdateJobListStatus j("Row"), execResult

        logRow = logRow + 1

        ' エラー時は停止
        If execResult("Status") <> "正常終了" And execResult("Status") <> "起動成功" Then
            success = False
            MsgBox "ジョブ「" & j("Path") & "」が失敗しました。" & vbCrLf & _
                   "処理を中断します。" & vbCrLf & vbCrLf & _
                   "詳細: " & execResult("Message") & vbCrLf & vbCrLf & _
                   "実行ログ: " & g_LogFilePath, vbCritical
            Exit For
        End If
    Next j

    Application.StatusBar = False
    Application.ScreenUpdating = True

    If success Then
        MsgBox "すべてのジョブが正常に完了しました。" & vbCrLf & vbCrLf & _
               "実行ログ: " & g_LogFilePath, vbInformation
    End If

    Worksheets(SHEET_LOG).Activate
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました。" & vbCrLf & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description & vbCrLf & _
           "発生場所: ExecuteCheckedJobs", vbCritical, "VBAエラー"
End Sub

Private Function GetOrderedJobs() As Collection
    Dim jobs As New Collection
    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_JOBLIST)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_JOBNET_PATH).End(xlUp).Row

    ' 順序が入力されている行を収集
    Dim orderedRows As New Collection
    Dim row As Long
    For row = ROW_JOBLIST_DATA_START To lastRow
        Dim orderValue As Variant
        orderValue = ws.Cells(row, COL_ORDER).Value

        ' 順序列に数字が入っている場合のみ対象
        If IsNumeric(orderValue) And orderValue <> "" Then
            Dim job As Object
            Set job = CreateObject("Scripting.Dictionary")
            job("Row") = row
            job("Path") = ws.Cells(row, COL_JOBNET_PATH).Value
            job("Order") = CLng(orderValue)
            ' 保留状態を取得
            job("IsHold") = (ws.Cells(row, COL_HOLD).Value = "保留中")

            orderedRows.Add job
        End If
    Next row

    ' 実行順でソート（単純なバブルソート）
    If orderedRows.Count = 0 Then
        Set GetOrderedJobs = jobs
        Exit Function
    End If

    Dim arr() As Variant
    ReDim arr(1 To orderedRows.Count)
    Dim i As Long
    Dim k As Long  ' ソート用ループ変数
    For i = 1 To orderedRows.Count
        Set arr(i) = orderedRows(i)
    Next i

    Dim temp As Object
    For i = 1 To UBound(arr) - 1
        For k = i + 1 To UBound(arr)
            If arr(i)("Order") > arr(k)("Order") Then
                Set temp = arr(i)
                Set arr(i) = arr(k)
                Set arr(k) = temp
            End If
        Next k
    Next i

    For i = 1 To UBound(arr)
        jobs.Add arr(i)
    Next i

    Set GetOrderedJobs = jobs
End Function

Private Function ExecuteSingleJob(ByVal config As Object, ByVal jobnetPath As String, ByVal isHold As Boolean, ByVal logFilePath As String) As Object
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

Private Function BuildExecuteJobScript(ByVal config As Object, ByVal jobnetPath As String, ByVal waitCompletion As Boolean, ByVal isHold As Boolean, ByVal logFilePath As String) As String
    Dim script As String

    script = "$ErrorActionPreference = 'Stop'" & vbCrLf
    script = script & vbCrLf

    ' UTF-8エンコーディング設定（日本語パス対応）
    script = script & "# UTF-8エンコーディング設定" & vbCrLf
    script = script & "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8" & vbCrLf
    script = script & "$OutputEncoding = [System.Text.Encoding]::UTF8" & vbCrLf
    script = script & "chcp 65001 | Out-Null" & vbCrLf
    script = script & vbCrLf

    ' ログ出力関数を定義
    script = script & "# ログ出力関数（ファイルとコンソール両方に出力）" & vbCrLf
    script = script & "$logFile = '" & Replace(logFilePath, "'", "''") & "'" & vbCrLf
    script = script & "function Write-Log {" & vbCrLf
    script = script & "  param([string]$Message)" & vbCrLf
    script = script & "  $timestamp = Get-Date -Format 'yyyy/MM/dd HH:mm:ss'" & vbCrLf
    script = script & "  $logLine = ""[$timestamp] $Message""" & vbCrLf
    script = script & "  Write-Host $logLine" & vbCrLf
    script = script & "  Add-Content -Path $logFile -Value $logLine -Encoding UTF8" & vbCrLf
    script = script & "}" & vbCrLf
    script = script & vbCrLf

    ' ローカルモードとリモートモードで処理を分岐
    If config("ExecMode") = "ローカル" Then
        ' ローカル実行モード（WinRM不使用）
        script = script & "try {" & vbCrLf
        script = script & "  Write-Log '--------------------------------------------------------------------------------'" & vbCrLf
        script = script & "  Write-Log 'ジョブネット: " & jobnetPath & "'" & vbCrLf
        script = script & "  Write-Log '--------------------------------------------------------------------------------'" & vbCrLf
        script = script & vbCrLf
        script = script & "  # JP1コマンドパスの検出" & vbCrLf
        script = script & "  $jp1BinPath = $null" & vbCrLf
        script = script & "  $searchPaths = @('C:\Program Files\HITACHI\JP1AJS3\bin','C:\Program Files (x86)\HITACHI\JP1AJS3\bin','C:\Program Files\Hitachi\JP1AJS2\bin','C:\Program Files (x86)\Hitachi\JP1AJS2\bin')" & vbCrLf
        script = script & "  foreach ($path in $searchPaths) {" & vbCrLf
        script = script & "    if (Test-Path ""$path\ajsentry.exe"") { $jp1BinPath = $path; break }" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & "  if (-not $jp1BinPath) {" & vbCrLf
        script = script & "    Write-Log '[ERROR] JP1コマンドが見つかりません'" & vbCrLf
        script = script & "    Write-Output ""ERROR: JP1コマンドが見つかりません。JP1/AJS3 Managerがインストールされているか確認してください。""" & vbCrLf
        script = script & "    exit 1" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & "  Write-Log ""JP1コマンドパス: $jp1BinPath""" & vbCrLf
        script = script & vbCrLf

        ' 保留解除処理（ローカル）
        If isHold Then
            script = script & "  # 保留解除" & vbCrLf
            script = script & "  Write-Log '[実行] ajsrelease - 保留解除'" & vbCrLf
            script = script & "  Write-Log ""コマンド: ajsrelease.exe -F " & config("SchedulerService") & " " & jobnetPath & """" & vbCrLf
            script = script & "  $releaseOutput = & ""$jp1BinPath\ajsrelease.exe"" -F " & config("SchedulerService") & " '" & jobnetPath & "' 2>&1" & vbCrLf
            script = script & "  Write-Log ""結果: $($releaseOutput -join ' ')""" & vbCrLf
            script = script & "  if ($LASTEXITCODE -ne 0) {" & vbCrLf
            script = script & "    Write-Log '[ERROR] 保留解除失敗'" & vbCrLf
            script = script & "    Write-Output ""RESULT_STATUS:保留解除失敗""" & vbCrLf
            script = script & "    Write-Output ""RESULT_MESSAGE:$($releaseOutput -join ' ')""" & vbCrLf
            script = script & "    exit" & vbCrLf
            script = script & "  }" & vbCrLf
            script = script & "  Write-Log '[成功] 保留解除完了'" & vbCrLf
            script = script & vbCrLf
        End If

        script = script & "  # ajsentry実行" & vbCrLf
        script = script & "  Write-Log '[実行] ajsentry - ジョブ起動'" & vbCrLf
        script = script & "  Write-Log ""コマンド: ajsentry.exe -F " & config("SchedulerService") & " " & jobnetPath & """" & vbCrLf
        script = script & "  $output = & ""$jp1BinPath\ajsentry.exe"" -F " & config("SchedulerService") & " '" & jobnetPath & "' 2>&1" & vbCrLf
        script = script & "  Write-Log ""結果: $($output -join ' ')""" & vbCrLf
        script = script & "  $exitCode = $LASTEXITCODE" & vbCrLf
        script = script & vbCrLf
        script = script & "  if ($exitCode -ne 0) {" & vbCrLf
        script = script & "    Write-Log '[ERROR] ジョブ起動失敗'" & vbCrLf
        script = script & "    Write-Output ""RESULT_STATUS:起動失敗""" & vbCrLf
        script = script & "    Write-Output ""RESULT_MESSAGE:$($output -join ' ')""" & vbCrLf
        script = script & "    exit" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & "  Write-Log '[成功] ジョブ起動完了'" & vbCrLf
        script = script & vbCrLf

        If waitCompletion Then
            ' 完了待ち
            script = script & "  Write-Log '[待機] ジョブ完了待ち開始...'" & vbCrLf
            script = script & "  $timeout = " & config("Timeout") & vbCrLf
            script = script & "  $interval = " & config("PollingInterval") & vbCrLf
            script = script & "  $startTime = Get-Date" & vbCrLf
            script = script & "  $isRunning = $true" & vbCrLf
            script = script & "  $pollCount = 0" & vbCrLf
            script = script & vbCrLf
            script = script & "  while ($isRunning) {" & vbCrLf
            script = script & "    $pollCount++" & vbCrLf
            script = script & "    if ($timeout -gt 0 -and ((Get-Date) - $startTime).TotalSeconds -ge $timeout) {" & vbCrLf
            script = script & "      Write-Log '[TIMEOUT] タイムアウトしました'" & vbCrLf
            script = script & "      Write-Output ""RESULT_STATUS:タイムアウト""" & vbCrLf
            script = script & "      break" & vbCrLf
            script = script & "    }" & vbCrLf
            script = script & vbCrLf
            script = script & "    # ajsshowでジョブネットの実行状態を取得" & vbCrLf
            script = script & "    $ajsshowPath = ""$jp1BinPath\ajsshow.exe""" & vbCrLf
            script = script & "    $statusResult = & $ajsshowPath -F " & config("SchedulerService") & " '" & jobnetPath & "' 2>&1" & vbCrLf
            script = script & "    $statusStr = $statusResult -join ' '" & vbCrLf
            script = script & "    $lastStatusStr = $statusStr" & vbCrLf
            script = script & "    Write-Log ""[ポーリング $pollCount] ステータス: $statusStr""" & vbCrLf
            script = script & vbCrLf
            script = script & "    # 日本語・英語両方のステータスに対応" & vbCrLf
            script = script & "    if ($statusStr -match '異常終了|ended abnormally|abnormal end|abend|killed|failed|キャンセル|中止') {" & vbCrLf
            script = script & "      Write-Log '[完了] 異常終了'" & vbCrLf
            script = script & "      Write-Output ""RESULT_STATUS:異常終了""" & vbCrLf
            script = script & "      $isRunning = $false" & vbCrLf
            script = script & "    } elseif ($statusStr -match '正常終了|ended normally|normal end|completed|end:') {" & vbCrLf
            script = script & "      Write-Log '[完了] 正常終了'" & vbCrLf
            script = script & "      Write-Output ""RESULT_STATUS:正常終了""" & vbCrLf
            script = script & "      $isRunning = $false" & vbCrLf
            script = script & "    } elseif ($statusStr -match '未実行|未登録|not registered|not found|KAVS0161') {" & vbCrLf
            script = script & "      Write-Log '[エラー] ユニットが見つかりません'" & vbCrLf
            script = script & "      Write-Output ""RESULT_STATUS:エラー""" & vbCrLf
            script = script & "      $isRunning = $false" & vbCrLf
            script = script & "    } else {" & vbCrLf
            script = script & "      Start-Sleep -Seconds $interval" & vbCrLf
            script = script & "    }" & vbCrLf
            script = script & "  }" & vbCrLf
            script = script & vbCrLf

            ' 最後のステータス情報から開始時間・終了時間を抽出
            script = script & "  Write-Log ""最終ステータス: $lastStatusStr""" & vbCrLf
            script = script & vbCrLf
            script = script & "  # ajsshow結果から時間を抽出（YYYY/MM/DD HH:MM形式）" & vbCrLf
            script = script & "  $timePattern = '\d{4}/\d{2}/\d{2}\s+\d{2}:\d{2}'" & vbCrLf
            script = script & "  $allTimes = [regex]::Matches($lastStatusStr, $timePattern)" & vbCrLf
            script = script & "  Write-Log ""検出した時間数: $($allTimes.Count)""" & vbCrLf
            script = script & "  if ($allTimes.Count -ge 1) {" & vbCrLf
            script = script & "    Write-Output ""RESULT_START:$($allTimes[0].Value)""" & vbCrLf
            script = script & "    Write-Log ""開始時間: $($allTimes[0].Value)""" & vbCrLf
            script = script & "  }" & vbCrLf
            script = script & "  if ($allTimes.Count -ge 2) {" & vbCrLf
            script = script & "    Write-Output ""RESULT_END:$($allTimes[1].Value)""" & vbCrLf
            script = script & "    Write-Log ""終了時間: $($allTimes[1].Value)""" & vbCrLf
            script = script & "  }" & vbCrLf
            script = script & vbCrLf
            script = script & "  # エラーメッセージを除いたメッセージを出力" & vbCrLf
            script = script & "  $cleanMsg = $lastStatusStr -replace 'KAVS\d+-[IEW][^\r\n]*', '' -replace '\s+', ' '" & vbCrLf
            script = script & "  Write-Output ""RESULT_MESSAGE:$cleanMsg""" & vbCrLf
        Else
            script = script & "  Write-Log '[完了] 起動成功（完了待ちなし）'" & vbCrLf
            script = script & "  Write-Output ""RESULT_STATUS:起動成功""" & vbCrLf
            script = script & "  Write-Output ""RESULT_MESSAGE:$($output -join ' ')""" & vbCrLf
        End If

        script = script & "} catch {" & vbCrLf
        script = script & "  Write-Log ""[EXCEPTION] $($_.Exception.Message)""" & vbCrLf
        script = script & "  Write-Output ""ERROR: $($_.Exception.Message)""" & vbCrLf
        script = script & "}" & vbCrLf
    Else
        ' リモート実行モード（WinRM使用）
        script = script & "Write-Log '--------------------------------------------------------------------------------'" & vbCrLf
        script = script & "Write-Log 'ジョブネット: " & jobnetPath & "'" & vbCrLf
        script = script & "Write-Log '接続先: " & config("JP1Server") & " (リモートモード)'" & vbCrLf
        script = script & "Write-Log '--------------------------------------------------------------------------------'" & vbCrLf
        script = script & vbCrLf

        ' 認証情報
        script = script & "$securePass = ConvertTo-SecureString '" & EscapePSString(config("RemotePassword")) & "' -AsPlainText -Force" & vbCrLf
        script = script & "$cred = New-Object System.Management.Automation.PSCredential('" & config("RemoteUser") & "', $securePass)" & vbCrLf
        script = script & vbCrLf

        ' WinRM設定の保存と自動設定
        script = script & "$originalTrustedHosts = $null" & vbCrLf
        script = script & "$winrmConfigChanged = $false" & vbCrLf
        script = script & "$winrmServiceWasStarted = $false" & vbCrLf
        script = script & vbCrLf

        script = script & "try {" & vbCrLf
        script = script & "  Write-Log '[準備] WinRM設定を確認中...'" & vbCrLf
        script = script & vbCrLf
        script = script & "  # WinRMサービスの起動確認（TrustedHosts取得前に起動が必要）" & vbCrLf
        script = script & "  $winrmService = Get-Service -Name WinRM -ErrorAction SilentlyContinue" & vbCrLf
        script = script & "  if ($winrmService.Status -ne 'Running') {" & vbCrLf
        script = script & "    Write-Log '[準備] WinRMサービスを起動'" & vbCrLf
        script = script & "    Start-Service -Name WinRM -ErrorAction Stop" & vbCrLf
        script = script & "    $winrmServiceWasStarted = $true" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & vbCrLf
        script = script & "  # 現在のTrustedHostsを取得（WinRMサービス起動後に取得）" & vbCrLf
        script = script & "  $originalTrustedHosts = (Get-Item WSMan:\localhost\Client\TrustedHosts -ErrorAction SilentlyContinue).Value" & vbCrLf
        script = script & vbCrLf
        script = script & "  # TrustedHostsに接続先を追加（必要な場合のみ）" & vbCrLf
        script = script & "  if ($originalTrustedHosts -notmatch '" & config("JP1Server") & "') {" & vbCrLf
        script = script & "    Write-Log '[準備] TrustedHostsに接続先を追加'" & vbCrLf
        script = script & "    if ($originalTrustedHosts) {" & vbCrLf
        script = script & "      Set-Item WSMan:\localhost\Client\TrustedHosts -Value ""$originalTrustedHosts," & config("JP1Server") & """ -Force" & vbCrLf
        script = script & "    } else {" & vbCrLf
        script = script & "      Set-Item WSMan:\localhost\Client\TrustedHosts -Value '" & config("JP1Server") & "' -Force" & vbCrLf
        script = script & "    }" & vbCrLf
        script = script & "    $winrmConfigChanged = $true" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & vbCrLf

        ' リモート実行
        script = script & "  Write-Log '[接続] リモートセッション作成中...'" & vbCrLf
        script = script & "  $session = New-PSSession -ComputerName '" & config("JP1Server") & "' -Credential $cred -ErrorAction Stop" & vbCrLf
        script = script & "  Write-Log '[接続] セッション確立完了'" & vbCrLf
        script = script & vbCrLf

        ' 保留解除処理（リモート）
        If isHold Then
            script = script & "  # 保留解除" & vbCrLf
            script = script & "  Write-Log '[実行] ajsrelease - 保留解除（リモート）'" & vbCrLf
            script = script & "  $releaseResult = Invoke-Command -Session $session -ScriptBlock {" & vbCrLf
            script = script & "    param($schedulerService, $jobnetPath)" & vbCrLf
            script = script & "    $ajsreleasePath = $null" & vbCrLf
            script = script & "    $searchPaths = @('C:\Program Files\HITACHI\JP1AJS3\bin\ajsrelease.exe','C:\Program Files (x86)\HITACHI\JP1AJS3\bin\ajsrelease.exe','C:\Program Files\Hitachi\JP1AJS2\bin\ajsrelease.exe','C:\Program Files (x86)\Hitachi\JP1AJS2\bin\ajsrelease.exe')" & vbCrLf
            script = script & "    foreach ($p in $searchPaths) { if (Test-Path $p) { $ajsreleasePath = $p; break } }" & vbCrLf
            script = script & "    if (-not $ajsreleasePath) { Write-Output 'ERROR: ajsrelease.exe not found'; return @{ ExitCode = 1; Output = 'ajsrelease.exe not found' } }" & vbCrLf
            script = script & "    $output = & $ajsreleasePath '-F' $schedulerService $jobnetPath 2>&1" & vbCrLf
            script = script & "    @{ ExitCode = $LASTEXITCODE; Output = ($output -join ' ') }" & vbCrLf
            script = script & "  } -ArgumentList '" & config("SchedulerService") & "', '" & jobnetPath & "'" & vbCrLf
            script = script & "  Write-Log ""結果: $($releaseResult.Output)""" & vbCrLf
            script = script & vbCrLf
            script = script & "  if ($releaseResult.ExitCode -ne 0) {" & vbCrLf
            script = script & "    Write-Log '[ERROR] 保留解除失敗'" & vbCrLf
            script = script & "    Write-Output ""RESULT_STATUS:保留解除失敗""" & vbCrLf
            script = script & "    Write-Output ""RESULT_MESSAGE:$($releaseResult.Output)""" & vbCrLf
            script = script & "    Remove-PSSession $session" & vbCrLf
            script = script & "    exit" & vbCrLf
            script = script & "  }" & vbCrLf
            script = script & "  Write-Log '[成功] 保留解除完了'" & vbCrLf
            script = script & vbCrLf
        End If

        ' ajsentry実行
        script = script & "  Write-Log '[実行] ajsentry - ジョブ起動（リモート）'" & vbCrLf
        script = script & "  $entryResult = Invoke-Command -Session $session -ScriptBlock {" & vbCrLf
        script = script & "    param($schedulerService, $jobnetPath)" & vbCrLf
        script = script & "    $ajsentryPath = $null" & vbCrLf
        script = script & "    $searchPaths = @('C:\Program Files\HITACHI\JP1AJS3\bin\ajsentry.exe','C:\Program Files (x86)\HITACHI\JP1AJS3\bin\ajsentry.exe','C:\Program Files\Hitachi\JP1AJS2\bin\ajsentry.exe','C:\Program Files (x86)\Hitachi\JP1AJS2\bin\ajsentry.exe')" & vbCrLf
        script = script & "    foreach ($p in $searchPaths) { if (Test-Path $p) { $ajsentryPath = $p; break } }" & vbCrLf
        script = script & "    if (-not $ajsentryPath) { Write-Output 'ERROR: ajsentry.exe not found'; return @{ ExitCode = 1; Output = 'ajsentry.exe not found' } }" & vbCrLf
        script = script & "    $output = & $ajsentryPath '-F' $schedulerService $jobnetPath 2>&1" & vbCrLf
        script = script & "    @{ ExitCode = $LASTEXITCODE; Output = ($output -join ' ') }" & vbCrLf
        script = script & "  } -ArgumentList '" & config("SchedulerService") & "', '" & jobnetPath & "'" & vbCrLf
        script = script & "  Write-Log ""結果: $($entryResult.Output)""" & vbCrLf
        script = script & vbCrLf

        script = script & "  if ($entryResult.ExitCode -ne 0) {" & vbCrLf
        script = script & "    Write-Log '[ERROR] ジョブ起動失敗'" & vbCrLf
        script = script & "    Write-Output ""RESULT_STATUS:起動失敗""" & vbCrLf
        script = script & "    Write-Output ""RESULT_MESSAGE:$($entryResult.Output)""" & vbCrLf
        script = script & "    Remove-PSSession $session" & vbCrLf
        script = script & "    exit" & vbCrLf
        script = script & "  }" & vbCrLf
        script = script & "  Write-Log '[成功] ジョブ起動完了'" & vbCrLf
        script = script & vbCrLf

        If waitCompletion Then
            ' 完了待ち
            script = script & "  Write-Log '[待機] ジョブ完了待ち開始...'" & vbCrLf
            script = script & "  $timeout = " & config("Timeout") & vbCrLf
            script = script & "  $interval = " & config("PollingInterval") & vbCrLf
            script = script & "  $startTime = Get-Date" & vbCrLf
            script = script & "  $isRunning = $true" & vbCrLf
            script = script & "  $pollCount = 0" & vbCrLf
            script = script & vbCrLf
            script = script & "  while ($isRunning) {" & vbCrLf
            script = script & "    $pollCount++" & vbCrLf
            script = script & "    if ($timeout -gt 0 -and ((Get-Date) - $startTime).TotalSeconds -ge $timeout) {" & vbCrLf
            script = script & "      Write-Log '[TIMEOUT] タイムアウトしました'" & vbCrLf
            script = script & "      Write-Output ""RESULT_STATUS:タイムアウト""" & vbCrLf
            script = script & "      break" & vbCrLf
            script = script & "    }" & vbCrLf
            script = script & vbCrLf
            script = script & "    # ajsshowでジョブネットの実行状態を取得（リモート）" & vbCrLf
            script = script & "    $statusResult = Invoke-Command -Session $session -ScriptBlock {" & vbCrLf
            script = script & "      param($schedulerService, $jobnetPath)" & vbCrLf
            script = script & "      $ajsshowPath = $null" & vbCrLf
            script = script & "      $searchPaths = @('C:\Program Files\HITACHI\JP1AJS3\bin\ajsshow.exe','C:\Program Files (x86)\HITACHI\JP1AJS3\bin\ajsshow.exe','C:\Program Files\Hitachi\JP1AJS2\bin\ajsshow.exe','C:\Program Files (x86)\Hitachi\JP1AJS2\bin\ajsshow.exe')" & vbCrLf
            script = script & "      foreach ($p in $searchPaths) { if (Test-Path $p) { $ajsshowPath = $p; break } }" & vbCrLf
            script = script & "      if (-not $ajsshowPath) { return 'ERROR: ajsshow.exe not found' }" & vbCrLf
            script = script & "      & $ajsshowPath '-F' $schedulerService $jobnetPath 2>&1" & vbCrLf
            script = script & "    } -ArgumentList '" & config("SchedulerService") & "', '" & jobnetPath & "'" & vbCrLf
            script = script & vbCrLf
            script = script & "    $statusStr = $statusResult -join ' '" & vbCrLf
            script = script & "    $lastStatusStr = $statusStr" & vbCrLf
            script = script & "    Write-Log ""[ポーリング $pollCount] ステータス: $statusStr""" & vbCrLf
            script = script & "    # 日本語・英語両方のステータスに対応" & vbCrLf
            script = script & "    if ($statusStr -match '異常終了|ended abnormally|abnormal end|abend|killed|failed|キャンセル|中止') {" & vbCrLf
            script = script & "      Write-Log '[完了] 異常終了'" & vbCrLf
            script = script & "      Write-Output ""RESULT_STATUS:異常終了""" & vbCrLf
            script = script & "      $isRunning = $false" & vbCrLf
            script = script & "    } elseif ($statusStr -match '正常終了|ended normally|normal end|completed|end:') {" & vbCrLf
            script = script & "      Write-Log '[完了] 正常終了'" & vbCrLf
            script = script & "      Write-Output ""RESULT_STATUS:正常終了""" & vbCrLf
            script = script & "      $isRunning = $false" & vbCrLf
            script = script & "    } elseif ($statusStr -match '未実行|未登録|not registered|not found|KAVS0161') {" & vbCrLf
            script = script & "      Write-Log '[エラー] ユニットが見つかりません'" & vbCrLf
            script = script & "      Write-Output ""RESULT_STATUS:エラー""" & vbCrLf
            script = script & "      $isRunning = $false" & vbCrLf
            script = script & "    } else {" & vbCrLf
            script = script & "      Start-Sleep -Seconds $interval" & vbCrLf
            script = script & "    }" & vbCrLf
            script = script & "  }" & vbCrLf
            script = script & vbCrLf

            ' 最後のステータス情報から開始時間・終了時間を抽出
            script = script & "  Write-Log ""最終ステータス: $lastStatusStr""" & vbCrLf
            script = script & vbCrLf
            script = script & "  # ajsshow結果から時間を抽出（YYYY/MM/DD HH:MM形式）" & vbCrLf
            script = script & "  $timePattern = '\d{4}/\d{2}/\d{2}\s+\d{2}:\d{2}'" & vbCrLf
            script = script & "  $allTimes = [regex]::Matches($lastStatusStr, $timePattern)" & vbCrLf
            script = script & "  Write-Log ""検出した時間数: $($allTimes.Count)""" & vbCrLf
            script = script & "  if ($allTimes.Count -ge 1) {" & vbCrLf
            script = script & "    Write-Output ""RESULT_START:$($allTimes[0].Value)""" & vbCrLf
            script = script & "    Write-Log ""開始時間: $($allTimes[0].Value)""" & vbCrLf
            script = script & "  }" & vbCrLf
            script = script & "  if ($allTimes.Count -ge 2) {" & vbCrLf
            script = script & "    Write-Output ""RESULT_END:$($allTimes[1].Value)""" & vbCrLf
            script = script & "    Write-Log ""終了時間: $($allTimes[1].Value)""" & vbCrLf
            script = script & "  }" & vbCrLf
            script = script & vbCrLf
            script = script & "  # エラーメッセージを除いたメッセージを出力" & vbCrLf
            script = script & "  $cleanMsg = $lastStatusStr -replace 'KAVS\d+-[IEW][^\r\n]*', '' -replace '\s+', ' '" & vbCrLf
            script = script & "  Write-Output ""RESULT_MESSAGE:$cleanMsg""" & vbCrLf
        Else
            script = script & "  Write-Log '[完了] 起動成功（完了待ちなし）'" & vbCrLf
            script = script & "  Write-Output ""RESULT_STATUS:起動成功""" & vbCrLf
            script = script & "  Write-Output ""RESULT_MESSAGE:$($entryResult.Output)""" & vbCrLf
        End If

        script = script & "  Write-Log '[クリーンアップ] セッション終了'" & vbCrLf
        script = script & "  Remove-PSSession $session" & vbCrLf
        script = script & "} catch {" & vbCrLf
        script = script & "  Write-Log ""[EXCEPTION] $($_.Exception.Message)""" & vbCrLf
        script = script & "  Write-Output ""ERROR: $($_.Exception.Message)""" & vbCrLf
        script = script & "} finally {" & vbCrLf
        script = script & "  # WinRM設定の復元" & vbCrLf
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

Private Sub UpdateJobListStatus(ByVal row As Long, ByVal result As Object)
    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_JOBLIST)

    ws.Cells(row, COL_LAST_STATUS).Value = result("Status")
    ws.Cells(row, COL_LAST_EXEC_TIME).Value = result("StartTime")
    ws.Cells(row, COL_LAST_END_TIME).Value = result("EndTime")

    ' 詳細メッセージを記録
    If result("Message") <> "" Then
        ws.Cells(row, COL_LAST_MESSAGE).Value = result("Message")
    End If

    ' 保留解除された場合（成功時）、保留列をクリアしてハイライトを解除
    If result("Status") = "正常終了" Or result("Status") = "起動成功" Then
        If ws.Cells(row, COL_HOLD).Value = "保留中" Then
            ws.Cells(row, COL_HOLD).Value = ""
            ws.Cells(row, COL_HOLD).Font.Bold = False
            ws.Cells(row, COL_HOLD).Font.Color = RGB(0, 0, 0)
            ' 行のハイライトを解除
            ws.Range(ws.Cells(row, COL_ORDER), ws.Cells(row, COL_LAST_MESSAGE)).Interior.ColorIndex = xlNone
        End If
    End If

    ' 色付け
    If result("Status") = "正常終了" Then
        ws.Cells(row, COL_LAST_STATUS).Interior.Color = RGB(198, 239, 206)
    ElseIf result("Status") = "起動成功" Then
        ws.Cells(row, COL_LAST_STATUS).Interior.Color = RGB(255, 235, 156)
    Else
        ws.Cells(row, COL_LAST_STATUS).Interior.Color = RGB(255, 199, 206)
    End If
End Sub

'==============================================================================
' 一覧クリア
'==============================================================================
Public Sub ClearJobList()
    If MsgBox("ジョブ一覧をクリアしますか？ (y/n)", vbYesNo + vbQuestion) = vbNo Then Exit Sub

    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_JOBLIST)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_JOBNET_PATH).End(xlUp).Row

    If lastRow >= ROW_JOBLIST_DATA_START Then
        ws.Range(ws.Cells(ROW_JOBLIST_DATA_START, COL_ORDER), ws.Cells(lastRow, COL_LAST_MESSAGE)).Clear
    End If

    MsgBox "クリアしました。", vbInformation
End Sub

'==============================================================================
' ユーティリティ
'==============================================================================
Private Function GetConfig() As Object
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

Private Function ExecutePowerShell(script As String) As String
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
    ' PowerShellはBOM付きUTF-8を自動認識するため、日本語パスが正しく処理される
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
    ' PowerShellウィンドウを直接表示して実行
    ' Start-Transcriptでログを取りながらリアルタイム表示
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

Private Function EscapePSString(str As String) As String
    ' PowerShell文字列内のシングルクォートをエスケープ
    EscapePSString = Replace(str, "'", "''")
End Function

'==============================================================================
' 管理者権限チェック
'==============================================================================
Private Function IsRunningAsAdmin() As Boolean
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

Private Function EnsureAdminForRemoteMode(config As Object) As Boolean
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

Private Sub RestartAsAdmin()
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

    ' 現在のExcelを終了
    Application.Quit
End Sub

'==============================================================================
' ログファイル関連
'==============================================================================
Private Function CreateLogFile() As String
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

Private Function GetLogFilePath() As String
    ' 現在のログファイルパスを返す
    GetLogFilePath = g_LogFilePath
End Function
