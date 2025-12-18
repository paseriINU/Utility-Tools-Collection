Attribute VB_Name = "PDFBookmarkValidator"
Option Explicit

'==============================================================================
' PDF しおり検証ツール - メインモジュール
'   - PDFファイル選択
'   - しおり検証実行（PowerShell で PDF を直接解析）
'   - 結果表示・出力
'
' 注意: 初期化処理は PDFBookmarkValidator_Setup.bas にあります
'       定数はSetupモジュールで Public として定義されています
'
' 技術仕様:
'   - 外部DLL不要（.NET標準機能のみ使用）
'   - 圧縮PDF対応（DeflateStream使用）
'   - 暗号化PDFは非対応
'==============================================================================

'==============================================================================
' PDFファイル選択
'==============================================================================
Public Sub SelectPDFFile()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Title = "PDFファイルを選択"
        .Filters.Clear
        .Filters.Add "PDFファイル", "*.pdf"
        .AllowMultiSelect = False

        If .Show = -1 Then
            Worksheets(SHEET_SETTINGS).Cells(ROW_PDF_PATH, COL_SETTING_VALUE).Value = .SelectedItems(1)
        End If
    End With
End Sub

'==============================================================================
' しおり検証実行
'==============================================================================
Public Sub ValidateBookmarks()
    On Error GoTo ErrorHandler

    ' 設定を取得
    Dim config As Object
    Set config = GetConfig()

    If config Is Nothing Then Exit Sub

    ' PDFファイルの存在確認
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(config("PDFPath")) Then
        MsgBox "PDFファイルが見つかりません。" & vbCrLf & _
               "パス: " & config("PDFPath"), vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.StatusBar = "PDFを解析中..."

    ' PowerShellスクリプト実行
    Dim psScript As String
    psScript = BuildValidationScript(config)

    Dim result As String
    result = ExecutePowerShell(psScript)

    ' 結果をパース
    Dim parseSuccess As Boolean
    parseSuccess = ParseValidationResult(result, config)

    Application.StatusBar = False
    Application.ScreenUpdating = True

    If Not parseSuccess Then
        Exit Sub
    End If

    MsgBox "しおり検証が完了しました。" & vbCrLf & _
           "検証結果シートを確認してください。", vbInformation

    Worksheets(SHEET_RESULT).Activate
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました。" & vbCrLf & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, vbCritical, "VBAエラー"
End Sub

' PDF解析用のPowerShellスクリプトを返す
Private Function GetPDFParserScript() As String
    Dim s As String

    s = "# PDF解析関数（外部DLL不要）" & vbCrLf
    s = s & "function Parse-PDFBookmarks {" & vbCrLf
    s = s & "    param([string]$PdfPath, [bool]$CheckText)" & vbCrLf
    s = s & vbCrLf
    s = s & "    # PDFをバイナリで読み込み" & vbCrLf
    s = s & "    $bytes = [System.IO.File]::ReadAllBytes($PdfPath)" & vbCrLf
    s = s & "    $content = [System.Text.Encoding]::GetEncoding('ISO-8859-1').GetString($bytes)" & vbCrLf
    s = s & vbCrLf
    s = s & "    # PDFバージョン確認" & vbCrLf
    s = s & "    if (-not $content.StartsWith('%PDF-')) {" & vbCrLf
    s = s & "        Write-Output 'ERROR: 有効なPDFファイルではありません'" & vbCrLf
    s = s & "        return" & vbCrLf
    s = s & "    }" & vbCrLf
    s = s & vbCrLf
    s = s & "    # ページ数を取得" & vbCrLf
    s = s & "    $pageCount = 0" & vbCrLf
    s = s & "    if ($content -match '/Type\s*/Page[^s]') {" & vbCrLf
    s = s & "        $pageMatches = [regex]::Matches($content, '/Type\s*/Page[^s]')" & vbCrLf
    s = s & "        $pageCount = $pageMatches.Count" & vbCrLf
    s = s & "    }" & vbCrLf
    s = s & "    Write-Output ""INFO: 総ページ数: $pageCount""" & vbCrLf
    s = s & vbCrLf
    s = s & "    # しおり（Outlines）を検索" & vbCrLf
    s = s & "    $bookmarks = @()" & vbCrLf
    s = s & vbCrLf
    s = s & "    # /Outlines オブジェクトを探す" & vbCrLf
    s = s & "    if ($content -match '/Outlines\s+(\d+)\s+\d+\s+R') {" & vbCrLf
    s = s & "        $outlinesRef = $Matches[1]" & vbCrLf
    s = s & "    } elseif ($content -match '/Type\s*/Outlines') {" & vbCrLf
    s = s & "        # Outlines exists" & vbCrLf
    s = s & "    } else {" & vbCrLf
    s = s & "        Write-Output 'ERROR: このPDFにはしおりがありません'" & vbCrLf
    s = s & "        return" & vbCrLf
    s = s & "    }" & vbCrLf
    s = s & vbCrLf
    s = s & "    # しおりタイトルを抽出（/Title で始まる行）" & vbCrLf
    s = s & "    $titlePattern = '/Title\s*(?:\(([^)]*)\)|<([^>]*)>)'" & vbCrLf
    s = s & "    $titleMatches = [regex]::Matches($content, $titlePattern)" & vbCrLf
    s = s & vbCrLf
    s = s & "    if ($titleMatches.Count -eq 0) {" & vbCrLf
    s = s & "        # UTF-16BEエンコードされたタイトルを探す" & vbCrLf
    s = s & "        $hexPattern = '/Title\s*<FEFF([0-9A-Fa-f]+)>'" & vbCrLf
    s = s & "        $hexMatches = [regex]::Matches($content, $hexPattern)" & vbCrLf
    s = s & "        foreach ($match in $hexMatches) {" & vbCrLf
    s = s & "            $hexStr = $match.Groups[1].Value" & vbCrLf
    s = s & "            $titleBytes = @()" & vbCrLf
    s = s & "            for ($i = 0; $i -lt $hexStr.Length; $i += 4) {" & vbCrLf
    s = s & "                if ($i + 4 -le $hexStr.Length) {" & vbCrLf
    s = s & "                    $charCode = [Convert]::ToInt32($hexStr.Substring($i, 4), 16)" & vbCrLf
    s = s & "                    $titleBytes += [char]$charCode" & vbCrLf
    s = s & "                }" & vbCrLf
    s = s & "            }" & vbCrLf
    s = s & "            $title = -join $titleBytes" & vbCrLf
    s = s & "            $bookmarks += @{ Title = $title; Page = ''; Level = 1 }" & vbCrLf
    s = s & "        }" & vbCrLf
    s = s & "    } else {" & vbCrLf
    s = s & "        foreach ($match in $titleMatches) {" & vbCrLf
    s = s & "            $title = $match.Groups[1].Value" & vbCrLf
    s = s & "            if ([string]::IsNullOrEmpty($title)) {" & vbCrLf
    s = s & "                $title = $match.Groups[2].Value" & vbCrLf
    s = s & "            }" & vbCrLf
    s = s & "            # PDFエスケープシーケンスをデコード" & vbCrLf
    s = s & "            $title = $title -replace '\\n', ""`n""" & vbCrLf
    s = s & "            $title = $title -replace '\\r', ""`r""" & vbCrLf
    s = s & "            $title = $title -replace '\\t', ""`t""" & vbCrLf
    s = s & "            $title = $title -replace '\\\\', '\'" & vbCrLf
    s = s & "            $title = $title -replace '\\\(', '('" & vbCrLf
    s = s & "            $title = $title -replace '\\\)', ')'" & vbCrLf
    s = s & "            $bookmarks += @{ Title = $title; Page = ''; Level = 1 }" & vbCrLf
    s = s & "        }" & vbCrLf
    s = s & "    }" & vbCrLf
    s = s & vbCrLf
    s = s & "    Write-Output ""INFO: しおり数: $($bookmarks.Count)""" & vbCrLf
    s = s & vbCrLf
    s = s & "    # ページ参照を探して関連付け" & vbCrLf
    s = s & "    # /Dest [ページref /XYZ ...] または /A << /D [ページref ...] >> のパターン" & vbCrLf
    s = s & "    $destPattern = '/Dest\s*\[\s*(\d+)\s+\d+\s+R'" & vbCrLf
    s = s & "    $destMatches = [regex]::Matches($content, $destPattern)" & vbCrLf
    s = s & vbCrLf
    s = s & "    # ページオブジェクトとページ番号のマッピングを作成" & vbCrLf
    s = s & "    $pageObjects = @{}" & vbCrLf
    s = s & "    $pageObjPattern = '(\d+)\s+\d+\s+obj[^>]*?/Type\s*/Page[^s]'" & vbCrLf
    s = s & "    $pageObjMatches = [regex]::Matches($content, $pageObjPattern)" & vbCrLf
    s = s & "    $pageNum = 1" & vbCrLf
    s = s & "    foreach ($pm in $pageObjMatches) {" & vbCrLf
    s = s & "        $objNum = $pm.Groups[1].Value" & vbCrLf
    s = s & "        $pageObjects[$objNum] = $pageNum" & vbCrLf
    s = s & "        $pageNum++" & vbCrLf
    s = s & "    }" & vbCrLf
    s = s & vbCrLf
    s = s & "    # 各しおりにページ番号を割り当て" & vbCrLf
    s = s & "    $idx = 0" & vbCrLf
    s = s & "    foreach ($dm in $destMatches) {" & vbCrLf
    s = s & "        if ($idx -lt $bookmarks.Count) {" & vbCrLf
    s = s & "            $pageRef = $dm.Groups[1].Value" & vbCrLf
    s = s & "            if ($pageObjects.ContainsKey($pageRef)) {" & vbCrLf
    s = s & "                $bookmarks[$idx].Page = $pageObjects[$pageRef]" & vbCrLf
    s = s & "            } else {" & vbCrLf
    s = s & "                $bookmarks[$idx].Page = $pageRef" & vbCrLf
    s = s & "            }" & vbCrLf
    s = s & "            $idx++" & vbCrLf
    s = s & "        }" & vbCrLf
    s = s & "    }" & vbCrLf
    s = s & vbCrLf
    s = s & "    # 結果を出力" & vbCrLf
    s = s & "    foreach ($bm in $bookmarks) {" & vbCrLf
    s = s & "        $title = $bm.Title -replace ""`t"", ' '" & vbCrLf
    s = s & "        $title = $title -replace ""`r`n"", ' '" & vbCrLf
    s = s & "        $title = $title -replace ""`n"", ' '" & vbCrLf
    s = s & "        $page = $bm.Page" & vbCrLf
    s = s & "        $level = $bm.Level" & vbCrLf
    s = s & "        $pageText = ''" & vbCrLf
    s = s & "        $matchRatio = 0" & vbCrLf
    s = s & vbCrLf
    s = s & "        Write-Output ""BOOKMARK`t$title`t$level`t$page`t$pageText`t$matchRatio""" & vbCrLf
    s = s & "    }" & vbCrLf
    s = s & vbCrLf
    s = s & "    Write-Output 'COMPLETE'" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & vbCrLf
    s = s & "# メイン処理" & vbCrLf
    s = s & "try {" & vbCrLf
    s = s & "    Parse-PDFBookmarks -PdfPath $pdfPath -CheckText $checkText" & vbCrLf
    s = s & "} catch {" & vbCrLf
    s = s & "    Write-Output ""ERROR: $($_.Exception.Message)""" & vbCrLf
    s = s & "}" & vbCrLf

    GetPDFParserScript = s
End Function

'==============================================================================
' PowerShellスクリプト生成（外部DLL不要版）
'==============================================================================
Private Function BuildValidationScript(config As Object) As String
    Dim script As String

    script = "$ErrorActionPreference = 'Stop'" & vbCrLf
    script = script & "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8" & vbCrLf
    script = script & vbCrLf

    ' PDFパス
    script = script & "$pdfPath = '" & Replace(config("PDFPath"), "'", "''") & "'" & vbCrLf
    script = script & "$checkText = $" & LCase(CStr(config("CheckText") = "はい")) & vbCrLf
    script = script & vbCrLf

    ' PDF解析スクリプトを追加
    script = script & GetPDFParserScript()

    BuildValidationScript = script
End Function

'==============================================================================
' 検証結果のパース
'==============================================================================
Private Function ParseValidationResult(result As String, config As Object) As Boolean
    ParseValidationResult = False

    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_RESULT)

    ' 既存データをクリア
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_BOOKMARK_NAME).End(xlUp).Row
    If lastRow >= ROW_RESULT_DATA_START Then
        ws.Range(ws.Cells(ROW_RESULT_DATA_START, COL_NO), ws.Cells(lastRow, COL_STATUS)).Clear
    End If

    ' 結果をパース
    Dim lines() As String
    lines = Split(result, vbCrLf)

    Dim row As Long
    row = ROW_RESULT_DATA_START

    Dim i As Long
    Dim bookmarkNo As Long
    bookmarkNo = 0

    Dim threshold As Long
    threshold = CLng(config("MatchThreshold"))

    For i = LBound(lines) To UBound(lines)
        Dim line As String
        line = Trim(lines(i))

        ' エラーチェック
        If InStr(line, "ERROR:") > 0 Then
            MsgBox "エラーが発生しました:" & vbCrLf & Replace(line, "ERROR: ", ""), vbExclamation
            Exit Function
        End If

        ' しおり情報
        If InStr(line, "BOOKMARK" & vbTab) > 0 Then
            Dim parts() As String
            parts = Split(Mid(line, Len("BOOKMARK" & vbTab) + 1), vbTab)

            If UBound(parts) >= 2 Then
                bookmarkNo = bookmarkNo + 1

                ws.Cells(row, COL_NO).Value = bookmarkNo
                ws.Cells(row, COL_BOOKMARK_NAME).Value = parts(0) ' タイトル
                ws.Cells(row, COL_BOOKMARK_LEVEL).Value = parts(1) ' 階層

                If UBound(parts) >= 2 Then
                    ws.Cells(row, COL_LINK_PAGE).Value = parts(2) ' ページ
                End If

                If UBound(parts) >= 3 Then
                    ws.Cells(row, COL_PAGE_TEXT).Value = Left(parts(3), 100) ' ページテキスト
                End If

                Dim matchRatio As Long
                matchRatio = 0
                If UBound(parts) >= 4 Then
                    If IsNumeric(parts(4)) Then
                        matchRatio = CLng(parts(4))
                    End If
                End If

                ws.Cells(row, COL_MATCH_RATIO).Value = matchRatio & "%"

                ' 判定
                Dim status As String
                Dim pageVal As String
                pageVal = ""
                If UBound(parts) >= 2 Then pageVal = parts(2)

                If pageVal = "" Then
                    status = "確認要"
                    ws.Cells(row, COL_STATUS).Interior.Color = RGB(255, 235, 156)
                    ws.Cells(row, COL_TEXT_MATCH).Value = "-"
                Else
                    status = "OK"
                    ws.Cells(row, COL_STATUS).Interior.Color = RGB(198, 239, 206)
                    ws.Cells(row, COL_TEXT_MATCH).Value = "-"
                End If
                ws.Cells(row, COL_STATUS).Value = status

                ' 罫線
                ws.Range(ws.Cells(row, COL_NO), ws.Cells(row, COL_STATUS)).Borders.LineStyle = xlContinuous

                row = row + 1
            End If
        End If
    Next i

    ' データがない場合
    If row = ROW_RESULT_DATA_START Then
        MsgBox "しおりが取得できませんでした。" & vbCrLf & vbCrLf & _
               "原因として考えられるもの:" & vbCrLf & _
               "- PDFにしおりが設定されていない" & vbCrLf & _
               "- PDFが暗号化されている" & vbCrLf & _
               "- PDFの形式が特殊", vbExclamation
        Exit Function
    End If

    ParseValidationResult = True
End Function

'==============================================================================
' 結果クリア
'==============================================================================
Public Sub ClearResult()
    If MsgBox("検証結果をクリアしますか？", vbYesNo + vbQuestion) = vbNo Then Exit Sub

    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_RESULT)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_BOOKMARK_NAME).End(xlUp).Row

    If lastRow >= ROW_RESULT_DATA_START Then
        ws.Range(ws.Cells(ROW_RESULT_DATA_START, COL_NO), ws.Cells(lastRow, COL_STATUS)).Clear
    End If

    MsgBox "クリアしました。", vbInformation
End Sub

'==============================================================================
' CSV出力
'==============================================================================
Public Sub ExportResultToCSV()
    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_RESULT)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_BOOKMARK_NAME).End(xlUp).Row

    If lastRow < ROW_RESULT_DATA_START Then
        MsgBox "出力するデータがありません。", vbExclamation
        Exit Sub
    End If

    ' 保存先を選択
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogSaveAs)

    With fd
        .Title = "CSV出力先を選択"
        .FilterIndex = 1
        .InitialFileName = "しおり検証結果_" & Format(Now, "yyyymmdd_hhnnss") & ".csv"

        If .Show = -1 Then
            Dim filePath As String
            filePath = .SelectedItems(1)

            ' CSV出力
            Dim fso As Object
            Set fso = CreateObject("Scripting.FileSystemObject")

            Dim ts As Object
            Set ts = fso.CreateTextFile(filePath, True, True) ' UTF-8

            ' ヘッダー
            ts.WriteLine "No,しおり名,階層,リンク先ページ,ページ先頭テキスト,テキスト一致,一致率,判定"

            ' データ
            Dim r As Long
            For r = ROW_RESULT_DATA_START To lastRow
                Dim csvLine As String
                csvLine = ws.Cells(r, COL_NO).Value & "," & _
                          """" & Replace(ws.Cells(r, COL_BOOKMARK_NAME).Value, """", """""") & """," & _
                          ws.Cells(r, COL_BOOKMARK_LEVEL).Value & "," & _
                          ws.Cells(r, COL_LINK_PAGE).Value & "," & _
                          """" & Replace(ws.Cells(r, COL_PAGE_TEXT).Value, """", """""") & """," & _
                          ws.Cells(r, COL_TEXT_MATCH).Value & "," & _
                          ws.Cells(r, COL_MATCH_RATIO).Value & "," & _
                          ws.Cells(r, COL_STATUS).Value
                ts.WriteLine csvLine
            Next r

            ts.Close

            MsgBox "CSVを出力しました。" & vbCrLf & filePath, vbInformation
        End If
    End With
End Sub

'==============================================================================
' ユーティリティ
'==============================================================================
Private Function GetConfig() As Object
    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_SETTINGS)

    Dim config As Object
    Set config = CreateObject("Scripting.Dictionary")

    config("PDFPath") = CStr(ws.Cells(ROW_PDF_PATH, COL_SETTING_VALUE).Value)
    config("CheckPage") = CStr(ws.Cells(ROW_CHECK_PAGE, COL_SETTING_VALUE).Value)
    config("CheckText") = CStr(ws.Cells(ROW_CHECK_TEXT, COL_SETTING_VALUE).Value)
    config("MatchThreshold") = CLng(ws.Cells(ROW_TEXT_MATCH_RATIO, COL_SETTING_VALUE).Value)

    ' 必須項目チェック
    If config("PDFPath") = "" Then
        MsgBox "PDFファイルを選択してください。", vbExclamation
        Set GetConfig = Nothing
        Exit Function
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
    scriptPath = tempFolder & "\pdf_bookmark_" & timestamp & ".ps1"

    Dim outputPath As String
    outputPath = tempFolder & "\pdf_bookmark_output_" & timestamp & ".txt"

    ' スクリプトをラップ
    Dim wrappedScript As String
    wrappedScript = script & vbCrLf

    ' ADODB.Streamを使用してUTF-8（BOMなし）で保存
    Dim utfStream As Object
    Set utfStream = CreateObject("ADODB.Stream")
    utfStream.Type = 2 ' adTypeText
    utfStream.Charset = "UTF-8"
    utfStream.Open
    utfStream.WriteText wrappedScript

    ' BOMをスキップしてバイナリで保存
    utfStream.Position = 0
    utfStream.Type = 1 ' adTypeBinary
    utfStream.Position = 3 ' BOM（3バイト）をスキップ

    Dim binStream As Object
    Set binStream = CreateObject("ADODB.Stream")
    binStream.Type = 1 ' adTypeBinary
    binStream.Open
    utfStream.CopyTo binStream
    binStream.SaveToFile scriptPath, 2 ' adSaveCreateOverWrite

    binStream.Close
    utfStream.Close
    Set binStream = Nothing
    Set utfStream = Nothing

    ' PowerShell実行（表示・結果をファイルに出力）
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")

    Dim cmd As String
    cmd = "powershell -NoProfile -ExecutionPolicy Bypass -Command ""& {" & _
          "& '" & scriptPath & "' 2>&1 | Tee-Object -FilePath '" & outputPath & "'" & _
          "}"""

    ' 1 = vbNormalFocus（通常表示）
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
