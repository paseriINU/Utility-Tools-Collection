Attribute VB_Name = "PDFBookmarkValidator"
Option Explicit

'==============================================================================
' PDF しおり検証ツール - メインモジュール
'   - PDFファイル選択
'   - しおり検証実行（PowerShell + iTextSharp）
'   - 結果表示・出力
'
' 注意: 初期化処理は PDFBookmarkValidator_Setup.bas にあります
'       定数はSetupモジュールで Public として定義されています
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

    ' iTextSharp.dllの確認
    Dim dllPath As String
    dllPath = ThisWorkbook.Path & "\lib\itextsharp.dll"

    If Not fso.FileExists(dllPath) Then
        MsgBox "iTextSharp.dll が見つかりません。" & vbCrLf & vbCrLf & _
               "以下のパスに配置してください:" & vbCrLf & _
               dllPath & vbCrLf & vbCrLf & _
               "詳細はREADME.mdを参照してください。", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.StatusBar = "PDFを解析中..."

    ' PowerShellスクリプト実行
    Dim psScript As String
    psScript = BuildValidationScript(config, dllPath)

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

'==============================================================================
' PowerShellスクリプト生成
'==============================================================================
Private Function BuildValidationScript(config As Object, dllPath As String) As String
    Dim script As String

    script = "$ErrorActionPreference = 'Stop'" & vbCrLf
    script = script & "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8" & vbCrLf
    script = script & vbCrLf

    ' iTextSharp DLLの読み込み
    script = script & "# iTextSharp DLLの読み込み" & vbCrLf
    script = script & "try {" & vbCrLf
    script = script & "    Add-Type -Path '" & Replace(dllPath, "'", "''") & "'" & vbCrLf
    script = script & "} catch {" & vbCrLf
    script = script & "    Write-Output ""ERROR: iTextSharp.dll の読み込みに失敗しました: $($_.Exception.Message)""" & vbCrLf
    script = script & "    exit 1" & vbCrLf
    script = script & "}" & vbCrLf
    script = script & vbCrLf

    ' PDFを開く
    script = script & "# PDFを開く" & vbCrLf
    script = script & "$pdfPath = '" & Replace(config("PDFPath"), "'", "''") & "'" & vbCrLf
    script = script & "try {" & vbCrLf
    script = script & "    $reader = New-Object iTextSharp.text.pdf.PdfReader($pdfPath)" & vbCrLf
    script = script & "} catch {" & vbCrLf
    script = script & "    Write-Output ""ERROR: PDFファイルを開けませんでした: $($_.Exception.Message)""" & vbCrLf
    script = script & "    exit 1" & vbCrLf
    script = script & "}" & vbCrLf
    script = script & vbCrLf

    ' しおり（アウトライン）の取得
    script = script & "# しおりの取得" & vbCrLf
    script = script & "$bookmarks = [iTextSharp.text.pdf.SimpleBookmark]::GetBookmark($reader)" & vbCrLf
    script = script & "if ($null -eq $bookmarks -or $bookmarks.Count -eq 0) {" & vbCrLf
    script = script & "    Write-Output ""ERROR: このPDFにはしおりがありません""" & vbCrLf
    script = script & "    $reader.Close()" & vbCrLf
    script = script & "    exit 1" & vbCrLf
    script = script & "}" & vbCrLf
    script = script & vbCrLf

    ' しおりを再帰的に処理する関数
    script = script & "# しおり処理関数" & vbCrLf
    script = script & "function Process-Bookmark {" & vbCrLf
    script = script & "    param($bookmark, $level, $reader, $checkText)" & vbCrLf
    script = script & vbCrLf
    script = script & "    $title = $bookmark['Title']" & vbCrLf
    script = script & "    $action = $bookmark['Action']" & vbCrLf
    script = script & "    $page = ''" & vbCrLf
    script = script & "    $pageText = ''" & vbCrLf
    script = script & "    $matchRatio = 0" & vbCrLf
    script = script & vbCrLf
    script = script & "    # ページ番号の取得" & vbCrLf
    script = script & "    if ($bookmark.ContainsKey('Page')) {" & vbCrLf
    script = script & "        $pageInfo = $bookmark['Page'] -split ' '" & vbCrLf
    script = script & "        $page = $pageInfo[0]" & vbCrLf
    script = script & "    } elseif ($bookmark.ContainsKey('Named')) {" & vbCrLf
    script = script & "        $page = 'Named: ' + $bookmark['Named']" & vbCrLf
    script = script & "    }" & vbCrLf
    script = script & vbCrLf

    ' テキスト一致確認
    If config("CheckText") = "はい" Then
        script = script & "    # ページテキストの取得と一致確認" & vbCrLf
        script = script & "    if ($checkText -and $page -match '^\d+$') {" & vbCrLf
        script = script & "        $pageNum = [int]$page" & vbCrLf
        script = script & "        if ($pageNum -le $reader.NumberOfPages) {" & vbCrLf
        script = script & "            $strategy = New-Object iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy" & vbCrLf
        script = script & "            $text = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($reader, $pageNum, $strategy)" & vbCrLf
        script = script & "            # 最初の200文字を取得" & vbCrLf
        script = script & "            if ($text.Length -gt 200) {" & vbCrLf
        script = script & "                $pageText = $text.Substring(0, 200) -replace '[\r\n]+', ' '" & vbCrLf
        script = script & "            } else {" & vbCrLf
        script = script & "                $pageText = $text -replace '[\r\n]+', ' '" & vbCrLf
        script = script & "            }" & vbCrLf
        script = script & vbCrLf
        script = script & "            # 一致率の計算（しおり名がテキストに含まれるか）" & vbCrLf
        script = script & "            $titleClean = $title -replace '\s+', ''" & vbCrLf
        script = script & "            $textClean = $text -replace '\s+', ''" & vbCrLf
        script = script & "            if ($textClean -like ""*$titleClean*"") {" & vbCrLf
        script = script & "                $matchRatio = 100" & vbCrLf
        script = script & "            } else {" & vbCrLf
        script = script & "                # 部分一致のチェック" & vbCrLf
        script = script & "                $words = $title -split '\s+'" & vbCrLf
        script = script & "                $matchCount = 0" & vbCrLf
        script = script & "                foreach ($word in $words) {" & vbCrLf
        script = script & "                    if ($word.Length -gt 1 -and $textClean -like ""*$word*"") {" & vbCrLf
        script = script & "                        $matchCount++" & vbCrLf
        script = script & "                    }" & vbCrLf
        script = script & "                }" & vbCrLf
        script = script & "                if ($words.Count -gt 0) {" & vbCrLf
        script = script & "                    $matchRatio = [math]::Round(($matchCount / $words.Count) * 100)" & vbCrLf
        script = script & "                }" & vbCrLf
        script = script & "            }" & vbCrLf
        script = script & "        }" & vbCrLf
        script = script & "    }" & vbCrLf
    End If

    script = script & vbCrLf
    script = script & "    # 結果出力（タブ区切り）" & vbCrLf
    script = script & "    Write-Output ""BOOKMARK`t$title`t$level`t$page`t$pageText`t$matchRatio""" & vbCrLf
    script = script & vbCrLf
    script = script & "    # 子しおりの処理" & vbCrLf
    script = script & "    if ($bookmark.ContainsKey('Kids')) {" & vbCrLf
    script = script & "        foreach ($child in $bookmark['Kids']) {" & vbCrLf
    script = script & "            Process-Bookmark -bookmark $child -level ($level + 1) -reader $reader -checkText $checkText" & vbCrLf
    script = script & "        }" & vbCrLf
    script = script & "    }" & vbCrLf
    script = script & "}" & vbCrLf
    script = script & vbCrLf

    ' メイン処理
    script = script & "# メイン処理" & vbCrLf
    script = script & "Write-Output ""INFO: 総ページ数: $($reader.NumberOfPages)""" & vbCrLf
    script = script & "Write-Output ""INFO: しおり数: $($bookmarks.Count)""" & vbCrLf
    script = script & vbCrLf
    script = script & "foreach ($bookmark in $bookmarks) {" & vbCrLf
    script = script & "    Process-Bookmark -bookmark $bookmark -level 1 -reader $reader -checkText $" & LCase(CStr(config("CheckText") = "はい")) & vbCrLf
    script = script & "}" & vbCrLf
    script = script & vbCrLf
    script = script & "$reader.Close()" & vbCrLf
    script = script & "Write-Output ""COMPLETE""" & vbCrLf

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

            If UBound(parts) >= 4 Then
                bookmarkNo = bookmarkNo + 1

                ws.Cells(row, COL_NO).Value = bookmarkNo
                ws.Cells(row, COL_BOOKMARK_NAME).Value = parts(0) ' タイトル
                ws.Cells(row, COL_BOOKMARK_LEVEL).Value = parts(1) ' 階層
                ws.Cells(row, COL_LINK_PAGE).Value = parts(2) ' ページ

                If UBound(parts) >= 3 Then
                    ws.Cells(row, COL_PAGE_TEXT).Value = Left(parts(3), 100) ' ページテキスト（100文字まで）
                End If

                If UBound(parts) >= 4 Then
                    Dim matchRatio As Long
                    If IsNumeric(parts(4)) Then
                        matchRatio = CLng(parts(4))
                    Else
                        matchRatio = 0
                    End If

                    ws.Cells(row, COL_MATCH_RATIO).Value = matchRatio & "%"

                    ' テキスト一致判定
                    If config("CheckText") = "はい" Then
                        If matchRatio >= threshold Then
                            ws.Cells(row, COL_TEXT_MATCH).Value = "一致"
                            ws.Cells(row, COL_TEXT_MATCH).Interior.Color = RGB(198, 239, 206)
                        ElseIf matchRatio > 0 Then
                            ws.Cells(row, COL_TEXT_MATCH).Value = "部分一致"
                            ws.Cells(row, COL_TEXT_MATCH).Interior.Color = RGB(255, 235, 156)
                        Else
                            ws.Cells(row, COL_TEXT_MATCH).Value = "不一致"
                            ws.Cells(row, COL_TEXT_MATCH).Interior.Color = RGB(255, 199, 206)
                        End If
                    End If
                End If

                ' 判定
                Dim status As String
                If parts(2) = "" Or InStr(parts(2), "Named:") > 0 Then
                    status = "確認要"
                    ws.Cells(row, COL_STATUS).Interior.Color = RGB(255, 235, 156)
                ElseIf config("CheckText") = "はい" And matchRatio < threshold Then
                    status = "NG"
                    ws.Cells(row, COL_STATUS).Interior.Color = RGB(255, 199, 206)
                Else
                    status = "OK"
                    ws.Cells(row, COL_STATUS).Interior.Color = RGB(198, 239, 206)
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
        MsgBox "しおりが取得できませんでした。", vbExclamation
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
            Dim row As Long
            For row = ROW_RESULT_DATA_START To lastRow
                Dim csvLine As String
                csvLine = ws.Cells(row, COL_NO).Value & "," & _
                          """" & Replace(ws.Cells(row, COL_BOOKMARK_NAME).Value, """", """""") & """," & _
                          ws.Cells(row, COL_BOOKMARK_LEVEL).Value & "," & _
                          ws.Cells(row, COL_LINK_PAGE).Value & "," & _
                          """" & Replace(ws.Cells(row, COL_PAGE_TEXT).Value, """", """""") & """," & _
                          ws.Cells(row, COL_TEXT_MATCH).Value & "," & _
                          ws.Cells(row, COL_MATCH_RATIO).Value & "," & _
                          ws.Cells(row, COL_STATUS).Value
                ts.WriteLine csvLine
            Next row

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
