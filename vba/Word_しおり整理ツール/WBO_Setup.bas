Attribute VB_Name = "WBO_Setup"
Option Explicit

' ============================================================================
' Word しおり整理ツール - セットアップモジュール
' シート作成、UI構築、フォーマット設定を行う
' ※ 定数はWBO_Configで定義。シート作成後は本モジュールを削除可能
' ============================================================================

' ============================================================================
' メイン初期化プロシージャ
' ============================================================================
Public Sub InitializeWordしおり整理ツール()
    Application.ScreenUpdating = False

    On Error GoTo ErrorHandler

    ' 既存シートがあれば削除
    Dim ws As Worksheet
    Dim sheetNames As Variant
    Dim sheetName As Variant
    sheetNames = Array(SHEET_MAIN, SHEET_SETTINGS)

    For Each sheetName In sheetNames
        For Each ws In ThisWorkbook.Worksheets
            If ws.Name = CStr(sheetName) Then
                Application.DisplayAlerts = False
                ws.Delete
                Application.DisplayAlerts = True
                Exit For
            End If
        Next ws
    Next sheetName

    ' 設定シート作成
    Dim wsSettings As Worksheet
    Set wsSettings = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsSettings.Name = SHEET_SETTINGS
    FormatSettingsSheet wsSettings

    ' メインシート作成
    Dim wsMain As Worksheet
    Set wsMain = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsMain.Name = SHEET_MAIN
    FormatMainSheet wsMain

    ' メインシートを表示
    wsMain.Activate

    Application.ScreenUpdating = True

    MsgBox "初期化が完了しました。" & vbCrLf & vbCrLf & _
           "「設定」シートでフォルダパスとスタイル設定を確認してください。", _
           vbInformation, "Word しおり整理ツール"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "初期化中にエラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub

' ============================================================================
' 設定シートのフォーマット
' ============================================================================
Private Sub FormatSettingsSheet(ByRef ws As Worksheet)
    Dim macroDir As String
    macroDir = ThisWorkbook.Path

    With ws
        .Cells.Interior.Color = RGB(255, 255, 255)

        ' === フォルダ設定セクション ===
        .Cells(SETTINGS_ROW_FOLDER_HEADER, SETTINGS_COL_LABEL).Value = "■ フォルダ設定"
        With .Cells(SETTINGS_ROW_FOLDER_HEADER, SETTINGS_COL_LABEL)
            .Font.Name = "Meiryo UI"
            .Font.Bold = True
            .Font.Size = 12
        End With

        ' 入力フォルダ
        .Cells(SETTINGS_ROW_INPUT_FOLDER, SETTINGS_COL_LABEL).Value = "入力フォルダ"
        .Cells(SETTINGS_ROW_INPUT_FOLDER, SETTINGS_COL_LABEL).Font.Name = "Meiryo UI"
        .Range(.Cells(SETTINGS_ROW_INPUT_FOLDER, SETTINGS_COL_VALUE), _
               .Cells(SETTINGS_ROW_INPUT_FOLDER, SETTINGS_COL_NOTE)).Merge
        .Cells(SETTINGS_ROW_INPUT_FOLDER, SETTINGS_COL_VALUE).Value = macroDir & "\Input\"
        .Cells(SETTINGS_ROW_INPUT_FOLDER, SETTINGS_COL_VALUE).Font.Name = "Meiryo UI"
        .Cells(SETTINGS_ROW_INPUT_FOLDER, SETTINGS_COL_VALUE).Interior.Color = RGB(255, 255, 204)

        ' 出力フォルダ
        .Cells(SETTINGS_ROW_OUTPUT_FOLDER, SETTINGS_COL_LABEL).Value = "出力フォルダ"
        .Cells(SETTINGS_ROW_OUTPUT_FOLDER, SETTINGS_COL_LABEL).Font.Name = "Meiryo UI"
        .Range(.Cells(SETTINGS_ROW_OUTPUT_FOLDER, SETTINGS_COL_VALUE), _
               .Cells(SETTINGS_ROW_OUTPUT_FOLDER, SETTINGS_COL_NOTE)).Merge
        .Cells(SETTINGS_ROW_OUTPUT_FOLDER, SETTINGS_COL_VALUE).Value = macroDir & "\Output\"
        .Cells(SETTINGS_ROW_OUTPUT_FOLDER, SETTINGS_COL_VALUE).Font.Name = "Meiryo UI"
        .Cells(SETTINGS_ROW_OUTPUT_FOLDER, SETTINGS_COL_VALUE).Interior.Color = RGB(255, 255, 204)

        ' === スタイル設定セクション ===
        .Cells(SETTINGS_ROW_STYLE_HEADER - 1, SETTINGS_COL_LABEL).Value = "■ スタイル設定（行を追加して設定を増やせます）"
        With .Cells(SETTINGS_ROW_STYLE_HEADER - 1, SETTINGS_COL_LABEL)
            .Font.Name = "Meiryo UI"
            .Font.Bold = True
            .Font.Size = 12
        End With

        ' ヘッダー行
        .Cells(SETTINGS_ROW_STYLE_HEADER, SETTINGS_COL_LABEL).Value = "種別"
        .Cells(SETTINGS_ROW_STYLE_HEADER, SETTINGS_COL_VALUE).Value = "レベル"
        .Cells(SETTINGS_ROW_STYLE_HEADER, SETTINGS_COL_PATTERN).Value = "パターン/テキスト"
        .Cells(SETTINGS_ROW_STYLE_HEADER, SETTINGS_COL_STYLE).Value = "適用スタイル"
        .Cells(SETTINGS_ROW_STYLE_HEADER, SETTINGS_COL_NOTE).Value = "備考"

        With .Range(.Cells(SETTINGS_ROW_STYLE_HEADER, SETTINGS_COL_LABEL), _
                    .Cells(SETTINGS_ROW_STYLE_HEADER, SETTINGS_COL_NOTE))
            .Font.Name = "Meiryo UI"
            .Font.Bold = True
            .Interior.Color = RGB(180, 198, 231)
            .HorizontalAlignment = xlCenter
        End With

        ' デフォルトのスタイル設定を追加
        Dim row As Long
        row = SETTINGS_ROW_STYLE_START

        AddStyleRow ws, row, "パターン", "1", "^第[0-9０-９]+部", "表題1", "第X部（ヘッダー空白時のみ）"
        row = row + 1
        AddStyleRow ws, row, "パターン", "2", "^第[0-9０-９]+章", "表題2", "第X章"
        row = row + 1
        AddStyleRow ws, row, "パターン", "3-節", "^第[0-9０-９]+節", "表題3", "第X節（節構造あり時）"
        row = row + 1
        AddStyleRow ws, row, "パターン", "3", "^[0-9]+-[0-9]+(?![,\.0-9])", "表題3", "X-X（節構造なし時）"
        row = row + 1
        AddStyleRow ws, row, "パターン", "4-節", "^[0-9]+-[0-9]+(?![,\.0-9])", "表題4", "X-X（節構造あり時）"
        row = row + 1
        AddStyleRow ws, row, "パターン", "4", "^[0-9]+-[0-9]+[,\.][0-9]+", "表題4", "X-X.X（節構造なし時）"
        row = row + 1
        AddStyleRow ws, row, "パターン", "5-節", "^[0-9]+-[0-9]+[,\.][0-9]+", "表題5", "X-X.X（節構造あり時）"
        row = row + 1
        AddStyleRow ws, row, "帳票", "", "\([A-Za-z][0-9]{3}\)", "表題5", "(X123)形式"
        row = row + 1
        AddStyleRow ws, row, "帳票", "", "\([A-Za-z]{2}[0-9]{2}\)", "表題5", "(XX12)形式"
        row = row + 1
        AddStyleRow ws, row, "特定", "1", "本書の記述について", "表題3", "完全一致、アウトラインレベル1"
        row = row + 1
        AddStyleRow ws, row, "特定", "1", "修正履歴", "表題3", "完全一致、アウトラインレベル1"
        row = row + 1
        AddStyleRow ws, row, "例外", "1", "", "本文", "パターン外で見出しスタイル適用済み"
        row = row + 1
        AddStyleRow ws, row, "例外", "2", "", "本文", "アウトライン設定済み"
        row = row + 1

        ' 追加用の空行を5行追加
        Dim i As Long
        For i = 1 To 5
            AddStyleRow ws, row, "", "", "", "", ""
            row = row + 1
        Next i

        ' テーブル罫線
        With .Range(.Cells(SETTINGS_ROW_STYLE_HEADER, SETTINGS_COL_LABEL), _
                    .Cells(row - 1, SETTINGS_COL_NOTE))
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With

        ' === オプション設定セクション ===
        .Cells(SETTINGS_ROW_OPTION_HEADER, SETTINGS_COL_LABEL).Value = "■ オプション設定"
        With .Cells(SETTINGS_ROW_OPTION_HEADER, SETTINGS_COL_LABEL)
            .Font.Name = "Meiryo UI"
            .Font.Bold = True
            .Font.Size = 12
        End With

        .Cells(SETTINGS_ROW_PDF_OUTPUT, SETTINGS_COL_LABEL).Value = "PDF出力"
        .Cells(SETTINGS_ROW_PDF_OUTPUT, SETTINGS_COL_LABEL).Font.Name = "Meiryo UI"
        .Cells(SETTINGS_ROW_PDF_OUTPUT, SETTINGS_COL_VALUE).Value = "はい"
        AddDropdown ws, .Cells(SETTINGS_ROW_PDF_OUTPUT, SETTINGS_COL_VALUE), "はい,いいえ"
        .Cells(SETTINGS_ROW_PDF_OUTPUT, SETTINGS_COL_VALUE).Interior.Color = RGB(255, 255, 204)
        .Cells(SETTINGS_ROW_PDF_OUTPUT, SETTINGS_COL_VALUE).Font.Name = "Meiryo UI"

        ' === 種別の説明 ===
        Dim noteRow As Long
        noteRow = SETTINGS_ROW_PDF_OUTPUT + 3

        .Cells(noteRow, SETTINGS_COL_LABEL).Value = "■ 種別の説明"
        With .Cells(noteRow, SETTINGS_COL_LABEL)
            .Font.Name = "Meiryo UI"
            .Font.Bold = True
            .Font.Size = 12
        End With
        noteRow = noteRow + 1

        .Cells(noteRow, SETTINGS_COL_LABEL).Value = "パターン"
        .Cells(noteRow, SETTINGS_COL_VALUE).Value = "正規表現でテキストを判定。レベル列に数字を指定。"
        noteRow = noteRow + 1
        .Cells(noteRow, SETTINGS_COL_LABEL).Value = ""
        .Cells(noteRow, SETTINGS_COL_VALUE).Value = "「X-節」は節構造あり時、数字のみは節構造なし時に適用。"
        noteRow = noteRow + 1
        .Cells(noteRow, SETTINGS_COL_LABEL).Value = "帳票"
        .Cells(noteRow, SETTINGS_COL_VALUE).Value = "1ページ目に「帳票」がある文書のみ適用。"
        noteRow = noteRow + 1
        .Cells(noteRow, SETTINGS_COL_LABEL).Value = "特定"
        .Cells(noteRow, SETTINGS_COL_VALUE).Value = "テキスト完全一致で適用。レベル列の数字でアウトラインレベルを設定。"
        noteRow = noteRow + 1
        .Cells(noteRow, SETTINGS_COL_LABEL).Value = "例外"
        .Cells(noteRow, SETTINGS_COL_VALUE).Value = "例外1=見出しスタイル適用済み、例外2=アウトライン設定済み"
        noteRow = noteRow + 1

        .Range(.Cells(SETTINGS_ROW_PDF_OUTPUT + 4, SETTINGS_COL_LABEL), _
               .Cells(noteRow - 1, SETTINGS_COL_VALUE)).Font.Name = "Meiryo UI"
        .Range(.Cells(SETTINGS_ROW_PDF_OUTPUT + 4, SETTINGS_COL_LABEL), _
               .Cells(noteRow - 1, SETTINGS_COL_VALUE)).Font.Size = 10

        ' === 列幅調整 ===
        .Columns("A").ColumnWidth = 3
        .Columns("B").ColumnWidth = 12
        .Columns("C").ColumnWidth = 10
        .Columns("D").ColumnWidth = 30
        .Columns("E").ColumnWidth = 15
        .Columns("F").ColumnWidth = 35

        .Range("A1").Select
    End With
End Sub

' ============================================================================
' スタイル設定行を追加
' ============================================================================
Private Sub AddStyleRow(ByRef ws As Worksheet, ByVal row As Long, _
                        ByVal category As String, ByVal Level As String, _
                        ByVal Pattern As String, ByVal styleName As String, _
                        ByVal note As String)
    With ws
        .Cells(row, SETTINGS_COL_LABEL).Value = category
        .Cells(row, SETTINGS_COL_VALUE).Value = Level
        .Cells(row, SETTINGS_COL_PATTERN).Value = Pattern
        .Cells(row, SETTINGS_COL_STYLE).Value = styleName
        .Cells(row, SETTINGS_COL_NOTE).Value = note

        .Range(.Cells(row, SETTINGS_COL_LABEL), .Cells(row, SETTINGS_COL_NOTE)).Font.Name = "Meiryo UI"

        .Cells(row, SETTINGS_COL_LABEL).Interior.Color = RGB(255, 255, 204)
        .Cells(row, SETTINGS_COL_VALUE).Interior.Color = RGB(255, 255, 204)
        .Cells(row, SETTINGS_COL_PATTERN).Interior.Color = RGB(255, 255, 204)
        .Cells(row, SETTINGS_COL_STYLE).Interior.Color = RGB(255, 255, 204)
        .Cells(row, SETTINGS_COL_NOTE).Interior.Color = RGB(230, 230, 230)

        If category <> "" Then
            AddDropdown ws, .Cells(row, SETTINGS_COL_LABEL), "パターン,帳票,特定,例外"
        End If
    End With
End Sub

' ============================================================================
' メインシートのフォーマット
' ============================================================================
Private Sub FormatMainSheet(ByRef ws As Worksheet)
    With ws
        .Cells.Interior.Color = RGB(255, 255, 255)

        ' === タイトルエリア ===
        .Range("B2:G3").Merge
        .Range("B2").Value = "Word しおり整理ツール"
        With .Range("B2:G3")
            .Font.Name = "Meiryo UI"
            .Font.Size = 20
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(68, 114, 196)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        .Rows(2).RowHeight = 35
        .Rows(3).RowHeight = 10

        ' === 説明エリア ===
        .Range("B5").Value = "段落テキストをパターンマッチでスタイル適用します。"
        .Range("B6").Value = "PDFエクスポート時に正しいしおり（ブックマーク）を生成します。"
        .Range("B5:B6").Font.Name = "Meiryo UI"
        .Range("B5:B6").Font.Size = 11

        .Range("B8").Value = "※ フォルダパス・スタイル設定は「設定」シートで変更してください"
        .Range("B8").Font.Name = "Meiryo UI"
        .Range("B8").Font.Size = 10
        .Range("B8").Font.Color = RGB(0, 112, 192)

        ' === ボタン配置 ===
        AddButton ws, .Range("B10"), 200, 40, "OrganizeWordBookmarks", "しおりを整理してPDF出力", RGB(68, 114, 196)

        ' === 使い方セクション ===
        .Range("B14").Value = "■ 使い方"
        .Range("B14").Font.Name = "Meiryo UI"
        .Range("B14").Font.Bold = True
        .Range("B14").Font.Size = 12

        .Range("B16").Value = "1. 「設定」シートでフォルダパスとスタイル設定を確認・編集します"
        .Range("B17").Value = "2. 処理したいWord文書(.docx/.doc)を入力フォルダに配置します"
        .Range("B18").Value = "3. 「しおりを整理してPDF出力」ボタンをクリックします"
        .Range("B19").Value = "4. 出力フォルダに処理済みのWord文書とPDFが出力されます"
        .Range("B16:B19").Font.Name = "Meiryo UI"
        .Range("B16:B19").Font.Size = 10

        ' === 動作説明セクション ===
        .Range("B22").Value = "■ 動作の説明"
        .Range("B22").Font.Name = "Meiryo UI"
        .Range("B22").Font.Bold = True
        .Range("B22").Font.Size = 12

        .Range("B24").Value = "【パターンマッチ方式】"
        .Range("B24").Font.Bold = True
        .Range("B25").Value = "  段落テキストを正規表現でパターンマッチし、該当するスタイルを適用します。"
        .Range("B26").Value = "  設定シートで自由にパターンとスタイルの組み合わせを追加できます。"

        .Range("B28").Value = "【スキップ条件】"
        .Range("B28").Font.Bold = True
        .Range("B29").Value = "  ・「参照」という文字を含む段落"
        .Range("B30").Value = "  ・「・」（中黒）で始まる段落（目次形式など）"
        .Range("B31").Value = "  ・ハイパーリンクを含む段落、表内の段落"

        .Range("B33").Value = "【節構造の自動判定】"
        .Range("B33").Font.Bold = True
        .Range("B34").Value = "  文書のヘッダーに「第X節」があるかを判定し、適用パターンを自動で切り替えます。"
        .Range("B35").Value = "  設定シートのレベル欄で「X-節」と指定すると節構造あり時のみ適用されます。"

        .Range("B24:B35").Font.Name = "Meiryo UI"
        .Range("B24:B35").Font.Size = 10

        ' === 列幅調整 ===
        .Columns("A").ColumnWidth = 3
        .Columns("B").ColumnWidth = 80

        .Rows(10).RowHeight = 45

        .Range("A1").Select
    End With
End Sub

' ============================================================================
' ドロップダウンリストの追加
' ============================================================================
Private Sub AddDropdown(ByRef ws As Worksheet, ByRef cell As Range, ByVal options As String)
    With cell.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:=options
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = False
        .ShowError = True
    End With
End Sub

' ============================================================================
' ボタンの追加（図形ボタン）
' ============================================================================
Private Sub AddButton(ByRef ws As Worksheet, ByRef cell As Range, _
                      ByVal width As Double, ByVal height As Double, _
                      ByVal macroName As String, ByVal caption As String, _
                      ByVal fillColor As Long)
    Dim btn As Shape
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
                                 cell.Left, cell.Top, width, height)

    With btn
        .Name = "btn" & macroName
        .Fill.ForeColor.RGB = fillColor
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Characters.Text = caption
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.Font.Size = 12
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .OnAction = macroName
        .Placement = xlFreeFloating
    End With
End Sub
