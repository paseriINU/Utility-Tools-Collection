Option Explicit

' ============================================================================
' Word しおり整理ツール - セットアップモジュール
' シート作成、UI構築、フォーマット設定を行う
' ============================================================================

' === シート名定数 ===
Public Const SHEET_MAIN As String = "Word_しおり整理ツール"

' === セル位置定数 ===
' パターン設定テーブル
Public Const ROW_PATTERN_HEADER As Long = 19
Public Const ROW_PATTERN_LEVEL1 As Long = 20
Public Const ROW_PATTERN_LEVEL2 As Long = 21
Public Const ROW_PATTERN_LEVEL3 As Long = 22   ' 第X節/X-X（節の有無で自動切替）
Public Const ROW_PATTERN_LEVEL4 As Long = 23   ' X-X/X-X,X（節の有無で自動切替）
Public Const ROW_PATTERN_LEVEL5 As Long = 24   ' X-X,X（節がある場合のみ使用）
Public Const ROW_PATTERN_EXCEPTION1 As Long = 25
Public Const ROW_PATTERN_EXCEPTION2 As Long = 26

Public Const COL_LEVEL As Long = 2          ' B列
Public Const COL_PATTERN_DESC As Long = 3   ' C列
Public Const COL_STYLE_NAME As Long = 4     ' D列

' オプション設定
Public Const ROW_OPTION_PDF_OUTPUT As Long = 28
Public Const COL_OPTION_LABEL As Long = 2   ' B列
Public Const COL_OPTION_VALUE As Long = 3   ' C列

' ボタン行
Public Const ROW_BUTTON As Long = 30

' フォルダパス表示
Public Const ROW_INPUT_FOLDER As Long = 10
Public Const ROW_OUTPUT_FOLDER As Long = 12

' ============================================================================
' メイン初期化プロシージャ
' ============================================================================
Public Sub InitializeWordしおり整理ツール()
    Application.ScreenUpdating = False

    On Error GoTo ErrorHandler

    ' 既存シートがあれば削除
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = SHEET_MAIN Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
            Exit For
        End If
    Next ws

    ' 新規シート作成
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = SHEET_MAIN

    ' フォーマット適用
    FormatMainSheet ws

    ' Input/Outputフォルダ作成
    CreateFolders

    Application.ScreenUpdating = True

    MsgBox "初期化が完了しました。" & vbCrLf & vbCrLf & _
           "処理したいWord文書をInputフォルダに配置してください。", _
           vbInformation, "Word しおり整理ツール"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "初期化中にエラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub

' ============================================================================
' メインシートのフォーマット
' ============================================================================
Private Sub FormatMainSheet(ByRef ws As Worksheet)
    Dim macroDir As String
    macroDir = ThisWorkbook.Path

    With ws
        ' 全体の背景色を白に
        .Cells.Interior.Color = RGB(255, 255, 255)

        ' === タイトルエリア（行1-3） ===
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

        ' === 説明エリア（行5-6） ===
        .Range("B5").Value = "段落テキストをパターンマッチでスタイル適用します（「参照」・リンク・「・」始まりはスキップ）。"
        .Range("B6").Value = "PDFエクスポート時に正しいしおり（ブックマーク）を生成します。"
        .Range("B5:B6").Font.Name = "Meiryo UI"
        .Range("B5:B6").Font.Size = 11

        ' === フォルダ設定セクション（行8-14） ===
        .Range("B8").Value = "■ フォルダ設定"
        .Range("B8").Font.Name = "Meiryo UI"
        .Range("B8").Font.Bold = True
        .Range("B8").Font.Size = 12

        .Range("B10").Value = "入力フォルダ:"
        .Range("B10").Font.Name = "Meiryo UI"
        .Range("C10:G10").Merge
        .Range("C10").Value = macroDir & "\Input\"
        .Range("C10").Font.Name = "Meiryo UI"
        .Range("C10").Interior.Color = RGB(242, 242, 242)

        .Range("B12").Value = "出力フォルダ:"
        .Range("B12").Font.Name = "Meiryo UI"
        .Range("C12:G12").Merge
        .Range("C12").Value = macroDir & "\Output\"
        .Range("C12").Font.Name = "Meiryo UI"
        .Range("C12").Interior.Color = RGB(242, 242, 242)

        ' === パターン設定セクション（行17-25） ===
        .Range("B17").Value = "■ スタイル設定"
        .Range("B17").Font.Name = "Meiryo UI"
        .Range("B17").Font.Bold = True
        .Range("B17").Font.Size = 12

        ' ヘッダー行
        .Cells(ROW_PATTERN_HEADER, COL_LEVEL).Value = "レベル"
        .Cells(ROW_PATTERN_HEADER, COL_PATTERN_DESC).Value = "テキストパターン"
        .Cells(ROW_PATTERN_HEADER, COL_STYLE_NAME).Value = "適用スタイル"

        With .Range(.Cells(ROW_PATTERN_HEADER, COL_LEVEL), .Cells(ROW_PATTERN_HEADER, COL_STYLE_NAME))
            .Font.Name = "Meiryo UI"
            .Font.Bold = True
            .Interior.Color = RGB(180, 198, 231)
            .HorizontalAlignment = xlCenter
        End With

        ' レベル1
        .Cells(ROW_PATTERN_LEVEL1, COL_LEVEL).Value = "1"
        .Cells(ROW_PATTERN_LEVEL1, COL_PATTERN_DESC).Value = "第X部"
        .Cells(ROW_PATTERN_LEVEL1, COL_STYLE_NAME).Value = "表題1"

        ' レベル2
        .Cells(ROW_PATTERN_LEVEL2, COL_LEVEL).Value = "2"
        .Cells(ROW_PATTERN_LEVEL2, COL_PATTERN_DESC).Value = "第X章"
        .Cells(ROW_PATTERN_LEVEL2, COL_STYLE_NAME).Value = "表題2"

        ' レベル3（節あり:第X節、節なし:X-X）
        .Cells(ROW_PATTERN_LEVEL3, COL_LEVEL).Value = "3"
        .Cells(ROW_PATTERN_LEVEL3, COL_PATTERN_DESC).Value = "第X節 / X-X"
        .Cells(ROW_PATTERN_LEVEL3, COL_STYLE_NAME).Value = "表題3"

        ' レベル4（節あり:X-X、節なし:X-X.X）
        .Cells(ROW_PATTERN_LEVEL4, COL_LEVEL).Value = "4"
        .Cells(ROW_PATTERN_LEVEL4, COL_PATTERN_DESC).Value = "X-X / X-X.X"
        .Cells(ROW_PATTERN_LEVEL4, COL_STYLE_NAME).Value = "表題4"

        ' レベル5（節がある場合のみ使用）
        .Cells(ROW_PATTERN_LEVEL5, COL_LEVEL).Value = "5"
        .Cells(ROW_PATTERN_LEVEL5, COL_PATTERN_DESC).Value = "X-X.X（※節あり時）"
        .Cells(ROW_PATTERN_LEVEL5, COL_STYLE_NAME).Value = "表題5"

        ' 例外1
        .Cells(ROW_PATTERN_EXCEPTION1, COL_LEVEL).Value = "例外1"
        .Cells(ROW_PATTERN_EXCEPTION1, COL_PATTERN_DESC).Value = "パターン外スタイル"
        .Cells(ROW_PATTERN_EXCEPTION1, COL_STYLE_NAME).Value = "本文"

        ' 例外2
        .Cells(ROW_PATTERN_EXCEPTION2, COL_LEVEL).Value = "例外2"
        .Cells(ROW_PATTERN_EXCEPTION2, COL_PATTERN_DESC).Value = "アウトライン設定済み"
        .Cells(ROW_PATTERN_EXCEPTION2, COL_STYLE_NAME).Value = "本文"

        ' テーブル全体のフォント設定
        With .Range(.Cells(ROW_PATTERN_LEVEL1, COL_LEVEL), .Cells(ROW_PATTERN_EXCEPTION2, COL_STYLE_NAME))
            .Font.Name = "Meiryo UI"
        End With

        ' テーブル罫線
        With .Range(.Cells(ROW_PATTERN_HEADER, COL_LEVEL), .Cells(ROW_PATTERN_EXCEPTION2, COL_STYLE_NAME))
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With

        ' 入力セルの背景色（黄色）
        With .Range(.Cells(ROW_PATTERN_LEVEL1, COL_STYLE_NAME), .Cells(ROW_PATTERN_EXCEPTION2, COL_STYLE_NAME))
            .Interior.Color = RGB(255, 255, 204)
        End With

        ' === オプション設定セクション（行27） ===
        .Cells(ROW_OPTION_PDF_OUTPUT, COL_OPTION_LABEL).Value = "PDF出力:"
        .Cells(ROW_OPTION_PDF_OUTPUT, COL_OPTION_LABEL).Font.Name = "Meiryo UI"
        .Cells(ROW_OPTION_PDF_OUTPUT, COL_OPTION_VALUE).Value = "はい"
        AddDropdown ws, .Cells(ROW_OPTION_PDF_OUTPUT, COL_OPTION_VALUE), "はい,いいえ"

        With .Cells(ROW_OPTION_PDF_OUTPUT, COL_OPTION_VALUE)
            .Interior.Color = RGB(255, 255, 204)
            .Font.Name = "Meiryo UI"
        End With

        ' === ボタン配置（行30） ===
        AddButton ws, .Range("B" & ROW_BUTTON), 180, 35, "OrganizeWordBookmarks", "しおりを整理してPDF出力", RGB(68, 114, 196)
        AddButton ws, .Range("F" & ROW_BUTTON), 80, 35, "ResetSettings", "リセット", RGB(128, 128, 128)

        ' === 使い方セクション（行34-39） ===
        .Range("B34").Value = "■ 使い方"
        .Range("B34").Font.Name = "Meiryo UI"
        .Range("B34").Font.Bold = True
        .Range("B34").Font.Size = 12

        .Range("B36").Value = "1. 処理したいWord文書(.docx/.doc)をInputフォルダに配置します"
        .Range("B37").Value = "2. 「適用スタイル」欄にWord文書で使用するスタイル名を入力します"
        .Range("B38").Value = "3. 「しおりを整理してPDF出力」ボタンをクリックします"
        .Range("B39").Value = "4. Outputフォルダに処理済みのWord文書とPDFが出力されます"
        .Range("B36:B39").Font.Name = "Meiryo UI"
        .Range("B36:B39").Font.Size = 10

        ' === 動作説明セクション（行43-60） ===
        .Range("B43").Value = "■ 動作の説明"
        .Range("B43").Font.Name = "Meiryo UI"
        .Range("B43").Font.Bold = True
        .Range("B43").Font.Size = 12

        .Range("B45").Value = "【パターンマッチ方式】（レベル1-5）"
        .Range("B45").Font.Bold = True
        .Range("B46").Value = "  段落テキストを正規表現でパターンマッチし、該当するスタイルを適用します。"
        .Range("B47").Value = "  例: 「第1章　概要」→ 第X章パターンに一致 → 表題2スタイル適用"
        .Range("B47").Font.Color = RGB(0, 112, 192)

        .Range("B49").Value = "【スキップ条件】"
        .Range("B49").Font.Bold = True
        .Range("B50").Value = "  「参照」を含む段落、「・」で始まる段落、ハイパーリンクを含む段落はスキップします。"
        .Range("B51").Value = "  第X部はヘッダー空欄時のみ処理します。"

        .Range("B53").Value = "【節構造の自動判定】"
        .Range("B53").Font.Bold = True
        .Range("B54").Value = "  ヘッダーに「第X節」があるかを事前に判定し、レベル構造を自動で切り替えます。"

        .Range("B56").Value = "【特定テキストの例外処理】"
        .Range("B56").Font.Bold = True
        .Range("B57").Value = "  「本書の記述について」「修正履歴」はレベル3スタイル適用、アウトラインはレベル1に設定。"

        .Range("B59").Value = "【ヘッダーフィールド更新】"
        .Range("B59").Font.Bold = True
        .Range("B60").Value = "  スタイル適用後、ヘッダー内のSTYLEREFフィールドのスタイル名を自動更新します。"

        .Range("B62").Value = "※ 図形（テキストボックス等）内のテキストも処理対象です"
        .Range("B62").Font.Color = RGB(0, 112, 192)

        .Range("B45:B62").Font.Name = "Meiryo UI"
        .Range("B45:B62").Font.Size = 10

        ' === 列幅調整 ===
        .Columns("A").ColumnWidth = 3
        .Columns("B").ColumnWidth = 18
        .Columns("C").ColumnWidth = 20
        .Columns("D").ColumnWidth = 15
        .Columns("E").ColumnWidth = 12
        .Columns("F").ColumnWidth = 12

        ' 行の高さ調整
        .Rows(ROW_BUTTON).RowHeight = 40

        ' 入力可能セルをアンロック（シート保護時用）
        .Range(.Cells(ROW_PATTERN_LEVEL1, COL_STYLE_NAME), .Cells(ROW_PATTERN_EXCEPTION2, COL_STYLE_NAME)).Locked = False
        .Cells(ROW_OPTION_PDF_OUTPUT, COL_OPTION_VALUE).Locked = False

        ' A1セルを選択
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
        .TextFrame2.TextRange.Font.Size = 11
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .OnAction = macroName
        ' セルサイズに依存しない固定配置
        .Placement = xlFreeFloating
    End With
End Sub

' ============================================================================
' Input/Outputフォルダの作成
' ============================================================================
Private Sub CreateFolders()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim macroDir As String
    macroDir = ThisWorkbook.Path

    ' Inputフォルダ
    If Not fso.FolderExists(macroDir & "\Input") Then
        fso.CreateFolder macroDir & "\Input"
    End If

    ' Outputフォルダ
    If Not fso.FolderExists(macroDir & "\Output") Then
        fso.CreateFolder macroDir & "\Output"
    End If

    Set fso = Nothing
End Sub

' ============================================================================
' 設定リセット
' ============================================================================
Public Sub ResetSettings()
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_MAIN)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "シートが見つかりません。初期化を実行してください。", vbExclamation
        Exit Sub
    End If

    With ws
        ' スタイル名をデフォルト値に戻す
        .Cells(ROW_PATTERN_LEVEL1, COL_STYLE_NAME).Value = "表題1"
        .Cells(ROW_PATTERN_LEVEL2, COL_STYLE_NAME).Value = "表題2"
        .Cells(ROW_PATTERN_LEVEL3, COL_STYLE_NAME).Value = "表題3"
        .Cells(ROW_PATTERN_LEVEL4, COL_STYLE_NAME).Value = "表題4"
        .Cells(ROW_PATTERN_LEVEL5, COL_STYLE_NAME).Value = "表題5"
        .Cells(ROW_PATTERN_EXCEPTION1, COL_STYLE_NAME).Value = "本文"
        .Cells(ROW_PATTERN_EXCEPTION2, COL_STYLE_NAME).Value = "本文"

        ' オプションをデフォルト値に戻す
        .Cells(ROW_OPTION_PDF_OUTPUT, COL_OPTION_VALUE).Value = "はい"
    End With

    MsgBox "設定をリセットしました。", vbInformation
End Sub
