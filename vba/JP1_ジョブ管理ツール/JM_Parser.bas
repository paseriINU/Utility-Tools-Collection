Attribute VB_Name = "JM_Parser"
Option Explicit

'==============================================================================
' JP1 ジョブ管理ツール - パーサーモジュール
' ajsprint出力の解析、データ抽出機能を提供
'==============================================================================

' ============================================================================
' ジョブ一覧結果のパース
' ============================================================================
Public Function ParseJobListResult(result As String, rootPath As String) As Boolean
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
    lastRow = ws.Cells(ws.Rows.Count, COL_JOBNET_PATH).End(xlUp).row
    If lastRow >= ROW_JOBLIST_DATA_START Then
        ws.Range(ws.Cells(ROW_JOBLIST_DATA_START, COL_SELECT), ws.Cells(lastRow, COL_LAST_MESSAGE)).Clear
    End If

    ' 結果をパース
    Dim lines() As String
    lines = Split(result, vbCrLf)

    Dim row As Long
    row = ROW_JOBLIST_DATA_START

    Dim i As Long

    ' ルートパスの末尾のスラッシュを正規化
    ' ただし、"/" のみの場合は空文字列にならないよう除外
    Dim basePath As String
    basePath = rootPath
    If Len(basePath) > 1 And Right(basePath, 1) = "/" Then
        basePath = Left(basePath, Len(basePath) - 1)
    End If

    ' ネスト対応のためスタック構造を使用
    ' 配列でスタックをシミュレート（最大ネスト深度10）
    Const MAX_DEPTH As Long = 10
    Dim unitStack(1 To MAX_DEPTH) As String   ' unit=...ヘッダーのスタック
    Dim blockStack(1 To MAX_DEPTH) As String  ' ブロック内容のスタック
    Dim pathStack(1 To MAX_DEPTH) As String   ' フルパスのスタック
    Dim rowStack(1 To MAX_DEPTH) As Long      ' 書き込み行番号のスタック（親を先に確保）
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
                            ' ただし、親パスが"/"の場合は"/"を追加しない（"//"を防ぐ）
                            If pathStack(stackDepth - 1) = "/" Then
                                pathStack(stackDepth) = "/" & unitName
                            Else
                                pathStack(stackDepth) = pathStack(stackDepth - 1) & "/" & unitName
                            End If
                        End If

                        ' 行番号を確保（親が先に行番号を取得するため、親が上に表示される）
                        rowStack(stackDepth) = row
                        row = row + 1
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
                Dim currentRow As Long
                currentHeader = unitStack(stackDepth)
                currentBlock = blockStack(stackDepth)
                currentFullPath = pathStack(stackDepth)
                currentRow = rowStack(stackDepth)  ' 事前に確保した行番号を使用

                ' ユニットタイプを抽出（ty=xxx; から xxx を取得）
                Dim unitType As String
                Dim unitTypeDisplay As String
                unitType = ExtractUnitType(currentBlock)
                unitTypeDisplay = GetUnitTypeDisplayName(unitType)

                ' ty=が存在し、グループ以外の場合に一覧に追加
                ' グループ(g, mg)は実行できないため除外
                If unitType <> "" And currentFullPath <> "" And unitType <> "g" And unitType <> "mg" Then
                    ws.Cells(currentRow, COL_SELECT).Value = ChrW(&H2610)  ' ☐（空のチェックボックス）
                    ws.Cells(currentRow, COL_ORDER).Value = ""
                    ' 種別を設定
                    ws.Cells(currentRow, COL_UNIT_TYPE).Value = unitTypeDisplay
                    ws.Cells(currentRow, COL_UNIT_TYPE).HorizontalAlignment = xlCenter
                    ' フルパス（ルートからのパス）を設定
                    ws.Cells(currentRow, COL_JOBNET_PATH).Value = currentFullPath
                    ' ユニット名を設定（unit=の最初のフィールド）
                    ws.Cells(currentRow, COL_JOBNET_NAME).Value = ExtractUnitName(currentHeader)
                    ws.Cells(currentRow, COL_COMMENT).Value = ExtractCommentFromBlock(currentBlock)
                    ' スクリプトファイル名 (sc=)
                    ws.Cells(currentRow, COL_SCRIPT).Value = ExtractAttributeFromBlock(currentBlock, "sc")
                    ' パラメーター (prm=)
                    ws.Cells(currentRow, COL_PARAMETER).Value = ExtractAttributeFromBlock(currentBlock, "prm")
                    ' ワークパス (wkp=)
                    ws.Cells(currentRow, COL_WORK_PATH).Value = ExtractAttributeFromBlock(currentBlock, "wkp")

                    ' 保留状態を解析
                    Dim isHold As Boolean
                    isHold = (InStr(currentBlock, "hd=h") > 0) Or (InStr(currentBlock, "hd=H") > 0)

                    If isHold Then
                        ws.Cells(currentRow, COL_HOLD).Value = "保留中"
                        ws.Cells(currentRow, COL_HOLD).HorizontalAlignment = xlCenter
                        ws.Cells(currentRow, COL_HOLD).Interior.Color = RGB(255, 235, 156)  ' 保留列のみ黄色
                        ws.Cells(currentRow, COL_HOLD).Font.Bold = True
                        ws.Cells(currentRow, COL_HOLD).Font.Color = RGB(156, 87, 0)
                    Else
                        ws.Cells(currentRow, COL_HOLD).Value = ""
                    End If

                    ' 選択列・順序列の書式
                    With ws.Cells(currentRow, COL_SELECT)
                        .HorizontalAlignment = xlCenter
                    End With
                    With ws.Cells(currentRow, COL_ORDER)
                        .HorizontalAlignment = xlCenter
                    End With

                    ' 罫線
                    ws.Range(ws.Cells(currentRow, COL_SELECT), ws.Cells(currentRow, COL_LAST_MESSAGE)).Borders.LineStyle = xlContinuous
                End If

                ' スタックをクリアしてポップ
                unitStack(stackDepth) = ""
                blockStack(stackDepth) = ""
                pathStack(stackDepth) = ""
                rowStack(stackDepth) = 0
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

    ' グループ除外により空になった行を削除（下から上に削除）
    Dim deleteRow As Long
    For deleteRow = row - 1 To ROW_JOBLIST_DATA_START Step -1
        If ws.Cells(deleteRow, COL_JOBNET_PATH).Value = "" Then
            ws.Rows(deleteRow).Delete
        End If
    Next deleteRow

    ' データがない場合（空行削除後に再チェック）
    Dim actualLastRow As Long
    actualLastRow = ws.Cells(ws.Rows.Count, COL_JOBNET_PATH).End(xlUp).row
    If actualLastRow < ROW_JOBLIST_DATA_START Then
        MsgBox "実行可能なユニットが見つかりませんでした。" & vbCrLf & _
               "（グループは除外されます）", vbExclamation
        Exit Function
    End If

    ' 成功
    ParseJobListResult = True
End Function

' ============================================================================
' グループ一覧結果のパース
' ============================================================================
Public Function ParseGroupListResult(result As String) As String
    ' グループ名を抽出してドロップダウン用のリストを作成（ネスト対応）
    ' 戻り値: カンマ区切りのパスリスト（例: /,/グループA,/グループA/サブグループ）
    ' ※アスタリスクなしで保存（ジョブ一覧取得時に-Rオプションで再帰取得）

    ' エラーチェック
    If InStr(result, "ERROR:") > 0 Then
        ParseGroupListResult = ""
        Exit Function
    End If

    Dim lines() As String
    lines = Split(result, vbCrLf)

    Dim groupPaths As String
    groupPaths = "/"  ' デフォルトで全件取得オプションを追加（ルート）

    ' ネスト対応のためスタック構造を使用
    Const MAX_DEPTH As Long = 20
    Dim pathStack(1 To MAX_DEPTH) As String   ' パスのスタック
    Dim typeStack(1 To MAX_DEPTH) As String   ' ユニットタイプのスタック
    Dim stackDepth As Long
    stackDepth = 0

    Dim i As Long
    Dim pendingUnitName As String
    pendingUnitName = ""

    For i = LBound(lines) To UBound(lines)
        Dim lineStr As String
        lineStr = Trim(Replace(lines(i), vbTab, ""))

        ' 空行はスキップ
        If lineStr = "" Then GoTo NextGroupLine

        ' unit=行からユニット名を取得
        If InStr(lineStr, "unit=") > 0 Then
            ' unit=名前,,...; から名前を抽出
            Dim parts() As String
            Dim unitPart As String
            unitPart = Mid(lineStr, InStr(lineStr, "unit=") + 5)
            parts = Split(unitPart, ",")
            If UBound(parts) >= 0 Then
                pendingUnitName = Trim(parts(0))
                ' 末尾のセミコロンを除去
                If Right(pendingUnitName, 1) = ";" Then
                    pendingUnitName = Left(pendingUnitName, Len(pendingUnitName) - 1)
                End If
            End If
            GoTo NextGroupLine
        End If

        ' { でネストレベルを上げる
        If Left(lineStr, 1) = "{" Then
            If pendingUnitName <> "" Then
                stackDepth = stackDepth + 1
                If stackDepth <= MAX_DEPTH Then
                    If stackDepth = 1 Then
                        pathStack(stackDepth) = "/" & pendingUnitName
                    Else
                        pathStack(stackDepth) = pathStack(stackDepth - 1) & "/" & pendingUnitName
                    End If
                    typeStack(stackDepth) = ""  ' まだタイプ未確定
                End If
                pendingUnitName = ""
            End If
            GoTo NextGroupLine
        End If

        ' ty=行でユニットタイプを確認
        If InStr(lineStr, "ty=") > 0 Then
            If stackDepth > 0 And stackDepth <= MAX_DEPTH Then
                ' ty=g;（グループ）の場合、パスリストに追加
                If InStr(lineStr, "ty=g;") > 0 Then
                    typeStack(stackDepth) = "g"
                    groupPaths = groupPaths & "," & pathStack(stackDepth)
                End If
            End If
            GoTo NextGroupLine
        End If

        ' } でネストレベルを下げる
        If Left(lineStr, 1) = "}" Then
            If stackDepth > 0 Then
                stackDepth = stackDepth - 1
            End If
            GoTo NextGroupLine
        End If

NextGroupLine:
    Next i

    ParseGroupListResult = groupPaths
End Function

' ============================================================================
' 抽出ユーティリティ関数
' ============================================================================
Public Function ExtractUnitName(line As String) As String
    ' unit=ユニット名,,admin,group; からユニット名を抽出
    ' 最初のフィールド（カンマまで）を返す
    ExtractUnitName = ExtractUnitPath(line)
End Function

Public Function ExtractUnitPath(line As String) As String
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

Public Function ExtractUnitType(blockContent As String) As String
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

Public Function ExtractCommentFromBlock(blockContent As String) As String
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

Public Function ExtractAttributeFromBlock(blockContent As String, attrName As String) As String
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

Public Function ExtractJobNameFromHeader(header As String) As String
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

Public Function ExtractJobName(line As String) As String
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

Public Function ExtractComment(line As String) As String
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

Public Function ExtractHoldStatus(line As String) As Boolean
    ' hd=y（保留）を検出
    ' JP1のajsprint出力で hd=y はホールド(保留)を示す
    ExtractHoldStatus = (InStr(line, ",hd=y") > 0 Or InStr(line, " hd=y") > 0)
End Function

Public Function GetLastPathComponent(path As String) As String
    Dim parts() As String
    parts = Split(path, "/")
    If UBound(parts) >= 0 Then
        GetLastPathComponent = parts(UBound(parts))
    End If
End Function

' ============================================================================
' ユニットタイプの日本語表示名変換
' ============================================================================
Public Function GetUnitTypeDisplayName(unitType As String) As String
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

