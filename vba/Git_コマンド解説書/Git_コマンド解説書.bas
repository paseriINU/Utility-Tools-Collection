'==============================================================================
' Git コマンド解説書 生成ツール
' モジュール名: Git_コマンド解説書
'==============================================================================
' 概要:
'   Git初心者向けのコマンド解説書をExcelで生成するツールです。
'   基本的なGitコマンドから実践的な使い方まで、わかりやすく解説します。
'
' 使い方:
'   1. このモジュールをExcelのVBAエディタにインポート
'   2. CreateGitCommandGuide マクロを実行
'   3. 解説書シートが自動生成されます
'
' 作成日: 2025-12-13
'==============================================================================

Option Explicit

'==============================================================================
' メイン: Git解説書を生成
'==============================================================================
Public Sub CreateGitCommandGuide()
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False

    ' シートを作成
    CreateSheet "Git基礎知識"
    CreateSheet "基本コマンド"
    CreateSheet "ブランチ操作"
    CreateSheet "リモート操作"
    CreateSheet "履歴・差分確認"
    CreateSheet "取り消し・修正"
    CreateSheet "実践シナリオ"
    CreateSheet "トラブル対処"

    ' 各シートの内容を作成
    FormatBasicsSheet
    FormatBasicCommandsSheet
    FormatBranchSheet
    FormatRemoteSheet
    FormatHistorySheet
    FormatUndoSheet
    FormatScenarioSheet
    FormatTroubleshootSheet

    ' 最初のシートをアクティブに
    Sheets("Git基礎知識").Activate

    Application.ScreenUpdating = True

    MsgBox "Git コマンド解説書を作成しました。" & vbCrLf & vbCrLf & _
           "【シート構成】" & vbCrLf & _
           "1. Git基礎知識 - Gitの概念を理解" & vbCrLf & _
           "2. 基本コマンド - よく使うコマンド" & vbCrLf & _
           "3. ブランチ操作 - ブランチの使い方" & vbCrLf & _
           "4. リモート操作 - GitHubとの連携" & vbCrLf & _
           "5. 履歴・差分確認 - 変更の確認方法" & vbCrLf & _
           "6. 取り消し・修正 - 間違いの修正方法" & vbCrLf & _
           "7. 実践シナリオ - 実際の作業フロー" & vbCrLf & _
           "8. トラブル対処 - よくある問題と解決法", _
           vbInformation, "作成完了"

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub

'==============================================================================
' シート作成ヘルパー
'==============================================================================
Private Sub CreateSheet(ByVal sheetName As String)
    Dim ws As Worksheet
    Dim exists As Boolean

    exists = False
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    If Not ws Is Nothing Then exists = True
    Err.Clear
    On Error GoTo 0

    If exists Then
        ws.Cells.Clear
        ws.Cells.Interior.ColorIndex = xlNone
    Else
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = sheetName
    End If
End Sub

'==============================================================================
' Git基礎知識シート
'==============================================================================
Private Sub FormatBasicsSheet()
    Dim ws As Worksheet
    Set ws = Sheets("Git基礎知識")

    With ws
        ' タイトル
        .Range("A1").Value = "Git 基礎知識 - まずはここから!"
        .Range("A1").Font.Size = 20
        .Range("A1").Font.Bold = True
        .Range("A1:H1").Merge
        .Range("A1").Interior.Color = RGB(64, 64, 64)
        .Range("A1").Font.Color = RGB(255, 255, 255)

        ' Gitとは
        .Range("A3").Value = "Gitとは?"
        .Range("A3").Font.Size = 14
        .Range("A3").Font.Bold = True
        .Range("A3").Interior.Color = RGB(68, 114, 196)
        .Range("A3").Font.Color = RGB(255, 255, 255)
        .Range("A3:H3").Merge

        .Range("A4").Value = "Git（ギット）は、ファイルの変更履歴を記録・管理するための「バージョン管理システム」です。"
        .Range("A4:H4").Merge
        .Range("A5").Value = "プログラムのソースコードだけでなく、ドキュメントや設定ファイルなど、あらゆるテキストファイルの管理に使えます。"
        .Range("A5:H5").Merge

        ' なぜGitを使うのか
        .Range("A7").Value = "なぜGitを使うのか?"
        .Range("A7").Font.Size = 14
        .Range("A7").Font.Bold = True
        .Range("A7").Interior.Color = RGB(68, 114, 196)
        .Range("A7").Font.Color = RGB(255, 255, 255)
        .Range("A7:H7").Merge

        Dim reasons As Variant
        reasons = Array( _
            "1. 変更履歴を残せる - いつ、誰が、何を変更したか記録される", _
            "2. 過去に戻れる - 間違えても以前の状態に戻せる", _
            "3. 複数人で作業できる - チームで同時に開発できる", _
            "4. 実験しやすい - ブランチで安全に新機能を試せる", _
            "5. バックアップになる - リモートにコピーを保存できる" _
        )

        Dim i As Long
        For i = 0 To UBound(reasons)
            .Range("A" & (8 + i)).Value = reasons(i)
            .Range("A" & (8 + i) & ":H" & (8 + i)).Merge
        Next i

        ' 重要な用語
        .Range("A14").Value = "覚えておきたい用語"
        .Range("A14").Font.Size = 14
        .Range("A14").Font.Bold = True
        .Range("A14").Interior.Color = RGB(68, 114, 196)
        .Range("A14").Font.Color = RGB(255, 255, 255)
        .Range("A14:H14").Merge

        .Range("A15").Value = "用語"
        .Range("B15").Value = "読み方"
        .Range("C15").Value = "説明"
        .Range("A15:C15").Font.Bold = True
        .Range("A15:C15").Interior.Color = RGB(180, 198, 231)

        Dim terms As Variant
        terms = Array( _
            Array("Repository", "リポジトリ", "Gitで管理されているフォルダ。プロジェクトの全履歴が保存される場所。"), _
            Array("Commit", "コミット", "変更を記録すること。セーブポイントを作るイメージ。"), _
            Array("Branch", "ブランチ", "開発の分岐。メインとは別の作業場所を作れる。"), _
            Array("Merge", "マージ", "ブランチを統合すること。別々の変更を1つにまとめる。"), _
            Array("Clone", "クローン", "リモートのリポジトリをローカルにコピーすること。"), _
            Array("Push", "プッシュ", "ローカルの変更をリモートに送ること。"), _
            Array("Pull", "プル", "リモートの変更をローカルに取り込むこと。"), _
            Array("Staging", "ステージング", "コミットする変更を選ぶ準備段階。"), _
            Array("HEAD", "ヘッド", "現在作業中のコミットを指すポインタ。"), _
            Array("Origin", "オリジン", "リモートリポジトリのデフォルト名。通常はGitHub等を指す。") _
        )

        For i = 0 To UBound(terms)
            .Range("A" & (16 + i)).Value = terms(i)(0)
            .Range("B" & (16 + i)).Value = terms(i)(1)
            .Range("C" & (16 + i)).Value = terms(i)(2)
            .Range("C" & (16 + i) & ":H" & (16 + i)).Merge
        Next i

        ' 3つの領域
        .Range("A27").Value = "Gitの3つの領域（これ重要!）"
        .Range("A27").Font.Size = 14
        .Range("A27").Font.Bold = True
        .Range("A27").Interior.Color = RGB(192, 80, 77)
        .Range("A27").Font.Color = RGB(255, 255, 255)
        .Range("A27:H27").Merge

        .Range("A28").Value = "領域"
        .Range("B28").Value = "説明"
        .Range("C28").Value = "状態"
        .Range("A28:C28").Font.Bold = True
        .Range("A28:C28").Interior.Color = RGB(230, 184, 183)

        Dim areas As Variant
        areas = Array( _
            Array("作業ディレクトリ", "実際にファイルを編集する場所", "ファイルを変更している状態"), _
            Array("ステージングエリア", "コミットする変更を選ぶ場所", "git add した状態"), _
            Array("リポジトリ", "変更履歴が保存される場所", "git commit した状態") _
        )

        For i = 0 To UBound(areas)
            .Range("A" & (29 + i)).Value = areas(i)(0)
            .Range("B" & (29 + i)).Value = areas(i)(1)
            .Range("C" & (29 + i)).Value = areas(i)(2)
            .Range("B" & (29 + i) & ":D" & (29 + i)).Merge
            .Range("C" & (29 + i) & ":H" & (29 + i)).Merge
        Next i

        ' 図解
        .Range("A33").Value = "【変更の流れ】"
        .Range("A33").Font.Bold = True
        .Range("A34").Value = "ファイル編集 → git add → git commit → git push"
        .Range("A34").Font.Name = "Consolas"
        .Range("A35").Value = "(作業)         (ステージ)  (ローカル保存)  (リモート送信)"
        .Range("A35").Font.Color = RGB(128, 128, 128)

        ' 列幅
        .Columns("A").ColumnWidth = 20
        .Columns("B").ColumnWidth = 15
        .Columns("C").ColumnWidth = 60
    End With
End Sub

'==============================================================================
' 基本コマンドシート
'==============================================================================
Private Sub FormatBasicCommandsSheet()
    Dim ws As Worksheet
    Set ws = Sheets("基本コマンド")

    With ws
        ' タイトル
        .Range("A1").Value = "基本コマンド - 毎日使うコマンド"
        .Range("A1").Font.Size = 20
        .Range("A1").Font.Bold = True
        .Range("A1:H1").Merge
        .Range("A1").Interior.Color = RGB(64, 64, 64)
        .Range("A1").Font.Color = RGB(255, 255, 255)

        ' ヘッダー
        .Range("A3").Value = "コマンド"
        .Range("B3").Value = "説明"
        .Range("C3").Value = "使用例"
        .Range("D3").Value = "ポイント"
        .Range("A3:D3").Font.Bold = True
        .Range("A3:D3").Interior.Color = RGB(68, 114, 196)
        .Range("A3:D3").Font.Color = RGB(255, 255, 255)

        Dim commands As Variant
        commands = Array( _
            Array("git init", "新しいリポジトリを作成", "git init", "空のフォルダで最初に1回だけ実行"), _
            Array("git clone", "リモートリポジトリをコピー", "git clone https://github.com/user/repo.git", "既存プロジェクトを取得する時に使う"), _
            Array("git status", "現在の状態を確認", "git status", "迷ったらまずこれ! 一番よく使う"), _
            Array("git add", "変更をステージング", "git add ファイル名" & vbLf & "git add .", "「.」で全ファイルを追加"), _
            Array("git commit", "変更を記録", "git commit -m ""メッセージ""", "メッセージは何を変更したか書く"), _
            Array("git push", "リモートに送信", "git push origin main", "ローカルの変更をGitHubへ"), _
            Array("git pull", "リモートから取得", "git pull origin main", "最新の変更を取り込む"), _
            Array("git log", "履歴を表示", "git log --oneline", "--onelineで1行表示"), _
            Array("git diff", "差分を表示", "git diff", "何が変わったか確認できる") _
        )

        Dim i As Long
        Dim row As Long
        row = 4
        For i = 0 To UBound(commands)
            .Range("A" & row).Value = commands(i)(0)
            .Range("A" & row).Font.Name = "Consolas"
            .Range("A" & row).Font.Bold = True
            .Range("A" & row).Interior.Color = RGB(242, 242, 242)
            .Range("B" & row).Value = commands(i)(1)
            .Range("C" & row).Value = commands(i)(2)
            .Range("C" & row).Font.Name = "Consolas"
            .Range("D" & row).Value = commands(i)(3)
            .Range("D" & row).Font.Color = RGB(0, 112, 192)
            row = row + 1
        Next i

        ' 基本の流れ
        .Range("A14").Value = "基本の作業フロー"
        .Range("A14").Font.Size = 14
        .Range("A14").Font.Bold = True
        .Range("A14").Interior.Color = RGB(0, 176, 80)
        .Range("A14").Font.Color = RGB(255, 255, 255)
        .Range("A14:H14").Merge

        Dim workflow As Variant
        workflow = Array( _
            "Step 1: git status          # 現在の状態を確認", _
            "Step 2: (ファイルを編集)", _
            "Step 3: git status          # 変更されたファイルを確認", _
            "Step 4: git add .           # 全ての変更をステージング", _
            "Step 5: git status          # ステージングされたか確認", _
            "Step 6: git commit -m ""変更内容""  # コミット", _
            "Step 7: git push origin main     # リモートに送信" _
        )

        For i = 0 To UBound(workflow)
            .Range("A" & (15 + i)).Value = workflow(i)
            .Range("A" & (15 + i)).Font.Name = "Consolas"
            .Range("A" & (15 + i) & ":H" & (15 + i)).Merge
        Next i

        ' Tips
        .Range("A23").Value = "初心者Tips"
        .Range("A23").Font.Size = 14
        .Range("A23").Font.Bold = True
        .Range("A23").Interior.Color = RGB(255, 192, 0)
        .Range("A23:H23").Merge

        Dim tips As Variant
        tips = Array( _
            "・迷ったら git status を実行! 今何が起きているかわかります", _
            "・コミットメッセージは「何をしたか」を日本語で書いてOK", _
            "・小さな単位でこまめにコミットするのがコツ", _
            "・pushする前にpullして最新を取り込む習慣をつけよう" _
        )

        For i = 0 To UBound(tips)
            .Range("A" & (24 + i)).Value = tips(i)
            .Range("A" & (24 + i) & ":H" & (24 + i)).Merge
        Next i

        ' 列幅
        .Columns("A").ColumnWidth = 18
        .Columns("B").ColumnWidth = 25
        .Columns("C").ColumnWidth = 45
        .Columns("D").ColumnWidth = 35
    End With
End Sub

'==============================================================================
' ブランチ操作シート
'==============================================================================
Private Sub FormatBranchSheet()
    Dim ws As Worksheet
    Set ws = Sheets("ブランチ操作")

    With ws
        ' タイトル
        .Range("A1").Value = "ブランチ操作 - 並行開発の要"
        .Range("A1").Font.Size = 20
        .Range("A1").Font.Bold = True
        .Range("A1:H1").Merge
        .Range("A1").Interior.Color = RGB(64, 64, 64)
        .Range("A1").Font.Color = RGB(255, 255, 255)

        ' ブランチとは
        .Range("A3").Value = "ブランチとは?"
        .Range("A3").Font.Size = 14
        .Range("A3").Font.Bold = True
        .Range("A3").Interior.Color = RGB(112, 48, 160)
        .Range("A3").Font.Color = RGB(255, 255, 255)
        .Range("A3:H3").Merge

        .Range("A4").Value = "ブランチは「枝分かれ」という意味。メインの開発ラインとは別に、独立した作業場所を作れます。"
        .Range("A4:H4").Merge
        .Range("A5").Value = "新機能の開発やバグ修正を、メインに影響を与えずに安全に行えます。"
        .Range("A5:H5").Merge

        ' コマンド一覧
        .Range("A7").Value = "ブランチ操作コマンド"
        .Range("A7").Font.Size = 14
        .Range("A7").Font.Bold = True
        .Range("A7").Interior.Color = RGB(112, 48, 160)
        .Range("A7").Font.Color = RGB(255, 255, 255)
        .Range("A7:H7").Merge

        .Range("A8").Value = "コマンド"
        .Range("B8").Value = "説明"
        .Range("C8").Value = "使用例"
        .Range("A8:C8").Font.Bold = True
        .Range("A8:C8").Interior.Color = RGB(204, 192, 218)

        Dim commands As Variant
        commands = Array( _
            Array("git branch", "ブランチ一覧を表示", "git branch"), _
            Array("git branch [名前]", "新しいブランチを作成", "git branch feature-login"), _
            Array("git checkout [名前]", "ブランチを切り替え", "git checkout feature-login"), _
            Array("git checkout -b [名前]", "作成と切り替えを同時に", "git checkout -b feature-login"), _
            Array("git switch [名前]", "ブランチを切り替え(新)", "git switch feature-login"), _
            Array("git switch -c [名前]", "作成と切り替え(新)", "git switch -c feature-login"), _
            Array("git merge [名前]", "ブランチを統合", "git merge feature-login"), _
            Array("git branch -d [名前]", "ブランチを削除", "git branch -d feature-login"), _
            Array("git branch -D [名前]", "強制削除", "git branch -D feature-login") _
        )

        Dim i As Long
        For i = 0 To UBound(commands)
            .Range("A" & (9 + i)).Value = commands(i)(0)
            .Range("A" & (9 + i)).Font.Name = "Consolas"
            .Range("B" & (9 + i)).Value = commands(i)(1)
            .Range("C" & (9 + i)).Value = commands(i)(2)
            .Range("C" & (9 + i)).Font.Name = "Consolas"
        Next i

        ' ブランチの使い方
        .Range("A19").Value = "ブランチを使った開発フロー"
        .Range("A19").Font.Size = 14
        .Range("A19").Font.Bold = True
        .Range("A19").Interior.Color = RGB(0, 176, 80)
        .Range("A19").Font.Color = RGB(255, 255, 255)
        .Range("A19:H19").Merge

        Dim workflow As Variant
        workflow = Array( _
            "1. git checkout -b feature-xxx    # 新機能用ブランチを作成", _
            "2. (ファイルを編集・コミットを繰り返す)", _
            "3. git checkout main              # mainブランチに戻る", _
            "4. git pull origin main           # 最新を取得", _
            "5. git merge feature-xxx          # 新機能をmainに統合", _
            "6. git push origin main           # リモートに送信", _
            "7. git branch -d feature-xxx      # 不要になったブランチを削除" _
        )

        For i = 0 To UBound(workflow)
            .Range("A" & (20 + i)).Value = workflow(i)
            .Range("A" & (20 + i)).Font.Name = "Consolas"
            .Range("A" & (20 + i) & ":H" & (20 + i)).Merge
        Next i

        ' ブランチ命名規則
        .Range("A28").Value = "ブランチ名の付け方（例）"
        .Range("A28").Font.Size = 14
        .Range("A28").Font.Bold = True
        .Range("A28").Interior.Color = RGB(255, 192, 0)
        .Range("A28:H28").Merge

        Dim naming As Variant
        naming = Array( _
            Array("feature/xxx", "新機能開発", "feature/user-login, feature/add-search"), _
            Array("bugfix/xxx", "バグ修正", "bugfix/login-error, bugfix/null-check"), _
            Array("hotfix/xxx", "緊急修正", "hotfix/security-patch"), _
            Array("release/xxx", "リリース準備", "release/v1.0.0") _
        )

        .Range("A29").Value = "プレフィックス"
        .Range("B29").Value = "用途"
        .Range("C29").Value = "例"
        .Range("A29:C29").Font.Bold = True
        .Range("A29:C29").Interior.Color = RGB(255, 230, 153)

        For i = 0 To UBound(naming)
            .Range("A" & (30 + i)).Value = naming(i)(0)
            .Range("A" & (30 + i)).Font.Name = "Consolas"
            .Range("B" & (30 + i)).Value = naming(i)(1)
            .Range("C" & (30 + i)).Value = naming(i)(2)
            .Range("C" & (30 + i)).Font.Name = "Consolas"
        Next i

        ' 列幅
        .Columns("A").ColumnWidth = 25
        .Columns("B").ColumnWidth = 25
        .Columns("C").ColumnWidth = 45
    End With
End Sub

'==============================================================================
' リモート操作シート
'==============================================================================
Private Sub FormatRemoteSheet()
    Dim ws As Worksheet
    Set ws = Sheets("リモート操作")

    With ws
        ' タイトル
        .Range("A1").Value = "リモート操作 - GitHubとの連携"
        .Range("A1").Font.Size = 20
        .Range("A1").Font.Bold = True
        .Range("A1:H1").Merge
        .Range("A1").Interior.Color = RGB(64, 64, 64)
        .Range("A1").Font.Color = RGB(255, 255, 255)

        ' リモートとは
        .Range("A3").Value = "リモートリポジトリとは?"
        .Range("A3").Font.Size = 14
        .Range("A3").Font.Bold = True
        .Range("A3").Interior.Color = RGB(0, 112, 192)
        .Range("A3").Font.Color = RGB(255, 255, 255)
        .Range("A3:H3").Merge

        .Range("A4").Value = "GitHub、GitLab、Bitbucketなどのサーバー上にあるリポジトリのこと。"
        .Range("A4:H4").Merge
        .Range("A5").Value = "チームでコードを共有したり、バックアップとして使ったりします。"
        .Range("A5:H5").Merge

        ' コマンド一覧
        .Range("A7").Value = "リモート操作コマンド"
        .Range("A7").Font.Size = 14
        .Range("A7").Font.Bold = True
        .Range("A7").Interior.Color = RGB(0, 112, 192)
        .Range("A7").Font.Color = RGB(255, 255, 255)
        .Range("A7:H7").Merge

        .Range("A8").Value = "コマンド"
        .Range("B8").Value = "説明"
        .Range("C8").Value = "使用例"
        .Range("A8:C8").Font.Bold = True
        .Range("A8:C8").Interior.Color = RGB(180, 198, 231)

        Dim commands As Variant
        commands = Array( _
            Array("git remote -v", "登録されているリモートを確認", "git remote -v"), _
            Array("git remote add", "リモートを登録", "git remote add origin https://..."), _
            Array("git push", "ローカルの変更を送信", "git push origin main"), _
            Array("git push -u", "上流ブランチを設定して送信", "git push -u origin main"), _
            Array("git pull", "リモートの変更を取得・統合", "git pull origin main"), _
            Array("git fetch", "リモートの情報だけ取得", "git fetch origin"), _
            Array("git clone", "リモートをコピー", "git clone https://github.com/...") _
        )

        Dim i As Long
        For i = 0 To UBound(commands)
            .Range("A" & (9 + i)).Value = commands(i)(0)
            .Range("A" & (9 + i)).Font.Name = "Consolas"
            .Range("B" & (9 + i)).Value = commands(i)(1)
            .Range("C" & (9 + i)).Value = commands(i)(2)
            .Range("C" & (9 + i)).Font.Name = "Consolas"
        Next i

        ' push vs pull
        .Range("A17").Value = "push と pull の違い"
        .Range("A17").Font.Size = 14
        .Range("A17").Font.Bold = True
        .Range("A17").Interior.Color = RGB(255, 192, 0)
        .Range("A17:H17").Merge

        .Range("A18").Value = "コマンド"
        .Range("B18").Value = "方向"
        .Range("C18").Value = "説明"
        .Range("A18:C18").Font.Bold = True
        .Range("A18:C18").Interior.Color = RGB(255, 230, 153)

        .Range("A19").Value = "git push"
        .Range("A19").Font.Name = "Consolas"
        .Range("B19").Value = "ローカル → リモート"
        .Range("C19").Value = "自分の変更をサーバーに送る"

        .Range("A20").Value = "git pull"
        .Range("A20").Font.Name = "Consolas"
        .Range("B20").Value = "リモート → ローカル"
        .Range("C20").Value = "サーバーの変更を自分に取り込む"

        .Range("A21").Value = "git fetch"
        .Range("A21").Font.Name = "Consolas"
        .Range("B21").Value = "リモート → ローカル(情報のみ)"
        .Range("C21").Value = "変更情報だけ取得(統合はしない)"

        ' fetch vs pull
        .Range("A23").Value = "fetch と pull の違い"
        .Range("A23").Font.Size = 14
        .Range("A23").Font.Bold = True
        .Range("A23").Interior.Color = RGB(0, 176, 80)
        .Range("A23").Font.Color = RGB(255, 255, 255)
        .Range("A23:H23").Merge

        .Range("A24").Value = "git fetch: リモートの情報を取得するだけ。作業中のファイルは変わらない。"
        .Range("A24:H24").Merge
        .Range("A25").Value = "git pull:  git fetch + git merge。取得と統合を一度に行う。"
        .Range("A25:H25").Merge
        .Range("A26").Value = ""
        .Range("A27").Value = "安全に確認したい時は fetch → 内容確認 → merge の流れがおすすめ"
        .Range("A27").Font.Color = RGB(0, 112, 192)
        .Range("A27:H27").Merge

        ' 列幅
        .Columns("A").ColumnWidth = 22
        .Columns("B").ColumnWidth = 30
        .Columns("C").ColumnWidth = 45
    End With
End Sub

'==============================================================================
' 履歴・差分確認シート
'==============================================================================
Private Sub FormatHistorySheet()
    Dim ws As Worksheet
    Set ws = Sheets("履歴・差分確認")

    With ws
        ' タイトル
        .Range("A1").Value = "履歴・差分確認 - 変更を追跡する"
        .Range("A1").Font.Size = 20
        .Range("A1").Font.Bold = True
        .Range("A1:H1").Merge
        .Range("A1").Interior.Color = RGB(64, 64, 64)
        .Range("A1").Font.Color = RGB(255, 255, 255)

        ' git log
        .Range("A3").Value = "履歴を見る (git log)"
        .Range("A3").Font.Size = 14
        .Range("A3").Font.Bold = True
        .Range("A3").Interior.Color = RGB(68, 114, 196)
        .Range("A3").Font.Color = RGB(255, 255, 255)
        .Range("A3:H3").Merge

        .Range("A4").Value = "コマンド"
        .Range("B4").Value = "説明"
        .Range("A4:B4").Font.Bold = True
        .Range("A4:B4").Interior.Color = RGB(180, 198, 231)

        Dim logCommands As Variant
        logCommands = Array( _
            Array("git log", "コミット履歴を表示"), _
            Array("git log --oneline", "1行で簡潔に表示（おすすめ）"), _
            Array("git log -n 5", "最新5件だけ表示"), _
            Array("git log --graph", "ブランチをグラフ表示"), _
            Array("git log --oneline --graph", "グラフを1行表示"), _
            Array("git log -p", "変更内容も表示"), _
            Array("git log --author=""名前""", "特定の人のコミットだけ表示"), _
            Array("git log --since=""2024-01-01""", "指定日以降のコミット"), _
            Array("git log ファイル名", "特定ファイルの履歴") _
        )

        Dim i As Long
        For i = 0 To UBound(logCommands)
            .Range("A" & (5 + i)).Value = logCommands(i)(0)
            .Range("A" & (5 + i)).Font.Name = "Consolas"
            .Range("B" & (5 + i)).Value = logCommands(i)(1)
            .Range("B" & (5 + i) & ":H" & (5 + i)).Merge
        Next i

        ' git diff
        .Range("A15").Value = "差分を見る (git diff)"
        .Range("A15").Font.Size = 14
        .Range("A15").Font.Bold = True
        .Range("A15").Interior.Color = RGB(0, 176, 80)
        .Range("A15").Font.Color = RGB(255, 255, 255)
        .Range("A15:H15").Merge

        .Range("A16").Value = "コマンド"
        .Range("B16").Value = "説明"
        .Range("A16:B16").Font.Bold = True
        .Range("A16:B16").Interior.Color = RGB(198, 224, 180)

        Dim diffCommands As Variant
        diffCommands = Array( _
            Array("git diff", "作業ディレクトリの変更を表示"), _
            Array("git diff --staged", "ステージングした変更を表示"), _
            Array("git diff HEAD", "最後のコミットとの差分"), _
            Array("git diff ブランチ1 ブランチ2", "ブランチ間の差分"), _
            Array("git diff コミットID1 コミットID2", "コミット間の差分"), _
            Array("git diff ファイル名", "特定ファイルの差分") _
        )

        For i = 0 To UBound(diffCommands)
            .Range("A" & (17 + i)).Value = diffCommands(i)(0)
            .Range("A" & (17 + i)).Font.Name = "Consolas"
            .Range("B" & (17 + i)).Value = diffCommands(i)(1)
            .Range("B" & (17 + i) & ":H" & (17 + i)).Merge
        Next i

        ' その他
        .Range("A24").Value = "その他の確認コマンド"
        .Range("A24").Font.Size = 14
        .Range("A24").Font.Bold = True
        .Range("A24").Interior.Color = RGB(255, 192, 0)
        .Range("A24:H24").Merge

        Dim otherCommands As Variant
        otherCommands = Array( _
            Array("git show コミットID", "特定コミットの詳細を表示"), _
            Array("git blame ファイル名", "各行を誰がいつ変更したか表示"), _
            Array("git shortlog", "ユーザーごとのコミット数"), _
            Array("git reflog", "HEADの移動履歴（復旧に便利）") _
        )

        .Range("A25").Value = "コマンド"
        .Range("B25").Value = "説明"
        .Range("A25:B25").Font.Bold = True
        .Range("A25:B25").Interior.Color = RGB(255, 230, 153)

        For i = 0 To UBound(otherCommands)
            .Range("A" & (26 + i)).Value = otherCommands(i)(0)
            .Range("A" & (26 + i)).Font.Name = "Consolas"
            .Range("B" & (26 + i)).Value = otherCommands(i)(1)
            .Range("B" & (26 + i) & ":H" & (26 + i)).Merge
        Next i

        ' 列幅
        .Columns("A").ColumnWidth = 35
        .Columns("B").ColumnWidth = 50
    End With
End Sub

'==============================================================================
' 取り消し・修正シート
'==============================================================================
Private Sub FormatUndoSheet()
    Dim ws As Worksheet
    Set ws = Sheets("取り消し・修正")

    With ws
        ' タイトル
        .Range("A1").Value = "取り消し・修正 - 間違いを直す"
        .Range("A1").Font.Size = 20
        .Range("A1").Font.Bold = True
        .Range("A1:H1").Merge
        .Range("A1").Interior.Color = RGB(64, 64, 64)
        .Range("A1").Font.Color = RGB(255, 255, 255)

        ' 警告
        .Range("A3").Value = "注意: これらのコマンドは変更を取り消します。慎重に使ってください!"
        .Range("A3").Font.Bold = True
        .Range("A3").Font.Color = RGB(192, 0, 0)
        .Range("A3:H3").Merge

        ' 作業ディレクトリの変更を取り消す
        .Range("A5").Value = "作業中の変更を取り消す"
        .Range("A5").Font.Size = 14
        .Range("A5").Font.Bold = True
        .Range("A5").Interior.Color = RGB(192, 80, 77)
        .Range("A5").Font.Color = RGB(255, 255, 255)
        .Range("A5:H5").Merge

        .Range("A6").Value = "状況"
        .Range("B6").Value = "コマンド"
        .Range("C6").Value = "説明"
        .Range("A6:C6").Font.Bold = True
        .Range("A6:C6").Interior.Color = RGB(230, 184, 183)

        Dim undoCommands As Variant
        undoCommands = Array( _
            Array("ファイルの変更を取り消したい", "git checkout -- ファイル名", "最後のコミット状態に戻す"), _
            Array("全ファイルの変更を取り消したい", "git checkout -- .", "全ファイルを戻す（危険）"), _
            Array("git addを取り消したい", "git reset HEAD ファイル名", "ステージングを解除"), _
            Array("git addを全部取り消したい", "git reset HEAD", "全てのステージングを解除") _
        )

        Dim i As Long
        For i = 0 To UBound(undoCommands)
            .Range("A" & (7 + i)).Value = undoCommands(i)(0)
            .Range("B" & (7 + i)).Value = undoCommands(i)(1)
            .Range("B" & (7 + i)).Font.Name = "Consolas"
            .Range("C" & (7 + i)).Value = undoCommands(i)(2)
        Next i

        ' コミットを修正
        .Range("A12").Value = "コミットを修正する"
        .Range("A12").Font.Size = 14
        .Range("A12").Font.Bold = True
        .Range("A12").Interior.Color = RGB(255, 192, 0)
        .Range("A12:H12").Merge

        Dim commitCommands As Variant
        commitCommands = Array( _
            Array("直前のコミットメッセージを修正", "git commit --amend", "エディタが開く"), _
            Array("直前のコミットに追加で変更を含める", "git add ファイル名" & vbLf & "git commit --amend --no-edit", "メッセージは変えない"), _
            Array("直前のコミットを取り消す(変更は残す)", "git reset --soft HEAD^", "ステージング状態に戻る"), _
            Array("直前のコミットを取り消す(変更も消す)", "git reset --hard HEAD^", "完全に消える（危険）") _
        )

        .Range("A13").Value = "状況"
        .Range("B13").Value = "コマンド"
        .Range("C13").Value = "説明"
        .Range("A13:C13").Font.Bold = True
        .Range("A13:C13").Interior.Color = RGB(255, 230, 153)

        For i = 0 To UBound(commitCommands)
            .Range("A" & (14 + i)).Value = commitCommands(i)(0)
            .Range("B" & (14 + i)).Value = commitCommands(i)(1)
            .Range("B" & (14 + i)).Font.Name = "Consolas"
            .Range("C" & (14 + i)).Value = commitCommands(i)(2)
        Next i

        ' reset の種類
        .Range("A19").Value = "git reset のオプション"
        .Range("A19").Font.Size = 14
        .Range("A19").Font.Bold = True
        .Range("A19").Interior.Color = RGB(68, 114, 196)
        .Range("A19").Font.Color = RGB(255, 255, 255)
        .Range("A19:H19").Merge

        Dim resetOptions As Variant
        resetOptions = Array( _
            Array("--soft", "コミットだけ取り消す", "変更はステージングに残る"), _
            Array("--mixed(デフォルト)", "コミットとステージングを取り消す", "変更は作業ディレクトリに残る"), _
            Array("--hard", "全て取り消す", "変更も消える（危険!）") _
        )

        .Range("A20").Value = "オプション"
        .Range("B20").Value = "動作"
        .Range("C20").Value = "結果"
        .Range("A20:C20").Font.Bold = True
        .Range("A20:C20").Interior.Color = RGB(180, 198, 231)

        For i = 0 To UBound(resetOptions)
            .Range("A" & (21 + i)).Value = resetOptions(i)(0)
            .Range("A" & (21 + i)).Font.Name = "Consolas"
            .Range("B" & (21 + i)).Value = resetOptions(i)(1)
            .Range("C" & (21 + i)).Value = resetOptions(i)(2)
        Next i

        ' revert
        .Range("A25").Value = "安全に取り消す (git revert)"
        .Range("A25").Font.Size = 14
        .Range("A25").Font.Bold = True
        .Range("A25").Interior.Color = RGB(0, 176, 80)
        .Range("A25").Font.Color = RGB(255, 255, 255)
        .Range("A25:H25").Merge

        .Range("A26").Value = "git revert コミットID"
        .Range("A26").Font.Name = "Consolas"
        .Range("A27").Value = "指定したコミットを打ち消す新しいコミットを作成します。"
        .Range("A27:H27").Merge
        .Range("A28").Value = "履歴は残るので、チーム開発では reset より revert が安全です。"
        .Range("A28").Font.Color = RGB(0, 112, 192)
        .Range("A28:H28").Merge

        ' 列幅
        .Columns("A").ColumnWidth = 35
        .Columns("B").ColumnWidth = 35
        .Columns("C").ColumnWidth = 35
    End With
End Sub

'==============================================================================
' 実践シナリオシート
'==============================================================================
Private Sub FormatScenarioSheet()
    Dim ws As Worksheet
    Set ws = Sheets("実践シナリオ")

    With ws
        ' タイトル
        .Range("A1").Value = "実践シナリオ - よくある作業フロー"
        .Range("A1").Font.Size = 20
        .Range("A1").Font.Bold = True
        .Range("A1:H1").Merge
        .Range("A1").Interior.Color = RGB(64, 64, 64)
        .Range("A1").Font.Color = RGB(255, 255, 255)

        Dim row As Long
        row = 3

        ' シナリオ1
        .Range("A" & row).Value = "シナリオ1: 新規プロジェクトを始める"
        .Range("A" & row).Font.Size = 14
        .Range("A" & row).Font.Bold = True
        .Range("A" & row).Interior.Color = RGB(68, 114, 196)
        .Range("A" & row).Font.Color = RGB(255, 255, 255)
        .Range("A" & row & ":H" & row).Merge
        row = row + 1

        Dim scenario1 As Variant
        scenario1 = Array( _
            "mkdir my-project          # プロジェクトフォルダを作成", _
            "cd my-project             # フォルダに移動", _
            "git init                  # Gitリポジトリを初期化", _
            "(ファイルを作成・編集)", _
            "git add .                 # 全ファイルをステージング", _
            "git commit -m ""Initial commit""  # 最初のコミット" _
        )

        Dim i As Long
        For i = 0 To UBound(scenario1)
            .Range("A" & row).Value = scenario1(i)
            .Range("A" & row).Font.Name = "Consolas"
            .Range("A" & row & ":H" & row).Merge
            row = row + 1
        Next i
        row = row + 1

        ' シナリオ2
        .Range("A" & row).Value = "シナリオ2: GitHubから既存プロジェクトを取得"
        .Range("A" & row).Font.Size = 14
        .Range("A" & row).Font.Bold = True
        .Range("A" & row).Interior.Color = RGB(0, 176, 80)
        .Range("A" & row).Font.Color = RGB(255, 255, 255)
        .Range("A" & row & ":H" & row).Merge
        row = row + 1

        Dim scenario2 As Variant
        scenario2 = Array( _
            "git clone https://github.com/xxx/yyy.git  # リポジトリをコピー", _
            "cd yyy                    # フォルダに移動", _
            "git branch -a             # ブランチ一覧を確認", _
            "(必要に応じてブランチを切り替え)" _
        )

        For i = 0 To UBound(scenario2)
            .Range("A" & row).Value = scenario2(i)
            .Range("A" & row).Font.Name = "Consolas"
            .Range("A" & row & ":H" & row).Merge
            row = row + 1
        Next i
        row = row + 1

        ' シナリオ3
        .Range("A" & row).Value = "シナリオ3: 新機能を開発する"
        .Range("A" & row).Font.Size = 14
        .Range("A" & row).Font.Bold = True
        .Range("A" & row).Interior.Color = RGB(112, 48, 160)
        .Range("A" & row).Font.Color = RGB(255, 255, 255)
        .Range("A" & row & ":H" & row).Merge
        row = row + 1

        Dim scenario3 As Variant
        scenario3 = Array( _
            "git checkout main         # mainブランチに移動", _
            "git pull origin main      # 最新を取得", _
            "git checkout -b feature/new-function  # 機能ブランチ作成", _
            "(機能を実装)", _
            "git add .                 # 変更をステージング", _
            "git commit -m ""Add new function""  # コミット", _
            "git push -u origin feature/new-function  # リモートに送信", _
            "(GitHubでPull Requestを作成)", _
            "(レビュー後、mainにマージ)" _
        )

        For i = 0 To UBound(scenario3)
            .Range("A" & row).Value = scenario3(i)
            .Range("A" & row).Font.Name = "Consolas"
            .Range("A" & row & ":H" & row).Merge
            row = row + 1
        Next i
        row = row + 1

        ' シナリオ4
        .Range("A" & row).Value = "シナリオ4: 朝イチの作業開始"
        .Range("A" & row).Font.Size = 14
        .Range("A" & row).Font.Bold = True
        .Range("A" & row).Interior.Color = RGB(255, 192, 0)
        .Range("A" & row & ":H" & row).Merge
        row = row + 1

        Dim scenario4 As Variant
        scenario4 = Array( _
            "git status                # 現在の状態確認", _
            "git checkout main         # mainに移動（作業ブランチにいた場合）", _
            "git pull origin main      # 最新を取得", _
            "git checkout feature/xxx  # 作業ブランチに戻る", _
            "git merge main            # mainの変更を取り込む（任意）", _
            "(作業開始)" _
        )

        For i = 0 To UBound(scenario4)
            .Range("A" & row).Value = scenario4(i)
            .Range("A" & row).Font.Name = "Consolas"
            .Range("A" & row & ":H" & row).Merge
            row = row + 1
        Next i

        ' 列幅
        .Columns("A").ColumnWidth = 80
    End With
End Sub

'==============================================================================
' トラブル対処シート
'==============================================================================
Private Sub FormatTroubleshootSheet()
    Dim ws As Worksheet
    Set ws = Sheets("トラブル対処")

    With ws
        ' タイトル
        .Range("A1").Value = "トラブル対処 - 困った時の解決法"
        .Range("A1").Font.Size = 20
        .Range("A1").Font.Bold = True
        .Range("A1:H1").Merge
        .Range("A1").Interior.Color = RGB(64, 64, 64)
        .Range("A1").Font.Color = RGB(255, 255, 255)

        ' ヘッダー
        .Range("A3").Value = "問題"
        .Range("B3").Value = "原因"
        .Range("C3").Value = "解決方法"
        .Range("A3:C3").Font.Bold = True
        .Range("A3:C3").Interior.Color = RGB(192, 80, 77)
        .Range("A3:C3").Font.Color = RGB(255, 255, 255)

        Dim troubles As Variant
        troubles = Array( _
            Array("git pushできない", "リモートに新しい変更がある", "git pull してから git push"), _
            Array("コンフリクト（競合）が発生", "同じ箇所を別々に変更した", "ファイルを編集して解決後 git add → git commit"), _
            Array("間違えてコミットした", "コミット内容に誤り", "git commit --amend（直前）または git revert"), _
            Array("変更を消してしまった", "git reset --hard を実行", "git reflog でコミットIDを探して git checkout"), _
            Array("ブランチを間違えた", "別のブランチで作業してた", "git stash → git checkout → git stash pop"), _
            Array("addし忘れたファイルがある", "コミット後に気づいた", "git add → git commit --amend --no-edit"), _
            Array("パスワードを聞かれる", "認証情報が未設定", "git config credential.helper store"), _
            Array("大きなファイルをpushできない", "GitHubの容量制限", "git reset HEAD^ でコミット取消し、.gitignoreに追加"), _
            Array("日本語ファイル名が文字化け", "文字コード設定", "git config --global core.quotepath false"), _
            Array("改行コードの警告が出る", "LF/CRLF混在", "git config --global core.autocrlf true (Windows)") _
        )

        Dim i As Long
        For i = 0 To UBound(troubles)
            .Range("A" & (4 + i)).Value = troubles(i)(0)
            .Range("B" & (4 + i)).Value = troubles(i)(1)
            .Range("C" & (4 + i)).Value = troubles(i)(2)
            .Range("C" & (4 + i)).Font.Name = "Consolas"
        Next i

        ' コンフリクト解決
        .Range("A15").Value = "コンフリクト（競合）の解決方法"
        .Range("A15").Font.Size = 14
        .Range("A15").Font.Bold = True
        .Range("A15").Interior.Color = RGB(255, 192, 0)
        .Range("A15:H15").Merge

        Dim conflict As Variant
        conflict = Array( _
            "1. コンフリクトが発生したファイルを開く", _
            "2. <<<<<<< HEAD と ======= と >>>>>>> の間の内容を確認", _
            "3. 正しい内容になるよう編集（マーカーも削除）", _
            "4. git add ファイル名", _
            "5. git commit（マージコミットが作成される）", _
            "", _
            "【ファイル内の見方】", _
            "<<<<<<< HEAD", _
            "自分の変更内容", _
            "'=======", _
            "相手の変更内容", _
            ">>>>>>> branch-name" _
        )

        For i = 0 To UBound(conflict)
            .Range("A" & (16 + i)).Value = conflict(i)
            If i >= 7 Then .Range("A" & (16 + i)).Font.Name = "Consolas"
            .Range("A" & (16 + i) & ":H" & (16 + i)).Merge
        Next i

        ' 困った時の基本
        .Range("A29").Value = "困った時の基本行動"
        .Range("A29").Font.Size = 14
        .Range("A29").Font.Bold = True
        .Range("A29").Interior.Color = RGB(0, 176, 80)
        .Range("A29").Font.Color = RGB(255, 255, 255)
        .Range("A29:H29").Merge

        Dim basicActions As Variant
        basicActions = Array( _
            "1. まず git status で状態確認", _
            "2. git log --oneline で履歴確認", _
            "3. 焦らない! 大抵の操作は取り消せる", _
            "4. わからなければ git reflog（最終手段）" _
        )

        For i = 0 To UBound(basicActions)
            .Range("A" & (30 + i)).Value = basicActions(i)
            .Range("A" & (30 + i) & ":H" & (30 + i)).Merge
        Next i

        ' 列幅
        .Columns("A").ColumnWidth = 30
        .Columns("B").ColumnWidth = 30
        .Columns("C").ColumnWidth = 50
    End With
End Sub
