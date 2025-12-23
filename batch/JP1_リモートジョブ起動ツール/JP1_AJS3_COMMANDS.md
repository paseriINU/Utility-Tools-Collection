# JP1/AJS3 コマンドリファレンス

JP1/AJS3（Job Management Partner 1 / Automatic Job Management System 3）で使用可能なコマンドの一覧と説明です。

---

## 目次

1. [コマンド一覧（カテゴリ別）](#コマンド一覧カテゴリ別)
2. [ジョブネット操作コマンド](#ジョブネット操作コマンド)
3. [情報取得コマンド](#情報取得コマンド)
4. [定義操作コマンド](#定義操作コマンド)
5. [スケジュール管理コマンド](#スケジュール管理コマンド)
6. [システム管理コマンド](#システム管理コマンド)
7. [共通オプション](#共通オプション)
8. [コマンドの配置場所](#コマンドの配置場所)

---

## コマンド一覧（カテゴリ別）

### ジョブネット操作

| コマンド | 説明 |
|----------|------|
| `ajsentry` | ジョブネットを即時実行 |
| `ajsrelease` | 保留を解除 |
| `ajspend` | 保留を設定 |
| `ajskilljob` | ジョブを強制終了 |
| `ajsrerun` | ジョブネットを再実行 |
| `ajssuspend` | 実行を一時停止 |
| `ajsresume` | 一時停止を解除 |
| `ajschgstat` | 実行状態を変更 |
| `ajsintrpt` | 実行を中断 |

### 情報取得

| コマンド | 説明 |
|----------|------|
| `ajsprint` | 定義情報を出力 |
| `ajsstatus` | 実行状態を取得 |
| `ajsshow` | 詳細情報を取得 |
| `ajsname` | 名前を解決 |
| `ajslogprint` | ログ情報を出力 |

### 定義操作

| コマンド | 説明 |
|----------|------|
| `ajsdefine` | ジョブネットを定義・登録 |
| `ajsdelete` | ジョブネットを削除 |
| `ajsimport` | 定義をインポート |
| `ajsexport` | 定義をエクスポート |
| `ajscopy` | 定義をコピー |
| `ajschgdef` | 定義を変更 |

### スケジュール管理

| コマンド | 説明 |
|----------|------|
| `ajsplan` | 実行予定を表示 |
| `ajscalendar` | カレンダーを管理 |
| `ajsschedule` | スケジュールを管理 |

### システム管理

| コマンド | 説明 |
|----------|------|
| `ajsstart` | JP1/AJS3サービスを起動 |
| `ajsstop` | JP1/AJS3サービスを停止 |
| `ajsstatus` | サービス状態を確認 |

---

## ジョブネット操作コマンド

### ajsentry（即時実行）

ジョブネットを即時実行（登録実行）するコマンドです。

```cmd
ajsentry [オプション] ジョブネットパス
```

**主なオプション**:

| オプション | 説明 |
|-----------|------|
| `-F スケジューラーサービス名` | スケジューラーサービスを指定（デフォルト: AJSROOT1） |
| `-h ホスト名` | JP1/AJS3マネージャーのホスト名を指定 |
| `-u ユーザー名` | JP1ユーザー名を指定 |
| `-p パスワード` | JP1パスワードを指定 |
| `-T 実行日時` | 実行開始日時を指定（YYYY/MM/DD HH:MM:SS形式） |
| `-X yes\|no` | 排他実行の指定（yes=排他あり） |

**実行例**:
```cmd
rem ローカル実行
ajsentry -F AJSROOT1 /main_unit/jobgroup1/daily_batch

rem リモート実行
ajsentry -h jp1server -u jp1admin -p password -F /main_unit/daily_batch

rem 日時指定実行
ajsentry -F AJSROOT1 -T "2025/12/20 10:00:00" /main_unit/daily_batch
```

**戻り値**:
- `0`: 正常終了
- `0以外`: 異常終了

**正常終了時メッセージ**:
```
KAVS1820-I ajsentryコマンドが正常終了しました。
```

---

### ajsrelease（保留解除）

保留中のジョブネットやジョブの保留を解除するコマンドです。

```cmd
ajsrelease [オプション] ジョブネットパス
```

**主なオプション**:

| オプション | 説明 |
|-----------|------|
| `-F スケジューラーサービス名` | スケジューラーサービスを指定 |
| `-h ホスト名` | JP1/AJS3マネージャーのホスト名を指定 |
| `-u ユーザー名` | JP1ユーザー名を指定 |
| `-p パスワード` | JP1パスワードを指定 |
| `-R` | 配下のユニットも再帰的に保留解除 |

**実行例**:
```cmd
rem 単一ジョブネットの保留解除
ajsrelease -F AJSROOT1 /main_unit/jobgroup1/weekly_batch

rem 配下のユニットも含めて保留解除
ajsrelease -F AJSROOT1 -R /main_unit/jobgroup1
```

**正常終了時メッセージ**:
```
KAVS1820-I ajsreleaseコマンドが正常終了しました。
```

---

### ajspend（保留設定）

ジョブネットやジョブに保留を設定するコマンドです。

```cmd
ajspend [オプション] ジョブネットパス
```

**主なオプション**:

| オプション | 説明 |
|-----------|------|
| `-F スケジューラーサービス名` | スケジューラーサービスを指定 |
| `-h ホスト名` | JP1/AJS3マネージャーのホスト名を指定 |
| `-u ユーザー名` | JP1ユーザー名を指定 |
| `-p パスワード` | JP1パスワードを指定 |
| `-R` | 配下のユニットも再帰的に保留設定 |

**実行例**:
```cmd
rem 保留設定
ajspend -F AJSROOT1 /main_unit/jobgroup1/weekly_batch
```

---

### ajskilljob（強制終了）

実行中のジョブを強制終了するコマンドです。

```cmd
ajskilljob [オプション] ジョブパス
```

**主なオプション**:

| オプション | 説明 |
|-----------|------|
| `-F スケジューラーサービス名` | スケジューラーサービスを指定 |
| `-h ホスト名` | JP1/AJS3マネージャーのホスト名を指定 |
| `-u ユーザー名` | JP1ユーザー名を指定 |
| `-p パスワード` | JP1パスワードを指定 |
| `-f` | 強制終了（確認なし） |

**実行例**:
```cmd
rem ジョブを強制終了
ajskilljob -F AJSROOT1 /main_unit/jobgroup1/daily_batch/job1
```

---

### ajsrerun（再実行）

異常終了したジョブネットやジョブを再実行するコマンドです。

```cmd
ajsrerun [オプション] ジョブネットパス
```

**主なオプション**:

| オプション | 説明 |
|-----------|------|
| `-F スケジューラーサービス名` | スケジューラーサービスを指定 |
| `-h ホスト名` | JP1/AJS3マネージャーのホスト名を指定 |
| `-u ユーザー名` | JP1ユーザー名を指定 |
| `-p パスワード` | JP1パスワードを指定 |
| `-b ユニットパス` | 再実行開始位置を指定 |

**実行例**:
```cmd
rem ジョブネット全体を再実行
ajsrerun -F AJSROOT1 /main_unit/jobgroup1/daily_batch

rem 特定のジョブから再実行
ajsrerun -F AJSROOT1 -b /main_unit/jobgroup1/daily_batch/job3 /main_unit/jobgroup1/daily_batch
```

---

### ajssuspend（一時停止）

実行中のジョブネットを一時停止するコマンドです。

```cmd
ajssuspend [オプション] ジョブネットパス
```

**主なオプション**:

| オプション | 説明 |
|-----------|------|
| `-F スケジューラーサービス名` | スケジューラーサービスを指定 |
| `-h ホスト名` | JP1/AJS3マネージャーのホスト名を指定 |

**実行例**:
```cmd
ajssuspend -F AJSROOT1 /main_unit/jobgroup1/daily_batch
```

---

### ajsresume（一時停止解除）

一時停止中のジョブネットの実行を再開するコマンドです。

```cmd
ajsresume [オプション] ジョブネットパス
```

**実行例**:
```cmd
ajsresume -F AJSROOT1 /main_unit/jobgroup1/daily_batch
```

---

### ajschgstat（状態変更）

ジョブネットやジョブの実行状態を変更するコマンドです。

```cmd
ajschgstat [オプション] -s 状態 ジョブネットパス
```

**主なオプション**:

| オプション | 説明 |
|-----------|------|
| `-F スケジューラーサービス名` | スケジューラーサービスを指定 |
| `-s 状態` | 変更後の状態（N=正常終了、W=警告終了、A=異常終了） |

**実行例**:
```cmd
rem 異常終了を正常終了に変更
ajschgstat -F AJSROOT1 -s N /main_unit/jobgroup1/daily_batch/job1
```

---

### ajsintrpt（中断）

実行中のジョブネットを中断するコマンドです。

```cmd
ajsintrpt [オプション] ジョブネットパス
```

**主なオプション**:

| オプション | 説明 |
|-----------|------|
| `-F スケジューラーサービス名` | スケジューラーサービスを指定 |
| `-k` | 強制中断（実行中のジョブも終了） |

**実行例**:
```cmd
rem 中断（実行中のジョブ完了後）
ajsintrpt -F AJSROOT1 /main_unit/jobgroup1/daily_batch

rem 強制中断
ajsintrpt -F AJSROOT1 -k /main_unit/jobgroup1/daily_batch
```

---

## 情報取得コマンド

### ajsprint（定義情報出力）

ジョブネットの定義情報を出力するコマンドです。ジョブ一覧の取得に使用します。

```cmd
ajsprint [オプション] ジョブネットパス
```

**主なオプション**:

| オプション | 説明 |
|-----------|------|
| `-F スケジューラーサービス名` | スケジューラーサービスを指定 |
| `-R` | 再帰的にサブユニットも取得 |
| `-a` | すべての属性を出力 |
| `-l` | 長い形式で出力 |

**実行例**:
```cmd
rem 全ジョブネットを再帰的に取得
ajsprint -F AJSROOT1 / -R

rem 特定グループ配下を取得
ajsprint -F AJSROOT1 /main_unit/jobgroup1 -R
```

**出力形式**:
```
unit=/main_unit/jobgroup1/daily_batch,daily_batch,ty=n,cm="日次バッチ処理";
unit=/main_unit/jobgroup1/weekly_batch,weekly_batch,ty=n,hd=y,cm="週次バッチ処理";
```

**出力フィールド**:

| フィールド | 説明 |
|-----------|------|
| `unit=` | ジョブネットのフルパス |
| `ty=n` | ユニットタイプ（n=ジョブネット、j=ジョブ、g=グループ） |
| `hd=y` | 保留状態（y=保留中） |
| `cm=` | コメント |
| `ex=` | 実行ホスト |
| `sc=` | スクリプトファイル |

---

### ajsstatus（実行状態取得）

ジョブネットの実行状態を取得するコマンドです。

```cmd
ajsstatus [オプション] ジョブネットパス
```

**主なオプション**:

| オプション | 説明 |
|-----------|------|
| `-F スケジューラーサービス名` | スケジューラーサービスを指定 |
| `-h ホスト名` | JP1/AJS3マネージャーのホスト名を指定 |
| `-u ユーザー名` | JP1ユーザー名を指定 |
| `-p パスワード` | JP1パスワードを指定 |
| `-l` | 詳細形式で出力 |

**実行例**:
```cmd
ajsstatus -F AJSROOT1 /main_unit/jobgroup1/daily_batch
```

**状態の種類**:

| 状態 | 説明 |
|------|------|
| `Now running` | 実行中 |
| `Running + Warning` | 警告付きで実行中 |
| `Wait for start time` | 開始時刻待ち |
| `Wait for prev to end` | 先行終了待ち |
| `Waiting` | 待機中 |
| `Queuing` | キューイング中 |
| `Ended normally` | 正常終了 |
| `Ended with warning` | 警告終了 |
| `Ended abnormally` | 異常終了 |
| `Killed` | 強制終了 |
| `Unknown end status` | 終了状態不明 |
| `Being held` | 保留中 |
| `Bypassed` | スキップ |
| `Not scheduled to execute` | 実行予定なし |

---

### ajsshow（詳細情報取得）

ジョブネットの詳細情報を取得するコマンドです。

```cmd
ajsshow [オプション] ジョブネットパス
```

**主なオプション**:

| オプション | 説明 |
|-----------|------|
| `-F スケジューラーサービス名` | スケジューラーサービスを指定 |
| `-h ホスト名` | JP1/AJS3マネージャーのホスト名を指定 |
| `-u ユーザー名` | JP1ユーザー名を指定 |
| `-p パスワード` | JP1パスワードを指定 |
| `-E` | 実行結果の詳細情報を取得 |
| `-i` | ユニット定義情報を取得 |
| `-g 世代番号` | 取得する実行世代を指定 |

**実行例**:
```cmd
rem 実行結果の詳細を取得（-Eオプション）
ajsshow -F AJSROOT1 -E /main_unit/jobgroup1/daily_batch

rem 状態を取得（-iオプション + 2バイト版フォーマット指示子）
ajsshow -F AJSROOT1 -g 1 -i "%CC" /main_unit/jobgroup1/daily_batch

rem 標準出力ファイルパスを取得（-iオプション + 2バイト版フォーマット指示子）
ajsshow -F AJSROOT1 -g 1 -i "%so" /main_unit/jobgroup1/daily_batch/job1
```

**出力例（-Eオプション）**:
```
UNIT-NAME       : /main_unit/jobgroup1/daily_batch
STATUS          : ENDED NORMALLY
START-TIME      : 2025/12/17 10:30:00
END-TIME        : 2025/12/17 10:32:35
RETURN-CODE     : 0
EXEC-HOST       : server01
```

---

### ajsname（名前解決）

ジョブネット名からパスを解決、またはその逆を行うコマンドです。

```cmd
ajsname [オプション] 名前またはパス
```

**主なオプション**:

| オプション | 説明 |
|-----------|------|
| `-F スケジューラーサービス名` | スケジューラーサービスを指定 |
| `-n` | パスから名前を取得 |
| `-p` | 名前からパスを取得 |

---

### ajslogprint（ログ出力）

ジョブの実行ログ情報を出力するコマンドです。

```cmd
ajslogprint [オプション] ジョブパス
```

**主なオプション**:

| オプション | 説明 |
|-----------|------|
| `-F スケジューラーサービス名` | スケジューラーサービスを指定 |
| `-o 出力ファイル` | ログの出力先ファイル |
| `-g 世代番号` | 取得する実行世代を指定 |

---

## 定義操作コマンド

### ajsdefine（定義登録）

ジョブネットの定義を登録するコマンドです。

```cmd
ajsdefine [オプション] 定義ファイル
```

**主なオプション**:

| オプション | 説明 |
|-----------|------|
| `-F スケジューラーサービス名` | スケジューラーサービスを指定 |
| `-r` | 既存の定義を上書き |

**実行例**:
```cmd
ajsdefine -F AJSROOT1 jobnet_definition.txt
```

---

### ajsdelete（定義削除）

ジョブネットの定義を削除するコマンドです。

```cmd
ajsdelete [オプション] ジョブネットパス
```

**主なオプション**:

| オプション | 説明 |
|-----------|------|
| `-F スケジューラーサービス名` | スケジューラーサービスを指定 |
| `-R` | 配下のユニットも再帰的に削除 |
| `-f` | 確認なしで削除 |

**実行例**:
```cmd
rem 確認なしで削除
ajsdelete -F AJSROOT1 -f /main_unit/jobgroup1/old_batch
```

---

### ajsimport（インポート）

エクスポートした定義ファイルをインポートするコマンドです。

```cmd
ajsimport [オプション] -i インポートファイル 登録先パス
```

**実行例**:
```cmd
ajsimport -F AJSROOT1 -i exported_jobs.txt /main_unit/jobgroup1
```

---

### ajsexport（エクスポート）

ジョブネットの定義をファイルにエクスポートするコマンドです。

```cmd
ajsexport [オプション] -o 出力ファイル ジョブネットパス
```

**主なオプション**:

| オプション | 説明 |
|-----------|------|
| `-F スケジューラーサービス名` | スケジューラーサービスを指定 |
| `-o 出力ファイル` | エクスポート先ファイル |
| `-R` | 配下のユニットも再帰的にエクスポート |

**実行例**:
```cmd
ajsexport -F AJSROOT1 -o backup.txt -R /main_unit/jobgroup1
```

---

### ajscopy（コピー）

ジョブネットの定義をコピーするコマンドです。

```cmd
ajscopy [オプション] コピー元パス コピー先パス
```

**実行例**:
```cmd
ajscopy -F AJSROOT1 /main_unit/jobgroup1/daily_batch /main_unit/jobgroup2/daily_batch_copy
```

---

### ajschgdef（定義変更）

ジョブネットの定義を変更するコマンドです。

```cmd
ajschgdef [オプション] ジョブネットパス
```

---

## スケジュール管理コマンド

### ajsplan（実行予定表示）

ジョブネットの実行予定を表示するコマンドです。

```cmd
ajsplan [オプション] ジョブネットパス
```

**主なオプション**:

| オプション | 説明 |
|-----------|------|
| `-F スケジューラーサービス名` | スケジューラーサービスを指定 |
| `-d 日付` | 対象日を指定（YYYY/MM/DD形式） |
| `-n 件数` | 表示する予定の件数 |

**実行例**:
```cmd
rem 今日の実行予定を表示
ajsplan -F AJSROOT1 -d today /main_unit/jobgroup1

rem 今後10件の実行予定を表示
ajsplan -F AJSROOT1 -n 10 /main_unit/jobgroup1/daily_batch
```

---

### ajscalendar（カレンダー管理）

カレンダー情報を管理するコマンドです。

```cmd
ajscalendar [オプション] カレンダー名
```

**サブコマンド**:

| サブコマンド | 説明 |
|------------|------|
| `-add` | カレンダーを追加 |
| `-del` | カレンダーを削除 |
| `-show` | カレンダー情報を表示 |

---

### ajsschedule（スケジュール管理）

ジョブネットのスケジュールを管理するコマンドです。

```cmd
ajsschedule [オプション] ジョブネットパス
```

---

## システム管理コマンド

### ajsstart（サービス起動）

JP1/AJS3サービスを起動するコマンドです。

```cmd
ajsstart [オプション]
```

**主なオプション**:

| オプション | 説明 |
|-----------|------|
| `-F スケジューラーサービス名` | 起動するスケジューラーサービスを指定 |
| `-c` | コールドスタート |
| `-w` | ウォームスタート |

---

### ajsstop（サービス停止）

JP1/AJS3サービスを停止するコマンドです。

```cmd
ajsstop [オプション]
```

**主なオプション**:

| オプション | 説明 |
|-----------|------|
| `-F スケジューラーサービス名` | 停止するスケジューラーサービスを指定 |
| `-f` | 強制停止 |
| `-n` | 計画停止（実行中のジョブ完了後） |

---

## 共通オプション

ほとんどのajsコマンドで使用できる共通オプションです。

| オプション | 説明 |
|-----------|------|
| `-F スケジューラーサービス名` | スケジューラーサービスを指定（デフォルト: AJSROOT1） |
| `-h ホスト名` | JP1/AJS3マネージャーのホスト名を指定（リモート接続時） |
| `-u ユーザー名` | JP1ユーザー名を指定（リモート接続時） |
| `-p パスワード` | JP1パスワードを指定（リモート接続時） |
| `-?` | ヘルプを表示 |

---

## コマンドの配置場所

JP1/AJS3コマンドは通常、以下のパスにインストールされています：

**Windows**:
```
C:\Program Files\HITACHI\JP1AJS3\bin\
C:\Program Files (x86)\HITACHI\JP1AJS3\bin\
C:\Program Files\Hitachi\JP1AJS2\bin\
C:\Program Files (x86)\Hitachi\JP1AJS2\bin\
```

**Linux/UNIX**:
```
/opt/jp1ajs3/bin/
/opt/jp1ajs2/bin/
```

---

## エラーメッセージ

JP1/AJS3コマンドの主なエラーメッセージ：

| メッセージID | 説明 |
|------------|------|
| KAVS0221-E | ジョブが異常終了しました |
| KAVS0222-E | ジョブネットが異常終了しました |
| KAVS1820-I | コマンドが正常終了しました |
| KAVS1821-E | コマンドが異常終了しました |
| KAVS4200-E | 認証に失敗しました |
| KAVS4201-E | ユーザーに権限がありません |

---

## 参考資料

- [JP1/AJS3 コマンドリファレンス](https://www.hitachi.co.jp/Prod/comp/soft1/manual/pc/d3K2211/AJSK01/EU210000.HTM)
- [JP1/AJS3 運用ガイド](https://www.hitachi.co.jp/Prod/comp/soft1/manual/pc/d3K2211/AJSK01/EU200000.HTM)
- [JP1 Version 13 製品マニュアル](https://itpfdoc.hitachi.co.jp/manuals/3021/30213D3200/JP1_manuals.html)

---

**作成日**: 2025-12-18
