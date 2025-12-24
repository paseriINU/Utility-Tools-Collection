# JP1/AJS3 Web Console REST API リファレンス

このドキュメントはJP1/AJS3 Web Console REST APIの使用方法をまとめたものです。

## 基本情報

### 接続先
- **HTTP**: `http://{Web Consoleサーバー}:22252`
- **HTTPS**: `https://{Web Consoleサーバー}:22253`

### 認証
- ヘッダー: `X-AJS-Authorization`
- 形式: Base64エンコードした `{JP1ユーザー}:{パスワード}`

### 共通パラメータ
| パラメータ | 説明 | 例 |
|-----------|------|-----|
| manager | JP1/AJS3 Managerホスト名 | localhost |
| serviceName | スケジューラーサービス名 | AJSROOT1 |
| location | ユニットパス | /jobnet1/job1 |

---

## API一覧（概要）

JP1/AJS3 Web Console REST APIで利用可能な全19個のAPIです。

| 番号 | API名 | 機能 | ドキュメント |
|------|-------|------|-------------|
| 7.1.1 | ユニット一覧の取得API | ユニットの一覧情報（状態・execID含む）を取得 | ✅ 詳細あり |
| 7.1.2 | ユニット情報の取得API | 単一ユニットの詳細情報を取得 | ✅ 詳細あり |
| 7.1.3 | 実行結果詳細の取得API | ジョブの実行結果詳細（標準エラー出力相当）を取得 | ✅ 詳細あり |
| 7.1.4 | 計画実行登録API | スケジュールに従った計画実行を登録 | ✅ 詳細あり |
| 7.1.5 | 確定実行登録API | 確定スケジュールでの実行を登録 | ✅ 詳細あり |
| 7.1.6 | 即時実行登録API | 即座にジョブネットを実行登録 | ✅ 詳細あり |
| 7.1.7 | 登録解除API | 登録済みの実行を解除 | ✅ 詳細あり |
| 7.1.8 | 保留属性変更API | ジョブ/ジョブネットの保留状態を変更 | ✅ 詳細あり |
| 7.1.9 | 遅延監視変更API | 遅延監視設定を変更 | ✅ 詳細あり |
| 7.1.10 | ジョブ状態変更API | ジョブの状態（正常終了/異常終了等）を変更 | ✅ 詳細あり |
| 7.1.11 | 計画一時変更（日時変更）API | スケジュール日時を一時的に変更 | ✅ 詳細あり |
| 7.1.12 | 計画一時変更（即時実行）API | スケジュールを即時実行に一時変更 | ✅ 詳細あり |
| 7.1.13 | 計画一時変更（実行中止）API | スケジュール実行を中止に一時変更 | ✅ 詳細あり |
| 7.1.14 | 計画一時変更（変更解除）API | 一時変更を解除して元に戻す | ✅ 詳細あり |
| 7.1.15 | 中断API | 実行中のジョブを中断 | ✅ 詳細あり |
| 7.1.16 | 強制終了API | 実行中のジョブを強制終了 | ✅ 詳細あり |
| 7.1.17 | 再実行API | ジョブを再実行 | ✅ 詳細あり |
| 7.1.18 | バージョン情報の取得API | Web Consoleのバージョン情報を取得 | ✅ 詳細あり |
| 7.1.19 | プロトコルバージョンの取得API | JP1/AJS3 - Managerのプロトコルバージョンを取得 | ✅ 詳細あり |

> **公式ドキュメント**: [JP1/AJS3 Web Console REST API](https://itpfdoc.hitachi.co.jp/manuals/3021/30213b1920/AJSO0278.HTM)

---

## API詳細

### 7.1.1 ユニット一覧の取得API

指定したユニットまたはユニット配下の，ジョブグループ，ジョブネット，およびジョブの情報を取得します。

#### 実行権限

ログインしたJP1ユーザーが，次のどれかのJP1権限を持つ必要があります：
- JP1_AJS_Admin権限
- JP1_AJS_Manager権限
- JP1_AJS_Editor権限
- JP1_AJS_Operator権限
- JP1_AJS_Guest権限

#### リクエスト形式

```
GET /ajs/api/v1/objects/statuses?query
```

#### パラメータ一覧

| パラメータ | 説明 | 必須/任意 |
|-----------|------|----------|
| mode | 固定で「search」を指定 | 必須 |
| manager | マネージャーホスト名またはIPアドレス（1〜255バイト） | 必須 |
| serviceName | スケジューラーサービス名（1〜30バイト） | 必須 |
| location | 取得したいユニットの上位ユニットのパス（1〜930バイト） | 必須 |
| searchLowerUnits | 直下1階層(NO)か配下すべて(YES)か | 任意（デフォルト: NO） |
| searchTarget | 定義のみ(DEFINITION)か定義と状態(DEFINITION_AND_STATUS)か | 任意（デフォルト: DEFINITION_AND_STATUS） |
| unitName | 取得したいユニットの名称 | 任意 |
| unitNameMatchMethods | ユニット名の比較方法（EQ/BW/EW/NE/CO/NC/RE/NO） | 任意（デフォルト: NO） |
| execID | 実行ID（@[mmmm]{A〜Z}nnnn形式、例: @10A200） | 任意 |
| unitType | ユニット種別（GROUP/ROOT/NET/JOB/NO） | 任意（デフォルト: NO） |
| generation | 世代（STATUS/EXECID/PERIOD） | 任意（デフォルト: STATUS） |
| periodBegin | 世代の開始日時（YYYY-MM-DDThh:mm形式） | 任意 |
| periodEnd | 世代の終了日時（YYYY-MM-DDThh:mm形式） | 任意 |
| status | ユニットの状態 | 任意（デフォルト: NO） |
| delayStatus | 遅延状態 | 任意（デフォルト: NO） |
| holdPlan | 保留予定の有無 | 任意（デフォルト: NO） |
| unitComment | ユニットのコメント | 任意 |
| unitCommentMatchMethods | コメントの比較方法 | 任意（デフォルト: NO） |
| execHost | 実行ホスト | 任意 |
| execHostMatchMethods | 実行ホスト名の比較方法 | 任意（デフォルト: NO） |
| releaseID | リリースID（1〜30バイト） | 任意 |
| releaseInfoSearchMethods | リリース情報の取得方法 | 任意（デフォルト: NO） |

#### ステータスコード

| コード | テキスト | 説明 |
|--------|---------|------|
| 200 | OK | ユニット一覧の取得に成功 |
| 400 | Bad Request | クエリ文字列が不正 |
| 401 | Unauthorized | 認証が必要 |
| 403 | Forbidden | 実行権限がない |
| 404 | Not found | リソースがない |
| 409 | Conflict | リクエストは現在のリソース状態と矛盾 |
| 412 | Precondition failed | Web Consoleサーバが利用できない |
| 500 | Server-side error | サーバ処理エラー |

#### レスポンス形式

```json
{
  "statuses": [ステータス監視のリソース,...],
  "all": すべての情報を取得できたかどうか
}
```

#### レスポンス詳細

| メンバー | データ型 | 説明 |
|---------|---------|------|
| statuses | object[] | ステータス監視のリソースの配列（最大1,000件） |
| all | boolean | 取得件数が1,000件を超えていない場合true |

#### レスポンス例

```json
{
  "statuses": [
    {
      "definition": {
        "owner": "jp1admin",
        "customJobType": "",
        "registerStatus": "YES",
        "rootJobnetName": "/JobGroup/Jobnet",
        "recoveryUnit": false,
        "unitType": "ROOTNET",
        "unitComment": "",
        "simpleUnitName": "Jobnet",
        "parameters": "",
        "execAgent": "",
        "execFileName": "",
        "wait": false,
        "jobnetReleaseUnit": false,
        "jp1ResourceGroup": "",
        "unitName": "/JobGroup/Jobnet",
        "unitID": 1560
      },
      "release": null,
      "unitStatus": {
        "status": "RUNNING",
        "execHost": "",
        "startDelayStatus": "NO",
        "nestStartDelayStatus": "NO",
        "endDelayStatus": "NO",
        "nestEndDelayStatus": "NO",
        "startDelayTime": "",
        "endDelayTime": "",
        "changeType": "NO",
        "registerTime": "",
        "jobNumber": -1,
        "retCode": "",
        "simpleUnitName": "Jobnet",
        "schStartTime": "2015-09-02T00:00:00+09:00",
        "reStartTime": "",
        "endTime": "",
        "holdAttr": "NO",
        "startTime": "2015-09-02T22:50:28+09:00",
        "unitName": "/JobGroup/Jobnet",
        "execID": "@A2959"
      }
    }
  ],
  "all": true
}
```

#### 使用例

ジョブグループ（/JobGroup）直下のユニット一覧（最新状態）を取得する場合：

```
GET /ajs/api/v1/objects/statuses?mode=search&manager=HOSTM&serviceName=AJSROOT1&location=%2FJobGroup HTTP/1.1
Host: HOSTW:22252
Accept-Language: ja
X-AJS-Authorization: dXNlcjpwYXNzd29yZA==
```

#### 注意事項

- 条件を満たすユニットが存在しない場合、0件のユニット一覧を返却
- 取得できるユニットは**最大1,000件**
- 参照権限がないユニットは取得結果に含まれない

---

### 7.1.3 実行結果詳細の取得API

実行が終了したユニットの，実行結果の詳細を取得します。

#### 実行権限

ログインしたJP1ユーザーが，次のどれかのJP1権限を持つ必要があります：
- JP1_AJS_Admin権限
- JP1_AJS_Manager権限
- JP1_AJS_Editor権限
- JP1_AJS_Operator権限
- JP1_AJS_Guest権限

#### リクエスト形式

```
GET /ajs/api/v1/objects/statuses/{unitName}:{execID}/actions/execResultDetails/invoke?query
```

#### リソース識別情報

| パラメータ | データ型 | 説明 | 必須/任意 |
|-----------|---------|------|----------|
| unitName | string | ユニット完全名（1〜930バイト） | 必須 |
| execID | string | 実行ID（@[mmmm]{A〜Z}nnnn形式、例: @10A200） | 必須 |

#### クエリパラメータ

| パラメータ | 説明 | 必須/任意 |
|-----------|------|----------|
| manager | マネージャーホスト名またはIPアドレス（1〜255バイト） | 必須 |
| serviceName | スケジューラーサービス名（1〜30バイト） | 必須 |

#### ステータスコード

| コード | メッセージ | 説明 |
|--------|----------|------|
| 200 | OK | 実行結果詳細の取得に成功 |
| 400 | Bad Request | クエリ文字列が不正 |
| 401 | Unauthorized | 認証が必要 |
| 403 | Forbidden | 実行権限がない |
| 404 | Not found | リソースがない、またはアクセス権限がない |
| 409 | Conflict | リクエストは現在のリソース状態と矛盾 |
| 412 | Precondition failed | Web Consoleサーバが利用できない |
| 500 | Server-side error | サーバ処理エラー |

#### レスポンス形式

```json
{
  "execResultDetails": "実行結果詳細",
  "all": すべての情報を取得できたかどうか
}
```

#### レスポンス詳細

| メンバー | データ型 | 説明 |
|---------|---------|------|
| execResultDetails | string | 実行結果詳細（**最大5MB**）。5MBを超える場合は切り捨て。改行コード（\n または \r\n）を含む。結果がない場合は空文字列。 |
| all | boolean | すべての実行結果詳細を取得できた場合true |

#### 使用例

指定したジョブの実行結果詳細を取得する場合：

```
GET /ajs/api/v1/objects/statuses/%2FJobGroup%2FJobnet%2FJob:%40A100/actions/execResultDetails/invoke?manager=HOSTM&serviceName=AJSROOT1 HTTP/1.1
Host: HOSTW:22252
Accept-Language: ja
X-AJS-Authorization: dXNlcjpwYXNzd29yZA==
```

#### レスポンス例

```json
{
  "execResultDetails": "実行結果詳細",
  "all": true
}
```

#### 注意事項

- **実行結果詳細**を取得（標準エラー出力相当）
- 最大サイズは**5MB**（超過分は切り捨て）
- 改行コードは実行環境依存（\n または \r\n）
- 実行IDは事前にユニット一覧取得API (7.1.1) で取得する必要がある
- URLエンコードが必要（例: `/` → `%2F`, `@` → `%40`）

---

### 7.1.2 ユニット情報の取得API

指定したユニットの情報を取得します。

#### リクエスト形式

```
GET /ajs/api/v1/objects/statuses/{unitName}:{execID}?query
```

#### パラメータ

| パラメータ | 説明 | 必須/任意 |
|-----------|------|----------|
| unitName | ユニット完全名（1〜930バイト） | 必須 |
| execID | 実行ID（@[mmmm]{A〜Z}nnnn形式） | 必須 |
| manager | マネージャーホスト名またはIPアドレス | 必須 |
| serviceName | スケジューラーサービス名 | 必須 |

#### レスポンス

成功時（200）: ステータス監視リソースをJSON形式で返却

---

### 7.1.4 計画実行登録API

指定したジョブネットを計画実行登録します。

#### リクエスト形式

```
POST /ajs/api/v1/objects/definitions/{unitName}/actions/registerPlannedExec/invoke
```

#### パラメータ

| パラメータ | 説明 | 必須/任意 |
|-----------|------|----------|
| unitName | ユニット完全名 | 必須 |
| manager | マネージャーホスト名またはIPアドレス | 必須 |
| serviceName | スケジューラーサービス名 | 必須 |
| passedDaemonStarts | スケジューラー起動時に予定時刻超過時の動作 | 必須 |
| passedRegForExe | 実行登録時に予定時刻超過時の動作 | 必須 |
| holding | ジョブネット実行を保留するか | 任意（デフォルト: false） |
| macro | マクロ変数設定（最大32個） | 任意 |

#### レスポンス

成功時（200）: メッセージボディなし

---

### 7.1.5 確定実行登録API

指定したジョブネットを確定実行登録します。

#### リクエスト形式

```
POST /ajs/api/v1/objects/definitions/{unitName}/actions/registerFixedExec/invoke
```

#### パラメータ

| パラメータ | 説明 | 必須/任意 |
|-----------|------|----------|
| unitName | ユニット完全名 | 必須 |
| manager | マネージャーホスト名またはIPアドレス | 必須 |
| serviceName | スケジューラーサービス名 | 必須 |
| fixedScheduleFrom / fixedScheduleTo | 確定期間（どれか1つ以上必須） | 条件付き必須 |
| futureGeneration | 未来世代数 | 条件付き必須 |
| startTime | 日時指定 | 条件付き必須 |
| holding | ジョブネット実行を保留するか | 任意（デフォルト: false） |
| macro | マクロ変数設定（最大32個） | 任意 |

#### レスポンス

- 確定期間/未来世代数指定時: メッセージボディなし
- 日時指定時: `execID`（実行ID）をJSON形式で返却

---

### 7.1.6 即時実行登録API

指定したジョブネットを即時実行登録します。

#### リクエスト形式

```
POST /ajs/api/v1/objects/definitions/{unitName}/actions/registerImmediateExec/invoke
```

#### パラメータ

| パラメータ | 説明 | 必須/任意 |
|-----------|------|----------|
| unitName | ユニット完全名 | 必須 |
| manager | マネージャーホスト名またはIPアドレス | 必須 |
| serviceName | スケジューラーサービス名 | 必須 |
| startCondition | 起動条件パラメーター | 任意 |
| holding | ジョブネット実行を保留するか | 任意（デフォルト: false） |
| macro | マクロ変数設定（最大32個） | 任意 |

#### レスポンス

成功時（200）: `execID`（実行ID）をJSON形式で返却

---

### 7.1.7 登録解除API

指定した実行登録中のジョブネットの登録を解除します。

#### リクエスト形式

```
POST /ajs/api/v1/objects/definitions/{unitName}/actions/cancelRegistration/invoke
```

#### パラメータ

| パラメータ | 説明 | 必須/任意 |
|-----------|------|----------|
| unitName | ユニット完全名 | 必須 |
| manager | マネージャーホスト名またはIPアドレス | 必須 |
| serviceName | スケジューラーサービス名 | 必須 |
| dateType | 暦日または実行日 | 任意（期間指定時は必須） |
| begin | 開始日（YYYY-MM-DD形式） | 任意（期間指定時は必須） |
| end | 終了日（YYYY-MM-DD形式） | 任意（期間指定時は必須） |

#### レスポンス

成功時（200）: メッセージボディなし

---

### 7.1.8 保留属性変更API

指定したユニットの保留状態を変更します。

#### リクエスト形式

```
POST /ajs/api/v1/objects/statuses/{unitName}:{execID}/actions/changeHold/invoke
```

#### パラメータ

| パラメータ | 説明 | 必須/任意 |
|-----------|------|----------|
| unitName | ユニット完全名 | 必須 |
| execID | 実行ID | 必須 |
| manager | マネージャーホスト名またはIPアドレス | 必須 |
| serviceName | スケジューラーサービス名 | 必須 |
| holdAttr | `SET`（保留設定）または `RELEASE`（保留解除） | 必須 |

#### レスポンス

成功時（200）: メッセージボディなし

---

### 7.1.9 遅延監視変更API

実行中のジョブネットの遅延監視設定を一時的に変更します。

#### リクエスト形式

```
POST /ajs/api/v1/objects/statuses/{unitName}:{execID}/actions/changeDelayMonitor/invoke
```

#### パラメータ

| パラメータ | 説明 | 必須/任意 |
|-----------|------|----------|
| unitName | ユニット完全名 | 必須 |
| execID | 実行ID | 必須 |
| manager | マネージャーホスト名またはIPアドレス | 必須 |
| serviceName | スケジューラーサービス名 | 必須 |
| delayedStart | 開始遅延の監視方法 | 任意（いずれか1つ以上必須） |
| delayedStartTime | 開始遅延監視開始時刻 | 任意 |
| delayedEnd | 終了遅延の監視方法 | 任意 |
| delayedEndTime | 終了遅延監視開始時刻 | 任意 |
| monitoringJobnet | 実行所要時間による監視方法 | 任意 |
| timeRequiredForExecution | ジョブネット実行所要時間 | 任意 |

#### レスポンス

成功時（200）: メッセージボディなし

---

### 7.1.10 ジョブ状態変更API

指定したジョブの状態を変更します。（バージョン11-10以降）

#### リクエスト形式

```
POST /ajs/api/v1/objects/statuses/{unitName}:{execID}/actions/changeStatus/invoke
```

#### パラメータ

| パラメータ | 説明 | 必須/任意 |
|-----------|------|----------|
| unitName | ユニット完全名 | 必須 |
| execID | 実行ID | 必須 |
| manager | マネージャーホスト名またはIPアドレス | 必須 |
| serviceName | スケジューラーサービス名 | 必須 |
| newStatus | 変更後の状態 | 任意（いずれか1つ以上必須） |
| newReturnCode | 変更後の終了コード | 任意 |

#### レスポンス

成功時（200）: メッセージボディなし

---

### 7.1.11 計画一時変更（日時変更）API

実行登録済みのジョブネットの実行開始予定日時を一時的に変更します。（バージョン11-10以降）

#### リクエスト形式

```
POST /ajs/api/v1/objects/statuses/{unitName}:{execID}/actions/changeStartTime/invoke
```

#### パラメータ

| パラメータ | 説明 | 必須/任意 |
|-----------|------|----------|
| unitName | ユニット完全名 | 必須 |
| execID | 実行ID | 必須 |
| manager | マネージャーホスト名またはIPアドレス | 必須 |
| serviceName | スケジューラーサービス名 | 必須 |
| kindTime | `ABSOLUTETIME`（絶対日時）または `RELATIVETIME`（相対日時） | 必須 |
| absoluteTime | 絶対日時（YYYY-MM-DDThh:mm形式）※kindTime=ABSOLUTETIME時は必須 | 条件付き必須 |
| relativeSign | 相対指定時の方向（`+` または `-`）※kindTime=RELATIVETIME時は必須 | 条件付き必須 |
| relativeDay | 相対指定時の相対日（0-99）※kindTime=RELATIVETIME時は必須 | 条件付き必須 |
| relativeTime | 相対指定時の相対時刻（hh:mm形式）※kindTime=RELATIVETIME時は必須 | 条件付き必須 |
| pushedAhead | 前倒し実行時の動作 | 任意 |
| changeLower | 配下ジョブネット連動変更の有無 | 任意 |

#### レスポンス

成功時（200）: メッセージボディなし

---

### 7.1.12 計画一時変更（即時実行）API

実行登録済みのジョブネットの実行スケジュールを一時的に変更し、即時実行します。

#### リクエスト形式

```
POST /ajs/api/v1/objects/statuses/{unitName}:{execID}/actions/execImmediate/invoke
```

#### パラメータ

| パラメータ | 説明 | 必須/任意 |
|-----------|------|----------|
| unitName | ユニット完全名 | 必須 |
| execID | 実行ID | 必須 |
| manager | マネージャーホスト名またはIPアドレス | 必須 |
| serviceName | スケジューラーサービス名 | 必須 |
| pushedAhead | 開始時間前倒し時の動作 | 任意（デフォルト: ADD） |
| changeLower | 配下ジョブネットの開始日時変更動作 | 任意（デフォルト: NOTSHIFT） |

#### レスポンス

成功時（200）: メッセージボディなし

---

### 7.1.13 計画一時変更（実行中止）API

実行登録済みのジョブネットの実行スケジュールを一時的に変更し、実行を中止します。

#### リクエスト形式

```
POST /ajs/api/v1/objects/statuses/{unitName}:{execID}/actions/cancelExecSchedule/invoke
```

#### パラメータ

| パラメータ | 説明 | 必須/任意 |
|-----------|------|----------|
| unitName | ユニット完全名 | 必須 |
| execID | 実行ID | 必須 |
| manager | マネージャーホスト名またはIPアドレス | 必須 |
| serviceName | スケジューラーサービス名 | 必須 |

#### レスポンス

成功時（200）: メッセージボディなし

---

### 7.1.14 計画一時変更（変更解除）API

実行開始日時の一時変更および実行中止を解除し、変更前の情報に戻します。

#### リクエスト形式

```
POST /ajs/api/v1/objects/statuses/{unitName}:{execID}/actions/restoreChangedSchedule/invoke
```

#### パラメータ

| パラメータ | 説明 | 必須/任意 |
|-----------|------|----------|
| unitName | ユニット完全名 | 必須 |
| execID | 実行ID | 必須 |
| manager | マネージャーホスト名またはIPアドレス | 必須 |
| serviceName | スケジューラーサービス名 | 必須 |

#### レスポンス

成功時（200）: メッセージボディなし

---

### 7.1.15 中断API

指定した実行中のルートジョブネットを中断します。（バージョン11-10以降）

#### リクエスト形式

```
POST /ajs/api/v1/objects/statuses/{unitName}:{execID}/actions/interrupt/invoke
```

#### パラメータ

| パラメータ | 説明 | 必須/任意 |
|-----------|------|----------|
| unitName | ユニット完全名 | 必須 |
| execID | 実行ID | 必須 |
| manager | マネージャーホスト名またはIPアドレス | 必須 |
| serviceName | スケジューラーサービス名 | 必須 |

#### レスポンス

成功時（200）: メッセージボディなし

---

### 7.1.16 強制終了API

指定した実行中のジョブおよびルートジョブネットを強制終了します。

#### リクエスト形式

```
POST /ajs/api/v1/objects/statuses/{unitName}:{execID}/actions/kill/invoke
```

#### パラメータ

| パラメータ | 説明 | 必須/任意 |
|-----------|------|----------|
| unitName | ユニット完全名 | 必須 |
| execID | 実行ID | 必須 |
| manager | マネージャーホスト名またはIPアドレス | 必須 |
| serviceName | スケジューラーサービス名 | 必須 |

#### レスポンス

成功時（200）: メッセージボディなし

---

### 7.1.17 再実行API

指定したユニットを再実行します。（バージョン11-10以降）

#### リクエスト形式

```
POST /ajs/api/v1/objects/statuses/{unitName}:{execID}/actions/rerun/invoke
```

#### パラメータ

| パラメータ | 説明 | 必須/任意 |
|-----------|------|----------|
| unitName | ユニット完全名 | 必須 |
| execID | 実行ID | 必須 |
| manager | マネージャーホスト名またはIPアドレス | 必須 |
| serviceName | スケジューラーサービス名 | 必須 |
| rerunMethod | 再実行方法（RootRerunType/RerunType定数） | 必須 |
| rerunOptions | 再実行オプション（保留状態にするか等） | 任意 |

#### レスポンス

成功時（200）: メッセージボディなし

---

### 7.1.18 バージョン情報の取得API

製品のバージョン情報を取得します。

#### リクエスト形式

```
GET /ajs/api/v1/version?query
```

#### パラメータ

| パラメータ | 説明 | 必須/任意 |
|-----------|------|----------|
| manager | マネージャーホスト名またはIPアドレス | 必須 |

#### レスポンス

成功時（200）: 以下をJSON形式で返却

| メンバー | データ型 | 説明 |
|---------|---------|------|
| productName | string | 製品名 |
| productVersion | string | 製品バージョン（VV-RR-SS形式） |
| displayProductVersion | string | 表示用バージョン（VV-RRまたはVV-RR-SS形式） |
| productVersionNumber | int | バージョン番号（VVRRSS形式）※**v11-10以降のみ** |

#### レスポンス例

```json
{
  "productName": "JP1/AJS3 - Web Console",
  "productVersion": "11-10-00",
  "displayProductVersion": "11-10",
  "productVersionNumber": 111000
}
```

---

### 7.1.19 プロトコルバージョンの取得API

JP1/AJS3 - Managerのプロトコルバージョンを取得します。

#### リクエスト形式

```
GET /ajs/api/v1/protocolVersion?query
```

#### パラメータ

| パラメータ | 説明 | 必須/任意 |
|-----------|------|----------|
| manager | マネージャーホスト名またはIPアドレス | 必須 |

#### レスポンス

成功時（200）: `protocolVersionNumber`（プロトコルバージョン番号）をJSON形式で返却

---

## 使用フロー

```
1. statuses API でユニット一覧と execID を取得
       ↓
2. execResultDetails API で実行結果詳細を取得
```

---

## 制限事項

1. **ユニット一覧取得API (7.1.1)**
   - 最大取得件数は1,000件
   - 参照権限がないユニットは取得結果に含まれない

2. **実行結果詳細取得API (7.1.3)**
   - 実行結果詳細を取得（標準エラー出力相当）
   - 最大5MBまで（超過分は切り捨て）
   - 標準出力の取得には ajsshow コマンド（WinRM経由）が必要

3. **認証**
   - JP1ユーザーの権限が必要
   - Web Console経由でManagerに接続

4. **URLエンコード**
   - パス内の特殊文字はURLエンコードが必要
   - `/` → `%2F`, `@` → `%40`

---

## 公式ドキュメント

- [JP1/AJS3 Web Console REST API 一覧](https://itpfdoc.hitachi.co.jp/manuals/3021/30213b1920/AJSO0278.HTM)
- [JP1/AJS3 Web Console REST API 詳細](https://itpfdoc.hitachi.co.jp/manuals/3021/30213b1920/AJSO0280.HTM)

