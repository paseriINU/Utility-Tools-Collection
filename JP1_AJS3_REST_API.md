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

## 7.4 定数の詳細

JP1/AJS3 Web Console REST APIで使用する定数の詳細です。

### 7.4.1 API共通で使用する定数

#### DateType

期間を指定する場合に、その期間を暦上の日付で指定するか、JP1/AJS3の実行日で指定するかを示します。

| 定数 | 説明 |
|------|------|
| CALENDAR_DATE | 期間を暦上の日付で指定します |
| EXECUTION_DATE | 期間を実行日で指定します |

---

### 7.4.2 ユニット一覧の取得APIで使用する定数

### (1) LowerType

ユニット配下のジョブを取得対象とするかを示します。

| 定数 | 説明 |
|------|------|
| YES | ユニットの配下をすべて取得対象にします |
| NO | ユニットの直下1階層にあるユニットだけを取得対象にします |

### (2) SearchTargetType

取得する情報の範囲を示します。

| 定数 | 説明 |
|------|------|
| DEFINITION | ユニットの定義だけを取得対象にします |
| DEFINITION_AND_STATUS | ユニットの実行状態や世代数などを考慮して、ユニットの定義および状態を取得対象にします |

### (3) MatchMethods

文字列の比較方法を示します。

| 定数 | 説明 |
|------|------|
| NO | この比較方法を検索条件にしません |
| EQ | 検索条件の値と完全に一致する（完全一致） |
| BW | 検索条件の値で始まる（前方一致） |
| EW | 検索条件の値で終わる（後方一致） |
| NE | 検索条件の値と一致しない（不一致） |
| CO | 検索条件の値が含まれる |
| NC | 検索条件の値が含まれない |
| RE | 検索条件の値を正規表現として扱います |

#### 正規表現で指定できる記号

| 記号 | 説明 | 指定例 | 取得例 |
|------|------|--------|--------|
| ? | 任意の1文字 | A? | AB, A1, A?, A*, A¥ |
| * | 任意の文字列 | B* | B, BC, B12, B?*¥ |
| \ | 直後の記号を文字として扱う | C\? | C? |

### (4) UnitType

情報を取得するユニットのユニット種別を示します。

| 定数 | 説明 |
|------|------|
| NO | ユニット種別を検索条件にしません |
| GROUP | ジョブグループの情報だけを取得（ジョブグループ、プランニンググループ、マネージャージョブグループ） |
| ROOT | ルートジョブネットの情報だけを取得（ルートジョブネット、ルートリモートジョブネット、ルートマネージャージョブネット） |
| NET | ジョブネットの情報だけを取得（ルート/ネストジョブネット、リモートジョブネット、マネージャージョブネット） |
| JOB | ジョブの情報だけを取得（標準ジョブ、イベントジョブ、アクションジョブ、カスタムジョブ、引き継ぎ情報設定ジョブ、HTTP接続ジョブ、フレキシブルジョブ） |

### (5) GenerationType

情報を取得するユニットの世代を示します。

| 定数 | 説明 |
|------|------|
| NO | 世代を検索条件にしません |
| STATUS | 最新状態の世代を取得します（VIEWSTATUSRANGEの設定値に従う） |
| RESULT | 最新結果の世代を取得します |
| PERIOD | 指定した期間に存在する世代を取得します |
| EXECID | 指定した実行IDの世代を取得します |

### (6) UnitStatus

情報を取得するユニットのユニット状態を示します。

#### 個別状態

| 定数 | 説明 |
|------|------|
| NO | ユニット状態を検索条件にしません |
| UNREGISTERED | 未登録 |
| NOPLAN | 未計画 |
| UNEXEC | 未実行終了 |
| BYPASS | 計画未実行 |
| EXECDEFFER | 繰越未実行 |
| SHUTDOWN | 閉塞 |
| TIMEWAIT | 開始時刻待ち |
| TERMWAIT | 先行終了待ち |
| EXECWAIT | 実行待ち |
| QUEUING | キューイング |
| CONDITIONWAIT | 起動条件待ち |
| HOLDING | 保留中 |
| RUNNING | 実行中 |
| WACONT | 警告検出実行中 |
| ABCONT | 異常検出実行中 |
| MONITORING | 監視中 |
| ABNORMAL | 異常検出終了 |
| INVALIDSEQ | 順序不正 |
| INTERRUPT | 中断 |
| KILL | 強制終了 |
| FAIL | 起動失敗 |
| UNKNOWN | 終了状態不正 |
| MONITORCLOSE | 監視打ち切り終了 |
| WARNING | 警告検出終了 |
| NORMAL | 正常終了 |
| NORMALFALSE | 正常終了-偽 |
| UNEXECMONITOR | 監視未起動終了 |
| MONITORINTRPT | 監視中断 |
| MONITORNORMAL | 監視正常終了 |

#### グループ状態

| 定数 | 説明 |
|------|------|
| GRP_WAIT | 開始時刻待ち、先行終了待ち、実行待ち、キューイング、起動条件待ち |
| GRP_RUN | 実行中、警告検出実行中、異常検出実行中、監視中 |
| GRP_ABNORMAL | 異常検出終了、順序不正、中断、強制終了、起動失敗、終了状態不明、監視打ち切り終了 |
| GRP_NORMAL | 正常終了、正常終了-偽、監視未起動終了、監視中断、監視正常終了 |

### (7) DelayType

情報を取得するユニットの、開始遅延または終了遅延の有無を示します。

| 定数 | 説明 |
|------|------|
| NO | 開始遅延または終了遅延の有無を検索条件にしません |
| START | 開始遅延のあるユニットの情報を取得します |
| END | 終了遅延のあるユニットの情報を取得します |
| YES | 開始遅延、または終了遅延のあるユニットの情報を取得します |

### (8) HoldPlan

情報を取得するユニットの、保留予定の有無を示します。

| 定数 | 説明 |
|------|------|
| NO | 保留予定の有無を検索条件にしません |
| PLAN_NONE | 保留予定のないユニットの情報を取得します |
| PLAN_YES | 保留予定のあるユニットの情報を取得します |

### (9) ReleaseInfoSearchMethods

情報を取得するユニットのリリースIDを示します。

| 定数 | 説明 |
|------|------|
| NO | リリースIDを検索条件にしません |
| ID | リリースIDを検索条件にします |

---

### 7.4.3 実行登録APIで使用する定数

#### PlannedTimePassedType

スケジューラーサービス起動時および実行登録時に、予定時刻超過時の動作を制御します。

| 定数 | 説明 |
|------|------|
| IMMEDIATE | 予定時刻を超過していた場合、ジョブネットをすぐに実行します |
| NEXTTIME | 予定時刻を超過していた場合、ジョブネットを実行しません。次回から実行します |

> **注意**: 日時指定の確定実行登録でNEXTTIMEを指定した場合、超過したジョブネットは実行されません。

---

### 7.4.4 ユニット定義情報オブジェクトで使用する定数

#### (1) Type

ユニットの種別を示します。

| 定数 | 説明 |
|------|------|
| GROUP | ジョブグループ |
| PLANGROUP | プランニンググループ |
| MGRGROUP | マネージャージョブグループ |
| ROOTNET | ルートジョブネット |
| ROOTRMTNET | ルートリモートジョブネット |
| ROOTMGRNET | ルートマネージャージョブネット |
| NET | ネストジョブネット |
| RMTNET | リモートジョブネット |
| MGRNET | マネージャージョブネット |
| JOB | UNIXジョブ |
| PJOB | PCジョブ |
| QJOB | QUEUEジョブ |
| JDJOB | 判定ジョブ |
| ORJOB | ORジョブ |
| EVWJB | ファイル監視ジョブ |
| FLWJB | ファイル監視ジョブ（拡張） |
| MLWJB | メール受信監視ジョブ |
| MSWJB | メッセージキュー受信監視ジョブ |
| LFWJB | ログファイル監視ジョブ |
| TMWJB | 時間監視ジョブ |
| EVSJB | イベント送信ジョブ |
| MLSJB | メール送信ジョブ |
| MSSJB | メッセージキュー送信ジョブ |
| PWLJB | JP1イベント受信監視ジョブ |
| PWRJB | 電源制御ジョブ |
| CJOB | カスタムジョブ |
| HTPJOB | HTTP接続ジョブ |
| CPJOB | 引き継ぎ情報設定ジョブ |
| FXJOB | フレキシブルジョブ |
| CUSTOM | カスタムジョブ（汎用） |

#### (2) RegisterStatus

ユニットの登録状態を示します。

| 定数 | 説明 |
|------|------|
| YES | 登録済みのユニットを示します |
| NO | 未登録のユニットを示します |
| NONE | 検索条件に該当するユニットがないことを示します |
| UNSUPPORTED | 未サポートを示します |

---

### 7.4.5 ステータス情報オブジェクトで使用する定数

#### (1) Status

ユニットの状態を示します。

| 定数 | 説明 |
|------|------|
| NONE | 検索条件に該当するユニットがないことを示します |
| UNSUPPORTED | 未サポートを示します |
| NOPLAN | 未計画 |
| UNEXEC | 未実行終了 |
| BYPASS | 計画未実行 |
| EXECDEFFER | 繰越未実行 |
| SHUTDOWN | 閉塞 |
| TIMEWAIT | 開始時刻待ち |
| TERMWAIT | 先行終了待ち |
| EXECWAIT | 実行待ち |
| QUEUING | キューイング |
| CONDITIONWAIT | 起動条件待ち |
| HOLDING | 保留中 |
| RUNNING | 実行中 |
| WACONT | 警告検出実行中 |
| ABCONT | 異常検出実行中 |
| MONITORING | 監視中 |
| ABNORMAL | 異常検出終了 |
| INVALIDSEQ | 順序不正 |
| INTERRUPT | 中断 |
| KILL | 強制終了 |
| FAIL | 起動失敗 |
| UNKNOWN | 終了状態不正 |
| MONITORCLOSE | 監視打ち切り終了 |
| WARNING | 警告検出終了 |
| NORMAL | 正常終了 |
| NORMALFALSE | 正常終了-偽 |
| UNEXECMONITOR | 監視未起動終了 |
| MONITORINTRPT | 監視中断 |
| MONITORNORMAL | 監視正常終了 |

#### (2) DelayStart / DelayEnd

開始遅延・終了遅延の有無を示します。

| 定数 | 説明 |
|------|------|
| NONE | 検索条件に該当するユニットがないことを示します |
| UNSUPPORTED | 未サポートを示します |
| YES | 遅延あり |
| NO | 遅延なし |

#### (3) ChangeType

計画一時変更の状態を示します。

| 定数 | 説明 |
|------|------|
| NONE | 検索条件に該当するユニットがないことを示します |
| UNSUPPORTED | 未サポートを示します |
| NO | 一時変更なし |
| TIME | 日時変更中 |
| CANCEL | 実行中止中 |
| MOVEMENT | 移動中 |

#### (4) HoldAttr

保留属性を示します。

| 定数 | 説明 |
|------|------|
| NONE | 検索条件に該当するユニットがないことを示します |
| UNSUPPORTED | 未サポートを示します |
| NO | 保留なし |
| YES | 保留中 |
| YES_ERR | 異常終了時に保留 |
| YES_WAR | 警告終了時に保留 |
| HOLD | 保留設定 |
| RELEASE | 保留解除 |

#### (5) TimeType

起動条件の時間種別を示します。

| 定数 | 説明 |
|------|------|
| NONE | 検索条件に該当するユニットがないことを示します |
| UNSUPPORTED | 未サポートを示します |
| NO | 時間種別なし |
| ABSOLUTE | 絶対時間 |
| RELATIVE | 相対時間 |
| UNLIMITED | 無制限 |

#### (6) DelayMonitor

ジョブネットの遅延監視方法を示します。

| 定数 | 説明 |
|------|------|
| NONE | 検索条件に該当するユニットがないことを示します |
| UNSUPPORTED | 未サポートを示します |
| NOT | 監視しない |
| ABSOLUTE | 絶対時刻での遅延監視 |
| ROOT | ルートジョブネットの開始予定時刻からの相対時刻での遅延監視 |
| TOP | 上位ジョブネットの開始予定時刻からの相対時刻での遅延監視 |
| OWN | 自ジョブネットの開始予定時刻からの相対時刻での遅延監視 |

#### (7) MonitoringJobnet

実行所要時間による監視方法を示します。

| 定数 | 説明 |
|------|------|
| NONE | 検索条件に該当するユニットがないことを示します |
| UNSUPPORTED | 未サポートを示します |
| TIME_REQUIRED_FOR_EXECUTION_NO | 実行所要時間による監視をしない |
| TIME_REQUIRED_FOR_EXECUTION_YES | 実行所要時間による監視をする |

---

### 7.4.6 リリース情報オブジェクトで使用する定数

#### ReleaseStatus

リリースの状態を示します。

| 定数 | 説明 |
|------|------|
| NONE | 検索条件に該当するユニットがないことを示します |
| UNSUPPORTED | 未サポートを示します |
| RELEASE_WAIT | リリース待ち |
| BEING_APPLIED | 適用中 |
| APPLIED | 適用終了 |
| DELETE_WAIT | 削除待ち |
| RELEASE_ENTRY_WAIT | リリース登録待ち |

---

### 7.4.7 起動条件パラメーターオブジェクトで使用する定数

#### TimeType

起動条件の有効範囲における時間種別を示します。

| 定数 | 説明 |
|------|------|
| ABSOLUTE | 絶対時間を示します |
| RELATIVE | 相対時間を示します |
| NO | 無制限を示します |

---

### 7.4.8 保留属性変更APIで使用する定数

#### ChangeHoldAttr

保留属性の変更方法を示します。

| 定数 | 説明 |
|------|------|
| SET | 保留属性設定を示します |
| RELEASE | 保留解除を示します |

---

### 7.4.9 計画一時変更APIで使用する定数

#### (1) ChangeStartTimeType

実行開始日時の種類を示します。

| 定数 | 説明 |
|------|------|
| ABSOLUTETIME | 絶対日時を指定します |
| RELATIVETIME | 相対日時を指定します |

#### (2) ChangePushedAheadType

開始時間前倒し時の動作を示します。

| 定数 | 説明 |
|------|------|
| ADD | 次回予定追加を示します |
| MOVE | 次回予定移動を示します |

#### (3) ChangeLowerType

配下ジョブネットの開始日時変更動作を示します。

| 定数 | 説明 |
|------|------|
| NOTSHIFT | 配下のジョブネットの開始日時は変更しません。指定したジョブネットだけ、開始日時を変更します |
| SHIFT | 配下のジョブネットの開始日時も連動して変更します。指定したジョブネットの配下にあるすべてのネストジョブネットの開始日時を、相対的に変更します |

---

### 7.4.10 再実行APIで使用する定数

#### (1) RootRerunType

ルートジョブネットの再実行方法を示します。

| 定数 | 説明 |
|------|------|
| TOP | 指定したルートジョブネットの先頭のジョブから再実行します |
| ABNORMAL_JOB | 配下ジョブ中の異常終了したジョブから再実行します |
| ABNORMAL_NEXT | 異常終了したジョブの次のジョブから再実行します |
| ABNORMAL_JOBNET | 配下のネストジョブネット中の異常終了ジョブネットから再実行します |
| WARNING | 警告終了したジョブだけ再実行します |

#### (2) RerunType

ネストジョブネット/ジョブの再実行方法を示します。

| 定数 | 説明 |
|------|------|
| FROM | 指定したユニットを再実行します。再実行が終了したら、後続ユニットの処理が続行されます |
| ONLY | 指定したユニットのみを再実行します |
| NEXT | 指定したユニットの次のユニットから再実行します |

#### (3) RerunOption

再実行オプションを示します。

| 定数 | 説明 | 適用条件 |
|------|------|----------|
| HOLD | 再実行するユニットを保留状態にします | FROM/ONLY指定時のみ有効 |
| WARNING | 異常状態の先行ユニットがある場合に、そのユニットの終了状態を警告終了にします | RootRerunTypeがABNORMAL_NEXTの場合のみ有効 |

---

### 7.4.11 遅延監視変更APIで使用する定数

#### (1) DelayMonitorType

ジョブネットの遅延監視方法を示します。

| 定数 | 説明 | 備考 |
|------|------|------|
| NOT | 監視しないことを示します | - |
| ABSOLUTE | 絶対時刻での遅延監視を示します | - |
| ROOT | ルートジョブネットの開始予定時刻からの相対時刻での遅延監視 | ルートジョブネット自体では指定不可 |
| TOP | 上位ジョブネットの開始予定時刻からの相対時刻での遅延監視 | ルートジョブネット自体では指定不可 |
| OWN | 自ジョブネットの開始予定時刻からの相対時刻での遅延監視 | - |

#### (2) MonitoringJobnetType

ジョブネットの実行所要時間に基づく監視方法を示します。

| 定数 | 説明 |
|------|------|
| NOT | 監視しないことを示します |
| TIME_REQUIRED_FOR_EXECUTION | ジョブネットの実行所要時間に基づく遅延監視を実施します |

---

### 7.4.12 ジョブ状態変更APIで使用する定数

#### ChangeStatus

ジョブの変更後の状態を示します。

| 定数 | 説明 |
|------|------|
| NORMAL | ジョブの状態を正常終了にします |
| FAIL | ジョブの状態を起動失敗にします |
| WARNING | ジョブの状態を警告検出終了にします |
| ABNORMAL | ジョブの状態を異常検出終了にします |
| BYPASS | ジョブの状態を計画未実行にします |
| RETURNCODE | 終了コードで判定し、指定した終了コードに変更します |

---

## 公式ドキュメント

- [JP1/AJS3 Web Console REST API 一覧](https://itpfdoc.hitachi.co.jp/manuals/3021/30213b1920/AJSO0278.HTM)
- [JP1/AJS3 Web Console REST API 詳細](https://itpfdoc.hitachi.co.jp/manuals/3021/30213b1920/AJSO0280.HTM)
- [JP1/AJS3 Web Console REST API 定数の詳細](https://itpfdoc.hitachi.co.jp/manuals/3021/30213b1920/AJSO0306.HTM)

