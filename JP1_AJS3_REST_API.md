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

## API一覧

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

## 使用フロー

```
1. statuses API でユニット一覧と execID を取得
       ↓
2. execResultDetails API で実行結果詳細を取得
```

---

## 制限事項

1. **ユニット一覧取得API (7.1.1) は実行登録中のジョブのみ対象**
   - 即時実行で終了済みのジョブは取得できない
   - 計画実行で登録中のジョブは取得可能
   - 最大取得件数は1,000件

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

- [JP1/AJS3 Web Console REST API](https://itpfdoc.hitachi.co.jp/manuals/3021/30213b1920/AJSO0280.HTM)

---

## 追加のAPIドキュメント

以下にAPIドキュメントを追加してください：

<!-- ここにAPIドキュメントを貼り付け -->

