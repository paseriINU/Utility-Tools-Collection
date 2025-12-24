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

### 1. ユニット状態取得 (statuses)

実行登録中のユニットの状態を取得します。

**エンドポイント:**
```
GET /ajs/api/v1/objects/statuses?manager={manager}&serviceName={serviceName}&location={path}&mode=search
```

**レスポンス例:**
```json
{
  "statuses": [
    {
      "path": "/jobnet1",
      "execId": "@A100",
      "status": "running"
    }
  ],
  "all": true
}
```

**注意:**
- 実行登録中のジョブのみ対象
- 即時実行で終了済みのジョブは取得できない

---

### 2. ユニット定義取得 (definitions)

ユニットの定義情報を取得します。

**エンドポイント:**
```
GET /ajs/api/v1/objects/definitions?manager={manager}&serviceName={serviceName}&location={path}
```

---

### 3. 実行結果詳細取得 (execResultDetails)

ジョブの実行結果詳細（標準エラー出力）を取得します。

**エンドポイント:**
```
GET /ajs/api/v1/objects/statuses/{path}:{execId}/actions/execResultDetails/invoke?manager={manager}&serviceName={serviceName}
```

**パラメータ:**
- `{path}`: ユニットパス（例: /jobnet1/job1）
- `{execId}`: 実行ID（例: @A100）

**レスポンス例:**
```json
{
  "execResultDetails": "エラーメッセージ内容...",
  "all": true
}
```

**注意:**
- **標準エラー出力のみ**取得可能（標準出力ではない）
- 実行IDはstatuses APIで事前に取得する必要がある

---

## 使用フロー

```
1. statuses API でユニット一覧と execID を取得
       ↓
2. execResultDetails API で実行結果詳細を取得
```

---

## 制限事項

1. **statuses API は実行登録中のジョブのみ対象**
   - 即時実行で終了済みのジョブは取得できない
   - 計画実行で登録中のジョブは取得可能

2. **execResultDetails API は標準エラー出力のみ**
   - 標準出力の取得には ajsshow コマンド（WinRM経由）が必要

3. **認証**
   - JP1ユーザーの権限が必要
   - Web Console経由でManagerに接続

---

## 公式ドキュメント

- [JP1/AJS3 Web Console REST API](https://itpfdoc.hitachi.co.jp/manuals/3021/30213b1920/AJSO0280.HTM)

---

## 追加のAPIドキュメント

以下にAPIドキュメントを追加してください：

<!-- ここにAPIドキュメントを貼り付け -->

