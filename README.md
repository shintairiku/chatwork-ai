Chatwork API と Google スプレッドシートを使ったリマインド機能です。

## 目的
- Chatwork のグループ最終メッセージ日時を取得して一覧化
- 最終連絡から一定日数が経過している場合、担当者へ通知（SMTPメール）

## スプレッドシートの形式
シート名は `Reminders` を既定として使用します。1行目は以下のヘッダーを設定してください。

```
group_id,customer_name,assignee_name,assignee_email,last_message_at
```

## 必要な環境変数
- `CHATWORK_TOKEN`: Chatwork API トークン
- `SPREADSHEET_ID`: 対象スプレッドシートID
- `SHEET_NAME`: シート名（省略時 `Reminders`）
- `GOOGLE_APPLICATION_CREDENTIALS`: サービスアカウントJSONのパス

通知メール用（SMTP）
- `SMTP_HOST`
- `SMTP_PORT`（省略時 `587`）
- `SMTP_USER`（任意）
- `SMTP_PASSWORD`（任意）
- `SMTP_FROM`
- `SMTP_USE_TLS`（省略時 `true`）
- `NOTIFY_DRY_RUN`（`true` で送信せずログのみ）

## 使い方
### 1) メッセージ取得〜スプレッドシート更新
```
python fetch_update.py
```

### 2) 通知送信
```
python notify.py --threshold-days 7
```

### 3) 契約解除の予兆予測
```
python churn_predict.py --room-id 123456 --limit 200
```

#### 予測モデル切り替え
環境変数 `LLM_HANDLER` で切り替えます。
- `openai`（既定、`OPENAI_API_KEY` がある場合）
- `rule`（ルールベースの簡易判定）

OpenAI 互換APIを使う場合の環境変数:
- `OPENAI_API_KEY`
- `OPENAI_MODEL`（省略時 `gpt-4o-mini`）
- `OPENAI_BASE_URL`（省略時 `https://api.openai.com`）
- `LLM_TEMPERATURE`（省略時 `0.2`）
- `LLM_TIMEOUT`（秒、省略時 `30`）

### テスト用グループID
`.env` に `TEST_GROUP_ID` または `ID` を記載すると、メッセージ取得はそのグループのみを対象にします。

## 参考: Cloud Scheduler で定期実行
Cloud Run / Cloud Functions などにデプロイし、毎日9時に HTTP 実行する構成を想定しています。
