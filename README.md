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
### 1) グループ一覧の取得
```
python main.py --list-rooms
```

### 2) グループIDをシートに追加
```
python main.py --sync-rooms
```

### 3) 日次実行（最終連絡の更新 + 期限超過通知）
```
python main.py --threshold-days 7
```

## 参考: Cloud Scheduler で定期実行
Cloud Run / Cloud Functions などにデプロイし、毎日9時に HTTP 実行する構成を想定しています。
