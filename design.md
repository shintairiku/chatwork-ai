# Chatworkリマインド機能 設計書テンプレート

## 1. 仕様概要
- 目的: 対象グループの最終メッセージ日時を取得し、1週間以上未連絡のグループについて担当者へ通知する。
- 対象: サポートデスクが所属する全グループ。
- 出力: スプレッドシートに group_id/顧客名/担当者名/担当者連絡先/最終連絡日時 を一覧化。

## 2. 機能仕様
- グループ取得: GET /rooms で所属グループ一覧を取得。
- 最終メッセージ取得: GET /rooms/{room_id}/messages?force=1 を用いて最終メッセージ日時（send_time）を算出。
- 日次更新: 毎日9時にシート記載の group_id を対象に最終連絡日時を更新。
- 通知条件: 最終連絡日時が「現在時刻 - 7日」以前であれば通知。
- 通知先: シートに登録された担当者連絡先へ通知。

## 3. データ仕様（スプレッドシート）
- 必須列:
  - group_id
  - customer_name
  - assignee_name
  - assignee_contact
  - last_message_at
- 役割:
  - group_id: 取得対象ルーム
  - customer_name: 顧客名
  - assignee_name: 担当者名
  - assignee_contact: 担当者連絡先
  - last_message_at: 最新メッセージ日時（自動更新）

## 4. 構成案
- 処理フロー:
  1) スプレッドシート読み込み（group_id と担当情報）
  2) Chatwork APIで最終メッセージ日時を取得
  3) スプレッドシートを更新
  4) 期限超過判定を行い通知送信

## 5. 次に決めるべき事項
- 通知手段（メール/Chatwork/Slack など）
