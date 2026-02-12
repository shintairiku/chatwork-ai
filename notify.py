import argparse
import datetime as dt
import os

from reminder_lib import (
    ChatworkClient,
    SheetsClient,
    build_header_map,
    ensure_sheet_header,
    load_env_file,
    load_required_env,
    parse_iso_datetime,
)

def notify_overdue(chatwork: ChatworkClient, sheets: SheetsClient, threshold_days: int) -> None:
    values = sheets.read_values("A1:Z")
    header, rows = ensure_sheet_header(sheets, values)
    header_map = build_header_map(header)
    required = [
        "顧客グループID",
        "顧客名",
        "担当者名",
        "担当者連絡先",
        "最終メッセージ日時",
    ]
    missing = [name for name in required if name not in header_map]
    if missing:
        raise RuntimeError(f"必要なヘッダーが不足しています: {', '.join(missing)}")

    jst = dt.timezone(dt.timedelta(hours=9))
    # print(f"jst: {jst}")
    now = dt.datetime.now(tz=jst)
    threshold = now - dt.timedelta(days=threshold_days)
    # print(f"threshold: {threshold}")
    notified = 0

    for row in rows:
        last_message_at = (
            row[header_map["最終メッセージ日時"]]
            if len(row) > header_map["最終メッセージ日時"]
            else ""
        )
        last_dt = parse_iso_datetime(last_message_at)
        # print(last_dt)
        if not last_dt:
            continue
        if last_dt.tzinfo is None:
            last_dt = last_dt.replace(tzinfo=jst)
        if last_dt > threshold:
            continue
        assignee_id = (
            row[header_map["担当者連絡先"]]
            if len(row) > header_map["担当者連絡先"]
            else ""
        )
        if not assignee_id:
            continue
        assignee_name = (
            row[header_map["担当者名"]]
            if len(row) > header_map["担当者名"]
            else ""
        )
        customer_name = (
            row[header_map["顧客名"]]
            if len(row) > header_map["顧客名"]
            else ""
        )
        body = (
        f"※通知テスト用メッセージ: 期限超過判定を1日に設定しています。\n\n"
        f"【連絡リマインド】\n\n"
        f"担当者: {assignee_name}様\n"
        f"顧客: {customer_name}様\n"
        f"最終連絡日時: {last_dt.strftime('%Y-%m-%d %H:%M:%S')}\n\n"
        "1週間以上連絡がないため、対応をご確認ください。\n"
        )
        if chatwork.send_messages(assignee_id, body):
            notified += 1
            print(f"通知送信\n顧客名:{customer_name}様\n担当者:{assignee_name}")
    print(f"{notified}件の通知を実行しました。")


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Chatwork通知送信")
    parser.add_argument(
        "--threshold-days", type=int, default=7, help="通知までの経過日数"
    )
    return parser


def main() -> None:
    load_env_file()
    args = build_arg_parser().parse_args()

    token = load_required_env("CHATWORK_TOKEN")
    spreadsheet_id = load_required_env("SPREADSHEET_ID")
    sheet_name = os.getenv("SHEET_NAME", "Reminders")
    env = os.getenv("ENV", "local")
    is_cloud_run = bool(os.getenv("K_SERVICE") or os.getenv("CLOUD_RUN_JOB"))
    if is_cloud_run:
        credentials_path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS", "")
    elif env == "local":
        credentials_path = load_required_env("GOOGLE_APPLICATION_CREDENTIALS")
    else:
        credentials_path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS", "")

    sheets = SheetsClient(spreadsheet_id, sheet_name, credentials_path)
    chatwork = ChatworkClient(token)
    notify_overdue(chatwork, sheets, args.threshold_days)


if __name__ == "__main__":
    main()
