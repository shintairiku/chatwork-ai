import argparse
import datetime as dt
import os

from reminder_lib import (
    SheetsClient,
    build_header_map,
    ensure_sheet_header,
    load_env_file,
    load_required_env,
    notify_if_needed,
    parse_iso_datetime,
)


def notify_overdue(sheets: SheetsClient, threshold_days: int) -> None:
    values = sheets.read_values("A1:Z")
    header, rows = ensure_sheet_header(sheets, values)
    header_map = build_header_map(header)
    required = [
        "グループID",
        "customer_name",
        "assignee_name",
        "assignee_email",
        "last_message_at",
    ]
    missing = [name for name in required if name not in header_map]
    if missing:
        raise RuntimeError(f"必要なヘッダーが不足しています: {', '.join(missing)}")

    now = dt.datetime.now(tz=dt.timezone.utc)
    threshold = now - dt.timedelta(days=threshold_days)
    notified = 0

    for row in rows:
        group_id = row[header_map["グループID"]] if len(row) > header_map["グループID"] else ""
        if not group_id:
            continue
        last_message_at = (
            row[header_map["last_message_at"]]
            if len(row) > header_map["last_message_at"]
            else ""
        )
        last_dt = parse_iso_datetime(last_message_at)
        if not last_dt:
            continue
        if last_dt.tzinfo is None:
            last_dt = last_dt.replace(tzinfo=dt.timezone.utc)
        if last_dt > threshold:
            continue
        assignee_email = (
            row[header_map["assignee_email"]]
            if len(row) > header_map["assignee_email"]
            else ""
        )
        if not assignee_email:
            continue
        assignee_name = (
            row[header_map["assignee_name"]]
            if len(row) > header_map["assignee_name"]
            else ""
        )
        customer_name = (
            row[header_map["customer_name"]]
            if len(row) > header_map["customer_name"]
            else ""
        )
        if notify_if_needed(
            assignee_email,
            assignee_name or "担当者",
            customer_name or "顧客",
            group_id,
            last_dt,
        ):
            notified += 1

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

    spreadsheet_id = load_required_env("SPREADSHEET_ID")
    sheet_name = os.getenv("SHEET_NAME", "Reminders")
    credentials_path = load_required_env("GOOGLE_APPLICATION_CREDENTIALS")

    sheets = SheetsClient(spreadsheet_id, sheet_name, credentials_path)
    notify_overdue(sheets, args.threshold_days)


if __name__ == "__main__":
    main()
