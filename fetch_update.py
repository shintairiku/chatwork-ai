import argparse
import os
from typing import List, Tuple
import datetime as dt

from reminder_lib import (
    ChatworkClient,
    SheetsClient,
    ensure_sheet_header,
    load_env_file,
    one_month_ago_ts,
    load_required_env,
    resolve_header_map,
)


def resolve_target_group_id() -> str:
    return os.getenv("TEST_GROUP_ID") or os.getenv("ID", "")


def update_sheet_last_message(
    chatwork: ChatworkClient, sheets: SheetsClient
) -> None:
    values = sheets.read_values("A1:Z")
    _, header, rows, data_start_row = ensure_sheet_header(sheets, values)
    header_map = resolve_header_map(
        header, ("customer_group_id", "last_message_at", "monthly_message_count")
    )

    updates_last_massage_ts: List[Tuple[int, int, str]] = []
    updates_monthly_message_count: List[Tuple[int, int, str]] = []
    since_ts = one_month_ago_ts()
    for idx, row in enumerate(rows, start=data_start_row):
        group_id = row[header_map["customer_group_id"]] if len(row) > header_map["customer_group_id"] else ""
        if not group_id:
            continue
        last_message_ts, count = chatwork.get_room_message_stats(group_id, since_ts)
        if not last_message_ts and count == 0:
            continue
        updates_monthly_message_count.append((idx, header_map["monthly_message_count"] + 1, str(count)))
        if last_message_ts:
            last_message_at = dt.datetime.fromtimestamp(last_message_ts).strftime("%Y/%m/%d %H:%M")
            print(last_message_at)
            current = (
                row[header_map["last_message_at"]]
                if len(row) > header_map["last_message_at"]
                else ""
            )
            if current != last_message_at:
                updates_last_massage_ts.append((idx, header_map["last_message_at"] + 1, last_message_at))

    sheets.update_values(updates_last_massage_ts)
    print(f"{len(updates_last_massage_ts)}件の最終連絡日時を更新しました。")
    sheets.update_values(updates_monthly_message_count)
    print(f"{len(updates_monthly_message_count)}件の月間コミュニケーション数を更新しました。")


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Chatworkメッセージ取得とスプレッドシート更新"
    )
    return parser


def main() -> None:
    load_env_file()
    build_arg_parser().parse_args()

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

    # target_group_id = resolve_target_group_id()
    # if target_group_id:
    #     print(f"テスト用グループIDを使用: {target_group_id}")

    chatwork = ChatworkClient(token)
    sheets = SheetsClient(spreadsheet_id, sheet_name, credentials_path)
    update_sheet_last_message(chatwork, sheets)


if __name__ == "__main__":
    main()
