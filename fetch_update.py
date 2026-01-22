import argparse
import os
from typing import List, Tuple
from dotenv import load_dotenv

from reminder_lib import (
    ChatworkClient,
    SheetsClient,
    build_header_map,
    ensure_sheet_header,
    format_iso_utc,
    load_env_file,
    load_required_env,
)


def resolve_target_group_id() -> str:
    return os.getenv("TEST_GROUP_ID") or os.getenv("ID", "")


def update_sheet_last_message(
    chatwork: ChatworkClient, sheets: SheetsClient, target_group_id: str
) -> None:
    values = sheets.read_values("A1:Z")
    header, rows = ensure_sheet_header(sheets, values)
    header_map = build_header_map(header)
    missing = [name for name in ("グループID", "最終メッセージ日時") if name not in header_map]
    if missing:
        raise RuntimeError(f"必要なヘッダーが不足しています: {', '.join(missing)}")

    updates: List[Tuple[int, int, str]] = []
    target_found = False
    for idx, row in enumerate(rows, start=2):
        group_id = row[header_map["グループID"]] if len(row) > header_map["グループID"] else ""
        if not group_id:
            continue
        if target_group_id and group_id != target_group_id:
            continue
        if target_group_id and group_id == target_group_id:
            target_found = True
        last_message_ts = chatwork.get_last_message_time(group_id)
        if not last_message_ts:
            continue
        last_message_at = format_iso_utc(last_message_ts)
        current = (
            row[header_map["last_message_at"]]
            if len(row) > header_map["last_message_at"]
            else ""
        )
        if current != last_message_at:
            updates.append((idx, header_map["last_message_at"] + 1, last_message_at))

    if target_group_id and not target_found:
        print("テスト用グループIDがシートに見つかりませんでした。")

    sheets.update_values(updates)
    print(f"{len(updates)}件の最終連絡日時を更新しました。")


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Chatworkメッセージ取得とスプレッドシート更新"
    )
    return parser


def main() -> None:
    load_env_file()
    load_dotenv(verbose=True)
    build_arg_parser().parse_args()

    token = load_required_env("CHATWORK_TOKEN")
    spreadsheet_id = load_required_env("SPREADSHEET_ID")
    sheet_name = os.getenv("SHEET_NAME", "Reminders")
    credentials_path = load_required_env("GOOGLE_APPLICATION_CREDENTIALS")

    # target_group_id = resolve_target_group_id()
    # if target_group_id:
    #     print(f"テスト用グループIDを使用: {target_group_id}")

    chatwork = ChatworkClient(token)
    sheets = SheetsClient(spreadsheet_id, sheet_name, credentials_path)
    update_sheet_last_message(chatwork, sheets, target_group_id)


if __name__ == "__main__":
    main()
