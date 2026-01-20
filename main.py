import argparse
import datetime as dt
import os
import smtplib
from email.message import EmailMessage
from typing import List, Optional, Sequence, Tuple

import requests
from google.auth.transport.requests import AuthorizedSession
from google.oauth2.service_account import Credentials

API_BASE = "https://api.chatwork.com/v2"
SHEET_HEADERS = [
    "group_id",
    "customer_name",
    "assignee_name",
    "assignee_email",
    "last_message_at",
]


class ChatworkClient:
    def __init__(self, token: str) -> None:
        self._session = requests.Session()
        self._session.headers.update(
            {"accept": "application/json", "x-chatworktoken": token}
        )

    def list_rooms(self) -> List[dict]:
        response = self._session.get(f"{API_BASE}/rooms")
        response.raise_for_status()
        return response.json()

    def list_messages(self, room_id: str) -> List[dict]:
        response = self._session.get(
            f"{API_BASE}/rooms/{room_id}/messages", params={"force": 1}
        )
        response.raise_for_status()
        return response.json()

    def get_last_message_time(self, room_id: str) -> Optional[int]:
        messages = self.list_messages(room_id)
        if not messages:
            return None
        return max(message.get("send_time", 0) for message in messages)


class SheetsClient:
    def __init__(self, spreadsheet_id: str, sheet_name: str, credentials_path: str) -> None:
        scopes = ["https://www.googleapis.com/auth/spreadsheets"]
        credentials = Credentials.from_service_account_file(credentials_path, scopes=scopes)
        self._session = AuthorizedSession(credentials)
        self._spreadsheet_id = spreadsheet_id
        self._sheet_name = sheet_name

    def _range(self, a1: str) -> str:
        sheet = self._sheet_name
        if " " in sheet:
            sheet = f"'{sheet}'"
        return f"{sheet}!{a1}"

    def read_values(self, a1: str) -> List[List[str]]:
        url = (
            f"https://sheets.googleapis.com/v4/spreadsheets/"
            f"{self._spreadsheet_id}/values/{self._range(a1)}"
        )
        response = self._session.get(url)
        response.raise_for_status()
        return response.json().get("values", [])

    def update_values(self, updates: Sequence[Tuple[int, int, str]]) -> None:
        if not updates:
            return
        data = []
        for row_idx, col_idx, value in updates:
            col = to_column_name(col_idx)
            cell_range = self._range(f"{col}{row_idx}")
            data.append({"range": cell_range, "values": [[value]]})
        url = (
            f"https://sheets.googleapis.com/v4/spreadsheets/"
            f"{self._spreadsheet_id}/values:batchUpdate"
        )
        response = self._session.post(
            url, json={"valueInputOption": "RAW", "data": data}
        )
        response.raise_for_status()

    def append_rows(self, rows: Sequence[Sequence[str]]) -> None:
        if not rows:
            return
        url = (
            f"https://sheets.googleapis.com/v4/spreadsheets/"
            f"{self._spreadsheet_id}/values/{self._range('A1')}:append"
        )
        response = self._session.post(
            url,
            params={"valueInputOption": "RAW", "insertDataOption": "INSERT_ROWS"},
            json={"values": rows},
        )
        response.raise_for_status()


def to_column_name(index: int) -> str:
    name = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def format_iso_utc(ts: int) -> str:
    return dt.datetime.fromtimestamp(ts, tz=dt.timezone.utc).isoformat()


def ensure_sheet_header(
    sheets: SheetsClient, values: List[List[str]]
) -> Tuple[List[str], List[List[str]]]:
    if not values:
        sheets.append_rows([SHEET_HEADERS])
        return SHEET_HEADERS, []
    header = values[0]
    return header, values[1:]


def build_header_map(header: Sequence[str]) -> dict:
    return {name: idx for idx, name in enumerate(header)}


def send_email(
    smtp_host: str,
    smtp_port: int,
    smtp_user: Optional[str],
    smtp_password: Optional[str],
    sender: str,
    recipient: str,
    subject: str,
    body: str,
    use_tls: bool,
) -> None:
    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = recipient
    msg["Subject"] = subject
    msg.set_content(body)
    with smtplib.SMTP(smtp_host, smtp_port) as server:
        if use_tls:
            server.starttls()
        if smtp_user and smtp_password:
            server.login(smtp_user, smtp_password)
        server.send_message(msg)


def notify_if_needed(
    assignee_email: str,
    assignee_name: str,
    customer_name: str,
    room_id: str,
    last_contact: dt.datetime,
) -> bool:
    smtp_host = os.getenv("SMTP_HOST")
    smtp_port = int(os.getenv("SMTP_PORT", "587"))
    smtp_user = os.getenv("SMTP_USER")
    smtp_password = os.getenv("SMTP_PASSWORD")
    sender = os.getenv("SMTP_FROM")
    use_tls = os.getenv("SMTP_USE_TLS", "true").lower() == "true"
    dry_run = os.getenv("NOTIFY_DRY_RUN", "false").lower() == "true"
    if not smtp_host or not sender:
        print("SMTP設定が不足しているため通知をスキップします。")
        return False
    subject = f"[Chatwork] {customer_name} 連絡リマインド"
    body = (
        f"{assignee_name} 様\n\n"
        f"顧客: {customer_name}\n"
        f"グループID: {room_id}\n"
        f"最終連絡日時: {last_contact.isoformat()}\n\n"
        "1週間以上連絡がないため、対応をご確認ください。\n"
    )
    if dry_run:
        print(f"DRY RUN通知: {assignee_email} 宛に送信予定")
        return True
    send_email(
        smtp_host,
        smtp_port,
        smtp_user,
        smtp_password,
        sender,
        assignee_email,
        subject,
        body,
        use_tls,
    )
    return True


def update_sheet_and_notify(
    chatwork: ChatworkClient, sheets: SheetsClient, threshold_days: int
) -> None:
    values = sheets.read_values("A1:Z")
    header, rows = ensure_sheet_header(sheets, values)
    header_map = build_header_map(header)
    missing = [name for name in SHEET_HEADERS if name not in header_map]
    if missing:
        raise RuntimeError(f"必要なヘッダーが不足しています: {', '.join(missing)}")

    updates: List[Tuple[int, int, str]] = []
    now = dt.datetime.now(tz=dt.timezone.utc)
    threshold = now - dt.timedelta(days=threshold_days)

    for idx, row in enumerate(rows, start=2):
        group_id = row[header_map["group_id"]] if len(row) > header_map["group_id"] else ""
        if not group_id:
            continue
        last_message_ts = chatwork.get_last_message_time(group_id)
        if not last_message_ts:
            continue
        last_message_at = format_iso_utc(last_message_ts)
        current = row[header_map["last_message_at"]] if len(row) > header_map["last_message_at"] else ""
        if current != last_message_at:
            updates.append((idx, header_map["last_message_at"] + 1, last_message_at))

        last_dt = dt.datetime.fromtimestamp(last_message_ts, tz=dt.timezone.utc)
        if last_dt <= threshold:
            assignee_email = (
                row[header_map["assignee_email"]]
                if len(row) > header_map["assignee_email"]
                else ""
            )
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
            if assignee_email:
                notify_if_needed(
                    assignee_email,
                    assignee_name or "担当者",
                    customer_name or "顧客",
                    group_id,
                    last_dt,
                )

    sheets.update_values(updates)


def list_rooms(chatwork: ChatworkClient) -> None:
    rooms = chatwork.list_rooms()
    for room in rooms:
        room_id = room.get("room_id")
        name = room.get("name")
        print(f"{room_id}\t{name}")


def sync_rooms(chatwork: ChatworkClient, sheets: SheetsClient) -> None:
    values = sheets.read_values("A1:Z")
    header, rows = ensure_sheet_header(sheets, values)
    header_map = build_header_map(header)
    if "group_id" not in header_map:
        raise RuntimeError("シートに group_id ヘッダーが必要です。")
    existing_ids = set()
    for row in rows:
        if len(row) > header_map["group_id"]:
            existing_ids.add(row[header_map["group_id"]])
    new_rows = []
    for room in chatwork.list_rooms():
        room_id = str(room.get("room_id"))
        if room_id in existing_ids:
            continue
        new_rows.append([room_id, room.get("name", ""), "", "", ""])
    if new_rows:
        sheets.append_rows(new_rows)
        print(f"{len(new_rows)}件のグループを追加しました。")
    else:
        print("追加対象のグループはありません。")


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Chatwork リマインド機能")
    parser.add_argument(
        "--list-rooms", action="store_true", help="所属グループ一覧を表示"
    )
    parser.add_argument(
        "--sync-rooms", action="store_true", help="所属グループをシートへ追加"
    )
    parser.add_argument(
        "--threshold-days", type=int, default=7, help="通知までの経過日数"
    )
    return parser


def load_required_env(name: str) -> str:
    value = os.getenv(name)
    if not value:
        raise RuntimeError(f"環境変数 {name} が設定されていません。")
    return value


def main() -> None:
    parser = build_arg_parser()
    args = parser.parse_args()

    token = load_required_env("CHATWORK_TOKEN")
    spreadsheet_id = load_required_env("SPREADSHEET_ID")
    sheet_name = os.getenv("SHEET_NAME", "Reminders")
    credentials_path = load_required_env("GOOGLE_APPLICATION_CREDENTIALS")

    chatwork = ChatworkClient(token)
    sheets = SheetsClient(spreadsheet_id, sheet_name, credentials_path)

    if args.list_rooms:
        list_rooms(chatwork)
        return
    if args.sync_rooms:
        sync_rooms(chatwork, sheets)
        return
    update_sheet_and_notify(chatwork, sheets, args.threshold_days)


if __name__ == "__main__":
    main()
