import datetime as dt
import os
import smtplib
from email.message import EmailMessage
from typing import List, Optional, Sequence, Tuple

import requests
import google.auth
from google.auth.transport.requests import AuthorizedSession
from google.oauth2.service_account import Credentials

API_BASE = "https://api.chatwork.com/v2"
SHEET_HEADERS = [
    "担当者名",
    "担当者連絡先",
    "グループID",
    "顧客名",
    "最終メッセージ日時",
]


def load_env_file(path: str = ".env") -> None:
    if not os.path.exists(path):
        return
    with open(path, "r", encoding="utf-8") as file:
        for line in file:
            stripped = line.strip()
            if not stripped or stripped.startswith("#"):
                continue
            if "=" not in stripped:
                continue
            key, value = stripped.split("=", 1)
            key = key.strip()
            value = value.strip().strip('"').strip("'")
            os.environ.setdefault(key, value)


class ChatworkClient:
    def __init__(self, token: str) -> None:
        self._session = requests.Session()
        self._session.headers.update(
            {"accept": "application/json", "x-chatworktoken": token}
        )

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
        return (messages[-1]["send_time"])


class SheetsClient:
    def __init__(self, spreadsheet_id: str, sheet_name: str, credentials_path: str) -> None:
        scopes = ["https://www.googleapis.com/auth/spreadsheets"]
        env = os.getenv("ENV", "local")
        if env == "local":
            credentials = Credentials.from_service_account_file(credentials_path, scopes=scopes)
        else:
            credentials, _ = google.auth.default(scopes=scopes)
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


def parse_iso_datetime(value: str) -> Optional[dt.datetime]:
    if not value:
        return None
    normalized = value.replace("Z", "+00:00")
    try:
        return dt.datetime.fromisoformat(normalized)
    except ValueError:
        return None


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


def load_required_env(name: str) -> str:
    value = os.getenv(name)
    if not value:
        raise RuntimeError(f"環境変数 {name} が設定されていません。")
    return value


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
