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
    "顧客グループID",
    "顧客名",
    "最終メッセージ日時",
    "月間コミュニケーション数"
]
HEADER_ALIASES = {
    "担当者名": ["担当者名"],
    "担当者連絡先": ["担当者連絡先"],
    "顧客グループID": ["顧客グループID"],
    "顧客名": ["顧客名"],
    "最終メッセージ日時": ["最終メッセージ日時"],
    "月間コミュニケーション数": ["月間コミュニケーション数", "月間交信数"],
}


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

def one_month_ago_ts(now: Optional[dt.datetime] = None) -> int:
    now = now or dt.datetime.now()
    return int((now - dt.timedelta(days=30)).timestamp())


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

    def count_messages_since(self, room_id: str, since_ts: int) -> int:
        messages = self.list_messages(room_id)
        return sum(1 for m in messages if int(m.get("send_time", 0)) >= since_ts)

    def get_room_message_stats(self, room_id: str, since_ts: int) -> Tuple[Optional[int], int]:
        messages = self.list_messages(room_id)
        if not messages:
            return None, 0
        last_message_ts = int(messages[-1].get("send_time", 0))
        monthly_count = sum(1 for m in messages if int(m.get("send_time", 0)) >= since_ts)
        return last_message_ts, monthly_count

    def get_last_message_time(self, room_id: str) -> Optional[int]:
        messages = self.list_messages(room_id)
        if not messages:
            return None
        return (messages[-1]["send_time"])

    def send_messages(self, room_id: str, text: str) -> bool:
        self._session.headers.update({"Content-Type": "application/x-www-form-urlencoded"})
        response = self._session.post(
            f"{API_BASE}/rooms/{room_id}/messages",
            data={"body": text}
        )
        response.raise_for_status()
        return True
        

class SheetsClient:
    def __init__(self, spreadsheet_id: str, sheet_name: str, credentials_path: str) -> None:
        scopes = ["https://www.googleapis.com/auth/spreadsheets"]
        env = os.getenv("ENV", "local")
        is_cloud_run = bool(os.getenv("K_SERVICE") or os.getenv("CLOUD_RUN_JOB"))
        if is_cloud_run:
            if credentials_path and not os.path.exists(credentials_path):
                os.environ.pop("GOOGLE_APPLICATION_CREDENTIALS", None)
            credentials, _ = google.auth.default(scopes=scopes)
        elif env == "local":
            if not credentials_path:
                raise RuntimeError("ローカル環境ではGOOGLE_APPLICATION_CREDENTIALSが必要です。")
            credentials = Credentials.from_service_account_file(credentials_path, scopes=scopes)
        else:
            if credentials_path:
                credentials = Credentials.from_service_account_file(credentials_path, scopes=scopes)
            else:
                credentials, _ = google.auth.default(scopes=scopes)
        sa_email = getattr(credentials, "service_account_email", None)
        print(
            "[auth] credentials=%s service_account_email=%s"
            % (type(credentials).__name__, sa_email or "unknown")
        )
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
    if value is None:
        return None
    normalized = str(value).strip().replace("Z", "+00:00")
    if not normalized:
        return None
    try:
        return dt.datetime.fromisoformat(normalized)
    except ValueError:
        pass
    for fmt in (
        "%Y/%m/%d %H:%M",
        "%Y/%m/%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d %H:%M:%S",
    ):
        try:
            return dt.datetime.strptime(normalized, fmt)
        except ValueError:
            continue
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


def resolve_header_map(
    header: Sequence[str], required: Sequence[str], aliases: Optional[dict] = None
) -> dict:
    aliases = aliases or HEADER_ALIASES
    raw_map = build_header_map(header)
    resolved = {}
    missing = []
    for name in required:
        candidates = aliases.get(name, [name])
        matched = next((candidate for candidate in candidates if candidate in raw_map), None)
        if matched is None:
            missing.append(name)
            continue
        resolved[name] = raw_map[matched]
    if missing:
        raise RuntimeError(f"必要なヘッダーが不足しています: {', '.join(missing)}")
    return resolved


def load_required_env(name: str) -> str:
    value = os.getenv(name)
    if not value:
        raise RuntimeError(f"環境変数 {name} が設定されていません。")
    return value
