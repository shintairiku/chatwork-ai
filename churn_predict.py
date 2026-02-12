import argparse
import datetime as dt
import json
import os
from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional

from reminder_lib import ChatworkClient, load_env_file, load_required_env

DEFAULT_OPENAI_BASE_URL = "https://api.openai.com"
DEFAULT_OPENAI_MODEL = "gpt-4o-mini"


@dataclass
class PredictionResult:
    risk_level: str
    score: float
    signals: List[str]
    summary: str
    model: str


class LLMHandler:
    name = "base"

    def predict(self, messages: List[dict]) -> PredictionResult:
        raise NotImplementedError


class RuleBasedHandler(LLMHandler):
    name = "rule"

    def predict(self, messages: List[dict]) -> PredictionResult:
        keywords = [
            "解約",
            "解除",
            "停止",
            "終了",
            "更新しない",
            "乗り換え",
            "不満",
            "高い",
            "高すぎ",
            "遅い",
            "解決しない",
        ]
        text = "\n".join(msg.get("body", "") or "" for msg in messages).lower()
        hits = [kw for kw in keywords if kw in text]
        score = min(0.1 + 0.15 * len(hits), 0.95)
        level = "high" if score >= 0.7 else "medium" if score >= 0.4 else "low"
        summary = "キーワード検出に基づく暫定判定です。"
        return PredictionResult(
            risk_level=level,
            score=round(score, 2),
            signals=hits or ["明確な解約ワードは未検出"],
            summary=summary,
            model="rule-based",
        )


class OpenAIChatHandler(LLMHandler):
    name = "openai"

    def __init__(
        self,
        api_key: str,
        model: str = DEFAULT_OPENAI_MODEL,
        base_url: str = DEFAULT_OPENAI_BASE_URL,
        temperature: float = 0.2,
        timeout: int = 30,
    ) -> None:
        self._api_key = api_key
        self._model = model
        self._base_url = base_url.rstrip("/")
        self._temperature = temperature
        self._timeout = timeout

    def predict(self, messages: List[dict]) -> PredictionResult:
        prompt = build_prompt(messages)
        payload = {
            "model": self._model,
            "temperature": self._temperature,
            "messages": [
                {
                    "role": "system",
                    "content": (
                        "あなたはChatworkメッセージ履歴から契約解除の予兆を評価するアナリストです。"
                        "出力は必ずJSONのみで返し、余計な文章は不要です。"
                    ),
                },
                {"role": "user", "content": prompt},
            ],
        }
        headers = {
            "Authorization": f"Bearer {self._api_key}",
            "Content-Type": "application/json",
        }

        import requests

        response = requests.post(
            f"{self._base_url}/v1/chat/completions",
            headers=headers,
            json=payload,
            timeout=self._timeout,
        )
        response.raise_for_status()
        data = response.json()
        content = data["choices"][0]["message"]["content"]
        parsed = parse_prediction_json(content)
        parsed.model = self._model
        return parsed


def build_prompt(messages: List[dict]) -> str:
    lines = []
    for msg in messages:
        body = msg.get("body", "") or ""
        account = msg.get("account", {}) or {}
        name = account.get("name") or account.get("account_id") or "unknown"
        ts = msg.get("send_time")
        if ts:
            dt_value = dt.datetime.fromtimestamp(int(ts)).strftime("%Y-%m-%d %H:%M")
        else:
            dt_value = "unknown"
        lines.append(f"[{dt_value}] ({name}) {body}")

    return (
        "以下のChatworkメッセージ履歴から契約解除の予兆を評価してください。\n"
        "要件:\n"
        "- 出力は必ずJSONのみ\n"
        "- フィールド: risk_level (low|medium|high), score (0-1), signals (配列), summary (1文)\n"
        "\n"
        "メッセージ履歴:\n"
        + "\n".join(lines)
    )


def parse_prediction_json(content: str) -> PredictionResult:
    text = content.strip()
    if text.startswith("```"):
        text = text.strip("`")
        text = text.replace("json", "", 1).strip()
    try:
        payload = json.loads(text)
    except json.JSONDecodeError:
        return PredictionResult(
            risk_level="unknown",
            score=0.0,
            signals=["モデル出力のJSON解析に失敗"],
            summary=content[:200],
            model="unknown",
        )

    return PredictionResult(
        risk_level=str(payload.get("risk_level", "unknown")),
        score=float(payload.get("score", 0.0)),
        signals=list(payload.get("signals", [])) or ["signals未設定"],
        summary=str(payload.get("summary", "")),
        model="unknown",
    )


def resolve_handler() -> LLMHandler:
    handler = os.getenv("LLM_HANDLER", "").strip().lower()
    api_key = os.getenv("OPENAI_API_KEY", "")
    if not handler:
        handler = "openai" if api_key else "rule"

    if handler == "openai":
        if not api_key:
            raise RuntimeError("OPENAI_API_KEYが設定されていません。")
        model = os.getenv("OPENAI_MODEL", DEFAULT_OPENAI_MODEL)
        base_url = os.getenv("OPENAI_BASE_URL", DEFAULT_OPENAI_BASE_URL)
        temperature = float(os.getenv("LLM_TEMPERATURE", "0.2"))
        timeout = int(os.getenv("LLM_TIMEOUT", "30"))
        return OpenAIChatHandler(
            api_key=api_key,
            model=model,
            base_url=base_url,
            temperature=temperature,
            timeout=timeout,
        )

    if handler == "rule":
        return RuleBasedHandler()

    raise RuntimeError(f"未対応のLLM_HANDLERです: {handler}")


def select_messages(messages: List[dict], limit: int) -> List[dict]:
    if limit <= 0:
        return messages
    return messages[-limit:]


def format_result(result: PredictionResult) -> str:
    payload = {
        "risk_level": result.risk_level,
        "score": result.score,
        "signals": result.signals,
        "summary": result.summary,
        "model": result.model,
    }
    return json.dumps(payload, ensure_ascii=False, indent=2)


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Chatworkメッセージ履歴から契約解除の予兆を予測"
    )
    parser.add_argument("--room-id", required=True, help="対象のChatworkルームID")
    parser.add_argument("--limit", type=int, default=200, help="解析に使う最新メッセージ件数")
    return parser


def main() -> None:
    load_env_file()
    args = build_arg_parser().parse_args()

    token = load_required_env("CHATWORK_TOKEN")
    chatwork = ChatworkClient(token)
    messages = chatwork.list_messages(args.room_id)
    target = select_messages(messages, args.limit)

    handler = resolve_handler()
    result = handler.predict(target)
    print(format_result(result))


if __name__ == "__main__":
    main()
