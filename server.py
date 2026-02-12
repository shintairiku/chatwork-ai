import json
import os
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse

from fetch_update import main as fetch_update_main

from notify import run as notify_run


def run_fetch_update() -> None:
    fetch_update_main()

def run_notify(threshold_days: int = 7) -> None:
    notify_run(threshold_days)

class Handler(BaseHTTPRequestHandler):
    def _send_json(self, status: int, payload: dict) -> None:
        data = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self) -> None:
        path = urlparse(self.path).path
        if path == "/health":
            self._send_json(HTTPStatus.OK, {"status": "ok"})
            return
        self._send_json(HTTPStatus.NOT_FOUND, {"error": "not_found"})

    def do_POST(self) -> None:
        path = urlparse(self.path).path
        if path not in ("/run"):
            self._send_json(HTTPStatus.NOT_FOUND, {"error": "not_found"})
            return
        try:
            run_fetch_update()
            run_notify(threshold_days=1)
        except Exception as exc:  # pragma: no cover - runtime error path
            self._send_json(HTTPStatus.INTERNAL_SERVER_ERROR, {"status": "error", "message": str(exc)})
            return
        self._send_json(HTTPStatus.OK, {"status": "ok"})

    def log_message(self, fmt: str, *args) -> None:
        message = fmt % args
        print(f"{self.address_string()} - {message}")


def main() -> None:
    port = int(os.getenv("PORT", "8080"))
    server = ThreadingHTTPServer(("0.0.0.0", port), Handler)
    print(f"Listening on 0.0.0.0:{port}")
    server.serve_forever()


if __name__ == "__main__":
    main()
