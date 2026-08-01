"""Vercel serverless-функция: расчёт метрик билетной программы.

POST /api/tickets  (JSON: {content_b64, capacity?, season?})
Принимает xlsx (base64), возвращает xlsx с добавленным листом «Расчеты».
"""
from __future__ import annotations

import base64
import json
import os
import sys
from http.server import BaseHTTPRequestHandler
from urllib.parse import quote

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from yandex_report.ticket_metrics import TicketError, build_calculations  # noqa: E402

XLSX_MIME = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)


class handler(BaseHTTPRequestHandler):
    def do_POST(self) -> None:  # noqa: N802
        length = int(self.headers.get("Content-Length", "0") or "0")
        raw = self.rfile.read(length) if length else b"{}"
        if "gzip" in (self.headers.get("Content-Encoding", "") or "").lower():
            try:
                import gzip

                raw = gzip.decompress(raw)
            except Exception:  # noqa: BLE001
                self._json(400, {"error": "Не удалось распаковать тело запроса."})
                return
        try:
            payload = json.loads(raw.decode("utf-8"))
            content = base64.b64decode(payload["content_b64"])
        except (ValueError, KeyError, UnicodeDecodeError):
            self._json(400, {"error": "Ожидается JSON с полем content_b64 (xlsx в base64)."})
            return

        capacity = payload.get("capacity")
        try:
            capacity = int(capacity) if capacity not in (None, "", 0, "0") else None
        except (TypeError, ValueError):
            capacity = None
        season = (payload.get("season") or "2025/2026").strip() or "2025/2026"

        try:
            data = build_calculations(content, capacity=capacity, season_label=season)
        except TicketError as exc:
            self._json(400, {"error": str(exc)})
            return
        except Exception as exc:  # noqa: BLE001
            self._json(500, {"error": f"{exc.__class__.__name__}: {exc}"})
            return

        filename = "tickets_calc.xlsx"
        self.send_response(200)
        self.send_header("Content-Type", XLSX_MIME)
        self.send_header(
            "Content-Disposition", f"attachment; filename*=UTF-8''{quote(filename)}"
        )
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _json(self, code: int, payload: dict) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)
