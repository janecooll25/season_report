"""Vercel serverless-функция: аналитическая справка по клубу (билетная программа).

POST /api/spravka
  • {list_clubs:true, monitoring_b64}                       → JSON {clubs:[...]}
  • {monitoring_b64, club, season?, prev_season?, aggregate_b64?,
     ticket_b64?, capacity?, to_rub?, income_b64?, income_year?} → docx

aggregate_b64 — общий файл «Средние по клубам» (данные текущего сезона по всем
клубам, включая агентскую комиссию); альтернатива отдельному билетному файлу.
"""
from __future__ import annotations

import base64
import json
import os
import sys
from http.server import BaseHTTPRequestHandler
from urllib.parse import quote

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from yandex_report.club_report import build_club_report, list_clubs  # noqa: E402
from yandex_report.ticket_metrics import TicketError  # noqa: E402

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


def _b64(payload: dict, key: str):
    v = payload.get(key)
    return base64.b64decode(v) if v else None


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
            monitoring = base64.b64decode(payload["monitoring_b64"])
        except (ValueError, KeyError, UnicodeDecodeError):
            self._json(400, {"error": "Ожидается JSON с полем monitoring_b64 (xlsx в base64)."})
            return

        # Режим: список клубов из файла мониторинга.
        if payload.get("list_clubs"):
            try:
                self._json(200, {"clubs": list_clubs(monitoring)})
            except Exception as exc:  # noqa: BLE001
                self._json(400, {"error": f"Не удалось прочитать файл мониторинга: {exc}"})
            return

        club = (payload.get("club") or "").strip()
        if not club:
            self._json(400, {"error": "Не выбран клуб."})
            return
        season = (payload.get("season") or "25/26").strip() or "25/26"
        prev_season = (payload.get("prev_season") or "24/25").strip() or "24/25"
        income_year = (payload.get("income_year") or "2025").strip() or "2025"

        def _num(key):
            v = payload.get(key)
            try:
                return float(v) if v not in (None, "", 0, "0") else None
            except (TypeError, ValueError):
                return None

        capacity = _num("capacity")
        to_rub = _num("to_rub") or 1.0

        try:
            data = build_club_report(
                monitoring, club, season=season, prev_season=prev_season,
                aggregate_bytes=_b64(payload, "aggregate_b64"),
                ticket_bytes=_b64(payload, "ticket_b64"),
                capacity=int(capacity) if capacity else None, to_rub=to_rub,
                region_income_bytes=_b64(payload, "income_b64"), income_year=income_year,
            )
        except TicketError as exc:
            self._json(400, {"error": str(exc)})
            return
        except Exception as exc:  # noqa: BLE001
            self._json(500, {"error": f"{exc.__class__.__name__}: {exc}"})
            return

        filename = f"spravka_{club}_{season.replace('/', '-')}.docx"
        self.send_response(200)
        self.send_header("Content-Type", DOCX_MIME)
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
