"""Vercel serverless-функция: генерация отчёта КХЛ по интернет-аудитории.

GET  /api/generate?from=YYYY-MM-DD&to=YYYY-MM-DD&season=2024/2025&llm=1
     — данные из API Метрики.
POST /api/generate  (JSON: {season?, llm?, files:[{name, content}]})
     — данные из загруженных CSV-выгрузок Метрики.
Отдаёт готовый .docx как вложение (или JSON с ошибкой).
"""
from __future__ import annotations

import asyncio
import json
import os
import sys
from datetime import datetime
from http.server import BaseHTTPRequestHandler
from urllib.parse import parse_qs, quote, urlparse

# Пакет лежит в src/ — добавляем в путь (bundled через vercel.json includeFiles).
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from yandex_report.config import DEFAULT_MODEL, Config  # noqa: E402
from yandex_report.csv_source import CsvError, load_blocks_from_files  # noqa: E402
from yandex_report.docx_builder import build_report_bytes  # noqa: E402
from yandex_report.metrika_client import MetrikaClient, MetrikaError  # noqa: E402
from yandex_report.report_data import collect, default_season  # noqa: E402
from yandex_report.report_writer import write_sections_parallel  # noqa: E402
from yandex_report.structure import SECTIONS  # noqa: E402

DOCX_MIME = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
)


def _prose(blocks, season, use_llm):
    if not use_llm:
        return {}
    key = os.environ.get("ANTHROPIC_API_KEY", "").strip()
    if not key:
        raise RuntimeError(
            "Не задан ANTHROPIC_API_KEY — он нужен для генерации текста. "
            "Добавьте его в переменные окружения проекта или снимите галочку «текст»."
        )
    model = os.environ.get("REPORT_MODEL", DEFAULT_MODEL).strip() or DEFAULT_MODEL
    return asyncio.run(write_sections_parallel(key, model, SECTIONS, blocks, season))


def _generate(params: dict[str, list[str]]) -> tuple[str, bytes]:
    use_llm = params.get("llm", ["1"])[0] != "0"
    config = Config.from_env(require_llm=use_llm)

    date_from = params.get("from", [""])[0]
    date_to = params.get("to", [""])[0]
    season = params.get("season", [""])[0]

    if date_from and date_to:
        d1, d2 = date_from, date_to
        season = season or f"{d1[:4]}/{d2[:4]}"
    else:
        d1, d2, auto_season = default_season()
        season = season or auto_season

    charts = _charts_from_query(params.get("charts", [""])[0])
    client = MetrikaClient(config.token, config.counter_id)
    blocks = collect(client, d1, d2)
    prose = _prose(blocks, season, use_llm)

    data = build_report_bytes(
        season=season,
        counter_id=config.counter_id,
        generated_at=datetime.now().strftime("%Y-%m-%d %H:%M"),
        blocks=blocks,
        prose=prose,
        charts=charts,
    )
    return f"khl_report_{season.replace('/', '_')}.docx", data


def _charts_from_query(raw: str) -> dict[str, str]:
    """'geo_countries:pie,os:bar' -> {'geo_countries':'pie', 'os':'bar'}."""
    out: dict[str, str] = {}
    for part in (raw or "").split(","):
        if ":" in part:
            k, v = part.split(":", 1)
            if k.strip() and v.strip():
                out[k.strip()] = v.strip()
    return out


def _generate_from_csv(payload: dict) -> tuple[str, bytes]:
    files = [
        (f.get("name", "file.csv"), f.get("content", ""))
        for f in payload.get("files", [])
    ]
    if not files:
        raise CsvError("Не приложено ни одного CSV-файла.")
    use_llm = bool(payload.get("llm", True))
    charts = payload.get("charts") or {}

    blocks, auto_season = load_blocks_from_files(files)
    season = (payload.get("season") or "").strip() or auto_season or default_season()[2]
    prose = _prose(blocks, season, use_llm)

    data = build_report_bytes(
        season=season,
        counter_id=os.environ.get("YANDEX_METRIKA_COUNTER_ID", "—"),
        generated_at=datetime.now().strftime("%Y-%m-%d %H:%M"),
        blocks=blocks,
        prose=prose,
        charts=charts if isinstance(charts, dict) else {},
    )
    return f"khl_report_{season.replace('/', '_')}.docx", data


class handler(BaseHTTPRequestHandler):
    def do_GET(self) -> None:  # noqa: N802 (Vercel требует это имя)
        params = parse_qs(urlparse(self.path).query)
        self._run(lambda: _generate(params))

    def do_POST(self) -> None:  # noqa: N802
        length = int(self.headers.get("Content-Length", "0") or "0")
        raw = self.rfile.read(length) if length else b"{}"
        try:
            payload = json.loads(raw.decode("utf-8"))
        except (ValueError, UnicodeDecodeError):
            self._json(400, {"error": "Некорректный JSON в теле запроса."})
            return
        self._run(lambda: _generate_from_csv(payload))

    def _run(self, produce) -> None:
        try:
            filename, data = produce()
        except (RuntimeError, CsvError) as exc:  # конфигурация / разбор CSV
            self._json(400, {"error": str(exc)})
            return
        except MetrikaError as exc:
            self._json(502, {"error": f"Ошибка Яндекс.Метрики: {exc}"})
            return
        except Exception as exc:  # noqa: BLE001
            self._json(500, {"error": f"{exc.__class__.__name__}: {exc}"})
            return

        self.send_response(200)
        self.send_header("Content-Type", DOCX_MIME)
        self.send_header(
            "Content-Disposition",
            f"attachment; filename*=UTF-8''{quote(filename)}",
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
