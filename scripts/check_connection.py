"""Проверка подключения к Яндекс.Метрике.

Запуск:
    PYTHONPATH=src python scripts/check_connection.py

Читает YANDEX_METRIKA_TOKEN и YANDEX_METRIKA_COUNTER_ID из .env / окружения,
проверяет доступ к счётчику и делает тестовый запрос статистики за 7 дней.
"""
from __future__ import annotations

import sys
from datetime import date, timedelta

import requests

from yandex_report.config import Config
from yandex_report.metrika_client import API_URL, MetrikaClient, MetrikaError

MGMT_URL = "https://api-metrika.yandex.net/management/v1/counter/{id}"


def main() -> int:
    try:
        config = Config.from_env(require_llm=False)
    except RuntimeError as exc:
        print(f"✗ {exc}", file=sys.stderr)
        return 2

    session = requests.Session()
    session.headers.update({"Authorization": f"OAuth {config.token}"})

    # 1. Доступ к счётчику (management API) — заодно узнаём имя/сайт.
    print(f"Проверяю доступ к счётчику {config.counter_id}…")
    try:
        resp = session.get(MGMT_URL.format(id=config.counter_id), timeout=30)
    except requests.RequestException as exc:
        print(f"✗ Сетевая ошибка: {exc}", file=sys.stderr)
        return 1

    if resp.status_code == 401:
        print("✗ 401 — токен недействителен или истёк. Получите новый.", file=sys.stderr)
        return 1
    if resp.status_code == 403:
        print("✗ 403 — у токена нет доступа к этому счётчику.", file=sys.stderr)
        return 1
    if resp.status_code == 404:
        print("✗ 404 — счётчик не найден. Проверьте YANDEX_METRIKA_COUNTER_ID.", file=sys.stderr)
        return 1
    if resp.status_code != 200:
        print(f"✗ Management API вернул {resp.status_code}: {resp.text[:300]}", file=sys.stderr)
        return 1

    counter = resp.json().get("counter", {})
    print(f"✓ Счётчик доступен: «{counter.get('name', '—')}» ({counter.get('site', '—')})")

    # 2. Тестовый запрос статистики за 7 дней.
    client = MetrikaClient(config.token, config.counter_id, session=session)
    d2 = date.today()
    d1 = d2 - timedelta(days=7)
    print(f"Тестовый запрос статистики за {d1} — {d2}…")
    try:
        data = client.query(
            "ym:s:visits,ym:s:users", None, d1.isoformat(), d2.isoformat(), limit=1
        )
    except MetrikaError as exc:
        print(f"✗ Ошибка запроса статистики: {exc}", file=sys.stderr)
        return 1

    totals = data.get("totals", [0, 0])
    print(f"✓ Данные получены: визитов {int(totals[0] or 0)}, "
          f"пользователей {int(totals[1] or 0)} за 7 дней")
    print("\nПодключение работает. Можно запускать генерацию отчёта.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
