"""Сбор данных из Яндекс.Метрики под структуру отчёта КХЛ.

Каждый раздел получает DataBlock: колонки, строки и краткая текстовая сводка
фактов (её читает Claude). Для колонки «Изменение» тянем два периода —
текущий сезон и прошлый — обычным эндпоинтом /stat/v1/data и сопоставляем
строки по названию. Все запросы идут параллельно.
"""
from __future__ import annotations

import os
from dataclasses import dataclass, field
from datetime import date
from typing import Any

from .metrika_client import MetrikaClient

RU_FILTER = "ym:s:regionCountryName=='Россия'"

# Метрика ограничивает число ОДНОВРЕМЕННЫХ запросов от пользователя
# (quota_parallel_requests_by_uid). Держим низкий параллелизм; на 429
# срабатывают ретраи в клиенте.
CONCURRENCY = int(os.environ.get("METRIKA_CONCURRENCY", "2"))

SUMMARY_METRICS = (
    "ym:s:visits,ym:s:users,ym:s:pageviews,ym:s:bounceRate,"
    "ym:s:avgVisitDurationSeconds,ym:s:pageDepth"
)


@dataclass
class DataBlock:
    columns: list[str]
    rows: list[list[str]]
    totals: dict[str, float] = field(default_factory=dict)

    def to_facts(self) -> str:
        """Компактное текстовое представление для промпта LLM."""
        lines = [" | ".join(self.columns)]
        for r in self.rows:
            lines.append(" | ".join(str(c) for c in r))
        if self.totals:
            tot = ", ".join(f"{k}: {v}" for k, v in self.totals.items())
            lines.append(f"ИТОГО: {tot}")
        return "\n".join(lines)


def default_season() -> tuple[str, str, str]:
    """Сезон КХЛ: 1 июля прошлого года — 30 июня текущего."""
    today = date.today()
    end_year = today.year
    start = date(end_year - 1, 7, 1)
    end = date(end_year, 6, 30)
    return start.isoformat(), end.isoformat(), f"{start.year}/{end.year}"


def prev_period(d1: str, d2: str) -> tuple[str, str]:
    """Тот же период годом ранее."""
    a = date.fromisoformat(d1)
    b = date.fromisoformat(d2)
    return (
        a.replace(year=a.year - 1).isoformat(),
        b.replace(year=b.year - 1).isoformat(),
    )


def _change(cur: float, prev: float) -> str:
    """Относительное изменение к прошлому сезону, напр. '+12.3%' / '-5.1%'."""
    if not prev:
        return "новое" if cur else "—"
    delta = (cur - prev) / prev * 100
    sign = "+" if delta >= 0 else "−"
    return f"{sign}{abs(round(delta, 1))}%"


def _change_points(cur: float, prev: float) -> str:
    """Изменение в абсолютных пунктах (для индексов)."""
    if not prev:
        return "новое" if cur else "—"
    d = round(cur - prev)
    return f"+{d}" if d >= 0 else f"−{abs(d)}"


def _first_metric(item: dict[str, Any]) -> float:
    m = item.get("metrics") or [0]
    try:
        return float(m[0] or 0)
    except (TypeError, ValueError):
        return 0.0


def _name(item: dict[str, Any]) -> str:
    dims = item.get("dimensions", [])
    return (dims[0].get("name") if dims else None) or "не задано"


def _prev_map(payload: dict[str, Any]) -> dict[str, float]:
    """name -> значение первой метрики за прошлый сезон."""
    return {_name(it): _first_metric(it) for it in payload.get("data", [])}


def _rows_share_change(
    cur: dict[str, Any], prev: dict[str, Any], total: float
) -> list[list[str]]:
    pm = _prev_map(prev)
    rows: list[list[str]] = []
    for item in cur.get("data", []):
        name = _name(item)
        cur_v = int(_first_metric(item))
        share = round(cur_v / total * 100, 1) if total else 0.0
        rows.append([name, str(cur_v), f"{share}%", _change(cur_v, pm.get(name, 0))])
    return rows


# Разделы (кроме сводки и интересов) — счётчик визитов.
_SPECS: dict[str, dict[str, Any]] = {
    "geo_countries": dict(metrics="ym:s:visits", dimensions="ym:s:regionCountry",
                          sort="-ym:s:visits", limit=15),
    "geo_regions": dict(metrics="ym:s:visits", dimensions="ym:s:regionCity",
                        sort="-ym:s:visits", limit=15, filters=RU_FILTER),
    "demography_age": dict(metrics="ym:s:visits", dimensions="ym:s:ageInterval",
                           sort="-ym:s:visits", limit=10),
    "demography_gender": dict(metrics="ym:s:visits", dimensions="ym:s:gender",
                              sort="-ym:s:visits", limit=5),
    "interests": dict(metrics="ym:s:affinityIndexInterests", dimensions="ym:s:interest",
                      sort="-ym:s:affinityIndexInterests", limit=10),
    "traffic_sources": dict(metrics="ym:s:visits", dimensions="ym:s:lastTrafficSource",
                            sort="-ym:s:visits", limit=10),
    "social": dict(metrics="ym:s:visits", dimensions="ym:s:lastSocialNetwork",
                   sort="-ym:s:visits", limit=10),
    "devices": dict(metrics="ym:s:visits", dimensions="ym:s:deviceCategory",
                    sort="-ym:s:visits", limit=10),
    "os": dict(metrics="ym:s:visits", dimensions="ym:s:operatingSystemRoot",
               sort="-ym:s:visits", limit=10),
}

_COLUMNS: dict[str, list[str]] = {
    "geo_countries": ["Страна", "Визиты", "Доля", "Изменение"],
    "geo_regions": ["Город", "Визиты", "Доля", "Изменение"],
    "demography_age": ["Возраст", "Визиты", "Доля", "Изменение"],
    "demography_gender": ["Пол", "Визиты", "Доля", "Изменение"],
    "traffic_sources": ["Источник трафика", "Визиты", "Доля", "Изменение"],
    "devices": ["Тип устройства", "Визиты", "Доля", "Изменение"],
    "os": ["Операционная система", "Визиты", "Доля", "Изменение"],
}


def _fetch_all(
    client: MetrikaClient, d1: str, d2: str, b1: str, b2: str,
    max_workers: int = CONCURRENCY,
) -> dict[tuple[str, str], dict[str, Any]]:
    """Тянет каждый раздел за оба периода параллельно. Ключ — (раздел, 'a'|'b')."""
    from concurrent.futures import ThreadPoolExecutor

    tasks: dict[tuple[str, str], tuple] = {}
    tasks[("summary", "a")] = (SUMMARY_METRICS, None, d1, d2, 1, None, None)
    tasks[("summary", "b")] = (SUMMARY_METRICS, None, b1, b2, 1, None, None)
    for key, spec in _SPECS.items():
        common = (spec["metrics"], spec["dimensions"])
        tail = (spec.get("limit", 10), spec.get("sort"), spec.get("filters"))
        tasks[(key, "a")] = (*common, d1, d2, *tail)
        tasks[(key, "b")] = (*common, b1, b2, *tail)

    def run(item_id: tuple[str, str]):
        m, dim, x1, x2, lim, srt, flt = tasks[item_id]
        payload = client.query(m, dim, x1, x2, limit=lim, sort=srt, filters=flt)
        return item_id, payload

    results: dict[tuple[str, str], dict[str, Any]] = {}
    with ThreadPoolExecutor(max_workers=max_workers) as ex:
        for item_id, payload in ex.map(run, list(tasks.keys())):
            results[item_id] = payload
    return results


def _totals(payload: dict[str, Any]) -> list[float]:
    t = payload.get("totals") or [0, 0, 0, 0, 0, 0]
    return [float(x or 0) for x in t]


def collect(client: MetrikaClient, d1: str, d2: str) -> dict[str, DataBlock]:
    """Собирает все блоки за сезон d1..d2 в сравнении с прошлым сезоном."""
    b1, b2 = prev_period(d1, d2)
    r = _fetch_all(client, d1, d2, b1, b2)
    blocks: dict[str, DataBlock] = {}

    # --- Сводка ---
    a = _totals(r[("summary", "a")])
    bp = _totals(r[("summary", "b")])
    visits = int(a[0])
    dur = int(round(a[4]))
    blocks["summary"] = DataBlock(
        columns=["Показатель", "Значение", "Изменение"],
        rows=[
            ["Визиты", str(visits), _change(a[0], bp[0])],
            ["Пользователи", str(int(a[1])), _change(a[1], bp[1])],
            ["Просмотры страниц", str(int(a[2])), _change(a[2], bp[2])],
            ["Отказы", f"{round(a[3], 1)}%", _change(a[3], bp[3])],
            ["Средняя длительность визита", f"{dur // 60} мин {dur % 60} сек",
             _change(a[4], bp[4])],
            ["Глубина просмотра", str(round(a[5], 2)), _change(a[5], bp[5])],
        ],
        totals={"визиты": visits, "пользователи": int(a[1])},
    )

    # --- Разделы с долей и изменением ---
    for key in ("geo_countries", "geo_regions", "demography_age",
                "demography_gender", "traffic_sources", "devices", "os"):
        blocks[key] = DataBlock(
            columns=_COLUMNS[key],
            rows=_rows_share_change(r[(key, "a")], r[(key, "b")], visits),
        )

    # --- Интересы (аффинити-индекс) ---
    pm = _prev_map(r[("interests", "b")])
    int_rows = []
    for item in r[("interests", "a")].get("data", []):
        name = _name(item)
        cur = round(_first_metric(item))
        int_rows.append([name, str(cur), _change_points(cur, pm.get(name, 0))])
    blocks["interests"] = DataBlock(
        columns=["Интерес", "Аффинити-индекс", "Изменение"],
        rows=int_rows,
    )

    # --- Социальные сети: доля внутри соцсетей ---
    social_rows = [
        row for row in _rows_share_change(r[("social", "a")], r[("social", "b")], visits)
        if row[0] != "не задано"
    ]
    social_total = sum(int(row[1]) for row in social_rows) or 1
    for row in social_rows:
        row[2] = f"{round(int(row[1]) / social_total * 100, 1)}%"
    blocks["social"] = DataBlock(
        columns=["Социальная сеть", "Визиты", "Доля", "Изменение"],
        rows=social_rows,
    )

    return blocks
