"""Сбор данных из Яндекс.Метрики под структуру отчёта КХЛ.

Каждый раздел получает DataBlock: заголовок, колонки, строки и краткая
текстовая сводка фактов (её читает Claude при написании прозы).
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from typing import Any

from .metrika_client import MetrikaClient

RU_FILTER = "ym:s:regionCountryName=='Россия'"


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
    """Сезон КХЛ: 1 июля прошлого года — 30 июня текущего.

    Возвращает (date1, date2, label), напр. ("2024-07-01","2025-06-30","2024/2025").
    """
    today = date.today()
    # Если сейчас до июля — текущий сезон ещё не закончился, берём прошлый.
    end_year = today.year if today.month >= 7 else today.year
    start = date(end_year - 1, 7, 1)
    end = date(end_year, 6, 30)
    return start.isoformat(), end.isoformat(), f"{start.year}/{end.year}"


def _rows_with_share(payload: dict[str, Any], total: float) -> list[list[str]]:
    rows: list[list[str]] = []
    for item in payload.get("data", []):
        dims = item.get("dimensions", [])
        name = (dims[0].get("name") if dims else None) or "не задано"
        metrics = item.get("metrics", [0])
        visits = int(metrics[0] or 0)
        share = round(visits / total * 100, 1) if total else 0.0
        rows.append([name, str(visits), f"{share}%"])
    return rows


# Спецификации запросов по разделам. Все, кроме сводки, независимы и
# выполняются параллельно — иначе последовательная выборка над периодом
# в сезон не укладывается в лимит времени serverless-функции.
_SPECS: dict[str, dict[str, Any]] = {
    "geo_countries": dict(
        metrics="ym:s:visits", dimensions="ym:s:regionCountry",
        sort="-ym:s:visits", limit=15,
    ),
    "geo_regions": dict(
        metrics="ym:s:visits", dimensions="ym:s:regionCity",
        sort="-ym:s:visits", limit=15, filters=RU_FILTER,
    ),
    "demography_age": dict(
        metrics="ym:s:visits,ym:s:bounceRate,ym:s:pageDepth,ym:s:avgVisitDurationSeconds",
        dimensions="ym:s:ageInterval", sort="-ym:s:visits", limit=10,
    ),
    "demography_gender": dict(
        metrics="ym:s:visits", dimensions="ym:s:gender",
        sort="-ym:s:visits", limit=5,
    ),
    "interests": dict(
        metrics="ym:s:affinityIndexInterests", dimensions="ym:s:interest",
        sort="-ym:s:affinityIndexInterests", limit=10,
    ),
    "traffic_sources": dict(
        metrics="ym:s:visits", dimensions="ym:s:lastTrafficSource",
        sort="-ym:s:visits", limit=10,
    ),
    "social": dict(
        metrics="ym:s:visits", dimensions="ym:s:lastSocialNetwork",
        sort="-ym:s:visits", limit=10,
    ),
    "devices": dict(
        metrics="ym:s:visits", dimensions="ym:s:deviceCategory",
        sort="-ym:s:visits", limit=10,
    ),
    "os": dict(
        metrics="ym:s:visits", dimensions="ym:s:operatingSystemRoot",
        sort="-ym:s:visits", limit=10,
    ),
}


def _fetch_parallel(
    client: MetrikaClient, d1: str, d2: str, max_workers: int = 5
) -> dict[str, dict[str, Any]]:
    from concurrent.futures import ThreadPoolExecutor

    def run(key: str, spec: dict[str, Any]) -> tuple[str, dict[str, Any]]:
        return key, client.query(
            spec["metrics"], spec["dimensions"], d1, d2,
            limit=spec.get("limit", 10), sort=spec.get("sort"),
            filters=spec.get("filters"),
        )

    results: dict[str, dict[str, Any]] = {}
    with ThreadPoolExecutor(max_workers=max_workers) as ex:
        for key, payload in ex.map(lambda kv: run(*kv), _SPECS.items()):
            results[key] = payload
    return results


def collect(client: MetrikaClient, d1: str, d2: str) -> dict[str, DataBlock]:
    """Собирает все блоки данных для отчёта за период d1..d2."""
    blocks: dict[str, DataBlock] = {}

    # --- Сводка (нужна первой: даёт общее число визитов для долей) ---
    summary = client.query(
        "ym:s:visits,ym:s:users,ym:s:pageviews,ym:s:bounceRate,"
        "ym:s:avgVisitDurationSeconds,ym:s:pageDepth",
        None,
        d1,
        d2,
        limit=1,
    )
    t = summary.get("totals", [0, 0, 0, 0, 0, 0])
    visits = int(t[0] or 0)
    dur = int(round(float(t[4] or 0)))
    blocks["summary"] = DataBlock(
        columns=["Показатель", "Значение"],
        rows=[
            ["Визиты", str(visits)],
            ["Пользователи", str(int(t[1] or 0))],
            ["Просмотры страниц", str(int(t[2] or 0))],
            ["Отказы", f"{round(float(t[3] or 0), 1)}%"],
            ["Средняя длительность визита", f"{dur // 60} мин {dur % 60} сек"],
            ["Глубина просмотра", str(round(float(t[5] or 0), 2))],
        ],
        totals={"визиты": visits, "пользователи": int(t[1] or 0)},
    )

    # --- Остальные разделы — параллельно ---
    p = _fetch_parallel(client, d1, d2)

    blocks["geo_countries"] = DataBlock(
        columns=["Страна", "Визиты", "Доля"],
        rows=_rows_with_share(p["geo_countries"], visits),
    )
    blocks["geo_regions"] = DataBlock(
        columns=["Город", "Визиты", "Доля"],
        rows=_rows_with_share(p["geo_regions"], visits),
    )

    age_rows = []
    for item in p["demography_age"].get("data", []):
        dims = item.get("dimensions", [])
        name = (dims[0].get("name") if dims else None) or "не задано"
        m = item.get("metrics", [0, 0, 0, 0])
        v = int(m[0] or 0)
        share = round(v / visits * 100, 1) if visits else 0.0
        age_rows.append(
            [name, str(v), f"{share}%", f"{round(float(m[1] or 0), 1)}%",
             str(round(float(m[2] or 0), 2))]
        )
    blocks["demography_age"] = DataBlock(
        columns=["Возраст", "Визиты", "Доля", "Отказы", "Глубина"],
        rows=age_rows,
    )

    blocks["demography_gender"] = DataBlock(
        columns=["Пол", "Визиты", "Доля"],
        rows=_rows_with_share(p["demography_gender"], visits),
    )

    int_rows = []
    for item in p["interests"].get("data", []):
        dims = item.get("dimensions", [])
        name = (dims[0].get("name") if dims else None) or "не задано"
        m = item.get("metrics", [0])
        int_rows.append([name, str(round(float(m[0] or 0)))])
    blocks["interests"] = DataBlock(
        columns=["Интерес", "Аффинити-индекс"],
        rows=int_rows,
    )

    blocks["traffic_sources"] = DataBlock(
        columns=["Источник трафика", "Визиты", "Доля"],
        rows=_rows_with_share(p["traffic_sources"], visits),
    )

    social_rows = [r for r in _rows_with_share(p["social"], visits) if r[0] != "не задано"]
    social_total = sum(int(r[1]) for r in social_rows) or 1
    for r in social_rows:  # доля внутри соцсетей
        r[2] = f"{round(int(r[1]) / social_total * 100, 1)}%"
    blocks["social"] = DataBlock(
        columns=["Социальная сеть", "Визиты", "Доля"],
        rows=social_rows,
    )

    blocks["devices"] = DataBlock(
        columns=["Тип устройства", "Визиты", "Доля"],
        rows=_rows_with_share(p["devices"], visits),
    )

    blocks["os"] = DataBlock(
        columns=["Операционная система", "Визиты", "Доля"],
        rows=_rows_with_share(p["os"], visits),
    )

    return blocks
