"""Сбор данных из Яндекс.Метрики под структуру отчёта КХЛ.

Каждый раздел получает DataBlock: колонки, строки и краткая текстовая сводка
фактов (её читает Claude при написании прозы). Данные собираются в сравнении
текущего сезона (A) с прошлым (B) — отсюда колонка «Изменение».
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
    """Изменение в абсолютных пунктах (для индексов/долей)."""
    d = round(cur - prev)
    return f"+{d}" if d >= 0 else f"−{abs(d)}"


def _ab(item: dict[str, Any]) -> tuple[list, list]:
    """Достаёт пару [значения A],[значения B] из строки comparison-ответа."""
    m = item.get("metrics") or [[0], [0]]
    a = m[0] if len(m) > 0 else [0]
    b = m[1] if len(m) > 1 else [0]
    return a, b


def _rows_share_change(payload: dict[str, Any], total: float) -> list[list[str]]:
    rows: list[list[str]] = []
    for item in payload.get("data", []):
        dims = item.get("dimensions", [])
        name = (dims[0].get("name") if dims else None) or "не задано"
        a, b = _ab(item)
        cur = int(a[0] or 0)
        prev = int(b[0] or 0)
        share = round(cur / total * 100, 1) if total else 0.0
        rows.append([name, str(cur), f"{share}%", _change(cur, prev)])
    return rows


# Разделы (кроме сводки и интересов) — счётчик визитов + доля + изменение.
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
        metrics="ym:s:visits", dimensions="ym:s:ageInterval",
        sort="-ym:s:visits", limit=10,
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

# Колонки таблиц по разделам (по умолчанию — как у большинства).
_COLUMNS: dict[str, list[str]] = {
    "geo_countries": ["Страна", "Визиты", "Доля", "Изменение"],
    "geo_regions": ["Город", "Визиты", "Доля", "Изменение"],
    "demography_age": ["Возраст", "Визиты", "Доля", "Изменение"],
    "demography_gender": ["Пол", "Визиты", "Доля", "Изменение"],
    "traffic_sources": ["Источник трафика", "Визиты", "Доля", "Изменение"],
    "social": ["Социальная сеть", "Визиты", "Доля", "Изменение"],
    "devices": ["Тип устройства", "Визиты", "Доля", "Изменение"],
    "os": ["Операционная система", "Визиты", "Доля", "Изменение"],
}


def _fetch_parallel(
    client: MetrikaClient, a1: str, a2: str, b1: str, b2: str, max_workers: int = 5
) -> dict[str, dict[str, Any]]:
    from concurrent.futures import ThreadPoolExecutor

    def run(item: tuple[str, dict[str, Any]]) -> tuple[str, dict[str, Any]]:
        key, spec = item
        return key, client.query_comparison(
            spec["metrics"], spec["dimensions"], a1, a2, b1, b2,
            limit=spec.get("limit", 10), sort=spec.get("sort"),
            filters=spec.get("filters"),
        )

    results: dict[str, dict[str, Any]] = {}
    with ThreadPoolExecutor(max_workers=max_workers) as ex:
        for key, payload in ex.map(run, _SPECS.items()):
            results[key] = payload
    return results


def collect(client: MetrikaClient, d1: str, d2: str) -> dict[str, DataBlock]:
    """Собирает все блоки данных за сезон d1..d2 в сравнении с прошлым сезоном."""
    b1, b2 = prev_period(d1, d2)
    blocks: dict[str, DataBlock] = {}

    # --- Сводка (нужна первой: даёт общее число визитов для долей) ---
    summary = client.query_comparison(
        "ym:s:visits,ym:s:users,ym:s:pageviews,ym:s:bounceRate,"
        "ym:s:avgVisitDurationSeconds,ym:s:pageDepth",
        None,
        d1, d2, b1, b2,
        limit=1,
    )
    tot = summary.get("totals") or [[0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0]]
    a = tot[0]
    bp = tot[1] if len(tot) > 1 else [0] * 6
    visits = int(a[0] or 0)
    dur = int(round(float(a[4] or 0)))
    blocks["summary"] = DataBlock(
        columns=["Показатель", "Значение", "Изменение"],
        rows=[
            ["Визиты", str(visits), _change(visits, int(bp[0] or 0))],
            ["Пользователи", str(int(a[1] or 0)), _change(int(a[1] or 0), int(bp[1] or 0))],
            ["Просмотры страниц", str(int(a[2] or 0)), _change(int(a[2] or 0), int(bp[2] or 0))],
            ["Отказы", f"{round(float(a[3] or 0), 1)}%",
             _change(float(a[3] or 0), float(bp[3] or 0))],
            ["Средняя длительность визита", f"{dur // 60} мин {dur % 60} сек",
             _change(float(a[4] or 0), float(bp[4] or 0))],
            ["Глубина просмотра", str(round(float(a[5] or 0), 2)),
             _change(float(a[5] or 0), float(bp[5] or 0))],
        ],
        totals={"визиты": visits, "пользователи": int(a[1] or 0)},
    )

    # --- Остальные разделы — параллельно ---
    p = _fetch_parallel(client, d1, d2, b1, b2)

    for key in ("geo_countries", "geo_regions", "demography_age",
                "demography_gender", "traffic_sources", "devices", "os"):
        blocks[key] = DataBlock(
            columns=_COLUMNS[key],
            rows=_rows_share_change(p[key], visits),
        )

    # --- Интересы (аффинити-индекс) ---
    int_rows = []
    for item in p["interests"].get("data", []):
        dims = item.get("dimensions", [])
        name = (dims[0].get("name") if dims else None) or "не задано"
        av, bv = _ab(item)
        cur = round(float(av[0] or 0))
        prev = round(float(bv[0] or 0))
        int_rows.append([name, str(cur), _change_points(cur, prev)])
    blocks["interests"] = DataBlock(
        columns=["Интерес", "Аффинити-индекс", "Изменение"],
        rows=int_rows,
    )

    # --- Социальные сети: доля внутри соцсетей ---
    social_rows = [r for r in _rows_share_change(p["social"], visits) if r[0] != "не задано"]
    social_total = sum(int(r[1]) for r in social_rows) or 1
    for r in social_rows:
        r[2] = f"{round(int(r[1]) / social_total * 100, 1)}%"
    blocks["social"] = DataBlock(
        columns=["Социальная сеть", "Визиты", "Доля", "Изменение"],
        rows=social_rows,
    )

    return blocks
