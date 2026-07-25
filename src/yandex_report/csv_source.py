"""Загрузка данных отчёта из CSV-выгрузок Яндекс.Метрики.

Позволяет строить отчёт из файлов, экспортированных в интерфейсе Метрики
(«Экспорт → CSV»), без обращения к API — удобно, когда API упирается в лимиты.

Каждый раздел — пара файлов (текущий сезон и прошлый). Тип файла определяется
по первому заголовку столбца, текущий/прошлый — по диапазону дат в имени файла.
"""
from __future__ import annotations

import csv
import io
import re
from pathlib import Path

from .report_data import DataBlock, _change

# Первый столбец заголовка -> идентификатор раздела.
_TYPE_BY_HEADER: dict[str, str] = {
    "Интервал дат визита": "summary",
    "Страна": "geo_countries",
    "Регион": "geo_regions",
    "Область": "geo_regions",
    "Город": "geo_regions",
    "Возраст": "demography_age",
    "Пол": "demography_gender",
    "Источник трафика": "traffic_sources",
    "Социальная сеть": "social",
    "Тип устройства": "devices",
    "Устройство": "devices",
    "Операционная система": "os",
}

_TOTAL_ROW = "Итого и средние"

_COLUMNS: dict[str, list[str]] = {
    "geo_countries": ["Страна", "Визиты", "Доля", "Изменение"],
    "geo_regions": ["Регион", "Визиты", "Доля", "Изменение"],
    "demography_age": ["Возраст", "Визиты", "Доля", "Изменение"],
    "demography_gender": ["Пол", "Визиты", "Доля", "Изменение"],
    "traffic_sources": ["Источник трафика", "Визиты", "Доля", "Изменение"],
    "devices": ["Тип устройства", "Визиты", "Доля", "Изменение"],
    "os": ["Операционная система", "Визиты", "Доля", "Изменение"],
    "social": ["Социальная сеть", "Визиты", "Доля", "Изменение"],
}


class CsvError(RuntimeError):
    pass


def _parse_text(name: str, text: str) -> tuple[list[str], list[list[str]]]:
    rows = [r for r in csv.reader(io.StringIO(text)) if r]
    if not rows:
        raise CsvError(f"Пустой файл: {name}")
    return rows[0], rows[1:]


def _date_range(name: str) -> tuple[str, str] | None:
    m = re.search(r"(\d{8})(\d{8})", name)
    if not m:
        return None
    return m.group(1), m.group(2)


def _num(s: str) -> float:
    try:
        return float((s or "0").replace(",", ".").replace("\xa0", "").replace(" ", ""))
    except ValueError:
        return 0.0


def _fmt_time(hms: str) -> str:
    parts = [int(x) for x in hms.split(":")] if ":" in hms else [0, 0, 0]
    while len(parts) < 3:
        parts.insert(0, 0)
    total = parts[0] * 3600 + parts[1] * 60 + parts[2]
    return f"{total // 60} мин {total % 60} сек"


def _totals_row(data_rows: list[list[str]]) -> list[str] | None:
    for r in data_rows:
        if r and r[0].strip() == _TOTAL_ROW:
            return r
    return None


def _data_only(data_rows: list[list[str]]) -> list[list[str]]:
    return [r for r in data_rows if r and r[0].strip() != _TOTAL_ROW]


def _visits_map(data_rows: list[list[str]]) -> dict[str, float]:
    return {r[0].strip(): _num(r[1]) for r in _data_only(data_rows) if len(r) > 1}


def _summary_block(cur_rows: list[list[str]], prev_rows: list[list[str]] | None) -> DataBlock:
    a = _totals_row(cur_rows)
    if not a:
        raise CsvError("В сводке нет строки «Итого и средние».")
    b = _totals_row(prev_rows) if prev_rows else None

    def val(row: list[str], i: int) -> float:
        return _num(row[i]) if row and len(row) > i else 0.0

    def chg(i: int) -> str:
        return _change(val(a, i), val(b, i)) if b else "—"

    time_str = a[7] if len(a) > 7 else "00:00:00"
    rows = [
        ["Визиты", str(int(val(a, 1))), chg(1)],
        ["Пользователи", str(int(val(a, 2))), chg(2)],
        ["Просмотры страниц", str(int(val(a, 3))), chg(3)],
        ["Доля новых посетителей", f"{round(val(a, 4) * 100, 1)}%", chg(4)],
        ["Отказы", f"{round(val(a, 5) * 100, 1)}%", chg(5)],
        ["Глубина просмотра", str(round(val(a, 6), 2)), chg(6)],
        ["Средняя длительность визита", _fmt_time(time_str), "—"],
    ]
    return DataBlock(
        columns=["Показатель", "Значение", "Изменение"],
        rows=rows,
        totals={"визиты": int(val(a, 1)), "пользователи": int(val(a, 2))},
    )


def _dim_block(
    block_id: str, cur_rows: list[list[str]], prev_rows: list[list[str]] | None,
    total_visits: float,
) -> DataBlock:
    prev_map = _visits_map(prev_rows) if prev_rows else {}
    rows: list[list[str]] = []
    for r in _data_only(cur_rows):
        if len(r) < 2:
            continue
        name = r[0].strip()
        cur = _num(r[1])
        share = round(cur / total_visits * 100, 1) if total_visits else 0.0
        rows.append([name, str(int(cur)), f"{share}%", _change(cur, prev_map.get(name, 0))])
    return DataBlock(columns=_COLUMNS.get(block_id, ["", "Визиты", "Доля", "Изменение"]), rows=rows)


def load_blocks_from_files(
    files: list[tuple[str, str]], limit_rows: int = 15
) -> tuple[dict[str, DataBlock], str]:
    """Строит блоки отчёта из пар (имя файла, содержимое CSV).

    Возвращает (blocks, season). Файлы одного раздела группируются, текущим
    считается файл с более поздним диапазоном дат в имени.
    """
    # раздел -> список (диапазон дат, строки данных)
    grouped: dict[str, list[tuple[tuple[str, str] | None, list[list[str]]]]] = {}
    for name, text in files:
        header, rows = _parse_text(name, text)
        block_id = _TYPE_BY_HEADER.get(header[0].strip())
        if not block_id:
            continue
        grouped.setdefault(block_id, []).append((_date_range(name), rows))

    if not grouped:
        raise CsvError(
            "Не распознан ни один файл. Ожидаются CSV-выгрузки Метрики "
            "с заголовками «Интервал дат визита», «Страна» и т. п."
        )

    def end_key(item):
        rng = item[0]
        return rng[1] if rng else ""

    season = ""
    parsed: dict[str, tuple[list[list[str]], list[list[str]] | None, tuple | None]] = {}
    for block_id, items in grouped.items():
        items.sort(key=end_key)  # позже = текущий сезон
        cur_rng, cur_rows = items[-1]
        prev_rows = items[-2][1] if len(items) > 1 else None
        parsed[block_id] = (cur_rows, prev_rows, cur_rng)
        if cur_rng and not season:
            season = f"{cur_rng[0][:4]}/{cur_rng[1][:4]}"

    # Общее число визитов для долей — из сводки, если она есть.
    total_visits = 0.0
    if "summary" in parsed:
        tr = _totals_row(parsed["summary"][0])
        total_visits = _num(tr[1]) if tr else 0.0

    blocks: dict[str, DataBlock] = {}
    for block_id, (cur_rows, prev_rows, _rng) in parsed.items():
        if block_id == "summary":
            blocks["summary"] = _summary_block(cur_rows, prev_rows)
            continue
        tr = _totals_row(cur_rows)
        tv = total_visits or (_num(tr[1]) if tr else 0.0)
        block = _dim_block(block_id, cur_rows, prev_rows, tv)
        block.rows = block.rows[:limit_rows]
        blocks[block_id] = block

    return blocks, season


def load_blocks(
    paths: list[str | Path], limit_rows: int = 15
) -> tuple[dict[str, DataBlock], str]:
    """Строит блоки из списка путей к CSV-файлам Метрики."""
    files = []
    for p in paths:
        path = Path(p)
        files.append((path.name, path.read_text(encoding="utf-8-sig")))
    return load_blocks_from_files(files, limit_rows=limit_rows)
