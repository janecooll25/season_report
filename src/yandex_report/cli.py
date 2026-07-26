from __future__ import annotations

import argparse
import os
import sys
from datetime import datetime
from pathlib import Path

from .config import DEFAULT_MODEL, Config
from .csv_source import CsvError, load_blocks
from .docx_builder import build_report
from .metrika_client import MetrikaClient, MetrikaError
from .report_data import collect, default_season
from .structure import SECTIONS


def _parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        prog="khl-report",
        description="Генерация docx-отчёта КХЛ по интернет-аудитории из Яндекс.Метрики.",
    )
    p.add_argument("--date-from", help="Начало периода YYYY-MM-DD (по умолчанию — сезон)")
    p.add_argument("--date-to", help="Конец периода YYYY-MM-DD")
    p.add_argument("--season", help="Метка сезона, напр. 2024/2025 (для заголовков)")
    p.add_argument("--output", help="Путь к выходному docx")
    p.add_argument(
        "--from-csv",
        nargs="+",
        metavar="CSV",
        help="Строить из CSV-выгрузок Метрики (файлы или каталог) вместо API",
    )
    p.add_argument(
        "--chart",
        action="append",
        default=[],
        metavar="РАЗДЕЛ=ТИП",
        help="Диаграмма в разделе, напр. --chart geo_countries=pie "
             "(типы: bar, barh, pie, line). Можно указать несколько раз.",
    )
    p.add_argument(
        "--no-llm",
        action="store_true",
        help="Черновик без текста: только заголовки и таблицы (не нужен ANTHROPIC_API_KEY)",
    )
    return p.parse_args(argv)


def _parse_charts(items: list[str]) -> dict[str, str]:
    charts: dict[str, str] = {}
    for item in items:
        if "=" in item:
            k, v = item.split("=", 1)
            if k.strip() and v.strip():
                charts[k.strip()] = v.strip()
    return charts


def _csv_paths(items: list[str]) -> list[Path]:
    paths: list[Path] = []
    for item in items:
        p = Path(item)
        if p.is_dir():
            paths.extend(sorted(p.glob("*.csv")))
        else:
            paths.append(p)
    return paths


def _write_prose(blocks, season, api_key, model) -> dict[str, str]:
    from .report_writer import ReportWriter

    writer = ReportWriter(api_key, model)
    prose: dict[str, str] = {}
    for section in SECTIONS:
        block = blocks.get(section.id)
        if not block or not block.rows:
            continue
        print(f"Пишу раздел: {section.id}…")
        prose[section.id] = writer.write_section(section, block, season)
    return prose


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(argv)

    # --- Режим CSV ---
    if args.from_csv:
        try:
            blocks, auto_season = load_blocks(_csv_paths(args.from_csv))
        except (CsvError, OSError) as exc:
            print(f"Ошибка чтения CSV: {exc}", file=sys.stderr)
            return 1
        season = args.season or auto_season or default_season()[2]
        counter_id = os.environ.get("YANDEX_METRIKA_COUNTER_ID", "—")
        charts = _parse_charts(args.chart)

        prose: dict[str, str] = {}
        if not args.no_llm:
            key = os.environ.get("ANTHROPIC_API_KEY", "").strip()
            if not key:
                print("Ошибка: нужен ANTHROPIC_API_KEY для текста (или --no-llm).",
                      file=sys.stderr)
                return 2
            model = os.environ.get("REPORT_MODEL", DEFAULT_MODEL).strip() or DEFAULT_MODEL
            prose = _write_prose(blocks, season, key, model)

        output_path = Path(args.output) if args.output else Path(
            f"output/khl_report_{season.replace('/', '_')}.docx"
        )
        result = build_report(
            season=season, counter_id=counter_id,
            generated_at=datetime.now().strftime("%Y-%m-%d %H:%M"),
            blocks=blocks, prose=prose, output_path=output_path, charts=charts,
        )
        print(f"Отчёт сохранён: {result}")
        return 0

    # --- Режим API ---
    try:
        config = Config.from_env(require_llm=not args.no_llm)
    except RuntimeError as exc:
        print(f"Ошибка конфигурации: {exc}", file=sys.stderr)
        return 2

    if args.date_from and args.date_to:
        d1, d2 = args.date_from, args.date_to
        season = args.season or f"{d1[:4]}/{d2[:4]}"
    else:
        d1, d2, season = default_season()
        if args.season:
            season = args.season

    output_path = Path(args.output) if args.output else Path(
        f"output/khl_report_{season.replace('/', '_')}.docx"
    )

    client = MetrikaClient(config.token, config.counter_id)
    try:
        print(f"Собираю данные Метрики за {d1} — {d2}…")
        blocks = collect(client, d1, d2)
    except MetrikaError as exc:
        print(f"Ошибка Метрики: {exc}", file=sys.stderr)
        return 1

    prose = {}
    if not args.no_llm:
        prose = _write_prose(blocks, season, config.anthropic_key, config.model)

    result = build_report(
        season=season,
        counter_id=config.counter_id,
        generated_at=datetime.now().strftime("%Y-%m-%d %H:%M"),
        blocks=blocks,
        prose=prose,
        output_path=output_path,
        charts=_parse_charts(args.chart),
    )
    print(f"Отчёт сохранён: {result}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
