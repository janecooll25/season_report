from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

from .config import Config
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
        "--no-llm",
        action="store_true",
        help="Черновик без текста: только заголовки и таблицы (не нужен ANTHROPIC_API_KEY)",
    )
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(argv)

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

    prose: dict[str, str] = {}
    if not args.no_llm:
        from .report_writer import ReportWriter

        writer = ReportWriter(config.anthropic_key, config.model)
        for section in SECTIONS:
            block = blocks.get(section.id)
            if not block or not block.rows:
                continue
            print(f"Пишу раздел: {section.id}…")
            prose[section.id] = writer.write_section(section, block, season)

    result = build_report(
        season=season,
        counter_id=config.counter_id,
        generated_at=datetime.now().strftime("%Y-%m-%d %H:%M"),
        blocks=blocks,
        prose=prose,
        output_path=output_path,
    )
    print(f"Отчёт сохранён: {result}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
