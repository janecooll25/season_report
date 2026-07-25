from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

from .config import Config
from .docx_builder import render_docx
from .metrika_client import MetrikaClient, MetrikaError
from .report_data import build_report, default_period

DEFAULT_TEMPLATE = Path("templates/report_template.docx")


def _parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="yandex-report",
        description="Генерация docx-отчёта из Яндекс.Метрики по шаблону.",
    )
    parser.add_argument("--date-from", help="Дата начала YYYY-MM-DD (по умолчанию 30 дней назад)")
    parser.add_argument("--date-to", help="Дата окончания YYYY-MM-DD (по умолчанию сегодня)")
    parser.add_argument("--days", type=int, default=30, help="Период в днях, если даты не заданы")
    parser.add_argument(
        "--template",
        default=str(DEFAULT_TEMPLATE),
        help="Путь к docx-шаблону (по умолчанию templates/report_template.docx)",
    )
    parser.add_argument(
        "--output",
        help="Путь к выходному docx (по умолчанию output/report_<даты>.docx)",
    )
    parser.add_argument("--title", help="Заголовок отчёта (переопределяет .env)")
    parser.add_argument("--top", type=int, default=10, help="Сколько строк в топах")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(argv)

    try:
        config = Config.from_env()
    except RuntimeError as exc:
        print(f"Ошибка конфигурации: {exc}", file=sys.stderr)
        return 2

    if args.date_from and args.date_to:
        date_from, date_to = args.date_from, args.date_to
    else:
        date_from, date_to = default_period(args.days)

    title = args.title or config.report_title
    output_path = Path(args.output) if args.output else Path(
        f"output/report_{date_from}_{date_to}.docx"
    )

    client = MetrikaClient(config.token, config.counter_id)
    try:
        report = build_report(
            client,
            title=title,
            date_from=date_from,
            date_to=date_to,
            generated_at=datetime.now().strftime("%Y-%m-%d %H:%M"),
            top_limit=args.top,
        )
    except MetrikaError as exc:
        print(f"Ошибка Метрики: {exc}", file=sys.stderr)
        return 1

    result = render_docx(args.template, report.as_context(), output_path)
    print(f"Отчёт сохранён: {result}")
    print(
        f"Визитов: {report.visits}, пользователей: {report.users}, "
        f"просмотров: {report.pageviews}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
