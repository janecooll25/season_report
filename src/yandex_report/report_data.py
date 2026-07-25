from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, timedelta
from typing import Any

from .metrika_client import MetrikaClient

# Метрики верхнего уровня (сводка)
_SUMMARY_METRICS = (
    "ym:s:visits,ym:s:users,ym:s:pageviews,"
    "ym:s:bounceRate,ym:s:avgVisitDurationSeconds,ym:s:pageDepth"
)


@dataclass
class Row:
    name: str
    visits: int
    users: int
    share: float  # доля визитов, %


@dataclass
class ReportData:
    title: str
    counter_id: str
    date_from: str
    date_to: str
    generated_at: str
    visits: int = 0
    users: int = 0
    pageviews: int = 0
    bounce_rate: float = 0.0
    avg_duration_sec: float = 0.0
    page_depth: float = 0.0
    sources: list[Row] = field(default_factory=list)
    top_pages: list[Row] = field(default_factory=list)
    devices: list[Row] = field(default_factory=list)

    @property
    def avg_duration_human(self) -> str:
        total = int(round(self.avg_duration_sec))
        return f"{total // 60} мин {total % 60} сек"

    def as_context(self) -> dict[str, Any]:
        """Плоский словарь для подстановки в docx-шаблон (docxtpl)."""
        return {
            "title": self.title,
            "counter_id": self.counter_id,
            "date_from": self.date_from,
            "date_to": self.date_to,
            "generated_at": self.generated_at,
            "visits": self.visits,
            "users": self.users,
            "pageviews": self.pageviews,
            "bounce_rate": round(self.bounce_rate, 1),
            "avg_duration": self.avg_duration_human,
            "page_depth": round(self.page_depth, 2),
            "sources": [r.__dict__ for r in self.sources],
            "top_pages": [r.__dict__ for r in self.top_pages],
            "devices": [r.__dict__ for r in self.devices],
        }


def default_period(days: int = 30) -> tuple[str, str]:
    today = date.today()
    start = today - timedelta(days=days)
    return start.isoformat(), today.isoformat()


def _to_rows(payload: dict[str, Any], total_visits: int) -> list[Row]:
    rows: list[Row] = []
    for item in payload.get("data", []):
        dims = item.get("dimensions", [])
        name = dims[0].get("name") if dims else "—"
        metrics = item.get("metrics", [0, 0])
        visits = int(metrics[0] or 0)
        users = int(metrics[1] or 0) if len(metrics) > 1 else 0
        share = round(visits / total_visits * 100, 1) if total_visits else 0.0
        rows.append(Row(name=name or "не задано", visits=visits, users=users, share=share))
    return rows


def build_report(
    client: MetrikaClient,
    *,
    title: str,
    date_from: str,
    date_to: str,
    generated_at: str,
    top_limit: int = 10,
) -> ReportData:
    report = ReportData(
        title=title,
        counter_id=client.counter_id,
        date_from=date_from,
        date_to=date_to,
        generated_at=generated_at,
    )

    summary = client.query(_SUMMARY_METRICS, None, date_from, date_to, limit=1)
    totals = summary.get("totals", [0, 0, 0, 0, 0, 0])
    report.visits = int(totals[0] or 0)
    report.users = int(totals[1] or 0)
    report.pageviews = int(totals[2] or 0)
    report.bounce_rate = float(totals[3] or 0.0)
    report.avg_duration_sec = float(totals[4] or 0.0)
    report.page_depth = float(totals[5] or 0.0)

    sources = client.query(
        "ym:s:visits,ym:s:users",
        "ym:s:lastTrafficSource",
        date_from,
        date_to,
        limit=top_limit,
        sort="-ym:s:visits",
    )
    report.sources = _to_rows(sources, report.visits)

    pages = client.query(
        "ym:s:visits,ym:s:users",
        "ym:s:startURLPathFull",
        date_from,
        date_to,
        limit=top_limit,
        sort="-ym:s:visits",
    )
    report.top_pages = _to_rows(pages, report.visits)

    devices = client.query(
        "ym:s:visits,ym:s:users",
        "ym:s:deviceCategory",
        date_from,
        date_to,
        limit=top_limit,
        sort="-ym:s:visits",
    )
    report.devices = _to_rows(devices, report.visits)

    return report
