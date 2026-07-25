from __future__ import annotations

from pathlib import Path

from yandex_report.docx_builder import render_docx
from yandex_report.report_data import build_report, default_period


class FakeClient:
    """Замоканный клиент Метрики: отдаёт фиксированные ответы по dimensions."""

    counter_id = "12345678"

    def query(self, metrics, dimensions, date1, date2, *, limit=100, sort=None):
        if dimensions is None:
            return {"totals": [1000, 800, 3000, 25.5, 130.0, 3.0]}
        if "lastTrafficSource" in dimensions:
            return {
                "data": [
                    {"dimensions": [{"name": "Переходы из поисковых систем"}], "metrics": [600, 500]},
                    {"dimensions": [{"name": "Прямые заходы"}], "metrics": [400, 300]},
                ]
            }
        if "startURLPathFull" in dimensions:
            return {"data": [{"dimensions": [{"name": "/"}], "metrics": [700, 550]}]}
        if "deviceCategory" in dimensions:
            return {"data": [{"dimensions": [{"name": "Смартфоны"}], "metrics": [650, 520]}]}
        return {"data": []}


def test_build_report_computes_totals_and_shares():
    report = build_report(
        FakeClient(),
        title="Тест",
        date_from="2026-01-01",
        date_to="2026-01-31",
        generated_at="2026-02-01 10:00",
    )
    assert report.visits == 1000
    assert report.users == 800
    assert report.bounce_rate == 25.5
    assert report.avg_duration_human == "2 мин 10 сек"
    assert report.sources[0].name == "Переходы из поисковых систем"
    assert report.sources[0].share == 60.0
    assert report.sources[1].share == 40.0


def test_default_period_orders_dates():
    date_from, date_to = default_period(7)
    assert date_from < date_to


def test_render_docx_produces_file(tmp_path: Path):
    report = build_report(
        FakeClient(),
        title="Тест",
        date_from="2026-01-01",
        date_to="2026-01-31",
        generated_at="2026-02-01 10:00",
    )
    out = render_docx(
        "templates/report_template.docx",
        report.as_context(),
        tmp_path / "out.docx",
    )
    assert out.exists()
    assert out.stat().st_size > 0
