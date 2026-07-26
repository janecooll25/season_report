from __future__ import annotations

from yandex_report.charts import render_chart
from yandex_report.report_data import DataBlock


def _block():
    return DataBlock(
        columns=["Страна", "Визиты", "Доля", "Изменение"],
        rows=[
            ["Россия", "73526805", "77.4%", "+9.7%"],
            ["Беларусь", "10332715", "10.9%", "+65.6%"],
            ["Казахстан", "2649888", "2.8%", "+57.7%"],
        ],
    )


def test_render_each_type_returns_png():
    for t in ("bar", "barh", "pie", "line"):
        png = render_chart(_block(), t, "Тест")
        assert png and png[:8] == b"\x89PNG\r\n\x1a\n", t


def test_unknown_type_returns_none():
    assert render_chart(_block(), "donut", "Тест") is None


def test_empty_data_returns_none():
    empty = DataBlock(columns=["A", "B"], rows=[])
    assert render_chart(empty, "bar", "Тест") is None


def test_build_report_embeds_chart(tmp_path):
    from docx import Document

    from yandex_report.docx_builder import build_report

    out = build_report(
        season="2025/2026", counter_id="1", generated_at="now",
        blocks={"geo_countries": _block()},
        prose={}, output_path=tmp_path / "r.docx",
        charts={"geo_countries": "pie"},
    )
    assert len(Document(str(out)).inline_shapes) == 1
