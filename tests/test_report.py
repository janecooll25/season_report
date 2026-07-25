from __future__ import annotations

from pathlib import Path

from yandex_report.docx_builder import build_report
from yandex_report.report_data import DataBlock, collect, default_season
from yandex_report.structure import SECTIONS


def _cmp(rows):
    """Оборачивает строки в comparison-формат: metrics=[[A],[B]]."""
    return {"data": [
        {"dimensions": [{"name": n}], "metrics": [[cur], [prev]]}
        for n, cur, prev in rows
    ]}


class FakeClient:
    """Замоканный клиент Метрики: comparison-ответы (текущий/прошлый сезон)."""

    counter_id = "12345678"

    def query_comparison(self, metrics, dimensions, a1, a2, b1, b2,
                         *, limit=100, sort=None, filters=None):
        if dimensions is None:  # сводка: totals=[[A],[B]]
            return {"totals": [[1000, 800, 3000, 25.5, 130.0, 3.0],
                               [900, 750, 2800, 27.0, 125.0, 2.9]]}
        if "regionCountry" in dimensions:
            return _cmp([("Россия", 900, 850), ("Беларусь", 100, 120)])
        if "regionCity" in dimensions:
            return _cmp([("Москва", 500, 480)])
        if "ageInterval" in dimensions:
            return _cmp([("25-34", 400, 380)])
        if "gender" in dimensions:
            return _cmp([("Мужской", 700, 690)])
        if "interest" in dimensions:
            return _cmp([("Финансы", 362, 350)])
        if "lastSocialNetwork" in dimensions:
            return _cmp([("ВКонтакте", 120, 100)])
        if "lastTrafficSource" in dimensions:
            return _cmp([("Поисковые системы", 600, 610)])
        if "deviceCategory" in dimensions:
            return _cmp([("Смартфоны", 650, 600)])
        if "operatingSystem" in dimensions:
            return _cmp([("Android", 500, 470)])
        return {"data": []}


def test_collect_builds_all_blocks():
    blocks = collect(FakeClient(), "2024-07-01", "2025-06-30")
    for section in SECTIONS:
        assert section.data_key in blocks, section.data_key
    assert blocks["summary"].totals["визиты"] == 1000
    # визиты выросли с 900 до 1000 → +11.1%
    assert blocks["summary"].rows[0] == ["Визиты", "1000", "+11.1%"]
    # доля страны + изменение к прошлому сезону (900 vs 850)
    assert blocks["geo_countries"].rows[0][:3] == ["Россия", "900", "90.0%"]
    assert blocks["geo_countries"].rows[0][3] == "+5.9%"
    # интересы: аффинити + изменение в пунктах (362 vs 350)
    assert blocks["interests"].rows[0] == ["Финансы", "362", "+12"]


def test_metrika_client_passes_filters():
    """Реальный клиент должен принимать filters и класть его в params."""
    import requests

    from yandex_report.metrika_client import MetrikaClient

    captured = {}

    class DummyResp:
        status_code = 200

        def json(self):
            return {"data": []}

    class DummySession(requests.Session):
        def get(self, url, params=None, timeout=None):
            captured["params"] = params
            return DummyResp()

    client = MetrikaClient("tok", "6571768", session=DummySession())
    client.query("ym:s:visits", "ym:s:regionCity", "2024-07-01", "2025-06-30",
                 filters="ym:s:regionCountryName=='Россия'", sort="-ym:s:visits")
    assert captured["params"]["filters"] == "ym:s:regionCountryName=='Россия'"
    assert captured["params"]["sort"] == "-ym:s:visits"


def test_default_season_is_july_to_june():
    d1, d2, label = default_season()
    assert d1.endswith("-07-01")
    assert d2.endswith("-06-30")
    assert "/" in label


def test_datablock_to_facts_serializes():
    block = DataBlock(columns=["A", "B"], rows=[["x", "1"]], totals={"n": 5})
    facts = block.to_facts()
    assert "A | B" in facts
    assert "x | 1" in facts
    assert "ИТОГО" in facts


def test_build_report_produces_docx(tmp_path: Path):
    blocks = collect(FakeClient(), "2024-07-01", "2025-06-30")
    prose = {"summary": "Тестовый аналитический абзац."}
    out = build_report(
        season="2024/2025",
        counter_id="12345678",
        generated_at="2026-07-25 12:00",
        blocks=blocks,
        prose=prose,
        output_path=tmp_path / "out.docx",
    )
    assert out.exists() and out.stat().st_size > 0
    from docx import Document
    doc = Document(str(out))
    headings = [p.text for p in doc.paragraphs if p.style.name.startswith("Heading")]
    assert any("Совокупные показатели" in h for h in headings)
