from __future__ import annotations

from pathlib import Path

from yandex_report.docx_builder import build_report
from yandex_report.report_data import DataBlock, collect, default_season
from yandex_report.structure import SECTIONS


class FakeClient:
    """Замоканный клиент Метрики: фиксированные ответы по dimensions."""

    counter_id = "12345678"

    def query(self, metrics, dimensions, d1, d2, *, limit=100, sort=None, filters=None):
        if dimensions is None:
            return {"totals": [1000, 800, 3000, 25.5, 130.0, 3.0]}
        if "regionCountry" in dimensions:
            return {"data": [
                {"dimensions": [{"name": "Россия"}], "metrics": [900]},
                {"dimensions": [{"name": "Беларусь"}], "metrics": [100]},
            ]}
        if "regionCity" in dimensions:
            return {"data": [{"dimensions": [{"name": "Москва"}], "metrics": [500]}]}
        if "ageInterval" in dimensions:
            return {"data": [
                {"dimensions": [{"name": "25-34"}], "metrics": [400, 20.0, 3.5, 150.0]},
            ]}
        if "gender" in dimensions:
            return {"data": [{"dimensions": [{"name": "Мужской"}], "metrics": [700]}]}
        if "interest" in dimensions:
            return {"data": [{"dimensions": [{"name": "Финансы"}], "metrics": [362]}]}
        if "lastSocialNetwork" in dimensions:
            return {"data": [{"dimensions": [{"name": "ВКонтакте"}], "metrics": [120]}]}
        if "lastTrafficSource" in dimensions:
            return {"data": [{"dimensions": [{"name": "Поисковые системы"}], "metrics": [600]}]}
        if "deviceCategory" in dimensions:
            return {"data": [{"dimensions": [{"name": "Смартфоны"}], "metrics": [650]}]}
        if "operatingSystem" in dimensions:
            return {"data": [{"dimensions": [{"name": "Android"}], "metrics": [500]}]}
        return {"data": []}


def test_collect_builds_all_blocks():
    blocks = collect(FakeClient(), "2024-07-01", "2025-06-30")
    for section in SECTIONS:
        assert section.data_key in blocks, section.data_key
    assert blocks["summary"].totals["визиты"] == 1000
    # доля страны считается от всех визитов
    assert blocks["geo_countries"].rows[0][0] == "Россия"
    assert blocks["geo_countries"].rows[0][2] == "90.0%"
    assert blocks["interests"].rows[0] == ["Финансы", "362"]


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
