from __future__ import annotations

from io import BytesIO

import openpyxl
import pytest

from test_ticket_metrics import _make_raw
from yandex_report.club_aggregate import (
    MAX_CLUBS, aggregate_clubs, build_aggregate_docx, compute_club_metrics,
)
from yandex_report.ticket_metrics import TicketError


def test_compute_club_metrics_basic():
    m = compute_club_metrics(_make_raw(), capacity=200)
    # 2 матча регулярки + 1 плей-офф = 172+190+265 = 627 проходов
    assert m["att_season"] == 627
    assert m["matches"] == 3
    assert m["income_total"] == 720000  # 200000+220000+300000
    assert m["online_share"] == pytest.approx(300 / 400)  # онлайн 300 из 400


def test_aggregate_has_four_params_with_grades():
    files = [("HC_A.xlsx", _make_raw()), ("HC_B.xlsx", _make_raw())]
    out = aggregate_clubs(files, capacity=200, season_label="2025/2026")
    ws = openpyxl.load_workbook(BytesIO(out)).active
    assert ws.title == "Средние по клубам"
    assert "Клубов в выборке: 2" in ws["A2"].value

    col_a = [ws.cell(r, 1).value for r in range(1, ws.max_row + 1)]
    # 4 параметра-секции
    assert any("Средняя цена билета (регулярный чемпионат)" in str(v) for v in col_a)
    assert any("Доля продаж билетов онлайн" in str(v) for v in col_a)
    assert any("Доля бесплатных билетов" in str(v) for v in col_a)
    assert any("Фактическое отклонение посещаемости" in str(v) for v in col_a)
    # клубы и градация присутствуют
    assert "HC A" in col_a and "HC B" in col_a
    grades = [ws.cell(r, 4).value for r in range(1, ws.max_row + 1)]
    assert any("показатели" in str(g) for g in grades)


def test_aggregate_docx_lists_params_and_places():
    from io import BytesIO as _B

    from docx import Document

    files = [("HC_A.xlsx", _make_raw()), ("HC_B.xlsx", _make_raw())]
    out = build_aggregate_docx(files, capacity=200, season_label="2025/2026")
    doc = Document(_B(out))
    headings = [p.text for p in doc.paragraphs if p.style.name.startswith("Heading")]
    assert any("Средняя цена билета" in h for h in headings)
    assert any("Доля продаж билетов онлайн" in h for h in headings)
    # таблицы с колонкой «Место» и клубами, ранжированными по местам
    assert doc.tables
    hdr = [c.text for c in doc.tables[0].rows[0].cells]
    assert hdr == ["Место", "Клуб", "Значение", "Градация"]
    places = [doc.tables[0].rows[i].cells[0].text for i in range(1, 3)]
    assert places == ["1", "2"]


def test_too_many_files_raises():
    files = [(f"c{i}.xlsx", _make_raw()) for i in range(MAX_CLUBS + 1)]
    with pytest.raises(TicketError):
        aggregate_clubs(files)


def test_empty_raises():
    with pytest.raises(TicketError):
        aggregate_clubs([])
