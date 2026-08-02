from __future__ import annotations

from io import BytesIO

import openpyxl
import pytest

from test_ticket_metrics import _make_raw
from yandex_report.club_aggregate import (
    MAX_CLUBS, aggregate_clubs, compute_club_metrics,
)
from yandex_report.ticket_metrics import TicketError


def test_compute_club_metrics_basic():
    m = compute_club_metrics(_make_raw(), capacity=200)
    # 2 матча регулярки + 1 плей-офф = 172+190+265 = 627 проходов
    assert m["att_season"] == 627
    assert m["matches"] == 3
    assert m["income_total"] == 720000  # 200000+220000+300000
    assert m["online_share"] == pytest.approx(300 / 400)  # онлайн 300 из 400


def test_aggregate_averages_across_clubs():
    # два «клуба» с одинаковыми данными → среднее = значение клуба
    files = [("HC_A.xlsx", _make_raw()), ("HC_B.xlsx", _make_raw())]
    out = aggregate_clubs(files, capacity=200, season_label="2025/2026")
    ws = openpyxl.load_workbook(BytesIO(out)).active
    assert ws.title == "Средние по клубам"
    assert "Клубов в выборке: 2" in ws["A2"].value
    # заголовки колонок клубов
    assert ws.cell(4, 4).value == "HC A"
    assert ws.cell(4, 5).value == "HC B"
    # строка «Посещаемость за сезон»: среднее = 627
    labels = {ws.cell(r, 1).value: r for r in range(5, ws.max_row + 1)}
    r = labels["Посещаемость за сезон (всего проходов)"]
    assert ws.cell(r, 2).value == 627


def test_too_many_files_raises():
    files = [(f"c{i}.xlsx", _make_raw()) for i in range(MAX_CLUBS + 1)]
    with pytest.raises(TicketError):
        aggregate_clubs(files)


def test_empty_raises():
    with pytest.raises(TicketError):
        aggregate_clubs([])
