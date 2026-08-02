from __future__ import annotations

from io import BytesIO

import openpyxl
from docx import Document

from test_ticket_metrics import _make_raw
from yandex_report.club_report import build_club_report, list_clubs

# сезоны в блоке параметра: 18/19 … 25/26 (индексы 0..7)
SEASONS = ["18/19", "19/20", "20/21", "21/22", "22/23", "23/24", "24/25", "25/26"]
# (стартовый столбец, значения по клубам для 23/24, 24/25)
PARAMS = {
    2: {"Альфа": (2, 2), "Бета": (3, 3), "Гамма": (1, 1)},        # отклонение (градация)
    10: {"Альфа": (0.80, 0.90), "Бета": (0.50, 0.60), "Гамма": (0.30, 0.40)},  # онлайн
    18: {"Альфа": (800, 900), "Бета": (600, 650), "Гамма": (400, 420)},        # цена
    26: {"Альфа": (0.20, 0.15), "Бета": (0.30, 0.25), "Гамма": (0.40, 0.35)},  # безвозм.
}


def _make_monitoring() -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Чек-лист мониторинг"
    ws["A5"] = "Отклонение данных официальных протоколов матчей от данных выгрузок"
    ws["J5"] = "Доля розничных продаж билетов через интернет"
    ws["R5"] = "Средняя цена коммерческой реализации билетов в рамках Регулярного чемпионата"
    ws["Z5"] = "Процент билетов и абонементов, распространяемых на безвозмездной основе"
    for start in (2, 10, 18, 26):
        for i, s in enumerate(SEASONS):
            ws.cell(6, start + i, s)
    clubs = ["Альфа", "Бета", "Гамма"]
    for ri, club in enumerate(clubs, start=7):
        ws.cell(ri, 1, club)
        for start, vals in PARAMS.items():
            v2324, v2425 = vals[club]
            ws.cell(ri, start + 5, v2324)  # 23/24
            ws.cell(ri, start + 6, v2425)  # 24/25
    ws.cell(10, 1, "Характеристика")
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_list_clubs():
    assert list_clubs(_make_monitoring()) == ["Альфа", "Бета", "Гамма"]


def _rows_by_param(out):
    doc = Document(BytesIO(out))
    return {r.cells[1].text: (r.cells[2].text, r.cells[3].text)
            for r in doc.tables[0].rows[1:]}


def test_report_from_monitoring_has_place_and_dynamics():
    out = build_club_report(_make_monitoring(), "Альфа", season="24/25", prev_season="23/24")
    doc = Document(BytesIO(out))
    assert "«Альфа»" in doc.paragraphs[0].text
    assert "2024/2025" in doc.paragraphs[0].text
    rows = _rows_by_param(out)
    # цена: Альфа 900 — лучшая из трёх → 1 место, среднее (900+650+420)/3≈657
    price_char = next(v[0] for k, v in rows.items() if "Средняя цена" in k)
    assert "900 руб" in price_char and "1 месте" in price_char
    # онлайн: динамика 80% → 90%
    online_rec = next(v[1] for k, v in rows.items() if "интернет" in k)
    assert "с 80% до 90%" in online_rec


def test_deviation_grade_from_monitoring_not_percent():
    out = build_club_report(_make_monitoring(), "Бета", season="24/25", prev_season="23/24")
    rows = _rows_by_param(out)
    dev_char = next(v[0] for k, v in rows.items() if "Отклонение" in k)
    # градация 3 → «хорошие показатели», без «300%»
    assert "300%" not in dev_char
    assert "хорошие показатели" in dev_char.lower()


def test_current_season_from_ticket_file():
    # 25/26 в мониторинге пуст → значения берутся из билетного файла
    out = build_club_report(_make_monitoring(), "Альфа", season="25/26",
                            prev_season="24/25", ticket_bytes=_make_raw(), capacity=200)
    rows = _rows_by_param(out)
    online_char = next(v[0] for k, v in rows.items() if "интернет" in k)
    # _make_raw: онлайн 300 из 400 = 75%
    assert "75%" in online_char
