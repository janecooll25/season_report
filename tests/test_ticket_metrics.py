from __future__ import annotations

from io import BytesIO

import openpyxl
import pytest

from yandex_report.ticket_metrics import (
    RZ, DH, AG, TicketError, _section_spans, build_calculations,
)


def _make_raw() -> bytes:
    """Мини-версия входного файла: 2 матча регулярки + 1 плей-офф."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    rz = wb.create_sheet(RZ)
    rz["A1"] = "Поматчевая реализация"
    rz.append(["№ игры", "Дата", "платн", "беспл", "аб.пл", "аб.беспл",
               "биз.пл", "биз.беспл", "ложи.пл", "ложи.беспл", "всего"])
    rz["A4"] = "РЕГУЛЯРНЫЙ ЧЕМПИОНАТ"
    rz.append([1, None, 100, 10, 50, 5, 2, 1, 3, 1, 172])   # row5
    rz.append([2, None, 120, 8, 52, 5, 0, 1, 3, 1, 190])    # row6
    rz["A7"] = "ПЛЕЙ-ОФФ"
    rz.append([3, None, 200, 4, 52, 5, 1, 0, 3, 0, 265])    # row8

    dh = wb.create_sheet(DH)
    dh["A1"] = "Поматчевая реализация"
    dh.append(["№ игры", "Дата", "всего", "платн", "абон", "биз", "ложи"])
    dh["A3"] = "РЕГУЛЯРНЫЙ ЧЕМПИОНАТ"
    dh.append([1, None, 200000, 120000, 60000, 5000, 15000])  # row4
    dh.append([2, None, 220000, 140000, 60000, 0, 15000])     # row5
    dh["A6"] = "ПЛЕЙ-ОФФ"
    dh.append([3, None, 300000, 220000, 60000, 3000, 15000])  # row7

    ag = wb.create_sheet(AG)
    ag["A9"] = "Канал продаж"
    ag["A10"] = "Онлайн"; ag["B10"] = 300
    ag["A11"] = "Оффлайн"; ag["B11"] = 100
    ag["A12"] = "Через агентов онлайн"; ag["B12"] = " - "
    ag["A13"] = "Через агентов оффлайн"; ag["B13"] = " - "

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_section_spans_detects_regular_and_playoff():
    wb = openpyxl.load_workbook(BytesIO(_make_raw()))
    reg, po = _section_spans(wb[RZ])
    assert reg == (5, 6)
    assert po == (8, 8)


def test_build_adds_calc_sheet_and_preserves_originals():
    out = build_calculations(_make_raw(), capacity=200, season_label="2025/2026")
    wb = openpyxl.load_workbook(BytesIO(out))
    assert wb.sheetnames[-1] == "Расчеты"
    assert RZ in wb.sheetnames and DH in wb.sheetnames


def test_calc_formulas_reference_correct_ranges():
    out = build_calculations(_make_raw(), capacity=200, season_label="2025/2026")
    ws = openpyxl.load_workbook(BytesIO(out))["Расчеты"]
    # Посещаемость за сезон = регулярка(K5:K6) + плей-офф(K8:K8)
    assert ws["C5"].value == "=(SUM('Реализованные билеты'!K5:K6)+SUM('Реализованные билеты'!K8:K8))"
    # Матчей регулярного чемпионата
    assert ws["C21"].value == "=COUNT('Реализованные билеты'!A5:A6)"
    # Вместимость арены — параметр в H2
    assert ws["H2"].value == 200


def test_protocol_column_and_deviation_metric():
    out = build_calculations(_make_raw(), capacity=200, season_label="2025/2026")
    wb = openpyxl.load_workbook(BytesIO(out))
    rz = wb[RZ]
    # заголовок столбца M и нули в строках матчей
    assert rz.cell(2, 13).value == "Посещаемость по протоколу"
    assert rz.cell(5, 13).value == 0
    assert rz.cell(6, 13).value == 0
    assert rz.cell(8, 13).value == 0
    calc = wb["Расчеты"]
    # абсолютное расхождение: факт(M) − заявлено(K)
    assert calc["B20"].value == "Расхождение: факт − заявлено (билеты)"
    assert calc["C20"].value == (
        "=((SUM('Реализованные билеты'!M5:M6)+SUM('Реализованные билеты'!M8:M8))"
        "-(SUM('Реализованные билеты'!K5:K6)+SUM('Реализованные билеты'!K8:K8)))"
    )
    # относительное расхождение, %
    assert calc["B26"].value == "Расхождение факта с заявленным, %"
    assert "IFERROR" in calc["C26"].value and calc["C26"].value.endswith(",0)")


def test_missing_required_sheet_raises():
    wb = openpyxl.Workbook()
    buf = BytesIO(); wb.save(buf)
    with pytest.raises(TicketError):
        build_calculations(buf.getvalue())
