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
    ag["A2"] = "Название агента"
    ag["D2"] = "Размер агентской комиссии, %"
    ag["A4"] = " - "; ag["D4"] = " - "   # текстовый плейсхолдер комиссии
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


def test_agent_commission_is_average():
    out = build_calculations(_make_raw(), capacity=200, season_label="2025/2026")
    calc = openpyxl.load_workbook(BytesIO(out))["Расчеты"]
    assert calc["B52"].value == "Агентская комиссия (средняя по агентам)"
    # средняя по столбцу комиссии агентов, /100, с запасом на пустую таблицу
    assert calc["C52"].value == (
        "=IFERROR(AVERAGE('Агенты и онлайн продажи'!D3:D8)/100,0)"
    )


def test_numeric_text_is_coerced_to_numbers():
    """Числа, записанные текстом, должны стать настоящими числами (иначе SUM=0)."""
    wb = openpyxl.load_workbook(BytesIO(_make_raw()))
    rz = wb[RZ]
    # портим данные: делаем часть чисел текстом (как в реальных выгрузках)
    rz["C5"] = "5 439"; rz["K5"] = "8 799"; rz["C6"] = "120"
    buf = BytesIO(); wb.save(buf)

    out = build_calculations(buf.getvalue(), capacity=200)
    r = openpyxl.load_workbook(BytesIO(out))[RZ]
    assert r["C5"].value == 5439 and isinstance(r["C5"].value, (int, float))
    assert r["K5"].value == 8799
    assert r["C6"].value == 120
    # даты и номера игр не пострадали
    assert isinstance(r["A5"].value, int)


def test_currency_conversion_adds_rate_cell_and_multiplies():
    out = build_calculations(_make_raw(), capacity=200, byn_to_rub=27.5)
    ws = openpyxl.load_workbook(BytesIO(out))["Расчеты"]
    assert ws["H3"].value == 27.5
    assert "BYN" in str(ws["G3"].value)
    # денежные метрики умножаются на курс из H3
    assert ws["C34"].value.endswith("*$H$3")
    assert "руб. РФ" in ws["B32"].value


def test_missing_required_sheet_raises():
    wb = openpyxl.Workbook()
    buf = BytesIO(); wb.save(buf)
    with pytest.raises(TicketError):
        build_calculations(buf.getvalue())
