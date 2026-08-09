"""Расчёт метрик билетной программы: из «сырого» xlsx собирает лист «Расчеты».

Повторяет шаблонный лист «Расчеты» формулами со ссылками на исходные листы
(диапазоны строк регулярки/плей-офф определяются динамически), поэтому лист
пересчитывается в Excel при изменении данных.
"""
from __future__ import annotations

import re
from io import BytesIO

import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

RZ = "Реализованные билеты"
DH = "Доход от реализации билетов"
AG = "Агенты и онлайн продажи"
CALC = "Расчеты"


class TicketError(RuntimeError):
    pass


def _is_int(v) -> bool:
    return isinstance(v, int) and not isinstance(v, bool)


_NUMISH = re.compile(r"^-?[\d\s .,]*\d[\d\s .,]*$")


def _to_number(s: str):
    """Число из текста ('1 234,5', '8 965', '12%') → float/int, иначе None."""
    raw = s.strip().replace("%", "").strip()
    if not _NUMISH.match(raw):
        return None
    t = raw.replace(" ", "").replace(" ", "")
    if t.count(",") == 1 and t.count(".") == 0:  # запятая-десятичная
        t = t.replace(",", ".")
    else:  # запятые/пробелы — разделители тысяч
        t = t.replace(",", "")
    try:
        f = float(t)
    except ValueError:
        return None
    return int(f) if f.is_integer() else f


def _coerce_numeric_text(ws) -> None:
    """Приводит числа, записанные текстом, к настоящим числам (иначе SUM = 0)."""
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            v = cell.value
            if isinstance(v, str):
                num = _to_number(v)
                if num is not None:
                    cell.value = num


def _is_data_row(row) -> bool:
    """Строка матча: есть № игры в гр. A ЛИБО числовые данные в гр. C+.

    В части выгрузок матчи плей-офф идут без порядкового номера (гр. A пустая),
    но с датой и числами реализации — такие строки тоже нужно учитывать.
    Текст в гр. A (маркеры секций, сноски «*…», заголовки) строкой данных не
    считается.
    """
    a = row[0] if row else None
    if _is_int(a):
        return True
    if isinstance(a, str) and a.strip():
        return False
    for v in row[2:]:  # № игры нет — матч определяем по числам в гр. C и далее
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            return True
    return False


def _section_spans(ws) -> tuple[tuple[int, int] | None, tuple[int, int] | None]:
    """Диапазоны строк данных: (регулярка, плей-офф) по номерам строк листа."""
    reg_marker = po_marker = None
    nums: list[int] = []
    for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
        a = row[0] if row else None
        if isinstance(a, str):
            u = a.strip().upper()
            # только короткие строки-заголовки секций; берём ПЕРВОЕ вхождение —
            # чтобы не спутать с примечаниями/акциями вроде «…(только Плей-офф)».
            if reg_marker is None and "РЕГУЛЯРН" in u and len(u) <= 25:
                reg_marker = i
                continue
            if po_marker is None and "ПЛЕЙ" in u and "ОФФ" in u and len(u) <= 25:
                po_marker = i
                continue
        if _is_data_row(row):
            nums.append(i)
    if reg_marker is None:
        raise TicketError(f"На листе «{ws.title}» не найдена секция «Регулярный чемпионат».")

    end = po_marker if po_marker else 10 ** 9
    reg = [i for i in nums if reg_marker < i < end]
    po = [i for i in nums if po_marker and i > po_marker]
    reg_span = (reg[0], reg[-1]) if reg else None
    po_span = (po[0], po[-1]) if po else None
    if not reg_span:
        raise TicketError(f"На листе «{ws.title}» нет строк матчей регулярного чемпионата.")
    return reg_span, po_span


def _sum(sheet: str, col: str, span: tuple[int, int] | None) -> str:
    if not span:
        return "0"
    return f"SUM('{sheet}'!{col}{span[0]}:{col}{span[1]})"


def _both(sheet: str, col: str, reg, po) -> str:
    return f"({_sum(sheet, col, reg)}+{_sum(sheet, col, po)})"


def _maxr(sheet: str, col: str, span) -> str:
    if not span:
        return "0"
    return f"MAX('{sheet}'!{col}{span[0]}:{col}{span[1]})"


def _count(sheet: str, span) -> str:
    # Считаем матчи по столбцу K («всего») — он числовой в каждой строке матча,
    # тогда как № игры (гр. A) у плей-офф в части выгрузок пустой.
    if not span:
        return "0"
    return f"COUNT('{sheet}'!K{span[0]}:K{span[1]})"


def _agent_rows(ws) -> dict[str, int]:
    """Находит строки каналов продаж и первой строки агента на листе «Агенты»."""
    rows: dict[str, int] = {}
    for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
        a = row[0] if row else None
        if not isinstance(a, str):
            continue
        u = a.strip().lower()
        if u == "онлайн":
            rows["online"] = i
        elif u == "оффлайн":
            rows["offline"] = i
        elif "через агентов онлайн" in u:
            rows["agent_online"] = i
        elif "через агентов оффлайн" in u:
            rows["agent_offline"] = i
        elif u == "название агента":
            rows["agent_first"] = i + 1
    return rows


def _num_or_zero(cell: str) -> str:
    return f"IF(ISNUMBER('{AG}'!{cell}),'{AG}'!{cell},0)"


def _agent_commission_range(ws) -> tuple[str, int, int] | None:
    """Диапазон столбца «Размер агентской комиссии, %» по строкам агентов.

    Возвращает (буква столбца, первая строка, последняя строка) для средней
    комиссии по агентам. AVERAGE сам игнорирует текст (« - ») и пустые ячейки.
    """
    header = channels = None
    for i in range(1, ws.max_row + 1):
        v = ws.cell(i, 1).value
        if isinstance(v, str):
            u = v.strip().lower()
            if u == "название агента":
                header = i
            elif channels is None and ("через каналы продаж" in u or u == "канал продаж"):
                channels = i
    if header is None:
        return None

    # Ищем именно столбец «Размер агентской комиссии, %», а не столбец дохода,
    # в заголовке которого тоже может встречаться слово «комиссия»
    # (напр. «Доход клуба … (в т.ч. агентская комиссия)»).
    comm_col = None
    for c in range(1, ws.max_column + 1):
        hv = ws.cell(header, c).value
        if not isinstance(hv, str):
            continue
        h = hv.lower()
        if "комисси" in h and "доход" not in h and ("размер" in h or "%" in h):
            comm_col = c
            break
    if comm_col is None:
        comm_col = 4  # D по умолчанию

    end = channels if channels else ws.max_row + 1
    # Первая строка данных — после объединённого заголовка.
    first = None
    for r in range(header + 1, end):
        if not isinstance(ws.cell(r, comm_col), MergedCell):
            first = r
            break
    if first is None:
        return None
    last = max(first, end - 1)
    return (get_column_letter(comm_col), first, last)


def _default_capacity(data: BytesIO) -> int:
    """Оценка вместимости = макс. посещаемость за матч (если не задана явно)."""
    wb = openpyxl.load_workbook(data, data_only=True, read_only=True)
    ws = wb[RZ]
    best = 0
    for row in ws.iter_rows(values_only=True):
        if row and _is_int(row[0]) and len(row) > 10 and isinstance(row[10], (int, float)):
            best = max(best, int(row[10]))
    wb.close()
    return best or 9000


# ── стили ──────────────────────────────────────────────────────────────────
_FONT = "Arial"
_H = Font(name=_FONT, bold=True, size=13)
_SEC = Font(name=_FONT, bold=True, size=11, color="FFFFFF")
_HDR = Font(name=_FONT, bold=True)
_REG = Font(name=_FONT)
_INPUT = Font(name=_FONT, bold=True, color="0000FF")
_SECFILL = PatternFill("solid", fgColor="0B5CAD")
_HDRFILL = PatternFill("solid", fgColor="E2E5EA")


PROTOCOL_HEADER = "Посещаемость по протоколу"
PROTOCOL_COL = "M"          # столбец для протокольной посещаемости
PROTOCOL_COL_IDX = 13       # M


def _prep_protocol_column(rz, reg, po) -> None:
    """Столбец M: заголовок «Посещаемость по протоколу», 0 в строках матчей."""
    # Заголовок — в строке с «№ игры».
    for i in range(1, rz.max_row + 1):
        a = rz.cell(i, 1).value
        if isinstance(a, str) and a.strip() == "№ игры":
            rz.cell(i, PROTOCOL_COL_IDX).value = PROTOCOL_HEADER
            break
    for span in (reg, po):
        if not span:
            continue
        for r in range(span[0], span[1] + 1):
            if rz.cell(r, PROTOCOL_COL_IDX).value in (None, ""):
                rz.cell(r, PROTOCOL_COL_IDX).value = 0


def _build_calc_sheet(wb, season_label: str, capacity_value: int,
                      fx: tuple[str, float | None] | None = None) -> None:
    rz = wb[RZ]
    dh = wb[DH]
    if AG in wb.sheetnames:
        ag_rows = _agent_rows(wb[AG])
        commission = _agent_commission_range(wb[AG])
    else:
        ag_rows = {}
        commission = None

    rz_reg, rz_po = _section_spans(rz)
    dh_reg, dh_po = _section_spans(dh)
    _prep_protocol_column(rz, rz_reg, rz_po)

    if CALC in wb.sheetnames:
        del wb[CALC]
    ws = wb.create_sheet(CALC)
    ws.column_dimensions["B"].width = 42
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 46

    MONEY = "#,##0"
    NUM = "#,##0"
    PCT = "0.0%"
    RUB = '#,##0" ₽"'

    # H2 — изменяемый параметр вместимости арены.
    ws["G2"] = "Вместимость арены, чел.:"
    ws["G2"].font = _HDR
    ws["H2"] = capacity_value
    ws["H2"].font = _INPUT
    ws["H2"].number_format = NUM
    CAP = "$H$2"

    # H3 — курс иностранной валюты к RUB (если суммы не в рублях РФ).
    # Редактируемый; значение можно оставить пустым — клуб впишет курс сам.
    if fx:
        code, rate_val = fx
        ws["G3"] = f"Курс {code}→RUB:"
        ws["G3"].font = _HDR
        if rate_val is not None:
            ws["H3"] = rate_val
        ws["H3"].font = _INPUT
        ws["H3"].number_format = "0.0000"
        RATE: str | None = "$H$3"
        money_unit = "руб. РФ"
    else:
        RATE = None
        money_unit = "руб."

    def rub(formula: str) -> str:
        return f"({formula})*{RATE}" if RATE else formula

    ws["B1"] = f"Расчёт метрик билетной программы — сезон {season_label}"
    ws["B1"].font = _H

    def section(r: int, text: str) -> None:
        c = ws.cell(r, 2, text)
        c.font = _SEC
        for col in range(2, 6):
            ws.cell(r, col).fill = _SECFILL

    def header(r: int) -> None:
        for col, t in zip(range(2, 6), ("Показатель", "Значение", "Ед.", "Формула / комментарий")):
            c = ws.cell(r, col, t)
            c.font = _HDR
            c.fill = _HDRFILL

    def metric(r, label, formula, unit, comment="", fmt=NUM):
        ws.cell(r, 2, label).font = _REG
        c = ws.cell(r, 3, f"={formula}")
        c.font = _REG
        c.number_format = fmt
        ws.cell(r, 4, unit).font = _REG
        if comment:
            ws.cell(r, 5, comment).font = Font(name=_FONT, size=9, italic=True, color="6B7280")

    # ── 1. Реализация билетов ────────────────────────────────────────────
    section(3, "1. Реализация билетов, шт.")
    header(4)
    metric(5, "Посещаемость за сезон (всего проходов)",
           _both(RZ, "K", rz_reg, rz_po), "чел.", "гр. K — сумма разовых, абонементов, лож")
    metric(6, "   регулярный чемпионат", _sum(RZ, "K", rz_reg), "чел.")
    metric(7, "   плей-офф", _sum(RZ, "K", rz_po), "чел.")
    metric(8, "Платные билеты на матч (разовые)", _both(RZ, "C", rz_reg, rz_po), "шт.", "гр. C")
    metric(9, "Бесплатные билеты на матч (разовые)", _both(RZ, "D", rz_reg, rz_po), "шт.", "гр. D")
    metric(10, "Платных абонементов/пакетов (продано)",
           f"MAX({_maxr(RZ, 'E', rz_reg)},{_maxr(RZ, 'E', rz_po)})", "шт.", "MAX действующих, гр. E")
    metric(11, "Бесплатных абонементов/пакетов (продано)",
           f"MAX({_maxr(RZ, 'F', rz_reg)},{_maxr(RZ, 'F', rz_po)})", "шт.", "MAX по гр. F")
    metric(12, "   абонемент-посещения за сезон (справочно)",
           f"{_both(RZ, 'E', rz_reg, rz_po)}+{_both(RZ, 'F', rz_reg, rz_po)}", "чел.")
    metric(13, "Платные билеты в бизнес-клубы/рестораны", _both(RZ, "G", rz_reg, rz_po), "шт.", "гр. G")
    metric(14, "Бесплатные билеты в бизнес-клубы/рестораны", _both(RZ, "H", rz_reg, rz_po), "шт.", "гр. H")
    metric(15, "Платные места в ложах", _both(RZ, "I", rz_reg, rz_po), "шт.", "гр. I")
    metric(16, "Бесплатные места в ложах", _both(RZ, "J", rz_reg, rz_po), "шт.", "гр. J")

    # Посещаемость по протоколу и отклонение (столбец L — заполняется клубом).
    metric(17, "Посещаемость по протоколу (за сезон)",
           _both(RZ, PROTOCOL_COL, rz_reg, rz_po),
           "чел.", "гр. M — из официального протокола матча")

    free_no_abon =(f"({_both(RZ, 'D', rz_reg, rz_po)}+{_both(RZ, 'H', rz_reg, rz_po)}"
                    f"+{_both(RZ, 'J', rz_reg, rz_po)})")
    paid_base = (f"({_both(RZ, 'K', rz_reg, rz_po)}-{_both(RZ, 'E', rz_reg, rz_po)}"
                 f"-{_both(RZ, 'F', rz_reg, rz_po)})")
    metric(18, "Всего бесплатных билетов (без абонементов)", free_no_abon, "шт.")
    metric(19, "Доля бесплатных билетов", f"IFERROR({free_no_abon}/{paid_base},0)", "%",
           "бесплатные / (проходы − абонементы)", fmt=PCT)
    metric(20, "Расхождение: факт − заявлено (билеты)",
           f"({_both(RZ, PROTOCOL_COL, rz_reg, rz_po)}-{_both(RZ, 'K', rz_reg, rz_po)})", "чел.",
           "посещаемость по протоколу − всего реализовано билетов")

    metric(21, "Матчей регулярного чемпионата", _count(RZ, rz_reg), "шт.")
    metric(22, "Матчей плей-офф", _count(RZ, rz_po), "шт.")
    metric(23, "Средняя посещаемость (регулярка)",
           f"IFERROR({_sum(RZ, 'K', rz_reg)}/{_count(RZ, rz_reg)},0)", "зрит./матч")
    metric(24, "Средняя посещаемость (плей-офф)",
           f"IFERROR({_sum(RZ, 'K', rz_po)}/{_count(RZ, rz_po)},0)", "зрит./матч")
    metric(25, "Средняя посещаемость (сезон)",
           f"IFERROR({_both(RZ, 'K', rz_reg, rz_po)}/({_count(RZ, rz_reg)}+{_count(RZ, rz_po)}),0)",
           "зрит./матч")
    metric(26, "Расхождение факта с заявленным, %",
           f"IFERROR(({_both(RZ, PROTOCOL_COL, rz_reg, rz_po)}-{_both(RZ, 'K', rz_reg, rz_po)})"
           f"/{_both(RZ, 'K', rz_reg, rz_po)},0)", "%",
           "(протокол − всего билетов) / всего билетов", fmt=PCT)

    metric(27, "Вместимость арены", CAP, "мест", "изменяемый параметр в ячейке H2")
    metric(28, "Суммарная посещаемость (регулярка)", _sum(RZ, "K", rz_reg), "зрит.")
    metric(29, "Теоретическая ёмкость (регулярка)", f"{CAP}*{_count(RZ, rz_reg)}", "мест")
    metric(30, "Заполняемость арены (регулярка)",
           f"IFERROR(({_sum(RZ, 'K', rz_reg)}/{_count(RZ, rz_reg)})/{CAP},0)", "%",
           "средняя посещаемость / вместимость", fmt=PCT)

    # ── 2. Доход ─────────────────────────────────────────────────────────
    conv = " (пересчёт по курсу H3)" if RATE else ""
    section(32, f"2. Доход от реализации билетов, {money_unit}{conv}")
    header(33)
    metric(34, "Доход — всего за сезон", rub(_both(DH, "C", dh_reg, dh_po)), money_unit,
           "гр. C" + conv, fmt=RUB)
    metric(35, "   регулярный чемпионат", rub(_sum(DH, "C", dh_reg)), money_unit, fmt=RUB)
    metric(36, "   плей-офф", rub(_sum(DH, "C", dh_po)), money_unit, fmt=RUB)
    metric(37, "Доход от платных билетов", rub(_both(DH, "D", dh_reg, dh_po)), money_unit,
           "гр. D" + conv, fmt=RUB)
    metric(38, "Доход от абонементов/пакетов", rub(_both(DH, "E", dh_reg, dh_po)), money_unit,
           "гр. E" + conv, fmt=RUB)
    metric(39, "Стоимость абонемента (за сезон)",
           rub(f"IFERROR({_both(DH, 'E', dh_reg, dh_po)}/MAX({_maxr(RZ, 'E', rz_reg)},{_maxr(RZ, 'E', rz_po)}),0)"),
           money_unit, "доход абон. / число владельцев" + conv, fmt=RUB)
    metric(40, "Доход от бизнес-клубов/ресторанов", rub(_both(DH, "F", dh_reg, dh_po)), money_unit,
           "гр. F" + conv, fmt=RUB)
    metric(41, "Доход от лож", rub(_both(DH, "G", dh_reg, dh_po)), money_unit, "гр. G" + conv, fmt=RUB)
    metric(43, "Средняя цена платного билета (регулярка)",
           rub(f"IFERROR({_sum(DH, 'D', dh_reg)}/{_sum(RZ, 'C', rz_reg)},0)"), money_unit,
           "доход платн. / платные билеты" + conv, fmt=RUB)
    metric(44, "Средняя цена платного билета (сезон)",
           rub(f"IFERROR({_both(DH, 'D', dh_reg, dh_po)}/{_both(RZ, 'C', rz_reg, rz_po)},0)"),
           money_unit, conv.strip(), fmt=RUB)

    # ── 3. Каналы продаж ─────────────────────────────────────────────────
    section(46, "3. Каналы продаж")
    header(47)
    on = _num_or_zero(f"B{ag_rows['online']}") if "online" in ag_rows else "0"
    on_ag = _num_or_zero(f"B{ag_rows['agent_online']}") if "agent_online" in ag_rows else "0"
    off = _num_or_zero(f"B{ag_rows['offline']}") if "offline" in ag_rows else "0"
    off_ag = _num_or_zero(f"B{ag_rows['agent_offline']}") if "agent_offline" in ag_rows else "0"
    online = f"({on}+{on_ag})"
    offline = f"({off}+{off_ag})"
    total_ch = f"({online}+{offline})"
    metric(48, "Продажи онлайн", online, "шт.")
    metric(49, "Продажи оффлайн", offline, "шт.")
    metric(50, "Всего по каналам", total_ch, "шт.")
    metric(51, "Доля онлайн-продаж", f'IF({total_ch}=0,"нет данных",{online}/{total_ch})', "%",
           fmt=PCT)
    if commission:
        col, first, last = commission
        rng = f"'{AG}'!{col}{first}:{col}{last}"
        metric(52, "Агентская комиссия (средняя по агентам)",
               f"IFERROR(AVERAGE({rng})/100,0)", "%",
               "среднее по столбцу комиссии; пустые/« - » не учитываются", fmt=PCT)
    else:
        ws.cell(52, 2, "Агентская комиссия (средняя по агентам)").font = _REG
        c = ws.cell(52, 3, 0)
        c.font = _REG
        c.number_format = PCT

    # ── 4. Диагностика данных (формулы-флаги по метрикам выше) ────────────
    section(54, "4. Диагностика данных")
    ws.cell(55, 2, "Показатель").font = _HDR
    ws.cell(55, 3, "Статус").font = _HDR
    ws.cell(55, 4, "Комментарий").font = _HDR
    diag = [
        (56, "Средний чек платного билета",
         'IF(AND(C44>=300,C44<=1600),"в норме","проверить")',
         '=IF(AND(C44>=300,C44<=1600),"норма 300–1600 ₽","вне нормы 300–1600 ₽")'),
        (57, "Доля бесплатных билетов",
         'IF(C19<=0.3,"в норме","проверить")', '="порог 30%"'),
        (58, "Заполняемость арены (регулярка)",
         'IF(C30<=1.05,"в норме","проверить")', '="при вместимости "&TEXT($H$2,"#,##0")'),
        (59, "Стоимость абонемента",
         'IF(AND(C39>=3000,C39<=150000),"в норме","проверить")', '="доход абон. / число владельцев"'),
        (60, "Средняя цена платного билета (сезон)",
         'IF(AND(C44>=300,C44<=1600),"в норме","проверить")', '="норма 300–1600 ₽"'),
    ]
    for r, label, status_f, comment_f in diag:
        ws.cell(r, 2, label).font = _REG
        ws.cell(r, 3, f"={status_f}").font = _REG
        ws.cell(r, 4, comment_f).font = Font(name=_FONT, size=9, italic=True, color="6B7280")

    ws.cell(62, 2, "Матчи плей-офф").font = _REG
    ws.cell(62, 3, f"={_count(RZ, rz_po)}").font = _REG

    ws.cell(64, 2,
            "Значения рассчитаны формулами со ссылками на исходные листы; "
            "метрики пересчитываются автоматически при изменении данных.").font = \
        Font(name=_FONT, size=9, italic=True, color="6B7280")

    for r in range(1, 65):
        for col in range(2, 6):
            ws.cell(r, col).alignment = Alignment(vertical="center", wrap_text=(col == 5))


def build_calculations(
    input_bytes: bytes, capacity: int | None = None, season_label: str = "2025/2026",
    byn_to_rub: float | None = None,
    currency: str | None = None, rate: float | None = None,
) -> bytes:
    """Из xlsx-байтов возвращает xlsx-байты с добавленным листом «Расчеты».

    Пересчёт валюты в рубли РФ:
      • currency — код валюты выгрузки («KZT», «BYN» …); если задан, денежные
        метрики считаются в рублях РФ (× курс из редактируемой ячейки H3);
      • rate — значение курса для H3; None — оставить ячейку пустой (клуб
        впишет курс сам);
      • byn_to_rub — устаревший алиас для currency="BYN", rate=byn_to_rub.
    """
    try:
        wb = openpyxl.load_workbook(BytesIO(input_bytes))
    except Exception as exc:  # noqa: BLE001
        raise TicketError(f"Не удалось открыть xlsx: {exc}") from exc

    for name in (RZ, DH):
        if name not in wb.sheetnames:
            raise TicketError(f"В файле нет обязательного листа «{name}».")

    # Числа, записанные текстом, → настоящие числа (иначе SUM даёт 0).
    for name in (RZ, DH, AG):
        if name in wb.sheetnames:
            _coerce_numeric_text(wb[name])

    fx: tuple[str, float | None] | None = None
    if currency:
        fx = (currency, rate)
    elif byn_to_rub is not None:
        fx = ("BYN", byn_to_rub)

    cap = capacity or _default_capacity(BytesIO(input_bytes))
    _build_calc_sheet(wb, season_label, cap, fx=fx)

    # Принудительный полный пересчёт формул при открытии (Excel/Sheets),
    # т.к. openpyxl не сохраняет кэш значений формул.
    try:
        wb.calculation.fullCalcOnLoad = True
    except Exception:  # noqa: BLE001
        pass

    out = BytesIO()
    wb.save(out)
    return out.getvalue()
