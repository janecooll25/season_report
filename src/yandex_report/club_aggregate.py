"""Агрегированный отчёт по клубам: средние показатели по нескольким файлам.

Считает метрики каждого клуба числами (в Python), затем усредняет по клубам
и собирает отдельную книгу с листом «Средние по клубам» (значение + разбивка
по клубам). Раскладка — первая версия, уточняется под пример заказчика.
"""
from __future__ import annotations

from io import BytesIO

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

from .ticket_metrics import (
    AG, DH, RZ, TicketError, _agent_commission_range, _agent_rows,
    _coerce_numeric_text, _section_spans,
)

MAX_CLUBS = 22


# ── чтение числовых значений ────────────────────────────────────────────────
def _num(ws, r: int, c: int) -> float:
    v = ws.cell(r, c).value
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else 0.0


def _sum_col(ws, col: int, span) -> float:
    if not span:
        return 0.0
    return sum(_num(ws, r, col) for r in range(span[0], span[1] + 1))


def _max_col(ws, col: int, span) -> float:
    if not span:
        return 0.0
    return max((_num(ws, r, col) for r in range(span[0], span[1] + 1)), default=0.0)


def _both(fn, ws, col, reg, po) -> float:
    return fn(ws, col, reg) + fn(ws, col, po)


# Метрики агрегата: (ключ, подпись, единица, доля?).
METRICS: list[tuple[str, str, str, bool]] = [
    ("att_season", "Посещаемость за сезон (всего проходов)", "чел.", False),
    ("att_avg_season", "Средняя посещаемость (сезон)", "зрит./матч", False),
    ("att_avg_reg", "Средняя посещаемость (регулярка)", "зрит./матч", False),
    ("matches", "Матчей за сезон", "шт.", False),
    ("paid_single", "Платные билеты (разовые)", "шт.", False),
    ("free_single", "Бесплатные билеты (разовые)", "шт.", False),
    ("free_share", "Доля бесплатных билетов", "%", True),
    ("abon_paid", "Платных абонементов (продано)", "шт.", False),
    ("fill_reg", "Заполняемость арены (регулярка)", "%", True),
    ("income_total", "Доход — всего за сезон", "руб.", False),
    ("income_paid", "Доход от платных билетов", "руб.", False),
    ("income_abon", "Доход от абонементов", "руб.", False),
    ("price_season", "Средняя цена платного билета (сезон)", "руб.", False),
    ("abon_price", "Стоимость абонемента (сезон)", "руб.", False),
    ("online_share", "Доля онлайн-продаж", "%", True),
    ("commission", "Агентская комиссия (средняя)", "%", True),
]


def compute_club_metrics(input_bytes: bytes, capacity: int | None = None) -> dict[str, float | None]:
    """Числовые метрики одного клуба (без формул)."""
    wb = openpyxl.load_workbook(BytesIO(input_bytes))
    for name in (RZ, DH):
        if name not in wb.sheetnames:
            raise TicketError(f"В файле нет обязательного листа «{name}».")
    for name in (RZ, DH, AG):
        if name in wb.sheetnames:
            _coerce_numeric_text(wb[name])

    rz, dh = wb[RZ], wb[DH]
    rz_reg, rz_po = _section_spans(rz)
    dh_reg, dh_po = _section_spans(dh)

    # Реализация билетов (колонки C..J = 3..10, всего = их сумма).
    att_reg = sum(_sum_col(rz, c, rz_reg) for c in range(3, 11))
    att_po = sum(_sum_col(rz, c, rz_po) for c in range(3, 11))
    att_season = att_reg + att_po
    n_reg = (rz_reg[1] - rz_reg[0] + 1) if rz_reg else 0
    n_po = (rz_po[1] - rz_po[0] + 1) if rz_po else 0
    n = n_reg + n_po

    paid_single = _both(_sum_col, rz, 3, rz_reg, rz_po)
    free_single = _both(_sum_col, rz, 4, rz_reg, rz_po)
    abon_paid_active = _both(_sum_col, rz, 5, rz_reg, rz_po)
    abon_free_active = _both(_sum_col, rz, 6, rz_reg, rz_po)
    biz_free = _both(_sum_col, rz, 8, rz_reg, rz_po)
    lodge_free = _both(_sum_col, rz, 10, rz_reg, rz_po)
    abon_paid = max(_max_col(rz, 5, rz_reg), _max_col(rz, 5, rz_po))

    free_total = free_single + biz_free + lodge_free
    paid_base = att_season - abon_paid_active - abon_free_active
    free_share = free_total / paid_base if paid_base else None

    cap = capacity or max(
        (sum(_num(rz, r, c) for c in range(3, 11))
         for span in (rz_reg, rz_po) if span
         for r in range(span[0], span[1] + 1)),
        default=0,
    ) or None
    fill_reg = (att_reg / n_reg / cap) if (n_reg and cap) else None

    # Доход. «Всего» — колонка C (3); если она формула без значения,
    # берём сумму компонент D..G (4..7).
    def _income_total(span) -> float:
        if not span:
            return 0.0
        tot = 0.0
        for r in range(span[0], span[1] + 1):
            c = _num(dh, r, 3)
            tot += c if c else (_num(dh, r, 4) + _num(dh, r, 5)
                                + _num(dh, r, 6) + _num(dh, r, 7))
        return tot

    income_paid = _both(_sum_col, dh, 4, dh_reg, dh_po)
    income_abon = _both(_sum_col, dh, 5, dh_reg, dh_po)
    income_total = _income_total(dh_reg) + _income_total(dh_po)

    price_reg_base = _sum_col(rz, 3, rz_reg)
    price_season_base = paid_single
    price_season = income_paid / price_season_base if price_season_base else None
    abon_price = income_abon / abon_paid if abon_paid else None

    # Каналы и комиссия.
    online = offline = 0.0
    online_share = commission = None
    if AG in wb.sheetnames:
        ag = wb[AG]
        rows = _agent_rows(ag)
        def chan(key):
            return _num(ag, rows[key], 2) if key in rows else 0.0
        online = chan("online") + chan("agent_online")
        offline = chan("offline") + chan("agent_offline")
        if online + offline:
            online_share = online / (online + offline)
        rng = _agent_commission_range(ag)
        if rng:
            col_letter, first, last = rng
            col = openpyxl.utils.column_index_from_string(col_letter)
            vals = [_num(ag, r, col) for r in range(first, last + 1)
                    if isinstance(ag.cell(r, col).value, (int, float))
                    and not isinstance(ag.cell(r, col).value, bool)]
            commission = (sum(vals) / len(vals) / 100) if vals else None

    return {
        "att_season": att_season,
        "att_avg_season": (att_season / n) if n else None,
        "att_avg_reg": (att_reg / n_reg) if n_reg else None,
        "matches": n,
        "paid_single": paid_single,
        "free_single": free_single,
        "free_share": free_share,
        "abon_paid": abon_paid,
        "fill_reg": fill_reg,
        "income_total": income_total,
        "income_paid": income_paid,
        "income_abon": income_abon,
        "price_season": price_season,
        "abon_price": abon_price,
        "online_share": online_share,
        "commission": commission,
    }


def _club_name(filename: str) -> str:
    stem = filename.rsplit("/", 1)[-1].rsplit(".", 1)[0]
    for pref in ("tickets_sales_", "tickets_", "raschety_"):
        if stem.lower().startswith(pref):
            stem = stem[len(pref):]
    return stem.replace("_", " ").strip() or filename


# ── стили ────────────────────────────────────────────────────────────────
_FONT = "Arial"
_H = Font(name=_FONT, bold=True, size=13)
_HDR = Font(name=_FONT, bold=True)
_REG = Font(name=_FONT)
_HDRFILL = PatternFill("solid", fgColor="E2E5EA")


def aggregate_clubs(
    files: list[tuple[str, bytes]], capacity: int | None = None,
    season_label: str = "2025/2026",
) -> bytes:
    """Собирает книгу с листом «Средние по клубам» по нескольким файлам."""
    if not files:
        raise TicketError("Не приложено ни одного файла.")
    if len(files) > MAX_CLUBS:
        raise TicketError(f"Слишком много файлов: {len(files)} (максимум {MAX_CLUBS}).")

    clubs: list[tuple[str, dict]] = []
    errors: list[str] = []
    for name, data in files:
        try:
            clubs.append((_club_name(name), compute_club_metrics(data, capacity)))
        except Exception as exc:  # noqa: BLE001
            errors.append(f"{name}: {exc}")
    if not clubs:
        raise TicketError("Ни один файл не удалось обработать. " + "; ".join(errors))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Средние по клубам"
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 10

    ws["A1"] = f"Агрегированный отчёт по клубам — сезон {season_label}"
    ws["A1"].font = _H
    ws["A2"] = f"Клубов в выборке: {len(clubs)}"
    ws["A2"].font = _REG

    # Заголовок таблицы: Показатель | Среднее | Ед. | <клуб1> | <клуб2> …
    hdr = ["Показатель", "Среднее по клубам", "Ед."] + [c[0] for c in clubs]
    for j, t in enumerate(hdr, start=1):
        cell = ws.cell(4, j, t)
        cell.font = _HDR
        cell.fill = _HDRFILL
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for j in range(4, 4 + len(clubs)):
        ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = 16

    money_fmt = '#,##0'
    pct_fmt = "0.0%"
    row = 5
    for key, label, unit, is_pct in METRICS:
        ws.cell(row, 1, label).font = _REG
        vals = [c[1].get(key) for c in clubs]
        present = [v for v in vals if isinstance(v, (int, float))]
        avg = (sum(present) / len(present)) if present else None
        fmt = pct_fmt if is_pct else money_fmt
        a = ws.cell(row, 2, round(avg, 4) if avg is not None else "—")
        a.font = Font(name=_FONT, bold=True)
        if avg is not None:
            a.number_format = fmt
        ws.cell(row, 3, unit).font = _REG
        for j, (_, m) in enumerate(clubs, start=4):
            v = m.get(key)
            c = ws.cell(row, j, round(v, 4) if isinstance(v, (int, float)) else "—")
            c.font = _REG
            if isinstance(v, (int, float)):
                c.number_format = fmt
        row += 1

    if errors:
        ws.cell(row + 1, 1, "Не обработаны:").font = _HDR
        for k, e in enumerate(errors, start=1):
            ws.cell(row + 1 + k, 1, e).font = Font(name=_FONT, size=9, color="A51C1C")

    out = BytesIO()
    wb.save(out)
    return out.getvalue()
