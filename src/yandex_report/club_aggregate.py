"""Агрегированный отчёт по клубам: средние показатели по нескольким файлам.

Считает метрики каждого клуба числами (в Python), затем усредняет по клубам
и собирает отдельную книгу с листом «Средние по клубам» (значение + разбивка
по клубам). Раскладка — первая версия, уточняется под пример заказчика.
"""
from __future__ import annotations

import re
from io import BytesIO

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import column_index_from_string, range_boundaries

from .ticket_metrics import (
    AG, DH, RZ, TicketError, _agent_commission_range, _agent_rows,
    _coerce_numeric_text, _section_spans, _to_number,
)

MAX_CLUBS = 22


# ── чтение числовых значений ────────────────────────────────────────────────
# В части выгрузок ячейки заданы формулами (напр. столбец «платные билеты» =
# «всего − бесплатные», =SUM(...)), а кэш значений в файле отсутствует. openpyxl
# формулы не считает, поэтому здесь — мини-вычислитель простых формул того же
# листа (SUM(диапазон) и арифметика +−*/ по ссылкам на ячейки).
_SUM_RE = re.compile(r"SUM\(([^)]+)\)", re.IGNORECASE)
_CELL_RE = re.compile(r"\$?([A-Z]{1,3})\$?(\d+)")
_SAFE_RE = re.compile(r"^[\d\s+\-*/().]*$")


def _eval_formula(ws, formula: str, seen: frozenset) -> float:
    expr = formula.lstrip("=").strip()

    def _sum_sub(mo: re.Match) -> str:
        min_c, min_r, max_c, max_r = range_boundaries(mo.group(1).replace("$", ""))
        total = sum(_cell_value(ws, r, c, seen)
                    for r in range(min_r, max_r + 1)
                    for c in range(min_c, max_c + 1))
        return repr(float(total))

    expr = _SUM_RE.sub(_sum_sub, expr)
    expr = _CELL_RE.sub(
        lambda mo: repr(_cell_value(ws, int(mo.group(2)),
                                    column_index_from_string(mo.group(1)), seen)),
        expr,
    )
    if not _SAFE_RE.match(expr):
        return 0.0
    try:
        return float(eval(expr, {"__builtins__": {}}, {}))  # noqa: S307 — expr санитизирован
    except Exception:  # noqa: BLE001
        return 0.0


def _cell_value(ws, r: int, c: int, seen: frozenset = frozenset()) -> float:
    v = ws.cell(r, c).value
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    if isinstance(v, str):
        s = v.strip()
        if s.startswith("="):
            key = (r, c)
            if key in seen:  # защита от циклических ссылок
                return 0.0
            return _eval_formula(ws, s, seen | {key})
        num = _to_number(s)  # число, записанное текстом
        return float(num) if num is not None else 0.0
    return 0.0


def _num(ws, r: int, c: int) -> float:
    return _cell_value(ws, r, c)


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


# Параметры агрегата: (ключ, подпись, единица, направление, доля?).
# direction: "desc" — выше = лучше (место 1 — максимум); "asc" — ниже = лучше.
PARAMS: list[tuple[str, str, str, str, bool]] = [
    ("price_reg", "Средняя цена билета (регулярный чемпионат)", "руб.", "desc", False),
    ("fill_reg", "Фактическая заполняемость арены (регулярка)", "%", "desc", True),
    ("income_total", "Общая выручка (за сезон)", "руб.", "desc", False),
    ("online_share", "Доля продаж билетов онлайн", "%", "desc", True),
    ("free_share", "Доля бесплатных билетов", "%", "asc", True),
    ("commission", "Агентская комиссия (средняя)", "%", "asc", True),
    ("deviation", "Фактическое отклонение посещаемости", "%", "asc", True),
]

# Параметры, ранжируемые по модулю (ближе к нулю — лучше), а показываемые со
# знаком. Отклонение: «+» протокол больше билетов, «−» меньше.
MAGNITUDE_KEYS = {"deviation"}


def _rank_key(key: str):
    return (lambda v: abs(v)) if key in MAGNITUDE_KEYS else (lambda v: v)


AGG_SHEET = "Средние по клубам"


def parse_aggregate_clubs(agg_bytes: bytes) -> dict[str, dict[str, float]]:
    """Из книги «Средние по клубам» → {клуб: {ключ метрики: значение}}.

    Позволяет брать данные текущего сезона по всем клубам из одного
    агрегированного файла вместо отдельного билетного файла на каждый клуб.
    """
    wb = openpyxl.load_workbook(BytesIO(agg_bytes), data_only=True)
    ws = wb[AGG_SHEET] if AGG_SHEET in wb.sheetnames else wb[wb.sheetnames[0]]
    label_key = {label: key for key, label, *_ in PARAMS}
    result: dict[str, dict[str, float]] = {}
    cur_key = None
    in_table = False
    for r in range(1, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if not isinstance(a, str) or not a.strip():
            continue
        s = a.strip()
        matched = next((k for lbl, k in label_key.items() if s.startswith(lbl)), None)
        if matched:                       # заголовок секции параметра
            cur_key, in_table = matched, False
            continue
        if s == "Клуб":                   # шапка таблицы клубов
            in_table = True
            continue
        if s in ("Среднее по Лиге", "Минимум", "Максимум") or s.startswith("Не обработаны"):
            in_table = False
            continue
        if in_table and cur_key:
            v = ws.cell(r, 2).value
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                result.setdefault(s, {})[cur_key] = float(v)
    return result


# Метрики агрегата → подстрока-метка на листе «Расчеты» (первое числовое значение).
_CALC_LABELS = {
    "price_reg": "средняя цена платного билета (регулярка)",
    "income_total": "всего за сезон",              # «Доход — всего за сезон»
    "online_share": "доля онлайн",
    "free_share": "доля бесплатных билетов",
    "commission": "агентская комиссия",
    "deviation": "расхождение факта с заявленным",
    "fill_reg": "заполняемость арены",             # берём готовой с листа «Расчеты»
}


def read_calc_metrics(input_bytes: bytes, capacity: int | None = None) -> dict | None:
    """Метрики клуба, взятые ГОТОВЫМИ с листа «Расчеты» (кэш значений Excel).

    Возвращает None, если листа нет или он не посчитан (нет кэша). Заполняемость
    считается по протоколам (в «Расчеты» она по билетам)."""
    try:
        wb = openpyxl.load_workbook(BytesIO(input_bytes), data_only=True)
    except Exception:  # noqa: BLE001
        return None
    if "Расчеты" not in wb.sheetnames:
        return None
    found: dict[str, float] = {}
    for row in wb["Расчеты"].iter_rows(values_only=True):
        if len(row) < 3 or not isinstance(row[1], str):
            continue
        v = row[2]
        if not isinstance(v, (int, float)) or isinstance(v, bool):
            continue
        lbl = row[1].strip().lower()
        for key, sub in _CALC_LABELS.items():
            if key not in found and sub in lbl:
                found[key] = float(v)
    if "price_reg" not in found and "income_total" not in found:
        return None                       # лист не посчитан
    return found


def _read_calc_capacity(input_bytes: bytes) -> float | None:
    """Вместимость арены с листа «Расчеты» (строка «Вместимость арены»)."""
    try:
        wb = openpyxl.load_workbook(BytesIO(input_bytes), data_only=True, read_only=True)
    except Exception:  # noqa: BLE001
        return None
    if "Расчеты" not in wb.sheetnames:
        return None
    for row in wb["Расчеты"].iter_rows(values_only=True):
        if len(row) >= 3 and isinstance(row[1], str) and "вместимост" in row[1].lower():
            v = row[2]
            if isinstance(v, (int, float)) and not isinstance(v, bool) and v > 0:
                return float(v)
    return None


def compute_club_metrics(input_bytes: bytes, capacity: int | None = None,
                         to_rub: float = 1.0) -> dict[str, float | None]:
    """Числовые метрики одного клуба (без формул).

    to_rub — множитель для пересчёта денежных метрик в рубли РФ (для клубов,
    чьи суммы в иностранной валюте, напр. BYN/KZT). По умолчанию 1.0 (RUB).
    """
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

    cap = capacity or _read_calc_capacity(input_bytes) or max(
        (sum(_num(rz, r, c) for c in range(3, 11))
         for span in (rz_reg, rz_po) if span
         for r in range(span[0], span[1] + 1)),
        default=0,
    ) or None
    # Фактическая заполняемость — по протоколам (столбец M «Посещаемость по
    # протоколу»), а не по реализованным билетам: средняя протокольная
    # посещаемость регулярки / вместимость арены.
    protocol_reg = _sum_col(rz, 13, rz_reg)
    fill_reg = (protocol_reg / n_reg / cap) if (n_reg and cap and protocol_reg) else None

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

    paid_reg = _sum_col(rz, 3, rz_reg)
    income_paid_reg = _sum_col(dh, 4, dh_reg)
    price_reg = income_paid_reg / paid_reg if paid_reg else None
    price_season = income_paid / paid_single if paid_single else None
    abon_price = income_abon / abon_paid if abon_paid else None

    # Фактическое отклонение посещаемости: (протокол (M=13) − билеты) / билеты.
    # Со знаком: «+» — протокол больше билетов, «−» — меньше. Для места/градации
    # берётся модуль (ближе к нулю — лучше), знак — только для отображения.
    protocol = _both(_sum_col, rz, 13, rz_reg, rz_po)
    deviation = (protocol - att_season) / att_season if (protocol and att_season) else None

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
            # комиссия — процент; отбрасываем неправдоподобные значения
            # (пустые/« - » → 0, а также абсолютные суммы, ошибочно внесённые
            # вместо процента).
            vals = [v for v in (_num(ag, r, col) for r in range(first, last + 1))
                    if 0 < v <= 100]
            commission = (sum(vals) / len(vals) / 100) if vals else None

    # Денежные метрики → рубли РФ (для клубов с суммами в иностранной валюте).
    def _rub(x):
        return x * to_rub if isinstance(x, (int, float)) else x

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
        "income_total": _rub(income_total),
        "income_paid": _rub(income_paid),
        "income_abon": _rub(income_abon),
        "price_season": _rub(price_season),
        "price_reg": _rub(price_reg),
        "abon_price": _rub(abon_price),
        "online_share": online_share,
        "commission": commission,
        "deviation": deviation,
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
_SEC = Font(name=_FONT, bold=True, size=11, color="FFFFFF")
_HDR = Font(name=_FONT, bold=True)
_REG = Font(name=_FONT)
_HDRFILL = PatternFill("solid", fgColor="E2E5EA")
_SECFILL = PatternFill("solid", fgColor="0B5CAD")
_GRADE_FILL = {
    "Хорошие показатели": PatternFill("solid", fgColor="E4F4E4"),
    "Удовлетворительные показатели": PatternFill("solid", fgColor="FCF4DD"),
    "Неудовлетворительные показатели": PatternFill("solid", fgColor="FBE4E4"),
}


def _grade(rank: int, n: int) -> str:
    import math
    if rank <= math.ceil(n / 3):
        return "Хорошие показатели"
    if rank <= math.ceil(2 * n / 3):
        return "Удовлетворительные показатели"
    return "Неудовлетворительные показатели"


def _is_belarus(club: str) -> bool:
    """Белорусский клуб (суммы в BYN): Динамо Минск (рус./лат. написание)."""
    c = club.lower()
    return any(k in c for k in ("минск", "minsk", "динамо мн", "dinamo mn"))


def _compute_all(
    files: list[tuple[str, bytes]], capacity: int | None, byn_to_rub: float | None = None,
) -> tuple[list[tuple[str, dict]], list[str]]:
    if not files:
        raise TicketError("Не приложено ни одного файла.")
    if len(files) > MAX_CLUBS:
        raise TicketError(f"Слишком много файлов: {len(files)} (максимум {MAX_CLUBS}).")
    clubs: list[tuple[str, dict]] = []
    errors: list[str] = []
    for name, data in files:
        club = _club_name(name)
        # белорусский клуб — деньги в BYN → переводим в рубли РФ по курсу
        rate = byn_to_rub if (byn_to_rub and _is_belarus(club)) else 1.0
        try:
            # Готовые значения с листа «Расчеты» (если посчитан в Excel),
            # иначе — считаем сами из исходных листов.
            m = read_calc_metrics(data, capacity)
            if m is not None:
                if rate != 1.0:
                    for k in ("price_reg", "income_total"):
                        if isinstance(m.get(k), (int, float)):
                            m[k] *= rate
            else:
                m = compute_club_metrics(data, capacity, to_rub=rate)
            clubs.append((club, m))
        except Exception as exc:  # noqa: BLE001
            errors.append(f"{name}: {exc}")
    if not clubs:
        raise TicketError("Ни один файл не удалось обработать. " + "; ".join(errors))
    return clubs, errors


def _rank_param(clubs, key: str, direction: str):
    """Возвращает ((среднее, мин, макс), [(место, клуб, значение, градация), …]).

    Клубы без значения идут в конце с местом None.
    """
    pairs = [(name, m.get(key)) for name, m in clubs]
    present = [(name, v) for name, v in pairs if isinstance(v, (int, float))]
    vals = [v for _, v in present]
    stats = (sum(vals) / len(vals), min(vals), max(vals)) if vals else (None, None, None)
    rk = _rank_key(key)
    ranked = sorted(present, key=lambda kv: rk(kv[1]), reverse=(direction == "desc"))
    rows = [(i, name, v, _grade(i, len(ranked)))
            for i, (name, v) in enumerate(ranked, start=1)]
    rows += [(None, name, None, "нет данных")
             for name, v in pairs if not isinstance(v, (int, float))]
    return stats, rows


def _fmt_val(v, unit: str, is_pct: bool, signed: bool = False) -> str:
    if not isinstance(v, (int, float)):
        return "—"
    if is_pct:
        s = f"{v * 100:.0f}%"                        # целые проценты
        return f"+{s}" if signed and v > 0 else s   # знак «+» для положительных
    s = f"{v:,.0f}".replace(",", " ")               # целые рубли
    return f"{s} ₽" if unit == "руб." else s


def aggregate_clubs(
    files: list[tuple[str, bytes]], capacity: int | None = None,
    season_label: str = "2025/2026", byn_to_rub: float | None = None,
) -> bytes:
    """Книга «Средние по клубам»: по каждому параметру — среднее/мин/макс и
    таблица клубов со значением, местом и градацией.

    byn_to_rub — курс белорусского рубля к RUB; применяется к клубу Динамо
    Минск (его суммы в BYN → рубли РФ)."""
    if not files:
        raise TicketError("Не приложено ни одного файла.")
    if len(files) > MAX_CLUBS:
        raise TicketError(f"Слишком много файлов: {len(files)} (максимум {MAX_CLUBS}).")

    clubs, errors = _compute_all(files, capacity, byn_to_rub)
    n = len(clubs)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Средние по клубам"
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 10
    ws.column_dimensions["D"].width = 30

    ws["A1"] = f"Агрегированный отчёт по клубам — сезон {season_label}"
    ws["A1"].font = _H
    ws["A2"] = f"Клубов в выборке: {n}"
    ws["A2"].font = _REG

    def _rnd(v):  # округляем всё до целого: проценты — до целого %, деньги — до рубля
        return round(v, 2) if is_pct else round(v)

    row = 4
    for key, label, unit, direction, is_pct in PARAMS:
        signed = key in MAGNITUDE_KEYS
        fmt = ("+0%;-0%;0%" if signed else "0%") if is_pct else "#,##0"
        pairs = [(name, m.get(key)) for name, m in clubs]
        present = [(name, v) for name, v in pairs if isinstance(v, (int, float))]
        vals = [v for _, v in present]

        # секция параметра
        sc = ws.cell(row, 1, f"{label}, {unit}")
        sc.font = _SEC
        for col in range(1, 5):
            ws.cell(row, col).fill = _SECFILL
        row += 1

        # среднее / мин / макс по Лиге
        if vals:
            avg = sum(vals) / len(vals)
            for j, (lbl, val) in enumerate((
                ("Среднее по Лиге", avg), ("Минимум", min(vals)), ("Максимум", max(vals)),
            )):
                ws.cell(row + j, 1, lbl).font = _REG
                c = ws.cell(row + j, 2, _rnd(val))
                c.font = _HDR
                c.number_format = fmt
            row += 3
        else:
            ws.cell(row, 1, "Нет данных по параметру").font = _REG
            row += 1

        # таблица клубов: Клуб | Значение | Место | Градация
        for j, t in enumerate(("Клуб", "Значение", "Место", "Градация"), start=1):
            hc = ws.cell(row, j, t)
            hc.font = _HDR
            hc.fill = _HDRFILL
        row += 1

        # ранжирование: место по направлению (desc — больше лучше;
        # отклонение — по модулю, ближе к нулю лучше)
        rk = _rank_key(key)
        ranked = sorted(present, key=lambda kv: rk(kv[1]), reverse=(direction == "desc"))
        rank_of = {name: i + 1 for i, (name, _) in enumerate(ranked)}
        for name, val in pairs:
            ws.cell(row, 1, name).font = _REG
            if isinstance(val, (int, float)):
                c = ws.cell(row, 2, _rnd(val))
                c.number_format = fmt
                c.font = _REG
                rk = rank_of[name]
                ws.cell(row, 3, rk).font = _REG
                grade = _grade(rk, len(ranked))
                gc = ws.cell(row, 4, grade)
                gc.font = _REG
                gc.fill = _GRADE_FILL[grade]
            else:
                ws.cell(row, 2, "—").font = _REG
                ws.cell(row, 3, "—").font = _REG
                ws.cell(row, 4, "нет данных").font = _REG
            row += 1
        row += 1  # пустая строка между параметрами

    if errors:
        ws.cell(row, 1, "Не обработаны:").font = _HDR
        for k, e in enumerate(errors, start=1):
            ws.cell(row + k, 1, e).font = Font(name=_FONT, size=9, color="A51C1C")

    out = BytesIO()
    wb.save(out)
    return out.getvalue()


def build_aggregate_docx(
    files: list[tuple[str, bytes]], capacity: int | None = None,
    season_label: str = "2025/2026", byn_to_rub: float | None = None,
) -> bytes:
    """Общий документ (docx): по каждому из 4 параметров — среднее/мин/макс по
    Лиге и таблица клубов, ранжированных по местам (1..N) с градацией.

    byn_to_rub — курс BYN→RUB для клуба Динамо Минск (суммы в белорусских рублях)."""
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    clubs, errors = _compute_all(files, capacity, byn_to_rub)

    doc = Document()
    title = doc.add_heading(
        f"Сводный анализ билетной программы клубов КХЛ — сезон {season_label}", level=0
    )
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(f"Клубов в выборке: {len(clubs)}")

    for key, label, unit, direction, is_pct in PARAMS:
        signed = key in MAGNITUDE_KEYS
        doc.add_heading(label, level=1)
        (avg, mn, mx), rows = _rank_param(clubs, key, direction)
        if avg is not None:
            better = ("ближе к нулю" if signed
                      else "выше значение" if direction == "desc" else "ниже значение")
            doc.add_paragraph(
                f"Среднее по Лиге: {_fmt_val(avg, unit, is_pct, signed)}; "
                f"минимум: {_fmt_val(mn, unit, is_pct, signed)}; "
                f"максимум: {_fmt_val(mx, unit, is_pct, signed)}. "
                f"Место 1 — лучший показатель ({better})."
            )
        else:
            doc.add_paragraph("Нет данных по параметру.")

        table = doc.add_table(rows=1, cols=4)
        table.style = "Light Grid Accent 1"
        for i, t in enumerate(("Место", "Клуб", "Значение", "Градация")):
            cell = table.rows[0].cells[i]
            cell.text = t
            for p in cell.paragraphs:
                for run in p.runs:
                    run.bold = True
        for place, name, v, grade in rows:
            c = table.add_row().cells
            c[0].text = str(place) if place else "—"
            c[1].text = name
            c[2].text = _fmt_val(v, unit, is_pct, signed)
            c[3].text = grade

    if errors:
        doc.add_heading("Не обработаны", level=2)
        for e in errors:
            doc.add_paragraph(e, style="List Bullet")

    from io import BytesIO as _BytesIO
    out = _BytesIO()
    doc.save(out)
    return out.getvalue()
