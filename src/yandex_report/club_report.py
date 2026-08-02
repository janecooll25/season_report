"""Аналитическая справка по клубу (билетная программа) — детерминированно.

Из файла мониторинга КХЛ («Чек-лист мониторинг») берётся история по сезонам и
контекст Лиги (среднее/мин/макс, место), а значения текущего сезона — либо из
того же файла (если колонка сезона заполнена), либо из билетного файла клуба
(через compute_club_metrics). Текст «Характеристика» и «Рекомендации»
собирается по правилам шаблона, без обращения к LLM.
"""
from __future__ import annotations

from io import BytesIO

import openpyxl

from .club_aggregate import compute_club_metrics
from .ticket_metrics import TicketError

MONITOR_SHEET = "Чек-лист мониторинг"
DIRECTION_TICKET = "Билетная программа Клуба"

# Параметры билетной программы: (столбец-начало блока, название как в шаблоне,
# вид значения, направление «лучше»). Столбцы соответствуют раскладке файла
# мониторинга (каждый блок — 8 сезонов, 18/19 … 25/26).
PARAM_ONLINE = "Доля розничных продаж билетов через интернет"
PARAM_PRICE = "Средняя цена коммерческой реализации билетов в рамках Регулярного чемпионата"
PARAM_FREE = "Процент билетов и абонементов, распространяемых на безвозмездной основе"
PARAM_DEV = ("Отклонение данных официальных протоколов матчей от данных выгрузок "
             "из билетной системы по выбранным играм")

TICKET_PARAMS = [
    # (col_start, name, kind, better)  kind: share|money|grade
    (2, PARAM_DEV, "grade", "asc"),
    (10, PARAM_ONLINE, "share", "desc"),
    (18, PARAM_PRICE, "money", "desc"),
    (26, PARAM_FREE, "share", "asc"),
]

# Ключи из compute_club_metrics для значений текущего сезона.
METRIC_KEY = {
    PARAM_ONLINE: "online_share",
    PARAM_PRICE: "price_reg",
    PARAM_FREE: "free_share",
    PARAM_DEV: "deviation",
}

GRADES = {3: "Хорошие показатели", 2: "Удовлетворительные показатели",
          1: "Неудовлетворительные показатели"}


class ReportError(TicketError):
    pass


# ── чтение файла мониторинга ────────────────────────────────────────────────
def _monitor_ws(wb):
    return wb[MONITOR_SHEET] if MONITOR_SHEET in wb.sheetnames else wb[wb.sheetnames[0]]


def _seasons(ws, col_start: int) -> list[str]:
    return [str(ws.cell(6, col_start + i).value or "").strip() for i in range(8)]


def _club_rows(ws) -> dict[str, int]:
    rows: dict[str, int] = {}
    for r in range(7, ws.max_row + 1):
        name = ws.cell(r, 1).value
        if not isinstance(name, str) or not name.strip():
            continue
        if name.strip() == "Характеристика":
            break
        rows[name.strip()] = r
    return rows


def _season_index(ws, col_start: int, season: str) -> int | None:
    for i, s in enumerate(_seasons(ws, col_start)):
        if s == season:
            return i
    return None


def _num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def _column_values(ws, rows: dict[str, int], col_start: int, idx: int) -> dict[str, float]:
    out = {}
    for name, r in rows.items():
        v = _num(ws.cell(r, col_start + idx).value)
        if v is not None:
            out[name] = v
    return out


def _latest_league_index(ws, rows, col_start: int, prefer: int | None) -> int | None:
    """Последний сезон с данными ≥3 клубов (для среднего/мин/макс/места).

    Если запрошенный сезон (prefer) уже имеет ≥3 клубов — берём его, иначе идём
    назад к последнему заполненному.
    """
    order = list(range(8))
    if prefer is not None:
        order = [prefer] + [i for i in range(7, -1, -1) if i != prefer]
    else:
        order = list(range(7, -1, -1))
    for i in order:
        if len(_column_values(ws, rows, col_start, i)) >= 3:
            return i
    return None


# ── форматирование ──────────────────────────────────────────────────────────
def _pct(v) -> str:
    return f"{round(v * 100)}%"


def _rub(v) -> str:
    return f"{round(v):,}".replace(",", " ") + " руб."


def _tidy(s: str) -> str:
    return s.replace("руб..", "руб.").replace("%.", "%.").replace("  ", " ").strip()


def _season_full(season: str) -> str:
    a, _, b = season.partition("/")
    return f"20{a}/20{b}" if a and b else season


def _grade_by_place(place: int, n: int) -> int:
    import math
    if place <= math.ceil(n / 3):
        return 3
    if place <= math.ceil(2 * n / 3):
        return 2
    return 1


# ── сборка текста по параметру ──────────────────────────────────────────────
def _place(better: str, league_all: dict[str, float], club: str, cur) -> int | None:
    if cur is None:
        return None
    pool = dict(league_all)
    pool[club] = cur  # клуб оценивается своим текущим значением
    ranked = sorted(pool.items(), key=lambda kv: kv[1], reverse=(better == "desc"))
    for i, (name, _) in enumerate(ranked, start=1):
        if name == club:
            return i
    return None


def _describe(param: str, kind: str, better: str, cur, prev, league_all: dict[str, float],
              club: str) -> tuple[str, str, str]:
    """Возвращает (характеристика, рекомендации, подпись-градация)."""
    vals = list(league_all.values())
    avg = sum(vals) / len(vals) if vals else None
    lo, hi = (min(vals), max(vals)) if vals else (None, None)

    if kind == "share":
        cur_s = _pct(cur) if cur is not None else "нет данных"
        if param == PARAM_ONLINE:
            rel = ("выше" if cur is not None and avg is not None and cur >= avg else "ниже")
            char = (f"Клуб имеет долю продаж билетов через интернет в размере {cur_s}"
                    + (f", что {rel} среднего значения среди клубов КХЛ." if avg is not None else ".")
                    + (f" Средняя доля продаж билетов через интернет составляет {_pct(avg)}, "
                       f"минимальный показатель – {_pct(lo)}, максимальный показатель – {_pct(hi)}."
                       if avg is not None else ""))
            tail = ("Рекомендации отсутствуют." if cur is not None and cur >= (avg or 0)
                    else "Клубу рекомендуется развивать онлайн-канал продаж.")
            if prev is not None and cur is not None and _pct(prev) != _pct(cur):
                trend = "увеличилась" if cur >= prev else "уменьшилась"
                rec = f"Доля продаж билетов через интернет {trend} с {_pct(prev)} до {cur_s}. {tail}"
            elif prev is not None and cur is not None:
                rec = ("Доля продаж билетов через интернет по сравнению с прошлым сезоном "
                       f"существенно не изменилась и составляет {cur_s}. {tail}")
            else:
                rec = tail
        else:  # PARAM_FREE
            char = (f"{cur_s} билетов на матчи Клуба в прошедшем сезоне распространялись "
                    f"на бесплатной основе."
                    + (f" Среднее значение по Лиге составило {_pct(avg)}, "
                       f"минимальный показатель – {_pct(lo)}, максимальный показатель – {_pct(hi)}."
                       if avg is not None else ""))
            if prev is not None and cur is not None:
                trend = "уменьшилась" if cur <= prev else "увеличилась"
                rec = (f"Доля бесплатных билетов {trend} с {_pct(prev)} до {cur_s} по сравнению "
                       f"с прошлым сезоном. "
                       + ("Показатель ниже среднего по Лиге. " if avg is not None and cur < avg else "")
                       + "Планомерное снижение доли бесплатных билетов позволит Клубу повысить "
                         "доходы от реализации билетной программы.")
            else:
                rec = "Рекомендации отсутствуют."

    elif kind == "money":
        cur_s = _rub(cur) if cur is not None else "нет данных"
        place = _place(better, league_all, club, cur)
        char = (f"Средняя стоимость билета в регулярном чемпионате составила {cur_s}"
                + (f", и Клуб находится на {place} месте в Лиге по этому показателю."
                   if place else ".")
                + (f" Средний показатель по клубам в рамках регулярного чемпионата составил "
                   f"{_rub(avg)}. При этом минимальный показатель составил {_rub(lo)}, "
                   f"максимальный – {_rub(hi)}." if avg is not None else ""))
        if prev is not None and cur is not None:
            trend = "увеличилась" if cur >= prev else "уменьшилась"
            rec = f"Средняя стоимость билета в регулярном чемпионате {trend} с {_rub(prev)} до {cur_s}"
        else:
            rec = "Рекомендации отсутствуют."

    else:  # grade — отклонение протоколов
        # из мониторинга приходит градация 1/2/3, из билетного файла — доля (0..1)
        is_grade = cur is not None and cur >= 1
        if cur is None:
            char = "Данные для оценки отклонения от официальных протоколов не предоставлены."
        elif is_grade:
            char = ("По итогам мониторинга данные по реализованным билетам и абонементам "
                    f"отнесены к категории «{GRADES.get(int(round(cur)), '—').lower()}».")
        else:
            char = ("Предоставленные данные по реализованным билетам и абонементам имеют "
                    f"отклонение от заявленной официальной посещаемости на {_pct(cur)}.")
        rec = ("Предоставление неверных данных о посещаемости матчей несёт репутационные риски "
               "для КХЛ и Клуба, а также риски наложения штрафа в соответствии с Дисциплинарным "
               "регламентом КХЛ. Согласно п.1.22 ст.54 Спортивного регламента КХЛ Клуб обязан "
               "предоставлять официальную информацию о количестве зрителей на основе данных "
               "электронной билетно-пропускной системы.")

    # градация/место
    grade_label = ""
    if kind in ("share", "money") and league_all:
        place = _place(better, league_all, club, cur)
        if place:
            grade_label = GRADES[_grade_by_place(place, len(league_all))]
    elif kind == "grade" and cur is not None:
        if cur >= 1:  # градация уже указана в тексте характеристики
            grade_label = ""
        else:         # доля из билетного файла: меньше — лучше
            grade_label = GRADES[3 if cur <= 0.02 else (2 if cur <= 0.05 else 1)]
    return _tidy(char), _tidy(rec), grade_label


# ── публичная функция ───────────────────────────────────────────────────────
def build_club_report(
    monitoring_bytes: bytes, club: str, season: str = "25/26",
    prev_season: str = "24/25", ticket_bytes: bytes | None = None,
    capacity: int | None = None, to_rub: float = 1.0,
) -> bytes:
    """docx-справка по билетной программе клуба за сезон.

    Значения текущего сезона: из ticket_bytes (если задан) или из колонки season
    файла мониторинга. Прошлый сезон, среднее/мин/макс и место — из мониторинга.
    """
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt

    wb = openpyxl.load_workbook(BytesIO(monitoring_bytes), data_only=True)
    ws = _monitor_ws(wb)
    rows = _club_rows(ws)
    if club not in rows:
        raise ReportError(f"Клуб «{club}» не найден в файле мониторинга. "
                          f"Доступны: {', '.join(sorted(rows))}.")

    metrics = None
    if ticket_bytes is not None:
        metrics = compute_club_metrics(ticket_bytes, capacity=capacity, to_rub=to_rub)

    doc = Document()
    h = doc.add_heading(
        f"Аналитическая справка по коммерческой деятельности "
        f"ХК «{club}» в сезоне {_season_full(season)} годов", level=1)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER

    table = doc.add_table(rows=1, cols=4)
    table.style = "Table Grid"
    for i, t in enumerate(("Направление деятельности", "Параметр",
                           "Характеристика / место Клуба в Лиге по параметру", "Рекомендации")):
        cell = table.rows[0].cells[i]
        cell.text = t
        for p in cell.paragraphs:
            for run in p.runs:
                run.bold = True

    for col_start, param, kind, better in TICKET_PARAMS:
        cur_idx = _season_index(ws, col_start, season)
        prev_idx = _season_index(ws, col_start, prev_season)
        # текущее значение
        cur = None
        if cur_idx is not None:
            cur = _num(ws.cell(rows[club], col_start + cur_idx).value)
        if cur is None and metrics is not None:
            cur = metrics.get(METRIC_KEY.get(param))
        prev = _num(ws.cell(rows[club], col_start + prev_idx).value) if prev_idx is not None else None

        # Лига по последнему заполненному сезону (для среднего/мин/макс/места)
        league_all = {}
        if kind in ("share", "money"):
            li = _latest_league_index(ws, rows, col_start, cur_idx)
            if li is not None:
                league_all = _column_values(ws, rows, col_start, li)

        char, rec, grade_label = _describe(param, kind, better, cur, prev, league_all, club)
        row = table.add_row().cells
        row[0].text = DIRECTION_TICKET
        row[1].text = param
        row[2].text = (char + (f"\n\n{grade_label}." if grade_label else ""))
        row[3].text = rec

    for p in doc.paragraphs:
        for run in p.runs:
            run.font.size = run.font.size or Pt(11)

    out = BytesIO()
    doc.save(out)
    return out.getvalue()


def list_clubs(monitoring_bytes: bytes) -> list[str]:
    wb = openpyxl.load_workbook(BytesIO(monitoring_bytes), data_only=True, read_only=True)
    ws = _monitor_ws(wb)
    return list(_club_rows(ws))
