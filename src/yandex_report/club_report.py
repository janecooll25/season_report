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

from .club_aggregate import compute_club_metrics, parse_aggregate_clubs
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
PARAM_COMMISSION = "Размер партнерской комиссии по договору на реализацию билетов"

TICKET_PARAMS = [
    # (col_start, name, kind, better)  kind: share|money|grade|commission
    # col_start=None — параметра нет в файле мониторинга (берётся из агрегата/билетов)
    (2, PARAM_DEV, "grade", "asc"),
    (10, PARAM_ONLINE, "share", "desc"),
    (None, PARAM_COMMISSION, "commission", "asc"),
    (18, PARAM_PRICE, "money", "desc"),
    (26, PARAM_FREE, "share", "asc"),
]

# Ключи из compute_club_metrics / агрегата для значений текущего сезона.
METRIC_KEY = {
    PARAM_ONLINE: "online_share",
    PARAM_PRICE: "price_reg",
    PARAM_FREE: "free_share",
    PARAM_DEV: "deviation",
    PARAM_COMMISSION: "commission",
}

GRADES = {3: "Хорошие показатели", 2: "Удовлетворительные показатели",
          1: "Неудовлетворительные показатели"}

# Привязка клубов КХЛ к субъектам РФ (для места региона по доходам населения).
# Зарубежные клубы (None) в ранжирование не входят.
CLUB_REGION: dict[str, str | None] = {
    "Авангард": "Омская область",
    "Автомобилист": "Свердловская область",
    "Адмирал": "Приморский край",
    "Ак Барс": "Республика Татарстан",
    "Амур": "Хабаровский край",
    "Витязь": "Московская область",
    "Динамо Москва": "г. Москва",
    "Лада": "Самарская область",
    "Локомотив": "Ярославская область",
    "Металлург": "Челябинская область",
    "Нефтехимик": "Республика Татарстан",
    "Салават Юлаев": "Республика Башкортостан",
    "Северсталь": "Вологодская область",
    "Сибирь": "Новосибирская область",
    "СКА": "г. Санкт-Петербург",
    "Сочи": "Краснодарский край",
    "Спартак": "г. Москва",
    "Торпедо": "Нижегородская область",
    "Трактор": "Челябинская область",
    "ЦСКА": "г. Москва",
    "Барыс": None, "Динамо Минск": None,
    "Куньлунь Ред Стар": None, "Шанхайские Драконы": None,
}
REGION_INCOME_SHEET = "СДД_субъекты"

# ── дизайн справки (по образцу шаблона) ─────────────────────────────────────
FONT_NAME = "Times New Roman"
FILL_HEADER = "BFBFBF"          # серый — шапка и легенда
FILL_WHITE = "FFFFFF"
GRADE_FILL = {                  # заливка ячейки «Характеристика» по градации
    "Хорошие показатели": "C5E0B3",            # зелёный
    "Удовлетворительные показатели": "FFE599",  # жёлтый
    "Неудовлетворительные показатели": "FF9999",  # красный
}
COL_WIDTHS = [1571, 2388, 4253, 7508]  # ширины столбцов, DXA (как в шаблоне)
HEADER_TITLES = ("Направление деятельности", "Параметр",
                 "Характеристика/ место Клуба в Лиге по параметру", "Рекомендации")


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
              club: str, region_note: str = "") -> tuple[str, str, str]:
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
            rec = f"Средняя стоимость билета в регулярном чемпионате {trend} с {_rub(prev)} до {cur_s}."
        else:
            rec = ""
        if region_note:
            rec = (rec + " " + region_note).strip()
        if not rec:
            rec = "Рекомендации отсутствуют."

    elif kind == "commission":
        pos = [v for v in vals if 0 < v <= 1]      # среднее по работающим с агентами
        avg_pos = sum(pos) / len(pos) if pos else None
        if not (isinstance(cur, (int, float)) and 0 < cur <= 1):
            cur = None
        if cur is None:
            char = "Клуб не работает с агентами для реализации билетов."
        else:
            char = (f"Размер агентской комиссии по договорам на реализацию билетов "
                    f"составляет {_pct(cur)}.")
        if avg_pos is not None:
            char += f" Среднее значение по Лиге составляет {_pct(avg_pos)}."
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
    if kind in ("share", "money", "commission") and league_all:
        place = _place(better, league_all, club, 0 if (kind == "commission" and not cur) else cur)
        if place:
            grade_label = GRADES[_grade_by_place(place, len(league_all))]
    elif kind == "grade" and cur is not None:
        if cur >= 1:  # градация из мониторинга (1/2/3)
            grade_label = GRADES.get(int(round(cur)), "")
        else:         # доля из билетного файла: меньше — лучше
            grade_label = GRADES[3 if cur <= 0.02 else (2 if cur <= 0.05 else 1)]
    return _tidy(char), _tidy(rec), grade_label


# ── публичная функция ───────────────────────────────────────────────────────
def build_club_report(
    monitoring_bytes: bytes, club: str, season: str = "25/26",
    prev_season: str = "24/25", ticket_bytes: bytes | None = None,
    capacity: int | None = None, to_rub: float = 1.0,
    region_income_bytes: bytes | None = None, income_year: str = "2025",
    aggregate_bytes: bytes | None = None,
) -> bytes:
    """docx-справка по билетной программе клуба за сезон.

    Значения текущего сезона (в порядке приоритета): из колонки season файла
    мониторинга → из агрегированного файла клубов (aggregate_bytes) → из
    отдельного билетного файла (ticket_bytes). Прошлый сезон и история — из
    мониторинга. aggregate_bytes — один общий файл «Средние по клубам» со всеми
    клубами (в т.ч. агентская комиссия), удобнее отдельного файла на клуб.
    region_income_bytes — файл Росстата (СДД по субъектам); если задан, в
    рекомендацию по цене добавляется место региона клуба по доходам населения.
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

    # Данные текущего сезона: из общего агрегата (по всем клубам) либо из
    # отдельного билетного файла клуба.
    metrics = None            # значения выбранного клуба
    agg_clubs: dict = {}      # значения всех клубов (для средней комиссии по Лиге)
    if aggregate_bytes is not None:
        agg_clubs = parse_aggregate_clubs(aggregate_bytes)
        metrics = agg_clubs.get(club, {})
    elif ticket_bytes is not None:
        metrics = compute_club_metrics(ticket_bytes, capacity=capacity, to_rub=to_rub)

    region_note = ""
    if region_income_bytes is not None:
        rr = region_rank(parse_region_income(region_income_bytes, income_year), club)
        if rr:
            region, place, n, val = rr
            region_note = (
                f"Согласно данным Федеральной службы государственной статистики за "
                f"{income_year} год {region} находится на {place}-м месте из {n} среди "
                f"регионов присутствия КХЛ по среднедушевым денежным доходам населения "
                f"({_rub(val)}/мес).")

    from docx.enum.section import WD_ORIENT
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Twips

    def _font(run, size_pt: float, bold: bool = False) -> None:
        run.font.name = FONT_NAME
        run.font.size = Pt(size_pt)
        run.font.bold = bold
        rpr = run._element.get_or_add_rPr()
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            rf = OxmlElement("w:rFonts")
            rpr.insert(0, rf)
        for attr in ("w:ascii", "w:hAnsi", "w:cs"):
            rf.set(qn(attr), FONT_NAME)

    def _shade(cell, hex6: str) -> None:
        tcpr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), hex6)
        tcpr.append(shd)

    def _fill_cell(cell, text: str, size: float = 10, bold: bool = False,
                   fill: str | None = None, center: bool = False) -> None:
        cell.text = ""
        para = cell.paragraphs[0]
        if center:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        lines = str(text).split("\n")
        run = para.add_run(lines[0])
        _font(run, size, bold)
        for extra in lines[1:]:
            run.add_break()
            run = para.add_run(extra)
            _font(run, size, bold)
        if fill:
            _shade(cell, fill)

    doc = Document()
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = Twips(16838), Twips(11906)
    sec.top_margin = sec.bottom_margin = Twips(567)
    sec.left_margin = sec.right_margin = Twips(567)

    # «ПРИЛОЖЕНИЕ №1» — справа сверху, как в шаблоне.
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p.add_run("ПРИЛОЖЕНИЕ №1\nк исх.№_______")
    _font(r, 14)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(f"Аналитическая справка по коммерческой деятельности "
                  f"ХК «{club}» в сезоне {_season_full(season)} годов")
    _font(r, 13, bold=True)

    table = doc.add_table(rows=0, cols=4)
    table.style = "Table Grid"
    table.autofit = False
    table.allow_autofit = False
    # фиксируем ширины столбцов в сетке таблицы (как в шаблоне)
    for gc, w in zip(table._tbl.tblGrid.findall(qn("w:gridCol")), COL_WIDTHS):
        gc.set(qn("w:w"), str(w))

    def _add_row(values, sizes=10, bolds=False, fills=(None, None, None, None)):
        cells = table.add_row().cells
        for i, cell in enumerate(cells):
            size = sizes[i] if isinstance(sizes, (list, tuple)) else sizes
            bold = bolds[i] if isinstance(bolds, (list, tuple)) else bolds
            _fill_cell(cell, values[i], size=size, bold=bold, fill=fills[i],
                       center=(i != 3))
            cell.width = Twips(COL_WIDTHS[i])
        return cells

    # шапка
    _add_row(HEADER_TITLES, sizes=10, bolds=True,
             fills=(FILL_HEADER, FILL_HEADER, FILL_HEADER, FILL_HEADER))
    # легенда градаций (цвет — в столбце «Характеристика»)
    for label, fill in (("Хорошие показатели", GRADE_FILL["Хорошие показатели"]),
                        ("Удовлетворительные показатели", GRADE_FILL["Удовлетворительные показатели"]),
                        ("Неудовлетворительные показатели", GRADE_FILL["Неудовлетворительные показатели"])):
        _add_row(("", "", label, ""), sizes=10, bolds=False,
                 fills=(FILL_HEADER, FILL_HEADER, fill, FILL_HEADER))

    first_data = len(table.rows)
    for col_start, param, kind, better in TICKET_PARAMS:
        cur_idx = _season_index(ws, col_start, season) if col_start else None
        prev_idx = _season_index(ws, col_start, prev_season) if col_start else None
        cur = None
        if col_start and cur_idx is not None:
            cur = _num(ws.cell(rows[club], col_start + cur_idx).value)
        if cur is None and metrics is not None:
            cur = metrics.get(METRIC_KEY.get(param))
        prev = (_num(ws.cell(rows[club], col_start + prev_idx).value)
                if col_start and prev_idx is not None else None)

        league_all = {}
        if kind in ("share", "money") and col_start:
            li = _latest_league_index(ws, rows, col_start, cur_idx)
            if li is not None:
                league_all = _column_values(ws, rows, col_start, li)
        elif kind == "commission" and agg_clubs:
            # средняя комиссия по Лиге — из агрегата; берём только правдоподобные
            # значения (доля 0..1, т.е. 0–100 %), отсеивая ошибочные выбросы
            league_all = {c: m["commission"] for c, m in agg_clubs.items()
                          if isinstance(m.get("commission"), (int, float))
                          and 0 < m["commission"] <= 1}

        char, rec, grade_label = _describe(
            param, kind, better, cur, prev, league_all, club,
            region_note=region_note if param == PARAM_PRICE else "")
        char_fill = GRADE_FILL.get(grade_label, FILL_WHITE)
        _add_row((DIRECTION_TICKET, param, char, rec or "Рекомендации отсутствуют."),
                 sizes=10, bolds=False, fills=(None, FILL_WHITE, char_fill, None))

    # объединяем столбец «Направление» по строкам билетной программы
    last_data = len(table.rows) - 1
    if last_data > first_data:
        merged = table.cell(first_data, 0).merge(table.cell(last_data, 0))
        _fill_cell(merged, DIRECTION_TICKET, size=10, center=True)

    out = BytesIO()
    doc.save(out)
    return out.getvalue()


def parse_region_income(income_bytes: bytes, year: str = "2025") -> dict[str, float]:
    """{субъект РФ: среднедушевой доход, руб./мес} за годовой столбец `year`.

    Столбец определяется автоматически: в строке-заголовке групп встречается
    год, а в подзаголовке — «год» (годовое значение). Строки федеральных
    округов и «Российская Федерация» исключаются.
    """
    wb = openpyxl.load_workbook(BytesIO(income_bytes), data_only=True)
    ws = wb[REGION_INCOME_SHEET] if REGION_INCOME_SHEET in wb.sheetnames else wb[wb.sheetnames[0]]
    # ищем годовой столбец нужного года (заголовок групп — обычно строка 6)
    target = None
    cur_year = None
    for c in range(2, ws.max_column + 1):
        y = ws.cell(6, c).value
        if y not in (None, ""):
            cur_year = str(y)
        q = ws.cell(7, c).value
        if cur_year and year in cur_year and isinstance(q, str) and q.strip().lower() == "год":
            target = c
            break
    if target is None:
        raise ReportError(f"В файле доходов не найден годовой столбец за {year} год.")

    out: dict[str, float] = {}
    for r in range(8, ws.max_row + 1):
        name = ws.cell(r, 1).value
        v = _num(ws.cell(r, target).value)
        if not isinstance(name, str) or v is None:
            continue
        n = name.strip()
        if "федеральн" in n.lower() or "Российская" in n:
            continue
        out[n] = v
    return out


def region_rank(income: dict[str, float], club: str) -> tuple[str, int, int, float] | None:
    """(регион, место, всего регионов КХЛ, доход) для клуба; None — если нет региона."""
    region = CLUB_REGION.get(club)
    if not region or region not in income:
        return None
    khl_regions = {r for c, r in CLUB_REGION.items() if r and r in income}
    ranked = sorted(khl_regions, key=lambda r: income[r], reverse=True)  # выше доход — выше место
    place = ranked.index(region) + 1
    return region, place, len(ranked), income[region]


def list_clubs(monitoring_bytes: bytes) -> list[str]:
    wb = openpyxl.load_workbook(BytesIO(monitoring_bytes), data_only=True, read_only=True)
    ws = _monitor_ws(wb)
    return list(_club_rows(ws))
