"""Генерация диаграмм для разделов отчёта (matplotlib → PNG для вставки в docx)."""
from __future__ import annotations

import io

from .report_data import DataBlock

# Поддерживаемые типы диаграмм (значение <select> на форме -> тип).
CHART_TYPES = {"bar", "barh", "pie", "line"}

CHART_LABELS_RU = {
    "bar": "Столбчатая",
    "barh": "Горизонтальная",
    "pie": "Круговая",
    "line": "Линейная",
}

# Приятная нейтральная палитра.
_PALETTE = [
    "#0b5cad", "#3b82f6", "#22a5b8", "#f59e0b", "#ef4444",
    "#8b5cf6", "#10b981", "#64748b", "#e879a6", "#84cc16",
]


def _to_float(raw: object) -> float | None:
    s = str(raw).replace("%", "").replace("\xa0", "").replace(" ", "").replace(",", ".")
    s = s.replace("−", "-")  # юникод-минус
    try:
        return float(s)
    except ValueError:
        return None


def _series(block: DataBlock, top: int) -> tuple[list[str], list[float]]:
    labels: list[str] = []
    values: list[float] = []
    for r in block.rows:
        if not r or r[0].strip() == "Итого и средние":
            continue
        val = _to_float(r[1]) if len(r) > 1 else None
        if val is None:
            continue
        labels.append(r[0])
        values.append(val)
    pairs = sorted(zip(labels, values), key=lambda x: -x[1])[:top]
    return [p[0] for p in pairs], [p[1] for p in pairs]


def render_chart(block: DataBlock, chart_type: str, title: str, top: int = 8) -> bytes | None:
    """Рисует диаграмму по первому числовому столбцу блока. None — если нечего рисовать."""
    if chart_type not in CHART_TYPES:
        return None
    labels, values = _series(block, top)
    if not values:
        return None

    # На Vercel домашняя папка недоступна для записи — кэш matplotlib в /tmp.
    import os
    os.environ.setdefault("MPLCONFIGDIR", "/tmp/mplconfig")
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    fig, ax = plt.subplots(figsize=(6.2, 3.6), dpi=150)
    colors = [_PALETTE[i % len(_PALETTE)] for i in range(len(values))]

    if chart_type == "pie":
        ax.pie(values, labels=labels, colors=colors, autopct="%1.1f%%",
               textprops={"fontsize": 8}, startangle=90)
        ax.axis("equal")
    elif chart_type == "barh":
        y = range(len(labels))
        ax.barh(list(y), values, color=colors)
        ax.set_yticks(list(y))
        ax.set_yticklabels(labels, fontsize=8)
        ax.invert_yaxis()  # крупнейшее сверху
    elif chart_type == "line":
        ax.plot(range(len(values)), values, marker="o", color=_PALETTE[0])
        ax.set_xticks(range(len(labels)))
        ax.set_xticklabels(labels, rotation=40, ha="right", fontsize=8)
        ax.grid(True, axis="y", alpha=0.3)
    else:  # bar
        x = range(len(labels))
        ax.bar(list(x), values, color=colors)
        ax.set_xticks(list(x))
        ax.set_xticklabels(labels, rotation=40, ha="right", fontsize=8)

    ax.set_title(title, fontsize=10)
    fig.tight_layout()

    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    return buf.getvalue()
