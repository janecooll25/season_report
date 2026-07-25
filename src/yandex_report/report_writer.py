"""Генерация аналитического текста разделов через Claude (Anthropic API)."""
from __future__ import annotations

import anthropic

from .report_data import DataBlock
from .structure import Section

SYSTEM_PROMPT = (
    "Ты — аналитик Континентальной хоккейной лиги (КХЛ). Пишешь официальный "
    "отчёт по интернет-аудитории для руководства Лиги. Тон — деловой, "
    "сдержанный, без маркетинговых клише и обращений к читателю. Пиши связными "
    "абзацами на русском языке. Опирайся ТОЛЬКО на переданные цифры — не "
    "выдумывай данные, которых нет. Если сравнение с прошлым сезоном невозможно "
    "(нет данных), не утверждай о росте или снижении. Объём — 1–2 абзаца."
)


class ReportWriter:
    def __init__(self, api_key: str, model: str):
        self.client = anthropic.Anthropic(api_key=api_key)
        self.model = model

    def write_section(self, section: Section, data: DataBlock, season: str) -> str:
        heading = section.heading.format(season=season)
        parts = [
            f"Раздел отчёта: «{heading}».",
            f"Сезон: {season}.",
            "",
            "Задача: " + section.guidance,
            "",
            "Данные из Яндекс.Метрики (таблица раздела):",
            data.to_facts(),
        ]
        if section.style_example:
            parts += [
                "",
                "Пример нужного стиля и тональности (НЕ копируй факты из него, "
                "только манеру изложения):",
                f"«{section.style_example}»",
            ]
        parts += [
            "",
            "Напиши аналитический текст для этого раздела. Без заголовка, только "
            "текст абзацев.",
        ]
        prompt = "\n".join(parts)

        message = self.client.messages.create(
            model=self.model,
            max_tokens=2000,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": prompt}],
        )
        if message.stop_reason == "refusal":
            return "[Текст не сгенерирован: запрос отклонён моделью.]"
        return "".join(
            b.text for b in message.content if getattr(b, "type", None) == "text"
        ).strip()
