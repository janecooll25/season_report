"""Генерация аналитического текста разделов через Claude (Anthropic API)."""
from __future__ import annotations

import asyncio
import os

import anthropic

from .report_data import DataBlock
from .structure import Section

# Короткие абзацы по готовым цифрам не требуют глубокого рассуждения —
# low ускоряет генерацию (важно для лимита времени на Vercel).
EFFORT = os.environ.get("REPORT_EFFORT", "low").strip() or "low"

SYSTEM_PROMPT = (
    "Ты — аналитик Континентальной хоккейной лиги (КХЛ). Пишешь официальный "
    "отчёт по интернет-аудитории для руководства Лиги. Тон — деловой, "
    "сдержанный, без маркетинговых клише и обращений к читателю. Пиши связными "
    "абзацами на русском языке. Опирайся ТОЛЬКО на переданные цифры — не "
    "выдумывай данные, которых нет. Если сравнение с прошлым сезоном невозможно "
    "(нет данных), не утверждай о росте или снижении. Объём — 1–2 абзаца."
)


def _build_prompt(section: Section, data: DataBlock, season: str) -> str:
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
    return "\n".join(parts)


def _extract(message) -> str:
    if message.stop_reason == "refusal":
        return "[Текст не сгенерирован: запрос отклонён моделью.]"
    return "".join(
        b.text for b in message.content if getattr(b, "type", None) == "text"
    ).strip()


class ReportWriter:
    def __init__(self, api_key: str, model: str):
        self.client = anthropic.Anthropic(api_key=api_key)
        self.model = model

    def write_section(self, section: Section, data: DataBlock, season: str) -> str:
        message = self.client.messages.create(
            model=self.model,
            max_tokens=2000,
            system=SYSTEM_PROMPT,
            output_config={"effort": EFFORT},
            messages=[{"role": "user", "content": _build_prompt(section, data, season)}],
        )
        return _extract(message)


async def write_sections_parallel(
    api_key: str,
    model: str,
    sections: list[Section],
    blocks: dict[str, DataBlock],
    season: str,
) -> dict[str, str]:
    """Пишет все разделы параллельно — критично для serverless-лимитов времени."""
    client = anthropic.AsyncAnthropic(api_key=api_key)

    async def one(section: Section) -> tuple[str, str]:
        block = blocks.get(section.id)
        if not block or not block.rows:
            return section.id, ""
        try:
            message = await client.messages.create(
                model=model,
                max_tokens=2000,
                system=SYSTEM_PROMPT,
                output_config={"effort": EFFORT},
                messages=[
                    {"role": "user", "content": _build_prompt(section, block, season)}
                ],
            )
            return section.id, _extract(message)
        except anthropic.APIError as exc:
            return section.id, f"[Текст не сгенерирован: {exc.__class__.__name__}.]"

    results = await asyncio.gather(*(one(s) for s in sections))
    await client.close()
    return {sid: text for sid, text in results if text}
