from __future__ import annotations

import os
from dataclasses import dataclass

from dotenv import load_dotenv

load_dotenv()

DEFAULT_MODEL = "claude-opus-5"


@dataclass(frozen=True)
class Config:
    token: str
    counter_id: str
    anthropic_key: str
    model: str

    @classmethod
    def from_env(cls, *, require_llm: bool = True) -> "Config":
        token = os.environ.get("YANDEX_METRIKA_TOKEN", "").strip()
        counter_id = os.environ.get("YANDEX_METRIKA_COUNTER_ID", "").strip()
        anthropic_key = os.environ.get("ANTHROPIC_API_KEY", "").strip()

        if not token:
            raise RuntimeError(
                "Не задан YANDEX_METRIKA_TOKEN. Скопируйте .env.example в .env и заполните."
            )
        if not counter_id:
            raise RuntimeError(
                "Не задан YANDEX_METRIKA_COUNTER_ID. Скопируйте .env.example в .env и заполните."
            )
        if require_llm and not anthropic_key:
            raise RuntimeError(
                "Не задан ANTHROPIC_API_KEY — он нужен для генерации аналитического текста. "
                "Заполните .env или запустите с --no-llp для черновика без текста."
            )
        return cls(
            token=token,
            counter_id=counter_id,
            anthropic_key=anthropic_key,
            model=os.environ.get("REPORT_MODEL", DEFAULT_MODEL).strip() or DEFAULT_MODEL,
        )
