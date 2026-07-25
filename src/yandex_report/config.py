from __future__ import annotations

import os
from dataclasses import dataclass

from dotenv import load_dotenv

load_dotenv()


@dataclass(frozen=True)
class Config:
    token: str
    counter_id: str
    report_title: str

    @classmethod
    def from_env(cls) -> "Config":
        token = os.environ.get("YANDEX_METRIKA_TOKEN", "").strip()
        counter_id = os.environ.get("YANDEX_METRIKA_COUNTER_ID", "").strip()
        if not token:
            raise RuntimeError(
                "Не задан YANDEX_METRIKA_TOKEN. Скопируйте .env.example в .env и заполните."
            )
        if not counter_id:
            raise RuntimeError(
                "Не задан YANDEX_METRIKA_COUNTER_ID. Скопируйте .env.example в .env и заполните."
            )
        return cls(
            token=token,
            counter_id=counter_id,
            report_title=os.environ.get("REPORT_TITLE", "Отчёт по трафику сайта").strip(),
        )
