from __future__ import annotations

from typing import Any

import requests

API_URL = "https://api-metrika.yandex.net/stat/v1/data"
REQUEST_TIMEOUT = 30


class MetrikaError(RuntimeError):
    """Ошибка обращения к API Яндекс.Метрики."""


class MetrikaClient:
    """Тонкая обёртка над Reporting API Яндекс.Метрики (stat/v1/data)."""

    def __init__(self, token: str, counter_id: str, session: requests.Session | None = None):
        self.counter_id = counter_id
        self.session = session or requests.Session()
        self.session.headers.update({"Authorization": f"OAuth {token}"})

    def query(
        self,
        metrics: str,
        dimensions: str | None,
        date1: str,
        date2: str,
        *,
        limit: int = 100,
        sort: str | None = None,
    ) -> dict[str, Any]:
        params: dict[str, Any] = {
            "ids": self.counter_id,
            "metrics": metrics,
            "date1": date1,
            "date2": date2,
            "limit": limit,
            "accuracy": "full",
        }
        if dimensions:
            params["dimensions"] = dimensions
        if sort:
            params["sort"] = sort

        try:
            resp = self.session.get(API_URL, params=params, timeout=REQUEST_TIMEOUT)
        except requests.RequestException as exc:
            raise MetrikaError(f"Сетевая ошибка при запросе к Метрике: {exc}") from exc

        if resp.status_code != 200:
            raise MetrikaError(
                f"Метрика вернула {resp.status_code}: {resp.text[:500]}"
            )
        return resp.json()
