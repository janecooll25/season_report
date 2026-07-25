from __future__ import annotations

import os
from typing import Any

import requests

API_URL = "https://api-metrika.yandex.net/stat/v1/data"
COMPARISON_URL = "https://api-metrika.yandex.net/stat/v1/data/comparison"

# accuracy=full над периодом в сезон считается очень долго — по умолчанию
# берём medium (быстро, достаточно для отчёта). Переопределяется через env.
DEFAULT_ACCURACY = os.environ.get("METRIKA_ACCURACY", "medium").strip() or "medium"
REQUEST_TIMEOUT = int(os.environ.get("METRIKA_TIMEOUT", "50"))


class MetrikaError(RuntimeError):
    """Ошибка обращения к API Яндекс.Метрики."""


class MetrikaClient:
    """Тонкая обёртка над Reporting API Яндекс.Метрики (stat/v1/data)."""

    def __init__(
        self,
        token: str,
        counter_id: str,
        session: requests.Session | None = None,
        *,
        accuracy: str = DEFAULT_ACCURACY,
        timeout: int = REQUEST_TIMEOUT,
    ):
        self.counter_id = counter_id
        self.session = session or requests.Session()
        self.session.headers.update({"Authorization": f"OAuth {token}"})
        self.accuracy = accuracy
        self.timeout = timeout

    def query(
        self,
        metrics: str,
        dimensions: str | None,
        date1: str,
        date2: str,
        *,
        limit: int = 100,
        sort: str | None = None,
        filters: str | None = None,
    ) -> dict[str, Any]:
        params: dict[str, Any] = {
            "ids": self.counter_id,
            "metrics": metrics,
            "date1": date1,
            "date2": date2,
            "limit": limit,
            "accuracy": self.accuracy,
        }
        if dimensions:
            params["dimensions"] = dimensions
        if sort:
            params["sort"] = sort
        if filters:
            params["filters"] = filters

        try:
            resp = self.session.get(API_URL, params=params, timeout=self.timeout)
        except requests.RequestException as exc:
            raise MetrikaError(f"Сетевая ошибка при запросе к Метрике: {exc}") from exc

        if resp.status_code != 200:
            raise MetrikaError(
                f"Метрика вернула {resp.status_code}: {resp.text[:500]}"
            )
        return resp.json()

    def query_comparison(
        self,
        metrics: str,
        dimensions: str | None,
        a1: str,
        a2: str,
        b1: str,
        b2: str,
        *,
        limit: int = 100,
        sort: str | None = None,
        filters: str | None = None,
    ) -> dict[str, Any]:
        """Сравнение двух периодов (A — текущий, B — прошлый) за один запрос.

        В ответе metrics и totals становятся парой [[значения A], [значения B]].
        """
        params: dict[str, Any] = {
            "ids": self.counter_id,
            "metrics": metrics,
            "date1_a": a1,
            "date2_a": a2,
            "date1_b": b1,
            "date2_b": b2,
            "limit": limit,
            "accuracy": self.accuracy,
        }
        if dimensions:
            params["dimensions"] = dimensions
        if sort:
            params["sort"] = sort
        if filters:
            params["filters"] = filters

        try:
            resp = self.session.get(COMPARISON_URL, params=params, timeout=self.timeout)
        except requests.RequestException as exc:
            raise MetrikaError(f"Сетевая ошибка при запросе к Метрике: {exc}") from exc

        if resp.status_code != 200:
            raise MetrikaError(
                f"Метрика вернула {resp.status_code}: {resp.text[:500]}"
            )
        return resp.json()
