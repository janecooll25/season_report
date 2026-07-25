from __future__ import annotations

import os
import time
from typing import Any

import requests

API_URL = "https://api-metrika.yandex.net/stat/v1/data"

# accuracy=full над периодом в сезон считается очень долго — по умолчанию
# берём medium (быстро, достаточно для отчёта). Переопределяется через env.
DEFAULT_ACCURACY = os.environ.get("METRIKA_ACCURACY", "medium").strip() or "medium"
REQUEST_TIMEOUT = int(os.environ.get("METRIKA_TIMEOUT", "50"))
MAX_RETRIES = int(os.environ.get("METRIKA_RETRIES", "5"))


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

        # Ретраи на 429 (лимит параллельных запросов) и 5xx — с бэкоффом.
        last_status = None
        last_text = ""
        for attempt in range(MAX_RETRIES):
            try:
                resp = self.session.get(API_URL, params=params, timeout=self.timeout)
            except requests.RequestException as exc:
                raise MetrikaError(
                    f"Сетевая ошибка при запросе к Метрике: {exc}"
                ) from exc

            if resp.status_code == 200:
                return resp.json()

            last_status, last_text = resp.status_code, resp.text[:500]
            if resp.status_code in (429, 503) and attempt < MAX_RETRIES - 1:
                retry_after = resp.headers.get("Retry-After")
                wait = float(retry_after) if retry_after else min(0.5 * 2 ** attempt, 8)
                time.sleep(wait)
                continue
            break

        raise MetrikaError(f"Метрика вернула {last_status}: {last_text}")
