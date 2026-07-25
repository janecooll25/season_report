from __future__ import annotations

import importlib.util
import os
from pathlib import Path

import pytest

from test_report import FakeClient

ROOT = Path(__file__).resolve().parent.parent


def _load_api():
    spec = importlib.util.spec_from_file_location("gen", ROOT / "api" / "generate.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def test_generate_no_llm(monkeypatch):
    monkeypatch.setenv("YANDEX_METRIKA_TOKEN", "t")
    monkeypatch.setenv("YANDEX_METRIKA_COUNTER_ID", "6571768")
    api = _load_api()
    # Подменяем клиент Метрики на фейковый — без сети.
    monkeypatch.setattr(api, "MetrikaClient", lambda *a, **k: FakeClient())

    filename, data = api._generate({"llm": ["0"], "season": ["2024/2025"]})
    assert filename == "khl_report_2024_2025.docx"
    assert data[:2] == b"PK" and len(data) > 1000


def test_generate_requires_token(monkeypatch):
    monkeypatch.delenv("YANDEX_METRIKA_TOKEN", raising=False)
    monkeypatch.delenv("YANDEX_METRIKA_COUNTER_ID", raising=False)
    api = _load_api()
    with pytest.raises(RuntimeError):
        api._generate({"llm": ["0"]})
