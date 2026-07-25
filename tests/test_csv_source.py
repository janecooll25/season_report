from __future__ import annotations

import pytest

from yandex_report.csv_source import CsvError, load_blocks_from_files

SUMMARY_CUR = (
    '"Интервал дат визита","Визиты","Посетители","Просмотры",'
    '"Доля новых посетителей","Отказы","Глубина просмотра","Время на сайте"\n'
    '"Итого и средние","95034161","20226924","259998808","0.98","0.17","2.74","00:03:45"\n'
    '"2025-09-01 - 2025-09-30","7030485","2064080","21306597","0.8","0.14","3.0","00:04:15"\n'
)
SUMMARY_PREV = (
    '"Интервал дат визита","Визиты","Посетители","Просмотры",'
    '"Доля новых посетителей","Отказы","Глубина просмотра","Время на сайте"\n'
    '"Итого и средние","80720891","17561513","244023289","0.99","0.13","3.02","00:04:16"\n'
)
COUNTRIES_CUR = (
    '"Страна","Визиты","Посетители","Отказы","Глубина просмотра","Время на сайте"\n'
    '"Итого и средние","95034161","20226924","0.17","2.74","00:03:45"\n'
    '"Россия","73526805","14738936","0.15","2.82","00:03:52"\n'
    '"Беларусь","10332715","1170346","0.14","2.17","00:03:22"\n'
)
COUNTRIES_PREV = (
    '"Страна","Визиты","Посетители","Отказы","Глубина просмотра","Время на сайте"\n'
    '"Итого и средние","80720891","17561513","0.13","3.02","00:04:16"\n'
    '"Россия","67031320","14386415","0.12","3.05","00:04:18"\n'
    '"Беларусь","6238830","864621","0.11","2.34","00:03:30"\n'
)


def _files():
    return [
        ("summary_2025070120260630.csv", SUMMARY_CUR),
        ("summary_2024070120250630.csv", SUMMARY_PREV),
        ("countries_2025070120260630.csv", COUNTRIES_CUR),
        ("countries_2024070120250630.csv", COUNTRIES_PREV),
    ]


def test_season_and_current_from_filename_dates():
    blocks, season = load_blocks_from_files(_files())
    assert season == "2025/2026"
    # текущий сезон — 95M визитов, рост к прошлым 80.7M
    assert blocks["summary"].rows[0] == ["Визиты", "95034161", "+17.7%"]


def test_countries_share_and_change():
    blocks, _ = load_blocks_from_files(_files())
    row = blocks["geo_countries"].rows[0]
    assert row[0] == "Россия"
    assert row[1] == "73526805"
    assert row[2] == "77.4%"          # доля от 95M визитов
    assert row[3] == "+9.7%"          # рост к 67.0M прошлого сезона
    # строка «Итого и средние» не попадает в данные
    assert all(r[0] != "Итого и средние" for r in blocks["geo_countries"].rows)


def test_time_reformatted():
    blocks, _ = load_blocks_from_files(_files())
    dur = [r for r in blocks["summary"].rows if r[0].startswith("Средняя")][0]
    assert dur[1] == "3 мин 45 сек"


def test_unknown_files_raise():
    with pytest.raises(CsvError):
        load_blocks_from_files([("junk.csv", '"Нечто","1"\n"a","2"\n')])
