# season_report — генератор отчётов из Яндекс.Метрики в docx

Сервис забирает статистику из Reporting API Яндекс.Метрики и заполняет
docx-шаблон (плейсхолдеры в стиле Jinja2 через [docxtpl](https://docxtpl.readthedocs.io/)).

## Что попадает в отчёт

- Сводка: визиты, пользователи, просмотры, отказы, средняя длительность визита, глубина просмотра
- Топ источников трафика (по визитам)
- Топ страниц входа
- Разбивка по типам устройств

## Установка

```bash
pip install -r requirements.txt
cp .env.example .env   # затем впишите токен и ID счётчика
```

Получить OAuth-токен: https://oauth.yandex.ru/ (приложению нужен доступ к API Метрики).

## Запуск

```bash
# за последние 30 дней, шаблон по умолчанию
PYTHONPATH=src python -m yandex_report.cli

# свой период и свой шаблон
PYTHONPATH=src python -m yandex_report.cli \
  --date-from 2026-06-01 --date-to 2026-06-30 \
  --template templates/my_template.docx \
  --output output/june.docx \
  --title "Отчёт за июнь"
```

Готовый файл появится в `output/`.

## Свой шаблон

Положите свой `.docx` в `templates/` и укажите путь через `--template`.
В шаблоне доступны такие плейсхолдеры:

| Тег | Значение |
|-----|----------|
| `{{ title }}` | Заголовок отчёта |
| `{{ counter_id }}` | ID счётчика |
| `{{ date_from }}` / `{{ date_to }}` | Границы периода |
| `{{ generated_at }}` | Дата формирования |
| `{{ visits }}` `{{ users }}` `{{ pageviews }}` | Сводные числа |
| `{{ bounce_rate }}` `{{ avg_duration }}` `{{ page_depth }}` | Сводные показатели |

Для таблиц используется row-цикл docxtpl. **Важно:** тег `{%tr ... %}` удаляет
строку целиком, поэтому `for`/`endfor` ставятся в отдельные строки-обёртки вокруг
строки с данными:

```
| {%tr for r in sources %}          |            |               |          |
| {{ r.name }} | {{ r.visits }}     | {{ r.users }} | {{ r.share }}          |
| {%tr endfor %}                    |            |               |          |
```

Доступные списки: `sources`, `top_pages`, `devices`. У каждой строки — поля
`name`, `visits`, `users`, `share`.

Пример-шаблон генерируется командой:

```bash
python scripts/make_sample_template.py
```

## Структура

```
src/yandex_report/
  config.py          # чтение .env
  metrika_client.py  # клиент Reporting API
  report_data.py     # запросы + агрегация в структуру отчёта
  docx_builder.py    # заполнение docx-шаблона
  cli.py             # точка входа
scripts/make_sample_template.py  # генератор примера-шаблона
templates/           # docx-шаблоны
output/              # готовые отчёты (в git не попадают)
tests/               # тесты на агрегацию и рендер
```

## Тесты

```bash
PYTHONPATH=src python -m pytest -q
```
