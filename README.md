# KP Generator

Веб‑приложение на Flask для расчёта коммерческих предложений (КП), генерации Word/Excel документов, ведения истории расчётов, работы со справочниками (пошлины, материалы GB, логистические города) и простого личного кабинета с ролями.

## Возможности
- Расчёт цены для одной и множественных позиций с единой целевой маржей
- Генерация документов (Excel + Word) и архивирование в ZIP
- История генераций с подробностями и повторной подстановкой
- Справочники (пошлины, материалы GB, логистика) с админ‑редактированием и версионированием
- Аналитика Excel (тендеры/продукция): предпросмотр, графики, выгрузки
- Аутентификация, профиль, базовая статистика

## Быстрый старт (Windows)
1) Python 3.12, PowerShell:
```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
alembic upgrade head
python app.py
```
2) Продакшн (Waitress):
```powershell
$env:USE_WAITRESS = "1"
python app.py
```
или
```powershell
waitress-serve --call "app:create_app"
```

## Структура
- `app/` — код приложения (роуты, сервисы, модели, БД, UI)
- `templates/`, `static/`, `templates_docs/` — фронт, ресурсы и шаблоны документов
- `config/` — JSON‑справочники и версии (`config/versions`)
- `migrations/` — миграции Alembic
- `tests/` — базовые тесты
- `docs/` — документация и журнал изменений

## Документация
- Руководство по установке: `docs/SETUP.md`
- Интеграция множественных позиций: `docs/MULTI_POSITION_INTEGRATION.md`
- Журнал изменений: `docs/CHANGELOG.md`

## Стиль и качество
- Форматирование: Black (`pyproject.toml`)
- Линтинг: Ruff (`pyproject.toml`)

## Конфигурация
- `.env` и `config/settings.json` (переменные, секреты, БД)
- По умолчанию SQLite: `sqlite:///kp_generator.db`
- Профили окружений: `config/environments/{development,staging,production}.json`
  - Выбор профиля переменной `APP_ENV` (по умолчанию `development`)
  - Значения `DATABASE_URL`, `SECRET_KEY`, `LOG_LEVEL`, `DEBUG` можно задать в профиле или через переменные окружения

## Миграции
- Статус/применение:
  - `python manage_migrations.py status`
  - `python manage_migrations.py upgrade [revision]`
  - `python manage_migrations.py downgrade <revision>`
  - `python manage_migrations.py history [--verbose]`
- Приложение при старте автоматически проверяет возможность применения и актуальность миграций Alembic; при несоответствии схема не запускается.

## Health-check
- `GET /health` или `/healthz` — JSON-отчёт о состоянии:
  - подключение к БД и актуальность миграций;
  - наличие шаблонов документов;
  - доступность справочников и кэшированных данных;
  - существование обязательных каталогов (`logs`, `templates_docs`).
- При сбоях возвращается HTTP 503 с деталями по каждому чек-пойнту.


