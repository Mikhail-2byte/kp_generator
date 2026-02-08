# Структура проекта KP Generator

Проект организован по принципам чистой архитектуры с четким разделением ответственности между слоями.

## Полная структура проекта

```
kp_generator/
├── app/                          # Основной код приложения
│   ├── __init__.py              # Фабрика приложения Flask
│   │
│   ├── core/                    # Ядро приложения
│   │   ├── config.py           # Загрузка конфигурации, настройка безопасности и логирования
│   │   ├── extensions.py       # Расширения Flask (CSRF, Login Manager, SQLAlchemy)
│   │   ├── errors.py           # Обработчики HTTP ошибок (404, 500)
│   │   ├── exceptions.py       # Кастомные исключения
│   │   ├── cache.py            # Кэширование данных
│   │   └── redis_session.py    # Настройка Redis сессий
│   │
│   ├── models/                  # Доменные модели
│   │   └── models.py           # ORM модели (UserRecord, GenerationHistoryRecord, AuditLogRecord)
│   │
│   ├── database/                # Работа с базой данных
│   │   ├── __init__.py         # Экспорт функций БД
│   │   ├── database.py         # CRUD операции, история генераций, статистика
│   │   └── service.py          # Сервисы работы с БД
│   │
│   ├── auth/                    # Аутентификация и авторизация
│   │   └── security.py         # Декораторы безопасности (admin_required)
│   │
│   ├── business/                # Бизнес-логика
│   │   ├── price_calculator.py  # Расчет продажных цен
│   │   ├── document_generator.py # Генерация документов (Excel, Word, ZIP)
│   │   └── interfaces.py       # Интерфейсы для бизнес-логики
│   │
│   ├── presentation/            # Слой представления
│   │   ├── forms.py           # Формы WTForms (LoginForm, RegistrationForm, GenerateForm и т.д.)
│   │   ├── ui.py              # UI хелперы (иконки, форматирование, контекст)
│   │   ├── helpers.py         # Вспомогательные функции (валидация, извлечение позиций)
│   │   └── validators.py      # Валидаторы данных
│   │
│   ├── routes/                  # Маршруты (Blueprints)
│   │   ├── __init__.py         # Регистрация всех blueprints
│   │   ├── main.py            # Основные страницы (генерация КП, история, аналитика)
│   │   ├── auth.py            # Авторизация, регистрация, профиль
│   │   ├── admin.py           # Административная панель
│   │   ├── api.py             # REST API endpoints
│   │   ├── api_docs.py        # Swagger документация API
│   │   └── health.py          # Health check endpoints
│   │
│   └── services/                # Сервисы и утилиты
│       ├── repositories.py     # Репозитории для работы с моделями
│       ├── datasets.py         # Управление справочниками (пошлины, материалы, логистика)
│       ├── multi_position_calculator.py # Калькулятор множественных позиций
│       ├── multi_position_processor.py  # Обработка множественных позиций в Excel
│       ├── word_multi_position_processor.py # Обработка множественных позиций в Word
│       ├── generation_orchestrator.py   # Оркестратор генерации КП
│       ├── logistics_calculator.py      # Расчет логистики
│       ├── excel_importer.py            # Импорт позиций из Excel
│       ├── analytics_service.py         # Сервис аналитики
│       ├── analytics_enhancements.py   # Расширенная аналитика
│       ├── audit_service.py             # Аудит действий пользователей
│       ├── content_manager.py           # Управление контентом
│       ├── export_service.py            # Экспорт данных
│       ├── feedback.py                  # Обратная связь
│       ├── healthcheck.py               # Проверка здоровья системы
│       └── datasets_validator.py        # Валидация справочников
│
├── ai_agent/                    # AI консультант
│   ├── agent.py                # Главный класс AIAgent
│   ├── analytics_helper.py     # Работа с данными 1С
│   ├── logistics_helper.py    # Расчет логистики через AI
│   ├── materials_helper.py     # Поиск материалов через AI
│   ├── duty_helper.py          # Расчет пошлин через AI
│   ├── intent_extractor.py     # Извлечение намерений из запросов
│   ├── validators.py           # Валидация параметров
│   ├── formatters.py           # Форматирование ответов
│   ├── suggestions.py          # Генерация предложений
│   ├── cache_manager.py        # Управление кешированием
│   ├── metrics.py              # Сбор метрик
│   ├── feedback.py             # Обратная связь
│   ├── exporters.py            # Экспорт результатов
│   ├── interactive.py          # Интерактивные элементы
│   ├── dashboard.py            # Dashboard с метриками
│   ├── datasource.py           # Абстракция источников данных
│   └── data/                   # Данные для AI агента
│
├── config/                      # Конфигурационные файлы
│   ├── settings.json          # Основные настройки приложения
│   ├── environments/           # Профили окружений
│   │   ├── development.json
│   │   ├── staging.json
│   │   └── production.json
│   ├── gb_materials.json      # Справочник материалов GB
│   ├── logistics_cities.json  # Справочник городов логистики
│   ├── logistics_ekb_rf_cities.json
│   ├── logistics_main_cities.json
│   ├── logistics_trail_cities.json
│   ├── orders_documents.json  # Документы заказов
│   ├── task_templates.json    # Шаблоны задач
│   ├── tnved_catalog.json     # Каталог ТН-ВЭД
│   └── versions/              # Версии справочников (версионирование)
│
├── migrations/                 # Миграции Alembic
│   ├── env.py                 # Конфигурация Alembic
│   ├── script.py.mako         # Шаблон миграции
│   └── versions/              # Файлы миграций
│
├── templates/                   # HTML шаблоны Jinja2
│   ├── base.html              # Базовый шаблон
│   ├── index.html             # Главная страница (генерация КП)
│   ├── history.html           # История генераций
│   ├── analytics.html         # Аналитика
│   ├── profile.html           # Профиль пользователя
│   └── admin/                 # Шаблоны админ-панели
│
├── static/                     # Статические файлы
│   ├── css/
│   │   └── style.css          # Основные стили
│   └── instructions/          # Инструкции
│
├── templates_docs/             # Шаблоны документов
│   ├── template.xlsx          # Шаблон Excel для КП
│   └── template.docx          # Шаблон Word для КП
│
├── tests/                      # Тесты
│   ├── conftest.py            # Фикстуры pytest
│   ├── test_smoke_health.py   # Smoke тесты
│   ├── business/              # Тесты бизнес-логики
│   ├── routes/                # Тесты маршрутов
│   └── services/             # Тесты сервисов
│
├── scripts/                    # Утилиты управления
│   ├── manage_users.py        # Управление пользователями
│   └── manage_migrations.py  # Управление миграциями
│
├── logs/                       # Логи (создается автоматически)
│
├── docs/                       # Документация
│   ├── SETUP.md               # Установка и настройка
│   ├── COMPLETE_GUIDE.md      # Полное руководство
│   ├── PROJECT_STRUCTURE.md   # Структура проекта (этот файл)
│   ├── USER_MANAGEMENT.md     # Управление пользователями
│   ├── REDIS_SETUP.md         # Настройка Redis
│   ├── AI_AGENT_SETUP.md      # Настройка AI агента
│   ├── AI_AGENT_USER_GUIDE.md # Руководство пользователя AI
│   └── CHANGELOG.md           # Журнал изменений
│
├── app.py                      # Точка входа приложения
├── alembic.ini                 # Конфигурация Alembic
├── requirements.txt           # Зависимости проекта
├── requirements-dev.txt       # Зависимости для разработки
├── pyproject.toml             # Настройки Black и Ruff
├── pytest.ini                 # Конфигурация pytest
└── README.md                   # Основной README проекта
```

## Описание основных директорий

### `app/core/` - Ядро приложения
Содержит базовую конфигурацию и расширения Flask:
- **config.py**: Загрузка конфигурации из файлов и переменных окружения, настройка безопасности и логирования
- **extensions.py**: Инициализация CSRF защиты, Login Manager и SQLAlchemy
- **errors.py**: Обработчики HTTP ошибок (404, 500, 503)
- **exceptions.py**: Кастомные исключения приложения
- **cache.py**: Кэширование данных справочников
- **redis_session.py**: Настройка Redis для хранения сессий Flask

### `app/models/` - Доменные модели
ORM модели для работы с базой данных:
- **models.py**: Определение таблиц и моделей:
  - `UserRecord` - пользователи системы
  - `GenerationHistoryRecord` - история генераций КП
  - `AuditLogRecord` - логи аудита действий пользователей

### `app/database/` - Работа с базой данных
Функции для работы с БД:
- **database.py**: CRUD операции, работа с историей генераций, статистика пользователей
- **service.py**: Сервисы работы с БД, инициализация схемы

### `app/auth/` - Аутентификация и авторизация
Безопасность и права доступа:
- **security.py**: Декораторы для проверки прав доступа (`admin_required`)

### `app/business/` - Бизнес-логика
Основная бизнес-логика приложения:
- **price_calculator.py**: Расчет продажных цен с учетом всех параметров (закуп, логистика, пошлина, маржа)
- **document_generator.py**: Генерация Excel и Word документов, создание ZIP архивов
- **interfaces.py**: Интерфейсы для бизнес-логики

### `app/presentation/` - Слой представления
Компоненты для работы с пользовательским интерфейсом:
- **forms.py**: Определение форм WTForms (LoginForm, RegistrationForm, GenerateForm и др.)
- **ui.py**: Вспомогательные функции для шаблонов (иконки, форматирование чисел, контекст)
- **helpers.py**: Валидация данных форм, извлечение позиций из формы, работа с файлами
- **validators.py**: Кастомные валидаторы для форм

### `app/routes/` - Маршруты (Blueprints)
Blueprint'ы для организации маршрутов:
- **main.py**: Основные страницы (генерация КП, история, аналитика)
- **auth.py**: Авторизация, регистрация, профиль пользователя
- **admin.py**: Административная панель (управление справочниками, пользователями, аудит)
- **api.py**: REST API endpoints для интеграции
- **api_docs.py**: Swagger документация API
- **health.py**: Health check endpoints для мониторинга

### `app/services/` - Сервисы и утилиты
Дополнительные сервисы и утилиты:
- **repositories.py**: Репозитории для работы с моделями (паттерн Repository)
- **datasets.py**: Управление справочниками (пошлины, материалы GB, логистика) с версионированием
- **multi_position_calculator.py**: Калькулятор множественных позиций с единой маржой
- **multi_position_processor.py**: Обработка множественных позиций в Excel документах
- **word_multi_position_processor.py**: Обработка множественных позиций в Word документах
- **generation_orchestrator.py**: Оркестратор генерации КП (валидация, расчет, генерация документов)
- **logistics_calculator.py**: Расчет логистики с учетом типа транспорта и расстояния
- **excel_importer.py**: Импорт позиций из Excel файлов
- **analytics_service.py**: Сервис аналитики (графики, статистика)
- **analytics_enhancements.py**: Расширенная аналитика (маржинальность, динамика курсов)
- **audit_service.py**: Аудит действий пользователей
- **content_manager.py**: Управление контентом (заказы, инструкции, шаблоны)
- **export_service.py**: Экспорт данных в различные форматы
- **feedback.py**: Обратная связь от пользователей
- **healthcheck.py**: Проверка здоровья системы (БД, шаблоны, справочники)
- **datasets_validator.py**: Валидация справочников при загрузке

### `ai_agent/` - AI консультант
Модуль AI агента для консультаций:
- **agent.py**: Главный класс AIAgent, оркестратор всех функций
- **analytics_helper.py**: Работа с данными 1С (CSV)
- **logistics_helper.py**: Расчет логистики через AI
- **materials_helper.py**: Поиск материалов через AI
- **duty_helper.py**: Расчет пошлин через AI
- **intent_extractor.py**: Извлечение намерений из запросов пользователей
- **validators.py**: Валидация параметров запросов
- **formatters.py**: Форматирование ответов AI
- **suggestions.py**: Генерация предложений и автодополнения
- **cache_manager.py**: Управление кешированием ответов AI
- **metrics.py**: Сбор метрик производительности
- **feedback.py**: Обратная связь по ответам AI
- **exporters.py**: Экспорт результатов в различные форматы
- **interactive.py**: Интерактивные элементы (графики, таблицы)
- **dashboard.py**: Dashboard с метриками использования AI
- **datasource.py**: Абстракция источников данных (CSV/БД)

### `config/` - Конфигурационные файлы
- **settings.json**: Основные настройки приложения (константы расчета, пагинация, Redis)
- **environments/**: Профили окружений (development, staging, production)
- **gb_materials.json**: Справочник материалов GB
- **logistics_*.json**: Справочники городов логистики
- **tnved_catalog.json**: Каталог ТН-ВЭД
- **versions/**: Версии справочников для версионирования изменений

## Принципы организации

### Разделение по слоям (Clean Architecture)

1. **Core Layer** (`app/core/`): Инфраструктурный слой - конфигурация, расширения Flask, обработка ошибок
2. **Domain Layer** (`app/models/`): Доменные модели - сущности бизнес-логики
3. **Data Layer** (`app/database/`): Слой данных - работа с БД, репозитории
4. **Business Layer** (`app/business/`): Бизнес-логика - расчеты, генерация документов
5. **Presentation Layer** (`app/presentation/`, `app/routes/`): Слой представления - формы, маршруты, UI

### Преимущества структуры

1. **Понятная группировка**: Файлы сгруппированы по функциональному назначению
2. **Легкая навигация**: Легко найти нужный файл по его назначению
3. **Разделение ответственности**: Четкое разделение на слои
4. **Масштабируемость**: Легко добавлять новые файлы в соответствующие директории
5. **Тестируемость**: Каждый слой можно тестировать независимо
6. **Следование принципам**: Структура соответствует принципам SOLID и KISS

## Основные паттерны

### Repository Pattern
Используется в `app/services/repositories.py` для абстракции работы с БД:
- `user_repository` - работа с пользователями
- `generation_repository` - работа с историей генераций

### Service Layer Pattern
Сервисы в `app/services/` инкапсулируют бизнес-логику:
- `generation_orchestrator` - оркестрация генерации КП
- `logistics_calculator` - расчет логистики
- `analytics_service` - аналитика и отчеты

### Factory Pattern
Используется для создания экземпляров:
- `create_app()` - фабрика приложения Flask
- `DataSourceFactory` - создание источников данных для AI агента

## Поток данных при генерации КП

```
1. Пользователь заполняет форму (routes/main.py)
   ↓
2. Валидация данных (presentation/helpers.py, presentation/validators.py)
   ↓
3. Оркестрация генерации (services/generation_orchestrator.py)
   ├─ Расчет цен (business/price_calculator.py, services/multi_position_calculator.py)
   ├─ Расчет логистики (services/logistics_calculator.py)
   └─ Генерация документов (business/document_generator.py)
      ├─ Excel (services/multi_position_processor.py)
      └─ Word (services/word_multi_position_processor.py)
   ↓
4. Сохранение в историю (database/database.py)
   ↓
5. Возврат результата пользователю
```

## Зависимости между слоями

```
routes/ → presentation/ → business/ → services/ → database/ → models/
   ↓         ↓              ↓           ↓            ↓
templates/  forms/      calculators/  repositories/  ORM models
```

Каждый слой зависит только от слоев ниже него, что обеспечивает:
- Низкую связанность (Low Coupling)
- Высокую связность внутри слоя (High Cohesion)
- Легкость тестирования
- Возможность замены реализации без изменения других слоев
