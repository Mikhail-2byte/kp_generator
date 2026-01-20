# Структура проекта

Проект реорганизован для лучшей навигации и понимания назначения файлов.

## Новая структура директорий

```
app/
├── core/                    # Ядро приложения
│   ├── config.py           # Конфигурация приложения (настройки, логирование, безопасность)
│   ├── extensions.py       # Расширения Flask (CSRF, Login Manager, SQLAlchemy)
│   └── errors.py           # Обработчики ошибок HTTP (404, 500)
│
├── models/                  # Доменные модели
│   └── models.py           # ORM модели (UserRecord, GenerationHistoryRecord, User)
│
├── database/               # Работа с базой данных
│   ├── __init__.py         # Экспорт функций БД
│   └── database.py         # Операции с БД (CRUD, история генераций, статистика)
│
├── auth/                   # Аутентификация и авторизация
│   └── security.py        # Декораторы безопасности (admin_required)
│
├── business/               # Бизнес-логика
│   ├── price_calculator.py    # Расчет цен (calculate_selling_price)
│   └── document_generator.py  # Генерация документов (Excel, Word, ZIP)
│
├── presentation/           # Слой представления
│   ├── forms.py           # Формы WTForms (LoginForm, RegistrationForm, и т.д.)
│   ├── ui.py              # UI хелперы (иконки, форматирование, контекст)
│   └── helpers.py         # Вспомогательные функции (валидация, извлечение позиций)
│
├── routes/                 # Маршруты (Blueprint)
│   ├── main.py            # Основные страницы (генерация КП, история)
│   ├── auth.py            # Авторизация и профиль
│   └── admin.py           # Административная панель
│
└── services/               # Сервисы
    ├── analytics_service.py
    ├── content_manager.py
    ├── datasets.py
    ├── excel_importer.py
    ├── feedback.py
    ├── multi_position_calculator.py
    ├── multi_position_processor.py
    ├── repositories.py
    └── word_multi_position_processor.py
```

## Описание директорий

### `core/` - Ядро приложения
Содержит базовую конфигурацию и расширения Flask:
- **config.py**: Загрузка конфигурации из файлов и переменных окружения, настройка безопасности и логирования
- **extensions.py**: Инициализация CSRF защиты, Login Manager и SQLAlchemy
- **errors.py**: Обработчики HTTP ошибок (404, 500)

### `models/` - Доменные модели
ORM модели для работы с базой данных:
- **models.py**: Определение таблиц и моделей (UserRecord, GenerationHistoryRecord, User)

### `database/` - Работа с базой данных
Функции для работы с БД:
- **database.py**: CRUD операции, работа с историей генераций, статистика пользователей

### `auth/` - Аутентификация и авторизация
Безопасность и права доступа:
- **security.py**: Декораторы для проверки прав доступа (admin_required)

### `business/` - Бизнес-логика
Основная бизнес-логика приложения:
- **price_calculator.py**: Расчет продажных цен с учетом всех параметров
- **document_generator.py**: Генерация Excel и Word документов, создание ZIP архивов

### `presentation/` - Слой представления
Компоненты для работы с пользовательским интерфейсом:
- **forms.py**: Определение форм WTForms для всех страниц
- **ui.py**: Вспомогательные функции для шаблонов (иконки, форматирование чисел, контекст)
- **helpers.py**: Валидация данных форм, извлечение позиций, работа с файлами

### `routes/` - Маршруты
Blueprint'ы для организации маршрутов:
- **main.py**: Основные страницы (генерация КП, история, аналитика)
- **auth.py**: Авторизация, регистрация, профиль
- **admin.py**: Административная панель (управление справочниками)

### `services/` - Сервисы
Дополнительные сервисы и утилиты:
- Различные сервисы для обработки данных, импорта, аналитики и т.д.

## Преимущества новой структуры

1. **Понятная группировка**: Файлы сгруппированы по функциональному назначению
2. **Легкая навигация**: Легко найти нужный файл по его назначению
3. **Разделение ответственности**: Четкое разделение на слои (core, models, database, business, presentation)
4. **Масштабируемость**: Легко добавлять новые файлы в соответствующие директории
5. **Следование принципам**: Структура соответствует принципам SOLID и KISS

## Миграция импортов

Все импорты обновлены для работы с новой структурой:
- `app.config` → `app.core.config`
- `app.extensions` → `app.core.extensions`
- `app.errors` → `app.core.errors`
- `app.models` → `app.models.models`
- `app.database` → `app.database` (через __init__.py)
- `app.security` → `app.auth.security`
- `app.calculate` → `app.business.price_calculator`
- `app.document_generator` → `app.business.document_generator`
- `app.forms` → `app.presentation.forms`
- `app.ui` → `app.presentation.ui`
- `app.helpers` → `app.presentation.helpers`

