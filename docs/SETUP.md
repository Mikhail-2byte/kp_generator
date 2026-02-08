# Установка и запуск KP Generator

Подробное руководство по установке и настройке системы генерации коммерческих предложений.

## Требования

- **Python**: 3.12 или выше
- **ОС**: Windows 10/11, Linux, macOS
- **База данных**: Microsoft SQL Server (обязательно). Пошаговая настройка: [MSSQL_SETUP.md](MSSQL_SETUP.md)
- **Redis**: Опционально, для хранения сессий (рекомендуется)

## Шаг 1: Подготовка окружения

### 1.1 Клонирование репозитория

```bash
git clone <repository-url>
cd kp_generator
```

### 1.2 Создание виртуального окружения

**Windows (PowerShell):**
```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
```

**Windows (CMD):**
```cmd
python -m venv venv
venv\Scripts\activate.bat
```

**Linux/macOS:**
```bash
python -m venv venv
source venv/bin/activate
```

### 1.3 Установка зависимостей

```bash
pip install -r requirements.txt
```

Для разработки дополнительно:
```bash
pip install -r requirements-dev.txt
```

## Шаг 2: Настройка конфигурации

### 2.1 Создание файла `.env`

Создайте файл `.env` в корне проекта со следующим содержимым:

```env
# Окружение
APP_ENV=development
FLASK_ENV=development
FLASK_DEBUG=True

# База данных — обязательно (см. docs/MSSQL_SETUP.md)
DATABASE_URL=mssql+pyodbc://USER:PASSWORD@host:1433/kp_generator?driver=ODBC+Driver+17+for+SQL+Server

# Секретный ключ (сгенерируйте случайную строку)
SECRET_KEY=your-secret-key-change-in-production

# Сервер
FLASK_RUN_HOST=0.0.0.0
FLASK_RUN_PORT=5000

# Логирование
LOG_LEVEL=INFO

# Redis (опционально, для сессий)
REDIS_HOST=localhost
REDIS_PORT=6379
REDIS_DB=0
REDIS_PASSWORD=

# Production режим
USE_WAITRESS=0
```

**Важно:** Для production обязательно измените `SECRET_KEY` на случайную строку!

### 2.2 Настройка `config/settings.json`

Основные настройки приложения находятся в `config/settings.json`:

```json
{
  "margin_percent": 20,
  "default_duty_percent": 0,
  "history_page_size": 25,
  "calculation_constants": {
    "conversion_rate": 11.5,
    "logistics_cnr_ratio": 0.3,
    "logistics_rf_ratio": 0.7,
    "conversion_fee_rate": 0.032,
    "credit_rate": 0.16,
    "vat_rate": 0.22
  }
}
```

## Шаг 3: Настройка базы данных

### 3.1 Применение миграций

```bash
alembic upgrade head
```

Проверка статуса миграций:
```bash
alembic current
alembic history
```

### 3.2 Создание первого пользователя

После запуска приложения:
1. Откройте `http://localhost:5000/profile`
2. Зарегистрируйтесь (первый пользователь автоматически становится администратором)
3. Или используйте скрипт управления пользователями:
   ```bash
   python scripts/manage_users.py list
   ```

## Шаг 4: Настройка Redis (опционально, но рекомендуется)

Redis используется для хранения сессий Flask, что решает проблему с большими cookie.

### Windows

**Вариант 1: Docker Desktop**
```powershell
docker run -d -p 6379:6379 --name redis redis:latest
```

**Вариант 2: Memurai**
Скачайте и установите [Memurai](https://www.memurai.com/) - нативная версия Redis для Windows.

### Linux
```bash
sudo apt update
sudo apt install redis-server
sudo systemctl start redis-server
sudo systemctl enable redis-server
```

### macOS
```bash
brew install redis
brew services start redis
```

**Проверка работы:**
```bash
redis-cli ping
# Должно вернуться: PONG
```

Подробнее см. [REDIS_SETUP.md](REDIS_SETUP.md)

## Шаг 5: Запуск приложения

### Режим разработки

```bash
python app.py
```

Приложение будет доступно по адресу: `http://localhost:5000`

### Production режим

**Вариант 1: Через переменную окружения**
```bash
# В .env установите
USE_WAITRESS=1

# Запустите
python app.py
```

**Вариант 2: Напрямую через Waitress**
```bash
waitress-serve --call "app:create_app"
```

**Вариант 3: Через Gunicorn (Linux/macOS)**
```bash
gunicorn -w 4 -b 0.0.0.0:5000 "app:create_app()"
```

## Шаг 6: Проверка работоспособности

### 6.1 Health Check

Откройте в браузере:
```
http://localhost:5000/health
```

Должен вернуться JSON с статусом всех компонентов системы.

### 6.2 Запуск тестов

```bash
# Все тесты
pytest

# С покрытием
pytest --cov=app --cov-report=html

# Конкретный тест
pytest tests/test_smoke_health.py -v
```

## Структура проекта

```
kp_generator/
├── app/                    # Основной код приложения
│   ├── core/              # Ядро (config, errors, extensions)
│   ├── models/            # Модели данных
│   ├── database/          # Работа с БД
│   ├── auth/              # Аутентификация
│   ├── business/          # Бизнес-логика
│   ├── presentation/      # Формы и UI
│   ├── routes/            # Маршруты (blueprints)
│   └── services/          # Сервисы
├── config/                # Конфигурационные файлы
├── migrations/            # Миграции Alembic
├── templates/             # HTML шаблоны
├── static/                # Статические файлы
├── templates_docs/        # Шаблоны документов (Excel/Word)
├── tests/                 # Тесты
├── scripts/               # Утилиты управления
├── logs/                  # Логи (создается автоматически)
├── app.py                 # Точка входа
├── requirements.txt       # Зависимости
└── alembic.ini           # Конфигурация Alembic
```

## Дополнительная документация

- **[Полное руководство](COMPLETE_GUIDE.md)** — подробная документация по всем функциям
- **[Структура проекта](PROJECT_STRUCTURE.md)** — описание архитектуры
- **[Управление пользователями](USER_MANAGEMENT.md)** — работа с пользователями через CLI
- **[Настройка Redis](REDIS_SETUP.md)** — подробная инструкция по Redis
- **[AI Agent Setup](AI_AGENT_SETUP.md)** — настройка AI консультанта
- **[AI Agent User Guide](AI_AGENT_USER_GUIDE.md)** — руководство пользователя AI агента
- **[Журнал изменений](CHANGELOG.md)** — история изменений проекта

## Устранение проблем

### База данных не создается

```bash
# Проверьте права доступа к директории
# Убедитесь, что миграции применены
alembic upgrade head
```

### Redis недоступен

Приложение автоматически переключится на cookie-based сессии. Проверьте:
- Redis запущен: `redis-cli ping`
- Настройки в `.env` корректны
- Порт 6379 не занят

### Ошибки импорта

```bash
# Убедитесь, что виртуальное окружение активировано
# Переустановите зависимости
pip install -r requirements.txt --force-reinstall
```

### Проблемы с миграциями

```bash
# Проверьте текущую версию
alembic current

# Просмотрите историю
alembic history

# Примените все миграции
alembic upgrade head
```

## Следующие шаги

1. Создайте первого пользователя через веб-интерфейс или скрипт
2. Настройте справочники через административную панель (`/admin`)
3. Загрузите шаблоны документов в `templates_docs/`
4. Настройте резервное копирование базы данных
5. Настройте мониторинг и логирование для production

## Поддержка

При возникновении проблем:
1. Проверьте логи в `logs/app.log`
2. Проверьте health check: `/health`
3. Убедитесь, что все зависимости установлены
4. Проверьте версию Python: `python --version` (должна быть 3.12+)


