# Полное руководство по проекту KP Generator

## Содержание

1. [Общее описание](#общее-описание)
2. [Установка и настройка](#установка-и-настройка)
3. [Структура проекта](#структура-проекта)
4. [Основная функциональность](#основная-функциональность)
5. [REST API](#rest-api)
6. [Расширенная аналитика](#расширенная-аналитика)
7. [Справочник контактов заказчиков](#справочник-контактов-заказчиков)
8. [Административная панель](#административная-панель)
9. [База данных и миграции](#база-данных-и-миграции)
10. [Тестирование](#тестирование)
11. [Развертывание](#развертывание)
12. [Примеры использования](#примеры-использования)

---

## Общее описание

**KP Generator** — веб-приложение для автоматической генерации коммерческих предложений (КП) с расчетом цен, маржинальности и формированием документов (Word, Excel).

### Основные возможности

- **Расчет цен** для одной или множественных позиций с единой целевой маржой
- **Генерация документов** (Word, Excel) с автоматическим заполнением шаблонов
- **История генераций** с версионированием по тендерам
- **Расчет логистики** с учетом типа транспорта и расстояния
- **Справочники** пошлин, материалов GB, городов логистики
- **Расширенная аналитика** с графиками и отчетами
- **REST API** для интеграции с внешними системами
- **Административная панель** для управления справочниками и пользователями
- **Справочник контактов** заказчиков с CRUD операциями

---

## Установка и настройка

### Требования

- Python 3.9+
- SQLite (или PostgreSQL для production)
- Git

### Шаг 1: Клонирование репозитория

```bash
git clone <repository-url>
cd kp_generator
```

### Шаг 2: Создание виртуального окружения

```bash
python -m venv venv

# Windows
venv\Scripts\activate

# Linux/Mac
source venv/bin/activate
```

### Шаг 3: Установка зависимостей

```bash
pip install -r requirements.txt
```

### Шаг 4: Настройка конфигурации

Создайте файл `.env` в корне проекта:

```env
# База данных
DATABASE_URL=sqlite:///kp_generator.db

# Секретный ключ (сгенерируйте случайную строку)
SECRET_KEY=your-secret-key-here

# Окружение
FLASK_ENV=development
FLASK_DEBUG=True

# Сервер
FLASK_RUN_HOST=0.0.0.0
FLASK_RUN_PORT=5000
```

### Шаг 5: Инициализация базы данных

```bash
# Применение миграций
python -m alembic upgrade head
```

### Шаг 6: Запуск приложения

```bash
# Режим разработки
python app.py

# Production режим (с Waitress)
set USE_WAITRESS=True
python app.py
```

Приложение будет доступно по адресу: `http://localhost:5000`

### Шаг 7: Создание первого пользователя

1. Откройте `http://localhost:5000/profile`
2. Зарегистрируйтесь (первый пользователь автоматически становится администратором)
3. Войдите в систему

---

## Структура проекта

```
kp_generator/
├── app/                          # Основное приложение
│   ├── __init__.py              # Фабрика приложения Flask
│   ├── business/                 # Бизнес-логика
│   │   ├── price_calculator.py  # Расчет цен
│   │   ├── document_generator.py # Генерация документов
│   │   └── interfaces.py        # Интерфейсы
│   ├── core/                    # Ядро приложения
│   │   ├── config.py           # Конфигурация
│   │   ├── exceptions.py       # Исключения
│   │   ├── extensions.py       # Расширения Flask
│   │   ├── errors.py           # Обработка ошибок
│   │   └── cache.py            # Кэширование
│   ├── database/                # Работа с БД
│   │   ├── database.py         # CRUD операции
│   │   └── service.py          # Сервис БД
│   ├── models/                  # Модели данных
│   │   └── models.py           # SQLAlchemy модели
│   ├── presentation/            # Представление
│   │   ├── forms.py            # WTForms формы
│   │   ├── helpers.py          # Вспомогательные функции
│   │   ├── validators.py       # Валидаторы
│   │   └── ui.py               # UI регистрация
│   ├── routes/                  # Маршруты
│   │   ├── main.py             # Основные маршруты
│   │   ├── auth.py             # Авторизация
│   │   ├── admin.py            # Админ-панель
│   │   ├── api.py              # REST API
│   │   ├── api_docs.py         # Swagger документация
│   │   └── health.py           # Health check
│   ├── services/               # Сервисы
│   │   ├── repositories.py     # Репозитории
│   │   ├── multi_position_calculator.py # Калькулятор позиций
│   │   ├── generation_orchestrator.py   # Оркестратор генерации
│   │   ├── logistics_calculator.py     # Расчет логистики
│   │   ├── excel_importer.py           # Импорт Excel
│   │   ├── analytics_service.py         # Аналитика
│   │   ├── analytics_enhancements.py   # Расширенная аналитика
│   │   ├── datasets.py                  # Справочники
│   │   └── audit_service.py             # Аудит
│   └── static/                  # Статические файлы
│   └── templates/              # Шаблоны Jinja2
├── config/                      # Конфигурационные файлы
│   ├── settings.json           # Основные настройки
│   ├── development.json        # Настройки разработки
│   ├── production.json         # Настройки production
│   └── ...                     # Справочники (JSON)
├── migrations/                  # Миграции Alembic
├── tests/                       # Тесты
├── docs/                        # Документация
├── logs/                        # Логи (создается автоматически)
├── templates_docs/             # Шаблоны документов
├── alembic.ini                  # Конфигурация Alembic
├── app.py                       # Точка входа
├── requirements.txt            # Зависимости
└── pyproject.toml              # Настройки линтера/форматтера
```

---

## Основная функциональность

### 1. Генерация коммерческих предложений

#### Расчет для одной позиции

1. Перейдите на главную страницу `/`
2. Заполните форму:
   - **Компания** — название заказчика
   - **Товар** — наименование товара
   - **Количество** — количество единиц
   - **Цена закупа** — стоимость за единицу (в юанях)
   - **Вес** — вес одной единицы (кг)
   - **Логистика** — стоимость доставки (руб)
   - **Пошлина** — процент пошлины
   - **Маржа** — целевая маржа (%)
   - **Время доставки** — дни
   - И другие параметры

3. Нажмите "Сгенерировать КП"
4. Система:
   - Рассчитает продажную цену
   - Сгенерирует документы Word и Excel
   - Сохранит в историю
   - Предоставит ZIP-архив для скачивания

#### Расчет для множественных позиций

1. На главной странице заполните первую позицию
2. Для дополнительных позиций используйте поля с суффиксами (`_2`, `_3`, ...):
   - `product_2`, `quantity_2`, `cost_price_2`, и т.д.
3. Или используйте импорт из Excel:
   - Нажмите "Импорт позиций из Excel"
   - Загрузите файл с колонками: `product`, `quantity`, `cost_price`, `weight`, `duty_percent`
4. Система рассчитает цены с **единой целевой маржой** для всех позиций

#### Импорт позиций из Excel

Система поддерживает два формата Excel файлов для импорта позиций:

**Формат 1: С заголовками (старый формат)**

Файл должен содержать строку заголовков с названиями колонок. Поддерживаемые названия колонок:
- `Номенклатура`, `Наименование товара`, `Наименование продукции` → `product`
- `Номер чертежа`, `Чертеж` → `drawing_number`
- `Материал`, `Материал изделия` → `material`
- `Цена закупа`, `Цена закупа, руб`, `Стоимость закупа` → `cost_price`
- `Цена закупа за кг`, `Цена за кг`, `Стоимость закупа за кг` → `cost_price_per_kg`
- `Количество, шт.`, `Количество шт.`, `Количество`, `шт.` → `quantity`
- `Вес за шт. (кг)`, `Вес за шт. кг`, `Вес, кг`, `Вес` → `weight`
- `Пошлина (%)`, `Пошлина %`, `Пошлина` → `duty_percent`

Пример:

| Номенклатура | Цена закупа | Количество | Вес за шт. (кг) | Пошлина (%) |
|--------------|-------------|------------|-----------------|-------------|
| Деталь А | 1200 | 10 | 3.5 | 5 |
| Деталь B | 800 | 5 | 2.0 | 10 |

**Формат 2: Фиксированные ячейки (новый формат, шаблон "Запрос.xlsx")**

Данные должны находиться в фиксированных ячейках начиная с 7-й строки:
- **B7, B8, ...** → Наименование (`product`)
- **D7, D8, ...** → Материал (`material`)
- **E7, E8, ...** → Номер чертежа (`drawing_number`)
- **K7, K8, ...** → Количество (`quantity`)
- **L7, L8, ...** → Цена закупа за шт. (`cost_price`)
- **M7, M8, ...** → Общая цена закупа (используется для валидации, не сохраняется)
- **N7, N8, ...** → Вес за шт. (`weight`)
- **P7, P8, ...** → Пошлина (`duty_percent`)
- **S7, S8, ...** → Цена закупа за кг (`cost_price_per_kg`, опционально)

Если указаны и `L` (цена за шт.), и `S` (цена за кг), система строго проверяет их согласованность.
Если указана только `S`, цена за шт. вычисляется автоматически по формуле `цена за кг × вес за шт.`.

Система автоматически определит формат файла и обработает его соответствующим образом. Если файл не содержит распознаваемых заголовков, будет использован формат с фиксированными ячейками.

### 2. История генераций

**Маршрут:** `/history`

**Возможности:**
- Просмотр всех генераций с пагинацией
- Фильтрация по датам (`date_from`, `date_to`)
- Группировка по тендерам (версионирование)
- Просмотр деталей каждой генерации
- Загрузка предыдущей генерации в форму для редактирования

**Использование:**
```
/history?page=1&date_from=2025-01-01&date_to=2025-01-31
```

### 3. Расчет логистики

**Маршрут:** `/api/logistics/calculate` (API) или через форму

**Параметры:**
- `weight_kg` — вес груза (кг)
- `city_price` — базовая стоимость доставки до города
- `transport_type` — тип транспорта (`truck` или `trail`)
- `city_name` — название города
- `is_main_route` — город на основном маршруте
- `distance_from_ekb_km` — расстояние от Екатеринбурга (км)
- `length_mm`, `width_mm`, `height_mm` — габариты (опционально)

**Результат:**
- Итоговая стоимость логистики с учетом всех параметров

### 4. Справочники

#### Поиск пошлин

**Маршрут:** `/duty?q=запрос`

Поиск по ставкам пошлин и категориям товаров.

#### Аналоги материалов GB

**Маршрут:** `/gb-analogs`

Просмотр справочника аналогов материалов GB с поиском.

#### Города логистики

Доступны через форму расчета логистики и API.

---

## REST API

### Документация Swagger

**Маршрут:** `/api/docs/`

Интерактивная документация API доступна по адресу: `http://localhost:5000/api/docs/`

### Базовый URL

```
http://localhost:5000/api/v1
```

### Аутентификация

Большинство endpoints требуют аутентификации через Flask-Login (сессии).

### Endpoints

#### 1. Health Check

```http
GET /api/v1/health
```

**Ответ:**
```json
{
  "status": "ok",
  "version": "1.0",
  "service": "kp_generator"
}
```

#### 2. Список генераций

```http
GET /api/v1/generations?page=1&per_page=25&date_from=2025-01-01&date_to=2025-01-31
```

**Параметры:**
- `page` — номер страницы (по умолчанию 1)
- `per_page` — записей на странице (по умолчанию 25)
- `date_from` — начальная дата (YYYY-MM-DD)
- `date_to` — конечная дата (YYYY-MM-DD)

**Ответ:**
```json
{
  "items": [
    {
      "id": 1,
      "company": "Компания",
      "tender_number": "Т-001",
      "timestamp": "2025-01-15T10:30:00",
      "final_price": 1500.0,
      "margin_percent": 30.0,
      ...
    }
  ],
  "pagination": {
    "page": 1,
    "per_page": 25,
    "total": 100,
    "pages": 4,
    "has_prev": false,
    "has_next": true
  }
}
```

#### 3. Детали генерации

```http
GET /api/v1/generations/{generation_id}
```

**Ответ:**
```json
{
  "id": 1,
  "company": "Компания",
  "product": "Товар",
  "quantity": 10,
  "cost_price": 1000.0,
  "final_price": 1500.0,
  "margin_percent": 30.0,
  "positions": [...],
  ...
}
```

#### 4. Расчет цен

```http
POST /api/v1/calculate
Content-Type: application/json

{
  "positions": [
    {
      "quantity": 10,
      "cost_price": 1000,
      "weight": 5,
      "duty_percent": 5
    }
  ],
  "logistics_rub": 50000,
  "delivery_time": 30,
  "margin_percent": 30
}
```

**Ответ:**
```json
{
  "positions": [
    {
      "position": {...},
      "final_price": 1500.0,
      "general_price": 15000.0,
      "margin": 30.0,
      "costs": {...}
    }
  ],
  "total_costs": 11500.0,
  "total_revenue": 15000.0,
  "target_margin": 30.0,
  "actual_margin": 30.0
}
```

#### 5. Расчет логистики

```http
POST /api/v1/logistics/calculate
Content-Type: application/json

{
  "weight_kg": 1000,
  "city_price": 50000,
  "transport_type": "truck",
  "city_name": "Москва",
  "is_main_route": true,
  "distance_from_ekb_km": 1800
}
```

**Ответ:**
```json
{
  "total_logistics": 55000.0,
  "breakdown": {
    "base_price": 50000,
    "distance_surcharge": 5000,
    ...
  }
}
```

#### 6. Поиск пошлин

```http
GET /api/v1/duty/search?query=сталь&limit=50
```

**Ответ:**
```json
{
  "items": [
    {
      "title": "Сталь",
      "code": "7201",
      "duty_percent": 5,
      ...
    }
  ],
  "total": 10,
  "query": "сталь"
}
```

#### 7. Материалы GB

```http
GET /api/v1/materials/gb?search=сталь
```

**Ответ:**
```json
{
  "items": [
    {
      "russian": "Сталь",
      "gb": "GB/T ...",
      "composition_search": "...",
      ...
    }
  ],
  "total": 50
}
```

#### 8. Города логистики

```http
GET /api/v1/logistics/cities
```

**Ответ:**
```json
{
  "items": [
    {
      "name": "Москва",
      "price": 50000,
      "is_main_route": true,
      ...
    }
  ],
  "total": 100
}
```

### Управление контактами заказчиков (API)

#### Список контактов

```http
GET /api/v1/customers?search=компания
```

#### Создание контакта

```http
POST /api/v1/customers
Content-Type: application/json

{
  "company_name": "ООО Компания",
  "contact_person": "Иван Иванов",
  "phone": "+7 999 123-45-67",
  "email": "info@company.ru",
  "address": "Москва, ул. Примерная, 1",
  "notes": "Примечания"
}
```

#### Получение контакта

```http
GET /api/v1/customers/{customer_id}
```

#### Обновление контакта

```http
PUT /api/v1/customers/{customer_id}
Content-Type: application/json

{
  "phone": "+7 999 999-99-99",
  "email": "new@company.ru"
}
```

#### Удаление контакта

```http
DELETE /api/v1/customers/{customer_id}
```

### Примеры использования API

#### Python (requests)

```python
import requests

BASE_URL = "http://localhost:5000/api/v1"

# Создание сессии с авторизацией
session = requests.Session()
session.post("http://localhost:5000/profile", data={
    "username": "user",
    "password": "password",
    "action": "login"
})

# Расчет цен
response = session.post(f"{BASE_URL}/calculate", json={
    "positions": [
        {"quantity": 10, "cost_price": 1000, "weight": 5, "duty_percent": 5}
    ],
    "logistics_rub": 50000,
    "delivery_time": 30,
    "margin_percent": 30
})
result = response.json()
print(result)
```

#### JavaScript (fetch)

```javascript
// Расчет логистики
fetch('http://localhost:5000/api/v1/logistics/calculate', {
  method: 'POST',
  headers: {
    'Content-Type': 'application/json',
  },
  body: JSON.stringify({
    weight_kg: 1000,
    city_price: 50000,
    transport_type: 'truck',
    city_name: 'Москва',
    is_main_route: true,
    distance_from_ekb_km: 1800
  })
})
.then(response => response.json())
.then(data => console.log(data));
```

---

## Расширенная аналитика

### Доступ

**Маршрут:** `/analytics`

### Возможности

1. **Анализ загруженных Excel файлов** (тендеры/продукция)
2. **Анализ маржинальности** — графики и метрики
3. **Динамика курсов валют** — графики изменения курсов
4. **Интерактивные отчеты** — комплексная аналитика

### 1. Анализ Excel файлов

**Использование:**
1. Перейдите на `/analytics`
2. Загрузите Excel файл
3. Система автоматически определит тип данных:
   - **Тендеры** — если найдены ключевые слова (Мониторинг, участвовали, Подались, и т.д.)
   - **Продукция** — если найдены ключевые слова (ч., заказчик, вес, цена, и т.д.)

**Результаты:**
- Метрики (количество записей, суммы, средние значения)
- Графики (распределения, топ-10, диаграммы)
- Статистика (описательная статистика по числовым полям)
- Превью данных (таблица)
- Экспорт (CSV, Excel)

### 2. Анализ маржинальности

**Использование:**
```
/analytics?margin=true&days=30
```

**Параметры:**
- `margin=true` — включить анализ маржи
- `days` — период анализа (по умолчанию 30 дней)

**Результаты:**
- **График динамики маржи** — изменение маржи по генерациям
- **Распределение маржи** — гистограмма распределения
- **Маржа vs Выручка** — scatter plot зависимости
- **Метрики:**
  - Средняя маржа
  - Медианная маржа
  - Минимальная/максимальная маржа

**Пример:**
```
/analytics?margin=true&days=90
```

### 3. Динамика курсов валют

**Использование:**
```
/analytics?exchange=true&days=90
```

**Параметры:**
- `exchange=true` — включить анализ курсов
- `days` — период анализа

**Результаты:**
- График динамики курса юаня к рублю
- Метрики (средний, минимальный, максимальный курс)

**Примечание:** В текущей версии используются примерные данные. Для реальных данных требуется интеграция с внешним API курсов валют.

### 4. Интерактивные отчеты

**Использование:**
```
/analytics?report=true&date_from=2025-01-01&date_to=2025-01-31
```

**Параметры:**
- `report=true` — включить отчет
- `date_from` — начальная дата (YYYY-MM-DD)
- `date_to` — конечная дата (YYYY-MM-DD)

**Результаты:**
- **Сводная статистика:**
  - Общее количество генераций
  - Общая выручка
  - Средняя маржа
  - Общее количество товара
  - Количество уникальных компаний

- **Графики:**
  - Выручка по дням
  - Топ-10 компаний по выручке

- **Таблицы:**
  - Статистика по компаниям (выручка, маржа, количество)

**Пример:**
```
/analytics?report=true&date_from=2025-01-01&date_to=2025-01-31
```

### Комбинированное использование

Можно комбинировать несколько типов аналитики:

```
/analytics?margin=true&exchange=true&report=true&days=60&date_from=2025-01-01&date_to=2025-01-31
```

---

## Справочник контактов заказчиков

### Доступ

**Web интерфейс:** через форму генерации КП (выпадающий список компаний)

**API:** `/api/v1/customers`

### Функциональность

#### 1. Создание контакта

**Через форму:**
1. На главной странице в поле "Компания" начните вводить название
2. Если контакт существует, он появится в выпадающем списке
3. Если нет — введите новое название и заполните дополнительные поля:
   - Контактное лицо
   - Телефон
   - Email
   - Адрес
   - Заметки

**Через API:**
```http
POST /api/v1/customers
Content-Type: application/json

{
  "company_name": "ООО Компания",
  "contact_person": "Иван Иванов",
  "phone": "+7 999 123-45-67",
  "email": "info@company.ru",
  "address": "Москва, ул. Примерная, 1",
  "notes": "Примечания"
}
```

#### 2. Просмотр контактов

**Через форму:**
- При вводе названия компании в форме генерации КП отображаются подсказки

**Через API:**
```http
GET /api/v1/customers?search=компания
```

#### 3. Редактирование контакта

**Через форму:**
- Контакты можно редактировать через административную панель (планируется)

**Через API:**
```http
PUT /api/v1/customers/{customer_id}
Content-Type: application/json

{
  "phone": "+7 999 999-99-99",
  "email": "new@company.ru"
}
```

#### 4. Удаление контакта

**Через API:**
```http
DELETE /api/v1/customers/{customer_id}
```

### Модель данных

```python
{
  "id": 1,
  "company_name": "ООО Компания",
  "contact_person": "Иван Иванов",
  "phone": "+7 999 123-45-67",
  "email": "info@company.ru",
  "address": "Москва, ул. Примерная, 1",
  "notes": "Примечания",
  "created_at": "2025-01-15T10:30:00",
  "updated_at": "2025-01-15T10:30:00"
}
```

### Интеграция с формой генерации

При вводе названия компании в форме генерации КП:
1. Система автоматически ищет совпадения в справочнике
2. Отображает выпадающий список с найденными контактами
3. При выборе контакта автоматически заполняются дополнительные поля (если доступны)
4. Если контакт не найден, можно создать новый прямо из формы

---

## Административная панель

### Доступ

**Маршрут:** `/admin`

**Требования:** Роль `admin`

### Разделы

#### 1. Статистика (`/admin/stats`)

- Активность пользователей
- Количество генераций
- Топ пользователей
- Графики активности

#### 2. Управление пользователями (`/admin/users`)

- Список пользователей с поиском и фильтрацией
- Создание пользователей
- Редактирование профилей
- Удаление пользователей
- Сброс паролей
- Изменение ролей

#### 3. Управление справочниками

**Пошлины** (`/admin/duty`):
- Добавление/редактирование/удаление ставок пошлин
- Версионирование данных

**Материалы GB** (`/admin/materials`):
- Управление справочником аналогов материалов

**Логистика** (`/admin/logistics`):
- Управление городами и тарифами логистики
- Сохранение расстояний между городами

#### 4. Управление контентом

**Заказы** (`/admin/orders`):
- Управление страницей заказов

**Шаблоны** (`/admin/templates`):
- Управление библиотекой шаблонов

**Инструкции** (`/admin/instructions`):
- Управление инструкциями

#### 5. Аудит (`/admin/audit`)

- Просмотр логов действий пользователей
- Фильтрация по типу действия, пользователю, дате
- Экспорт логов
- Статистика по действиям
- Топ пользователей по активности

---

## База данных и миграции

### Модели данных

#### UserRecord (Пользователи)

```python
- id: Integer (PK)
- username: String (unique)
- password_hash: String
- last_name: String
- first_name: String
- role: String (admin/user)
- contact_info: Text
- created_at: DateTime
- last_login: DateTime
```

#### GenerationHistoryRecord (История генераций)

```python
- id: Integer (PK)
- tender_number: String
- company: String
- product: String
- quantity: Integer
- cost_price: Float
- weight: Float
- logistics: Float
- margin_percent: Float
- final_price: Float
- drawing_number: String
- material: String
- delivery_address: String
- duty_percent: Float
- delivery_time: Integer
- payment_terms: String
- proposal_validity: String
- warranty_period: String
- comment: Text
- user_id: Integer (FK -> users.id)
- timestamp: DateTime
- positions_data: Text (JSON)
- total_general_price: Float
- positions_count: Integer
```

#### AuditLogRecord (Логи аудита)

```python
- id: Integer (PK)
- user_id: Integer (FK -> users.id)
- username: String
- action_type: String
- description: Text
- resource_type: String
- resource_id: String
- created_at: DateTime
```

#### CustomerContactRecord (Контакты заказчиков)

```python
- id: Integer (PK)
- company_name: String
- contact_person: String
- phone: String
- email: String
- address: Text
- notes: Text
- created_at: DateTime
- updated_at: DateTime
```

### Миграции

#### Создание новой миграции

```bash
python -m alembic revision --autogenerate -m "Описание изменений"
```

#### Применение миграций

```bash
python -m alembic upgrade head
```

#### Откат миграции

```bash
python -m alembic downgrade -1
```

#### Проверка статуса

```bash
python -m alembic current
python -m alembic history
```

### Индексы

Для оптимизации запросов созданы индексы на:
- `generation_history.timestamp`
- `generation_history.user_id`
- `generation_history.drawing_number`
- `generation_history.tender_number`
- `generation_history.company`
- `generation_history.user_id, timestamp` (составной)
- `users.username` (unique)
- `users.role`
- `customer_contacts.company_name`
- `audit_logs.user_id, created_at` (составной)
- `audit_logs.resource_type, resource_id` (составной)

---

## Тестирование

### Запуск тестов

```bash
# Все тесты
pytest

# С покрытием
pytest --cov=app --cov-report=html

# Конкретный файл
pytest tests/test_multi_position_calculator.py

# Конкретный тест
pytest tests/test_multi_position_calculator.py::test_calculate_positions_global_margin
```

### Структура тестов

```
tests/
├── conftest.py                          # Фикстуры
├── test_api_health.py                   # Health check
├── test_auth.py                         # Авторизация
├── test_database_service.py             # БД сервис
├── test_document_generation_service.py  # Генерация документов
├── test_excel_importer_service.py      # Импорт Excel
├── test_generate.py                     # Генерация КП
├── test_history_pagination.py           # Пагинация истории
├── test_logistics_admin.py              # Логистика (админ)
├── test_logistics_api_integration.py    # Логистика (API)
├── test_logistics_updated.py            # Логистика (обновления)
├── test_multi_position_calculator.py    # Калькулятор позиций
├── test_rf_tariffs_demo.py             # Тарифы РФ
├── test_smoke_health.py                # Smoke тесты
├── test_validators.py                   # Валидаторы
├── test_generation_orchestrator.py      # Оркестратор
├── test_customer_contact_repository.py  # Контакты
└── test_repositories_integration.py     # Интеграционные тесты
```

### Покрытие

Целевое покрытие: **80%+**

Проверка покрытия:
```bash
pytest --cov=app --cov-report=term-missing
```

---

## Развертывание

### Production настройки

1. **Создайте `.env` файл:**
```env
DATABASE_URL=postgresql://user:password@localhost/kp_generator
SECRET_KEY=<strong-random-secret-key>
FLASK_ENV=production
FLASK_DEBUG=False
USE_WAITRESS=True
```

2. **Примените миграции:**
```bash
python -m alembic upgrade head
```

3. **Запустите с Waitress:**
```bash
python app.py
```

### Docker (планируется)

```dockerfile
FROM python:3.9-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt
COPY . .
CMD ["python", "app.py"]
```

### Nginx reverse proxy (пример)

```nginx
server {
    listen 80;
    server_name kp-generator.example.com;

    location / {
        proxy_pass http://127.0.0.1:5000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
    }
}
```

---

## Примеры использования

### Пример 1: Генерация КП для одной позиции

1. Откройте `/`
2. Заполните форму:
   - Компания: "ООО Заказчик"
   - Товар: "Деталь А"
   - Количество: 10
   - Цена закупа: 1000 (юаней)
   - Вес: 5 (кг)
   - Логистика: 50000 (руб)
   - Пошлина: 5%
   - Маржа: 30%
   - Время доставки: 30 дней
3. Нажмите "Сгенерировать КП"
4. Скачайте ZIP-архив с документами

### Пример 2: Генерация КП для множественных позиций

1. Откройте `/`
2. Заполните первую позицию
3. Заполните вторую позицию в полях с `_2`:
   - `product_2`: "Деталь B"
   - `quantity_2`: 5
   - `cost_price_2`: 800
   - `weight_2`: 3
   - `duty_percent_2`: 10
4. Нажмите "Сгенерировать КП"
5. Система рассчитает цены с единой маржой 30% для всех позиций

### Пример 3: Импорт позиций из Excel

1. Подготовьте Excel файл:
```
| product    | quantity | cost_price | weight | duty_percent |
|------------|----------|------------|--------|--------------|
| Деталь А   | 10       | 1200       | 3.5    | 5            |
| Деталь B   | 5        | 800        | 2.0    | 10           |
```

2. На главной странице нажмите "Импорт позиций из Excel"
3. Загрузите файл
4. Позиции автоматически заполнят форму
5. Нажмите "Сгенерировать КП"

### Пример 4: Использование REST API

```python
import requests

BASE_URL = "http://localhost:5000/api/v1"

# 1. Расчет цен
response = requests.post(f"{BASE_URL}/calculate", json={
    "positions": [
        {"quantity": 10, "cost_price": 1000, "weight": 5, "duty_percent": 5},
        {"quantity": 5, "cost_price": 800, "weight": 3, "duty_percent": 10}
    ],
    "logistics_rub": 50000,
    "delivery_time": 30,
    "margin_percent": 30
})
print(response.json())

# 2. Расчет логистики
response = requests.post(f"{BASE_URL}/logistics/calculate", json={
    "weight_kg": 1000,
    "city_price": 50000,
    "transport_type": "truck",
    "city_name": "Москва",
    "is_main_route": True,
    "distance_from_ekb_km": 1800
})
print(response.json())

# 3. Создание контакта
response = requests.post(f"{BASE_URL}/customers", json={
    "company_name": "ООО Компания",
    "contact_person": "Иван Иванов",
    "phone": "+7 999 123-45-67",
    "email": "info@company.ru"
})
print(response.json())
```

### Пример 5: Анализ маржинальности

1. Откройте `/analytics?margin=true&days=90`
2. Просмотрите графики:
   - Динамика маржи
   - Распределение маржи
   - Маржа vs Выручка
3. Изучите метрики (средняя, медианная, мин/макс маржа)

### Пример 6: Интерактивный отчет

1. Откройте `/analytics?report=true&date_from=2025-01-01&date_to=2025-01-31`
2. Просмотрите:
   - Сводную статистику
   - График выручки по дням
   - Топ-10 компаний
   - Таблицу статистики по компаниям

---

## Дополнительная информация

### Логирование

Логи сохраняются в папке `logs/`:
- `app.log` — основные логи приложения
- Ротация логов (максимальный размер: 10MB, количество файлов: 5)

### Конфигурация

Основные настройки в `config/settings.json`:
- Параметры расчета (курсы, коэффициенты)
- Настройки пагинации
- Параметры логирования
- И другие

### Безопасность

- CSRF защита (Flask-WTF)
- Хеширование паролей (werkzeug)
- Валидация входных данных
- Аудит действий пользователей

### Производительность

- Кэширование справочников (LRU cache)
- Оптимизация запросов к БД (индексы, joinedload)
- Пагинация для больших списков

---

## Поддержка

При возникновении проблем:
1. Проверьте логи в `logs/app.log`
2. Убедитесь, что миграции применены: `python -m alembic current`
3. Проверьте конфигурацию в `config/settings.json`
4. Запустите тесты: `pytest`

---

**Версия документа:** 1.0  
**Дата обновления:** 2025-01-15

