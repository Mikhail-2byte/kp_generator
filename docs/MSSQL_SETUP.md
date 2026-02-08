# Подключение Microsoft SQL Server к KP Generator

Приложение работает **только с Microsoft SQL Server**. Ниже пошаговая настройка и подключение.

## Требования

- **Python**: 3.12+
- **ODBC-драйвер**: на машине, где запускается приложение, должен быть установлен драйвер для SQL Server (например, **ODBC Driver 17 for SQL Server** или **ODBC Driver 18 for SQL Server**).
- **Доступ к SQL Server**: сеть до экземпляра MS SQL и учётные данные (Windows-аутентификация или логин/пароль).

---

## Шаг 1. Установка ODBC-драйвера

### Windows

- Скачайте и установите [Microsoft ODBC Driver 17 for SQL Server](https://docs.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server) или [ODBC Driver 18](https://docs.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server).
- Или: `winget install Microsoft.SqlServer.ODBC.17` (или 18).

### Linux

- Инструкции для RHEL/Ubuntu: [документация Microsoft](https://docs.microsoft.com/en-us/sql/connect/odbc/linux-mac/installing-the-microsoft-odbc-driver-for-sql-server).

Проверка: в списке драйверов должно быть имя вида `ODBC Driver 17 for SQL Server` (Windows — «Диспетчер источников данных ODBC» → «Драйверы»; Linux — `odbcinst -q -d`).

---

## Шаг 2. Установка зависимостей Python

В корне проекта:

```bash
pip install -r requirements.txt
```

Устанавливается в том числе **pyodbc**. При ошибках сборки pyodbc на Windows может понадобиться [Microsoft C++ Build Tools](https://visualstudio.microsoft.com/visual-cpp-build-tools/).

---

## Шаг 3. Создание базы данных на SQL Server

Базу нужно создать **один раз** на экземпляре MS SQL. Таблицы создаст приложение при первом запуске (миграции Alembic).

### 3.1. Где выполнять команды

Выберите один из способов:

| Способ | Когда удобно |
|--------|----------------|
| **SQL Server Management Studio (SSMS)** | Обычная администрирование на Windows |
| **Azure Data Studio** | Кроссплатформенно, бесплатно |
| **sqlcmd** | Командная строка, скрипты, CI |

- **SSMS**: [скачать](https://docs.microsoft.com/en-us/sql/ssms/download-sql-server-management-studio-ssms). Подключитесь к серверу → «Создать запрос» (New Query) → вставьте SQL ниже.
- **Azure Data Studio**: [скачать](https://docs.microsoft.com/en-us/sql/azure-data-studio/download). Подключение к серверу → New Query.
- **sqlcmd**: в папке с установленным SQL Server выполните, подставив свой сервер и режим аутентификации:
  ```bash
  sqlcmd -S localhost -E -Q "CREATE DATABASE kp_generator;"
  ```
  Или с логином и паролем:
  ```bash
  sqlcmd -S localhost -U sa -P YourPassword -Q "CREATE DATABASE kp_generator;"
  ```

### 3.2. Создание базы (минимальный вариант)

В окне запроса (SSMS или Azure Data Studio) выполните:

```sql
CREATE DATABASE kp_generator;
GO
```

Этого достаточно: при первом запуске KP Generator миграции создадут все таблицы в этой базе.

Если нужна база с явной кодировкой для русского языка (обычно не обязательно):

```sql
CREATE DATABASE kp_generator
  COLLATE Cyrillic_General_CI_AS;
GO
```

### 3.3. Учётная запись для приложения

Строка подключения в `.env` (шаг 4) должна использовать учётную запись, у которой есть права **на эту базу**. Возможные варианты:

1. **Встроенная учётная запись** (например, `sa` или доменный пользователь с правами на сервер) — уже имеет права на создание таблиц.
2. **Отдельный логин для приложения** (рекомендуется в production): создайте логин и пользователя, выдайте ему роль `db_owner` в базе `kp_generator`.

Пример создания отдельного логина и пользователя (выполнять под `sa` или другим администратором):

```sql
-- Логин на уровне сервера (режим SQL Server authentication)
CREATE LOGIN kp_user WITH PASSWORD = 'YourStrongPassword123!';
GO

USE kp_generator;
GO

-- Пользователь в базе kp_generator
CREATE USER kp_user FOR LOGIN kp_user;

-- Права на создание/изменение таблиц (миграции) и полный доступ к данным
ALTER ROLE db_owner ADD MEMBER kp_user;
GO
```

Дальше в `.env` укажите:

```env
DATABASE_URL=mssql+pyodbc://kp_user:YourStrongPassword123!@your-server:1433/kp_generator?driver=ODBC+Driver+18+for+SQL+Server
```

Если в пароле есть спецсимволы (`@`, `#`, `%` и т.д.), закодируйте их в URL (например, `@` → `%40`).

### 3.4. Проверка

- В SSMS/Azure Data Studio в дереве «Databases» должна появиться база **kp_generator**.
- После настройки `DATABASE_URL` и запуска приложения (`python app.py`) таблицы появятся в этой базе автоматически.

---

## Шаг 4. Настройка DATABASE_URL в .env

В корне проекта создайте или отредактируйте файл **`.env`** и задайте **`DATABASE_URL`** в формате **`mssql+pyodbc://...`**.

### Логин и пароль

Подставьте свой хост, порт, имя БД, логин и пароль. В параметре `driver` укажите точное имя драйвера из списка ODBC (пробелы в URL замените на `+`).

**С портом (например, 1433):**

```env
DATABASE_URL=mssql+pyodbc://USERNAME:PASSWORD@hostname:1433/database_name?driver=ODBC+Driver+17+for+SQL+Server
```

**Пример для БД `kp_generator`, драйвер 18:**

```env
DATABASE_URL=mssql+pyodbc://kp_user:YourPassword@sql.company.local:1433/kp_generator?driver=ODBC+Driver+18+for+SQL+Server
```

Спецсимволы в пароле кодируйте для URL (`@` → `%40`, `#` → `%23` и т.д.).

### Windows-аутентификация (Trusted_Connection)

```env
DATABASE_URL=mssql+pyodbc://hostname:1433/database_name?driver=ODBC+Driver+17+for+SQL+Server&Trusted_Connection=yes
```

### Шифрование (при необходимости)

Для серверов с самоподписанным сертификатом можно добавить:

```env
DATABASE_URL=mssql+pyodbc://user:pass@host:1433/kp_generator?driver=ODBC+Driver+18+for+SQL+Server&Encrypt=yes&TrustServerCertificate=yes
```

---

## Шаг 5. Запуск приложения и миграции

1. В `.env` обязательно задайте **`DATABASE_URL`** (см. шаг 4). Остальные переменные (Redis, `SECRET_KEY` и т.д.) — по необходимости.
2. Запустите приложение из корня проекта:

   ```bash
   python app.py
   ```

   Или в production с Waitress:

   ```bash
   set USE_WAITRESS=1
   python app.py
   ```

3. При первом запуске выполняются миграции Alembic — таблицы создаются в указанной БД. Дальнейшие запуски только подключаются к существующей схеме.

---

## Шаг 6. Пул соединений (опционально)

При большой нагрузке можно увеличить пул в `.env`:

```env
DATABASE_POOL_SIZE=10
DATABASE_MAX_OVERFLOW=20
DATABASE_POOL_TIMEOUT=30
```

По умолчанию: 5, 10 и 30 соответственно.

---

## Шаг 7. Проверка подключения

В браузере откройте:

```
http://localhost:5000/health
```

В ответе должен быть статус компонентов. Успешная проверка БД означает, что `DATABASE_URL` и драйвер настроены верно.

Типичные ошибки:

- **"Data source name not found"** / **"Driver not found"** — не установлен или не виден ODBC-драйвер; проверьте имя в `driver=...` в `DATABASE_URL`.
- **"Требуется DATABASE_URL в формате mssql+pyodbc://..."** — приложение запущено без заданного `DATABASE_URL` в `.env` или переменных окружения.

---

## Шаг 8. Управление пользователями

- **Веб**: профиль пользователя и админ-панель (вход, регистрация, смена пароля, роли).
- **CLI** (используют тот же `DATABASE_URL` из `.env`):

  ```bash
  python scripts/manage_users.py list
  python scripts/manage_users.py list --detailed
  python scripts/manage_users.py reset-password <username> <new_password>
  python scripts/manage_users.py set-role <username> admin
  python scripts/manage_roles.py list
  python scripts/manage_roles.py set-role <username> user
  ```

Скрипты подключаются к Microsoft SQL Server через `DATABASE_URL`; отдельный путь к файлу БД не используется.

---

## Запуск тестов без MSSQL (для разработки)

Чтобы запускать тесты без установленного SQL Server, задайте переменные и запустите pytest:

```bash
set USE_TEST_SQLITE=1
set DATABASE_URL=sqlite:///:memory:
pytest
```

В production и при обычном запуске приложения используется только **Microsoft SQL Server** и `DATABASE_URL` в формате `mssql+pyodbc://...`.
