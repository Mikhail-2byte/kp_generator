# Настройка Redis для Flask сессий

## Обзор

Приложение использует Redis для хранения сессий Flask, что решает проблему с большими cookie (>4093 байт) и улучшает производительность.

## Установка Redis

### Windows

**Вариант 1: Docker Desktop (рекомендуется)**

1. Установите и запустите Docker Desktop: https://www.docker.com/products/docker-desktop/
2. Дождитесь полного запуска Docker Desktop (иконка в трее должна быть зеленая)
3. Запустите Redis контейнер:
   ```powershell
   docker run -d -p 6379:6379 --name redis redis:latest
   ```
4. Проверьте, что контейнер запущен:
   ```powershell
   docker ps
   ```

**Вариант 2: WSL2 с Redis (если установлен WSL)**

1. Установите WSL2, если еще не установлен:
   ```powershell
   wsl --install
   ```
2. В WSL установите Redis:
   ```bash
   sudo apt update
   sudo apt install redis-server
   sudo service redis-server start
   ```
3. Redis будет доступен на localhost:6379 из Windows

**Вариант 3: Memurai (нативная версия Redis для Windows)**

1. Скачайте и установите Memurai: https://www.memurai.com/get-memurai
2. Запустите Memurai (он работает как служба Windows)
3. Используйте стандартные настройки подключения (localhost:6379)

**Вариант 4: Использовать fallback (без Redis)**

Приложение будет работать с cookie-based сессиями, если Redis недоступен. 
Просто запустите приложение - оно автоматически определит отсутствие Redis и переключится на fallback режим.

### Linux (Ubuntu/Debian)

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

## Проверка работы Redis

После установки проверьте, что Redis запущен:

```bash
redis-cli ping
```

Должен вернуться ответ: `PONG`

## Конфигурация

### Переменные окружения

Вы можете настроить подключение к Redis через переменные окружения:

```bash
# Хост Redis (по умолчанию: localhost)
REDIS_HOST=localhost

# Порт Redis (по умолчанию: 6379)
REDIS_PORT=6379

# Пароль Redis (опционально)
REDIS_PASSWORD=your_password

# Номер базы данных (по умолчанию: 0)
REDIS_DB=0

# Тип сессии для fallback (redis/filesystem/null)
SESSION_TYPE=redis
```

### Файлы конфигурации

Настройки Redis также можно задать в файлах конфигурации:

**config/settings.json:**
```json
{
  "redis": {
    "host": "localhost",
    "port": 6379,
    "password": null,
    "db": 0,
    "socket_timeout": 5,
    "socket_connect_timeout": 5,
    "retry_on_timeout": true
  },
  "session": {
    "type": "redis",
    "permanent": true,
    "lifetime": 86400
  }
}
```

## Fallback механизм

Если Redis недоступен, приложение автоматически переключится на cookie-based сессии Flask. В логах вы увидите предупреждение:

```
Не удалось подключиться к Redis: ... Используется fallback на cookie-based сессии.
```

## Установка зависимостей Python

Убедитесь, что установлены все необходимые пакеты:

```bash
pip install -r requirements.txt
```

Требуемые пакеты:
- `redis==5.0.1`
- `hiredis==2.2.3` (оптимизированный парсер для Redis)
- `Flask-Session==0.5.0`

## Проверка работы

1. **Запустите Redis сервер**

2. **Запустите приложение:**
   ```bash
   python app.py
   ```

3. **Проверьте логи:**
   - При успешном подключении: `Redis подключен успешно: localhost:6379/0`
   - При использовании Redis сессий: `Flask-Session с Redis бэкендом настроен успешно`
   - При fallback: `Используются cookie-based сессии (Redis недоступен)`

4. **Проверьте отсутствие предупреждений:**
   - Не должно быть предупреждений о размере cookie >4093 байт

## Устранение проблем

### Redis не запускается

**Windows:**
- Убедитесь, что порт 6379 не занят другим процессом
- Проверьте файрвол

**Linux:**
```bash
sudo systemctl status redis-server
sudo journalctl -u redis-server
```

### Ошибка подключения

1. Проверьте, что Redis запущен: `redis-cli ping`
2. Проверьте настройки хоста и порта
3. Проверьте файрвол
4. Проверьте логи приложения

### Сессии не сохраняются

1. Убедитесь, что Redis доступен
2. Проверьте настройки `SESSION_PERMANENT` и `PERMANENT_SESSION_LIFETIME`
3. Проверьте логи на наличие ошибок

## Производственное развертывание

Для production рекомендуется:

1. **Настроить пароль для Redis:**
   ```bash
   # В redis.conf
   requirepass your_strong_password
   ```

2. **Использовать переменные окружения:**
   ```bash
   export REDIS_HOST=your-redis-host
   export REDIS_PORT=6379
   export REDIS_PASSWORD=your_strong_password
   ```

3. **Настроить персистентность Redis:**
   - Включить RDB или AOF в redis.conf

4. **Мониторинг:**
   - Использовать Redis CLI для мониторинга: `redis-cli info`
   - Настроить алерты на использование памяти

## Дополнительные возможности

### Кэширование данных

Redis можно использовать для кэширования:
- Результатов расчетов логистики
- Данных каталогов (ТН-ВЭД, материалы)
- Часто используемых запросов к БД

### Очереди задач

Redis можно использовать для фоновых задач:
- Генерация документов в фоне
- Отправка уведомлений
- Обработка импорта данных

## Полезные команды Redis

```bash
# Подключиться к Redis CLI
redis-cli

# Проверить подключение
PING

# Посмотреть все ключи сессий
KEYS session:*

# Посмотреть информацию о Redis
INFO

# Очистить все данные (осторожно!)
FLUSHALL

# Мониторинг команд в реальном времени
MONITOR
```

