# Структура тестов проекта

## 📁 tests/
Все pytest тесты находятся здесь:

### tests/ai_agent/
- `test_ai_agent_improvements.py` - Unit тесты AI агента (pytest + mock)
- `test_ai_manual.py` - Ручное тестирование без mock (для отладки)

### tests/api/
- API endpoints тесты

### tests/auth/
- Тесты авторизации

### tests/business/
- Тесты бизнес-логики

### tests/database/
- Тесты репозиториев и БД

### tests/datasets/
- Тесты загрузки данных

### tests/generation/
- Тесты генерации документов

### tests/history/
- Тесты истории

### tests/logistics/
- Тесты логистики

### tests/migrations/
- Тесты миграций

### tests/presentation/
- Тесты UI/валидации

## 🔧 Диагностические скрипты

### ai_agent/test_imports.py
- Проверка всех импортов AI агента
- Диагностика проблем с зависимостями
- Используется при первом запуске/отладке

**Запуск:**
```bash
python ai_agent/test_imports.py
```

## 🧪 Запуск тестов

### Все тесты:
```bash
pytest
```

### Только AI агент:
```bash
pytest tests/ai_agent/
```

### Ручное тестирование AI:
```bash
python tests/ai_agent/test_ai_manual.py
```

### Проверка импортов:
```bash
python ai_agent/test_imports.py
```


