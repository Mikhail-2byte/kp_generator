import sys
from pathlib import Path

# Добавляем корень проекта в путь для возможности прямого запуска
project_root = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(project_root))

import pytest

from app.database import get_migration_status, apply_migrations


def test_migrations_are_up_to_date():
    """Проверяем, что миграции Alembic могут быть применены до head без ошибок."""
    # Пытаемся явно применить миграции до последней версии.
    # Если история ревизий неконсистентна, этот вызов вызовет исключение.
    assert apply_migrations('head') is True

    status = get_migration_status()
    assert status['configured'] is True
    assert status['pending'] is False
    assert status['head_revision'] in status['head_revisions']


if __name__ == "__main__":
    # Запуск тестов при прямом выполнении файла
    pytest.main([__file__, "-v"])
