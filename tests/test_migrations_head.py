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


