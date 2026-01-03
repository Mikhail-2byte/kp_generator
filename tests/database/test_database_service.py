import sys
from datetime import datetime, timedelta
from pathlib import Path

# Добавляем корень проекта в путь для возможности прямого запуска
project_root = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(project_root))

import pytest

from app.database import database_service


def test_database_service_user_lifecycle(app):
    service = database_service

    user_id = service.create_user(
        username='service-user',
        password_hash='hash',
        last_name='Иванов',
        first_name='Иван',
        role='admin',
        contact_info='test@example.com',
    )

    assert user_id is not None

    by_username = service.get_user_by_username('service-user')
    assert by_username is not None
    assert by_username[1] == 'service-user'
    assert by_username[8] == 'admin'

    by_id = service.get_user_by_id(int(user_id))
    assert by_id is not None
    assert by_id[0] == user_id

    updated = service.update_last_login(user_id)
    assert updated is True

    with_login = service.get_user_by_id(user_id)
    assert isinstance(with_login[4], datetime)
    assert with_login[4].tzinfo is not None
    assert with_login[4].tzinfo.utcoffset(with_login[4]) == timedelta(0)

    stats = service.get_user_statistics(user_id)
    assert stats['total_generations'] == 0

    deleted = service.delete_user(user_id)
    assert deleted is True
    assert service.get_user_by_id(user_id) is None


if __name__ == "__main__":
    # Запуск тестов при прямом выполнении файла
    pytest.main([__file__, "-v"])

