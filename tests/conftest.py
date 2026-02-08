import os
from contextlib import contextmanager
from typing import Iterator

import pytest
from flask import Flask

from app import create_app
from app.core.extensions import SessionLocal


@contextmanager
def _override_env(**kwargs) -> Iterator[None]:
    original = {key: os.environ.get(key) for key in kwargs}
    try:
        for key, value in kwargs.items():
            if value is None and key in os.environ:
                del os.environ[key]
            elif value is not None:
                os.environ[key] = value
        yield
    finally:
        for key, value in original.items():
            if value is None:
                os.environ.pop(key, None)
            else:
                os.environ[key] = value


@pytest.fixture
def app(tmp_path) -> Iterator[Flask]:
    db_path = tmp_path / "test.db"

    with _override_env(
        DATABASE_URL=f"sqlite:///{db_path}",
        USE_TEST_SQLITE="1",
        SECRET_KEY="test-secret",
        FLASK_ENV="testing",
        FLASK_DEBUG="False",
        USE_WAITRESS=None,
    ):
        flask_app = create_app()
        flask_app.config.update(
            TESTING=True,
            WTF_CSRF_ENABLED=False,
            SERVER_NAME="localhost",  # для url_for в тестах
        )
        yield flask_app

    SessionLocal.remove()


@pytest.fixture
def client(app):
    return app.test_client()


@pytest.fixture
def logged_in_client(client, admin_user, app):
    """Клиент с авторизованным администратором."""
    with app.app_context():
        with client.session_transaction() as sess:
            sess['_user_id'] = str(admin_user.id)
            sess['_fresh'] = True
    return client


@pytest.fixture
def admin_user(app):
    """Создает тестового администратора."""
    from app.database.service import DatabaseService
    from app.models.models import User
    from werkzeug.security import generate_password_hash
    
    with app.app_context():
        db_service = DatabaseService()
        
        # Проверяем, существует ли уже пользователь
        existing_user_data = db_service.get_user_by_username('testadmin')
        if existing_user_data:
            # get_user_by_username возвращает кортеж, преобразуем в User
            return User.from_row(existing_user_data)
        
        # Создаем нового администратора
        password_hash = generate_password_hash('admin123')
        user_id = db_service.create_user(
            username='testadmin',
            password_hash=password_hash,
            role='admin'
        )
        
        # Получаем созданного пользователя и преобразуем в User
        user_data = db_service.get_user_by_id(user_id)
        return User.from_row(user_data)