import os
from contextlib import contextmanager
from typing import Iterator

import pytest

from app import create_app
from app.extensions import SessionLocal
from flask import Flask


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

