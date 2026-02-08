import os
from typing import Optional

from flask_login import LoginManager
from flask_wtf import CSRFProtect
from sqlalchemy import create_engine
from sqlalchemy.orm import scoped_session, sessionmaker
from sqlalchemy.pool import StaticPool

csrf = CSRFProtect()

login_manager = LoginManager()
login_manager.login_view = 'auth.profile'
login_manager.login_message_category = 'info'


_engine = None
SessionLocal = scoped_session(
    sessionmaker(
        autocommit=False,
        autoflush=False,
        expire_on_commit=False,
        future=True,
    )
)


def get_database_url() -> str:
    """Возвращает строку подключения к БД из окружения (только Microsoft SQL Server в production)."""
    return os.environ.get('DATABASE_URL', '')


def get_engine():
    """Возвращает текущий движок SQLAlchemy (после вызова init_db_engine)."""
    if _engine is None:
        raise RuntimeError('Database engine not initialized. Call init_db_engine(app) first.')
    return _engine


def init_db_engine(app) -> Optional[object]:
    """Создаёт движок SQLAlchemy и регистрирует очистку сессий в жизненном цикле Flask."""
    global _engine
    if _engine is not None:
        return _engine

    database_url = (get_database_url() or '').strip()
    use_test_sqlite = os.environ.get('USE_TEST_SQLITE', '').strip().lower() in ('1', 'true', 'yes', 'on')

    # Разрешаем SQLite только при явном USE_TEST_SQLITE (для pytest без MSSQL)
    if use_test_sqlite and database_url.startswith('sqlite'):
        engine_options = {
            'future': True,
            'pool_pre_ping': True,
            'connect_args': {'check_same_thread': False},
        }
        if database_url == 'sqlite:///:memory:':
            engine_options['poolclass'] = StaticPool
    elif database_url.startswith('mssql+pyodbc'):
        def _safe_int(name: str, default: int) -> int:
            value = os.environ.get(name)
            if value is None:
                return default
            try:
                return int(value)
            except (TypeError, ValueError):  # pragma: no cover - defensive
                return default

        engine_options = {
            'future': True,
            'pool_pre_ping': True,
            'pool_size': _safe_int('DATABASE_POOL_SIZE', 5),
            'max_overflow': _safe_int('DATABASE_MAX_OVERFLOW', 10),
            'pool_timeout': _safe_int('DATABASE_POOL_TIMEOUT', 30),
        }
    else:
        raise RuntimeError(
            'Требуется DATABASE_URL в формате mssql+pyodbc://... '
            'См. docs/MSSQL_SETUP.md. Для запуска тестов без MSSQL задайте USE_TEST_SQLITE=1 и DATABASE_URL=sqlite:///:memory:'
        )

    _engine = create_engine(database_url, **engine_options)
    SessionLocal.configure(bind=_engine)

    @app.teardown_appcontext
    def remove_session(exception=None):  # pragma: no cover - Flask teardown hook
        """Закрывает сессию после обработки запроса, предотвращая утечки соединений."""
        SessionLocal.remove()

    return _engine