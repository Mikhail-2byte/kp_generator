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
    """Возвращает строку подключения к БД из окружения или использует SQLite по умолчанию."""
    return os.environ.get('DATABASE_URL', 'sqlite:///kp_generator.db')


def init_db_engine(app) -> Optional[object]:
    """Создаёт движок SQLAlchemy и регистрирует очистку сессий в жизненном цикле Flask."""
    global _engine
    if _engine is not None:
        return _engine

    database_url = get_database_url()

    engine_options = {
        'future': True,
        'pool_pre_ping': True,
    }

    def _safe_int(name: str, default: int) -> int:
        value = os.environ.get(name)
        if value is None:
            return default
        try:
            return int(value)
        except (TypeError, ValueError):  # pragma: no cover - defensive
            return default

    if database_url.startswith('sqlite'):  # SQLite needs special handling for multi-threaded servers
        engine_options['connect_args'] = {'check_same_thread': False}

        # In-memory SQLite requires StaticPool for shared connections
        if database_url == 'sqlite:///:memory:':
            engine_options['poolclass'] = StaticPool
    else:
        engine_options['pool_size'] = _safe_int('DATABASE_POOL_SIZE', 5)
        engine_options['max_overflow'] = _safe_int('DATABASE_MAX_OVERFLOW', 10)
        engine_options['pool_timeout'] = _safe_int('DATABASE_POOL_TIMEOUT', 30)

    _engine = create_engine(database_url, **engine_options)
    SessionLocal.configure(bind=_engine)

    @app.teardown_appcontext
    def remove_session(exception=None):  # pragma: no cover - Flask teardown hook
        """Закрывает сессию после обработки запроса, предотвращая утечки соединений."""
        SessionLocal.remove()

    return _engine
