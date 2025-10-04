import os
from typing import Optional

from flask_login import LoginManager
from flask_wtf import CSRFProtect
from sqlalchemy import create_engine
from sqlalchemy.orm import scoped_session, sessionmaker


csrf = CSRFProtect()

login_manager = LoginManager()
login_manager.login_view = 'auth.profile'
login_manager.login_message_category = 'info'


_engine = None
SessionLocal = scoped_session(sessionmaker(autocommit=False, autoflush=False, future=True))


def get_database_url() -> str:
    """Возвращает строку подключения к БД из окружения или использует SQLite по умолчанию."""
    return os.environ.get('DATABASE_URL', 'sqlite:///kp_generator.db')


def init_db_engine(app) -> Optional[object]:
    """Создаёт движок SQLAlchemy и регистрирует очистку сессий в жизненном цикле Flask."""
    global _engine
    if _engine is not None:
        return _engine

    database_url = get_database_url()
    _engine = create_engine(database_url, future=True)
    SessionLocal.configure(bind=_engine)

    @app.teardown_appcontext
    def remove_session(exception=None):  # pragma: no cover - Flask teardown hook
        """Закрывает сессию после обработки запроса, предотвращая утечки соединений."""
        SessionLocal.remove()

    return _engine
