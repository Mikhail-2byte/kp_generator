from pathlib import Path

from flask import Flask

from app.core.config import load_config, setup_app_security, setup_logging
from app.database import init_db
from app.core.errors import register_error_handlers
from app.core.extensions import csrf, init_db_engine, login_manager
from app.routes import register_blueprints
from app.services import datasets
from app.services.repositories import user_repository
from app.presentation.ui import register_ui


PROJECT_ROOT = Path(__file__).resolve().parents[1]


def create_app() -> Flask:
    """Создаёт и настраивает экземпляр приложения Flask со всеми зависимостями."""
    app = Flask(
        __name__,
        template_folder=str(PROJECT_ROOT / 'templates'),
        static_folder=str(PROJECT_ROOT / 'static')
    )

    # Загружаем конфигурацию из файлов/переменных окружения и сохраняем в app.config
    app_config = load_config(app)
    app.config['APP_SETTINGS'] = app_config

    # Настраиваем защиту, логирование и подключаем расширения к приложению
    setup_app_security(app, app_config)
    setup_logging(app, app_config)

    init_db_engine(app)
    csrf.init_app(app)
    login_manager.init_app(app)
    login_manager.login_view = 'auth.profile'
    login_manager.login_message_category = 'info'

    @login_manager.user_loader
    def load_user(user_id):  # pragma: no cover - Flask callback
        """Загружает пользователя по идентификатору для менеджера сессий Flask-Login."""
        return user_repository.get_by_id(user_id)

    register_ui(app)
    register_blueprints(app)
    register_error_handlers(app)

    with app.app_context():
        # Инициализируем вспомогательные сервисы и создаём необходимые каталоги
        datasets.init_app(app)

        for folder in ['logs', 'templates_docs']:
            target = PROJECT_ROOT / folder
            target.mkdir(parents=True, exist_ok=True)

        # Производим первичную инициализацию базы данных (например, миграции)
        init_db()

    return app
