from pathlib import Path

from flask import Flask

from app.config import load_config, setup_app_security, setup_logging
from app.database import get_user_by_id, init_db
from app.errors import register_error_handlers
from app.extensions import csrf, login_manager
from app.models import User
from app.routes import register_blueprints
from app.services import datasets
from app.ui import register_ui


PROJECT_ROOT = Path(__file__).resolve().parents[1]


def create_app() -> Flask:
    app = Flask(
        __name__,
        template_folder=str(PROJECT_ROOT / 'templates'),
        static_folder=str(PROJECT_ROOT / 'static')
    )

    app_config = load_config(app)
    app.config['APP_SETTINGS'] = app_config

    setup_app_security(app, app_config)
    setup_logging(app, app_config)

    csrf.init_app(app)
    login_manager.init_app(app)
    login_manager.login_view = 'auth.profile'
    login_manager.login_message_category = 'info'

    @login_manager.user_loader
    def load_user(user_id):  # pragma: no cover - Flask callback
        row = get_user_by_id(user_id)
        return User.from_row(row) if row else None

    register_ui(app)
    register_blueprints(app)
    register_error_handlers(app)

    with app.app_context():
        datasets.init_app(app)

        for folder in ['logs', 'templates_docs']:
            target = PROJECT_ROOT / folder
            target.mkdir(parents=True, exist_ok=True)

        init_db()

    return app
