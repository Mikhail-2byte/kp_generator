from .auth import auth_bp
from .main import main_bp
from .admin import admin_bp
from .health import health_bp
from .api import api_bp
from .api_docs import api_docs_bp

__all__ = ['register_blueprints']


def register_blueprints(app):
    """Подключает все блюпринты приложения к экземпляру Flask."""
    app.register_blueprint(main_bp)
    app.register_blueprint(auth_bp)
    app.register_blueprint(admin_bp)
    app.register_blueprint(health_bp)
    app.register_blueprint(api_bp)
    app.register_blueprint(api_docs_bp)
