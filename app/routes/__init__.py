from .auth import auth_bp
from .main import main_bp
from .admin import admin_bp

__all__ = ['register_blueprints']


def register_blueprints(app):
    """Подключает все блюпринты приложения к экземпляру Flask."""
    app.register_blueprint(main_bp)
    app.register_blueprint(auth_bp)
    app.register_blueprint(admin_bp)
