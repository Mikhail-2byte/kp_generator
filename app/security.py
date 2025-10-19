from functools import wraps

from flask import current_app, flash, redirect, request, url_for
from flask_login import current_user


def admin_required(view_func):
    """Декоратор, ограничивающий доступ к маршруту только администраторам."""
    @wraps(view_func)
    def wrapped_view(*args, **kwargs):
        """Проверяет роль текущего пользователя и перенаправляет при отсутствии прав."""
        if _is_debugger_resource_request():
            return view_func(*args, **kwargs)

        if not current_user.is_authenticated:
            return current_app.login_manager.unauthorized()

        if not current_user.is_admin:
            flash('Недостаточно прав для доступа к этой странице.', 'danger')
            return redirect(url_for('auth.profile'))
        return view_func(*args, **kwargs)

    return wrapped_view


def _is_debugger_resource_request() -> bool:
    """Разрешает запросы встроенного отладчика Flask, чтобы он мог загрузить свои ресурсы."""
    try:
        return (
            current_app and current_app.debug
            and request.args.get('__debugger__') == 'yes'
            and request.args.get('cmd') == 'resource'
        )
    except RuntimeError:
        return False
