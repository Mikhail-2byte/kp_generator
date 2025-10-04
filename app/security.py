from functools import wraps

from flask import flash, redirect, url_for
from flask_login import current_user, login_required


def admin_required(view_func):
    """Декоратор, ограничивающий доступ к маршруту только администраторам."""
    @wraps(view_func)
    @login_required
    def wrapped_view(*args, **kwargs):
        """Проверяет роль текущего пользователя и перенаправляет при отсутствии прав."""
        if not current_user.is_admin:
            flash('Недостаточно прав для доступа к этой странице.', 'danger')
            return redirect(url_for('auth.profile'))
        return view_func(*args, **kwargs)

    return wrapped_view
