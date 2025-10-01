from datetime import datetime

from flask import Blueprint, current_app, flash, redirect, render_template, request, url_for
from flask_login import current_user, login_required, login_user, logout_user
from werkzeug.security import check_password_hash, generate_password_hash

from app.database import (
    create_user,
    delete_user,
    get_user_by_id,
    get_user_by_username,
    get_user_statistics,
    update_last_login,
    update_user_profile
)
from app.forms import (
    DeleteAccountForm,
    LoginForm,
    ProfileUpdateForm,
    RegistrationForm
)
from app.models import User
from app.ui import build_context


auth_bp = Blueprint('auth', __name__)


@auth_bp.route('/profile', methods=['GET', 'POST'])
def profile():
    login_form = LoginForm(prefix='login')
    register_form = RegistrationForm(prefix='register')
    update_form = ProfileUpdateForm(prefix='update') if current_user.is_authenticated else None
    delete_form = DeleteAccountForm(prefix='delete') if current_user.is_authenticated else None
    show_registration_modal = False
    show_update_form = bool(update_form and update_form.submit_update.data)

    if request.method == 'POST':
        if login_form.submit_login.data:
            if login_form.validate():
                username = (login_form.username.data or '').strip()
                login_form.username.data = username
                user_row = get_user_by_username(username)
                if user_row and check_password_hash(user_row[2], login_form.password.data):
                    user = User.from_row(user_row)
                    login_user(user, remember=login_form.remember_me.data)
                    update_last_login(int(user.id))
                    user.set_last_login_now()
                    flash('Вы успешно вошли в систему.', 'success')
                    return redirect(url_for('auth.profile'))
                flash('Неверный логин или пароль.', 'danger')
        elif register_form.submit_register.data:
            if register_form.validate():
                username = (register_form.username.data or '').strip()
                last_name = (register_form.last_name.data or '').strip()
                first_name = (register_form.first_name.data or '').strip()
                register_form.username.data = username
                register_form.last_name.data = last_name
                register_form.first_name.data = first_name
                if get_user_by_username(username):
                    register_form.username.errors.append('Пользователь с таким логином уже существует.')
                    show_registration_modal = True
                else:
                    password_hash = generate_password_hash((register_form.password.data or '').strip())
                    new_user_id = create_user(username, password_hash, last_name, first_name)
                    if new_user_id:
                        user = User.from_row(get_user_by_id(new_user_id))
                        login_user(user)
                        update_last_login(int(user.id))
                        user.set_last_login_now()
                        flash('Аккаунт успешно создан.', 'success')
                        return redirect(url_for('auth.profile'))
                    flash('Не удалось создать пользователя. Попробуйте позже.', 'danger')
                    show_registration_modal = True
            else:
                show_registration_modal = True
        elif current_user.is_authenticated and update_form and update_form.submit_update.data:
            if update_form.validate():
                username = (update_form.username.data or '').strip()
                last_name = (update_form.last_name.data or '').strip()
                first_name = (update_form.first_name.data or '').strip()
                new_password = (update_form.new_password.data or '').strip()
                update_form.username.data = username
                update_form.last_name.data = last_name
                update_form.first_name.data = first_name

                existing_user = get_user_by_username(username)
                if existing_user and str(existing_user[0]) != str(current_user.id):
                    update_form.username.errors.append('Пользователь с таким логином уже существует.')
                else:
                    password_hash = generate_password_hash(new_password) if new_password else None
                    updated = update_user_profile(
                        int(current_user.id),
                        username,
                        last_name,
                        first_name,
                        password_hash
                    )
                    if updated:
                        current_user.username = username
                        current_user.last_name = last_name
                        current_user.first_name = first_name
                        flash('Профиль успешно обновлён.', 'success')
                        return redirect(url_for('auth.profile'))
                    flash('Не удалось обновить профиль. Попробуйте позже.', 'danger')
                    show_update_form = True
        elif current_user.is_authenticated and delete_form and delete_form.submit_delete.data:
            if delete_form.validate():
                if delete_user(int(current_user.id)):
                    logout_user()
                    flash('Аккаунт и связанные данные удалены.', 'info')
                    return redirect(url_for('auth.profile'))
                flash('Не удалось удалить аккаунт. Попробуйте позже.', 'danger')

    stats = None
    if current_user.is_authenticated:
        stats = get_user_statistics(int(current_user.id))
        last_gen = stats.get('last_generation_at')
        if last_gen:
            try:
                stats['last_generation_at_formatted'] = datetime.strptime(
                    last_gen, '%Y-%m-%d %H:%M:%S'
                ).strftime('%d.%m.%Y %H:%M')
            except ValueError:
                stats['last_generation_at_formatted'] = last_gen
        else:
            stats['last_generation_at_formatted'] = None

        formatted_recent = []
        for record in stats.get('recent_generations', []):
            record_id, company, product, margin_percent, final_price, timestamp_value = record
            try:
                formatted_ts = datetime.strptime(
                    timestamp_value, '%Y-%m-%d %H:%M:%S'
                ).strftime('%d.%m.%Y %H:%M')
            except ValueError:
                formatted_ts = timestamp_value
            formatted_recent.append({
                'id': record_id,
                'company': company,
                'product': product,
                'margin_percent': margin_percent,
                'final_price': final_price,
                'timestamp': formatted_ts
            })
        stats['recent_generations'] = formatted_recent

        if request.method != 'POST' and update_form:
            update_form.username.data = current_user.username
            update_form.last_name.data = current_user.last_name or ''
            update_form.first_name.data = current_user.first_name or ''

    if update_form and update_form.errors:
        show_update_form = True

    return render_template(
        'profile.html',
        **build_context(
            'profile',
            'Профиль',
            login_form=login_form,
            register_form=register_form,
            stats=stats,
            show_registration_modal=show_registration_modal,
            update_form=update_form,
            delete_form=delete_form,
            show_update_form=show_update_form
        )
    )


@auth_bp.route('/logout')
@login_required
def logout():
    logout_user()
    flash('Вы вышли из системы.', 'info')
    return redirect(url_for('auth.profile'))
