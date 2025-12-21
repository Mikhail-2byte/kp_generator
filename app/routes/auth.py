from datetime import datetime
from typing import Union

from flask import Blueprint, Response, flash, redirect, render_template, request, url_for
from flask_login import current_user, login_required, login_user, logout_user
from werkzeug.security import check_password_hash, generate_password_hash

from app.presentation.forms import (
    ContactInfoForm,
    DeleteAccountForm,
    LoginForm,
    ProfileUpdateForm,
    RegistrationForm
)
from app.presentation.ui import build_context
from app.services.audit_service import log_create, log_delete, log_login, log_logout, log_update
from app.services.repositories import user_repository


auth_bp = Blueprint('auth', __name__)  # Блюпринт для операций авторизации и профиля


@auth_bp.route('/profile', methods=['GET', 'POST'])
def profile() -> Union[str, Response]:
    """Обрабатывает авторизацию, регистрацию и управление профилем пользователя."""
    login_form = LoginForm(prefix='login')
    register_form = RegistrationForm(prefix='register')
    update_form = ProfileUpdateForm(prefix='update') if current_user.is_authenticated else None
    delete_form = DeleteAccountForm(prefix='delete') if current_user.is_authenticated else None
    contact_form = ContactInfoForm(prefix='contact') if current_user.is_authenticated else None
    show_registration_modal = False
    show_update_form = bool(update_form and update_form.submit_update.data)
    show_contact_form = bool(contact_form and contact_form.submit.data)

    if request.method == 'POST':
        if login_form.submit_login.data:
            if login_form.validate():
                username = (login_form.username.data or '').strip()
                login_form.username.data = username
                user = user_repository.get_by_username(username)
                if user and check_password_hash(user.password_hash, login_form.password.data):
                    login_user(user, remember=login_form.remember_me.data)
                    user_repository.record_login(user.id)
                    user.set_last_login_now()
                    log_login(user.username, int(user.id) if user.id else None)
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
                if user_repository.get_by_username(username):
                    register_form.username.errors.append('Пользователь с таким логином уже существует.')
                    show_registration_modal = True
                else:
                    password_hash = generate_password_hash((register_form.password.data or '').strip())
                    user = user_repository.create_user(username, password_hash, last_name, first_name)
                    if user:
                        login_user(user)
                        user_repository.record_login(user.id)
                        user.set_last_login_now()
                        log_create(
                            resource_type='user',
                            resource_id=str(user.id) if user.id else None,
                            description=f'Регистрация нового пользователя: {username}',
                            data={'username': username, 'last_name': last_name, 'first_name': first_name},
                        )
                        log_login(user.username, int(user.id) if user.id else None)
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

                existing_user = user_repository.get_by_username(username)
                if existing_user and str(existing_user.id) != str(current_user.id):
                    update_form.username.errors.append('Пользователь с таким логином уже существует.')
                else:
                    # Сохраняем данные до изменения для логирования
                    data_before = {
                        'username': current_user.username,
                        'last_name': current_user.last_name,
                        'first_name': current_user.first_name,
                    }
                    
                    password_hash = generate_password_hash(new_password) if new_password else None
                    updated = user_repository.update_profile(
                        current_user.id,
                        username,
                        last_name,
                        first_name,
                        current_user.contact_info,  # сохраняем текущую контактную информацию
                        password_hash
                    )
                    if updated:
                        data_after = {
                            'username': username,
                            'last_name': last_name,
                            'first_name': first_name,
                            'password_changed': bool(new_password),
                        }
                        log_update(
                            resource_type='user',
                            resource_id=str(current_user.id),
                            description=f'Обновление профиля пользователя: {username}',
                            data_before=data_before,
                            data_after=data_after,
                        )
                        current_user.username = username
                        current_user.last_name = last_name
                        current_user.first_name = first_name
                        flash('Профиль успешно обновлён.', 'success')
                        return redirect(url_for('auth.profile'))
                    flash('Не удалось обновить профиль. Попробуйте позже.', 'danger')
                    show_update_form = True
        elif current_user.is_authenticated and delete_form and delete_form.submit_delete.data:
            if delete_form.validate():
                user_id = current_user.id
                username = current_user.username
                user_data = {
                    'username': username,
                    'last_name': current_user.last_name,
                    'first_name': current_user.first_name,
                }
                if user_repository.delete(user_id):
                    log_delete(
                        resource_type='user',
                        resource_id=str(user_id),
                        description=f'Удаление аккаунта пользователя: {username}',
                        data=user_data,
                    )
                    logout_user()
                    flash('Аккаунт и связанные данные удалены.', 'info')
                    return redirect(url_for('auth.profile'))
                flash('Не удалось удалить аккаунт. Попробуйте позже.', 'danger')
        elif current_user.is_authenticated and contact_form and contact_form.submit.data:
            if contact_form.validate():
                contact_info = (contact_form.contact_info.data or '').strip()
                contact_form.contact_info.data = contact_info

                data_before = {'contact_info': current_user.contact_info}
                updated = user_repository.update_profile(
                    current_user.id,
                    current_user.username,  # сохраняем текущие значения
                    current_user.last_name,
                    current_user.first_name,
                    contact_info,  # обновляем только контактную информацию
                    None  # не меняем пароль
                )
                
                if updated:
                    log_update(
                        resource_type='user',
                        resource_id=str(current_user.id),
                        description=f'Обновление контактной информации пользователя: {current_user.username}',
                        data_before=data_before,
                        data_after={'contact_info': contact_info},
                    )
                    current_user.contact_info = contact_info
                    flash('Контактная информация успешно обновлена.', 'success')
                    return redirect(url_for('auth.profile'))
                else:
                    flash('Не удалось обновить контактную информацию. Попробуйте позже.', 'danger')
                    show_contact_form = True

    stats = None
    user_created_at_formatted = None
    user_last_login_formatted = None
    
    if current_user.is_authenticated:
        stats = user_repository.get_statistics(current_user.id)
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
        
        # Форматируем дату регистрации
        if current_user.created_at:
            if isinstance(current_user.created_at, str):
                try:
                    created_dt = datetime.strptime(current_user.created_at, '%Y-%m-%d %H:%M:%S')
                    user_created_at_formatted = created_dt.strftime('%d.%m.%Y')
                except ValueError:
                    user_created_at_formatted = current_user.created_at[:10] if len(current_user.created_at) >= 10 else current_user.created_at
            else:
                user_created_at_formatted = current_user.created_at.strftime('%d.%m.%Y')
        
        # Форматируем дату последнего входа
        if current_user.last_login:
            if isinstance(current_user.last_login, str):
                try:
                    login_dt = datetime.strptime(current_user.last_login, '%Y-%m-%d %H:%M:%S')
                    user_last_login_formatted = login_dt.strftime('%d.%m.%Y %H:%M')
                except ValueError:
                    user_last_login_formatted = current_user.last_login[:16] if len(current_user.last_login) >= 16 else current_user.last_login
            else:
                user_last_login_formatted = current_user.last_login.strftime('%d.%m.%Y %H:%M')

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

    if request.method != 'POST' and contact_form:
        contact_form.contact_info.data = current_user.contact_info or ''

    if update_form and update_form.errors:
        show_update_form = True
    if contact_form and contact_form.errors:
        show_contact_form = True

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
            contact_form=contact_form,
            show_update_form=show_update_form,
            show_contact_form=show_contact_form,
            user_created_at_formatted=user_created_at_formatted,
            user_last_login_formatted=user_last_login_formatted
        )
    )


@auth_bp.route('/logout')
@login_required
def logout() -> Response:
    """Завершает пользовательскую сессию и возвращает на страницу профиля."""
    user_id = int(current_user.id) if current_user.id else None
    username = current_user.username
    logout_user()
    log_logout(username, user_id)
    flash('Вы вышли из системы.', 'info')
    return redirect(url_for('auth.profile'))
