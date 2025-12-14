"""Сервис для логирования действий пользователей в системе аудита."""

from __future__ import annotations

from typing import Any, Dict, Optional

from flask import request
from flask_login import current_user

from app.services.repositories import audit_log_repository


def _get_client_info() -> tuple[Optional[str], Optional[str]]:
    """Извлекает IP-адрес и User-Agent из запроса."""
    ip_address = None
    user_agent = None

    if request:
        # Получаем реальный IP (с учетом прокси)
        if request.headers.get('X-Forwarded-For'):
            ip_address = request.headers.get('X-Forwarded-For').split(',')[0].strip()
        elif request.headers.get('X-Real-IP'):
            ip_address = request.headers.get('X-Real-IP')
        else:
            ip_address = request.remote_addr

        user_agent = request.headers.get('User-Agent')

    return ip_address, user_agent


def _get_current_user_info() -> tuple[Optional[int], str]:
    """Получает информацию о текущем пользователе."""
    if current_user and current_user.is_authenticated:
        user_id = int(current_user.id) if current_user.id else None
        username = current_user.username or 'unknown'
        return user_id, username
    return None, 'anonymous'


def log_action(
    action_type: str,
    description: str,
    resource_type: Optional[str] = None,
    resource_id: Optional[str] = None,
    changes_before: Optional[Dict[str, Any]] = None,
    changes_after: Optional[Dict[str, Any]] = None,
    user_id: Optional[int] = None,
    username: Optional[str] = None,
) -> bool:
    """
    Логирует действие пользователя.

    Args:
        action_type: Тип действия (login, logout, create, update, delete, view, etc.)
        description: Описание действия
        resource_type: Тип ресурса (user, generation, duty, material, logistics, etc.)
        resource_id: ID ресурса
        changes_before: Данные до изменения (словарь)
        changes_after: Данные после изменения (словарь)
        user_id: ID пользователя (если не указан, берется из current_user)
        username: Имя пользователя (если не указано, берется из current_user)

    Returns:
        True если логирование успешно, False в противном случае
    """
    if user_id is None or username is None:
        current_user_id, current_username = _get_current_user_info()
        user_id = user_id or current_user_id
        username = username or current_username

    ip_address, user_agent = _get_client_info()

    return audit_log_repository.create_log(
        user_id=user_id,
        username=username,
        action_type=action_type,
        description=description,
        resource_type=resource_type,
        resource_id=resource_id,
        ip_address=ip_address,
        user_agent=user_agent,
        changes_before=changes_before,
        changes_after=changes_after,
    )


def log_login(username: str, user_id: Optional[int] = None) -> bool:
    """Логирует вход пользователя в систему."""
    return log_action(
        action_type='login',
        description=f'Вход в систему: {username}',
        user_id=user_id,
        username=username,
    )


def log_logout(username: str, user_id: Optional[int] = None) -> bool:
    """Логирует выход пользователя из системы."""
    return log_action(
        action_type='logout',
        description=f'Выход из системы: {username}',
        user_id=user_id,
        username=username,
    )


def log_create(
    resource_type: str,
    resource_id: Optional[str],
    description: Optional[str] = None,
    data: Optional[Dict[str, Any]] = None,
) -> bool:
    """Логирует создание ресурса."""
    desc = description or f'Создан {resource_type}'
    if resource_id:
        desc += f' (ID: {resource_id})'

    return log_action(
        action_type='create',
        description=desc,
        resource_type=resource_type,
        resource_id=resource_id,
        changes_after=data,
    )


def log_update(
    resource_type: str,
    resource_id: Optional[str],
    description: Optional[str] = None,
    data_before: Optional[Dict[str, Any]] = None,
    data_after: Optional[Dict[str, Any]] = None,
) -> bool:
    """Логирует обновление ресурса."""
    desc = description or f'Обновлен {resource_type}'
    if resource_id:
        desc += f' (ID: {resource_id})'

    return log_action(
        action_type='update',
        description=desc,
        resource_type=resource_type,
        resource_id=resource_id,
        changes_before=data_before,
        changes_after=data_after,
    )


def log_delete(
    resource_type: str,
    resource_id: Optional[str],
    description: Optional[str] = None,
    data: Optional[Dict[str, Any]] = None,
) -> bool:
    """Логирует удаление ресурса."""
    desc = description or f'Удален {resource_type}'
    if resource_id:
        desc += f' (ID: {resource_id})'

    return log_action(
        action_type='delete',
        description=desc,
        resource_type=resource_type,
        resource_id=resource_id,
        changes_before=data,
    )


def log_view(
    resource_type: str,
    resource_id: Optional[str] = None,
    description: Optional[str] = None,
) -> bool:
    """Логирует просмотр ресурса."""
    desc = description or f'Просмотр {resource_type}'
    if resource_id:
        desc += f' (ID: {resource_id})'

    return log_action(
        action_type='view',
        description=desc,
        resource_type=resource_type,
        resource_id=resource_id,
    )


def log_generation_created(generation_id: int, company: str, tender_number: Optional[str] = None) -> bool:
    """Логирует создание генерации КП."""
    desc = f'Создана генерация КП для компании "{company}"'
    if tender_number:
        desc += f' (тендер: {tender_number})'

    return log_action(
        action_type='create',
        description=desc,
        resource_type='generation',
        resource_id=str(generation_id),
        changes_after={'generation_id': generation_id, 'company': company, 'tender_number': tender_number},
    )


def log_document_export(export_type: str, filters: Optional[Dict[str, Any]] = None) -> bool:
    """Логирует экспорт документов."""
    desc = f'Экспорт {export_type}'
    if filters:
        desc += f' (фильтры: {filters})'

    return log_action(
        action_type='export',
        description=desc,
        resource_type='export',
        changes_after={'export_type': export_type, 'filters': filters},
    )

