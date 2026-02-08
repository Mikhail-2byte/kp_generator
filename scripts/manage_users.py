#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""CLI-утилита для управления пользователями в проекте KP Generator (Microsoft SQL Server)."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Iterable

from werkzeug.security import generate_password_hash

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

# Загрузка .env и инициализация приложения (DATABASE_URL должен быть задан)
try:
    from dotenv import load_dotenv
    load_dotenv(PROJECT_ROOT / ".env")
except ImportError:
    pass

from app import create_app
from app.database import database_service
from app.database.database import _session_scope
from app.models.models import User, UserRecord


def _ensure_app_context():
    """Создаёт приложение и контекст для работы с БД (используется при первом вызове)."""
    app = create_app()
    return app.app_context()


def list_users(detailed: bool = False) -> None:
    """Выводит список всех пользователей (вызывать внутри app context)."""
    if detailed:
        with _session_scope() as session:
            users = session.query(UserRecord).order_by(UserRecord.username).all()
        if not users:
            print("Пользователи не найдены.")
            return
        print(f"\n{'='*100}")
        print(f"Найдено пользователей: {len(users)}")
        print(f"{'='*100}\n")
        for i, user in enumerate(users, 1):
            full_name = f"{user.last_name or ''} {user.first_name or ''}".strip() or "Не указано"
            created = user.created_at.strftime('%Y-%m-%d %H:%M:%S') if user.created_at else "—"
            last_login = user.last_login.strftime('%Y-%m-%d %H:%M:%S') if user.last_login else "—"
            print(f"Пользователь #{i}")
            print(f"  ID: {user.id}")
            print(f"  Логин: {user.username}")
            print(f"  Роль: {user.role}")
            print(f"  ФИО: {full_name}")
            print(f"  Контактная информация: {user.contact_info or 'Не указана'}")
            print(f"  Дата создания: {created}")
            print(f"  Последний вход: {last_login}")
            ph = user.password_hash or ""
            print(f"  Хэш пароля: {ph[:60]}..." if len(ph) > 60 else f"  Хэш пароля: {ph}")
            print(f"{'-'*100}\n")
    else:
        data = database_service.get_users_list(page=1, per_page=10000)
        items = data.get("items") or []
        if not items:
            print("Пользователи не найдены.")
            return
        print(f"Найдено {len(items)} пользователей:")
        for u in items:
            created = u.get("created_at") or "—"
            last_login = u.get("last_login") or "—"
            print(f"- {u['username']} | роль: {u['role']} | создан: {created} | послед. вход: {last_login}")


def reset_password(username: str, new_password: str) -> bool:
    """Сбрасывает пароль для указанного пользователя (вызывать внутри app context)."""
    if not new_password or len(new_password.strip()) < 3:
        print("Ошибка: пароль должен содержать минимум 3 символа.")
        return False

    row = database_service.get_user_by_username(username)
    user = User.from_row(row) if row else None
    if not user:
        print(f"Ошибка: пользователь '{username}' не найден.")
        return False
    user_id = int(user.id)
    password_hash = generate_password_hash(new_password.strip())
    ok = database_service.update_user_profile(
        user_id,
        user.username,
        user.last_name,
        user.first_name,
        user.contact_info,
        password_hash,
    )
    if not ok:
        print("Ошибка при обновлении пароля.", file=sys.stderr)
        return False
    print("[OK] Пароль успешно сброшен для пользователя:", user.username)
    print("  Роль:", user.role)
    print("  Новый пароль:", new_password)
    return True


def update_role(username: str, role: str) -> None:
    """Изменяет роль пользователя (вызывать внутри app context)."""
    with _session_scope() as session:
        user = session.query(UserRecord).filter(UserRecord.username == username).one_or_none()
        if user is None:
            raise ValueError(f"Пользователь '{username}' не найден.")
        user.role = role.lower()


def parse_args(args: Iterable[str] | None = None) -> argparse.Namespace:
    """Парсит аргументы командной строки."""
    parser = argparse.ArgumentParser(
        description="Управление пользователями (просмотр, сброс паролей, управление ролями). Использует DATABASE_URL из .env (Microsoft SQL Server).",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    list_parser = subparsers.add_parser("list", help="Показать всех пользователей.")
    list_parser.add_argument(
        "--detailed",
        action="store_true",
        help="Показать подробную информацию, включая хэши паролей.",
    )

    reset_parser = subparsers.add_parser("reset-password", help="Сбросить пароль пользователя.")
    reset_parser.add_argument("username", help="Логин пользователя (username).")
    reset_parser.add_argument("new_password", help="Новый пароль (минимум 3 символа).")

    set_role_parser = subparsers.add_parser("set-role", help="Изменить роль пользователя.")
    set_role_parser.add_argument("username", help="Логин пользователя (username).")
    set_role_parser.add_argument(
        "role",
        choices=["admin", "user"],
        help="Роль, которую нужно назначить.",
    )

    return parser.parse_args(args or sys.argv[1:])


def main(argv: Iterable[str] | None = None) -> int:
    """Главная функция утилиты."""
    args = parse_args(argv)

    try:
        with _ensure_app_context():
            if args.command == "list":
                list_users(detailed=getattr(args, "detailed", False))
            elif args.command == "reset-password":
                success = reset_password(args.username, args.new_password)
                return 0 if success else 1
            elif args.command == "set-role":
                update_role(args.username, args.role)
                print(f"Роль пользователя '{args.username}' изменена на '{args.role}'.")
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        return 1
    except Exception as exc:
        print(f"Ошибка работы с базой данных: {exc}", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
