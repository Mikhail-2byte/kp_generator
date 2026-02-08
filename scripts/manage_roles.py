#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Утилита для управления ролями пользователей (Microsoft SQL Server).

Использует DATABASE_URL из .env. Для полного управления пользователями
используйте: python scripts/manage_users.py (list, reset-password, set-role).
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Iterable

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

try:
    from dotenv import load_dotenv
    load_dotenv(PROJECT_ROOT / ".env")
except ImportError:
    pass

from app import create_app
from app.database import database_service
from app.database.database import _session_scope
from app.models.models import UserRecord


def _ensure_app_context():
    """Создаёт приложение и контекст для работы с БД."""
    app = create_app()
    return app.app_context()


def list_users() -> None:
    """Выводит список пользователей и их ролей (вызывать внутри app context)."""
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


def update_role(username: str, role: str) -> None:
    """Изменяет роль пользователя (вызывать внутри app context)."""
    with _session_scope() as session:
        user = session.query(UserRecord).filter(UserRecord.username == username).one_or_none()
        if user is None:
            raise ValueError(f"Пользователь '{username}' не найден.")
        user.role = role.lower()


def parse_args(args: Iterable[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Управление ролями пользователей (admin/user). Использует DATABASE_URL из .env.",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)
    subparsers.add_parser("list", help="Показать всех пользователей и их роли.")
    set_role_parser = subparsers.add_parser("set-role", help="Изменить роль пользователя.")
    set_role_parser.add_argument("username", help="Логин пользователя (username).")
    set_role_parser.add_argument(
        "role",
        choices=["admin", "user"],
        help="Роль, которую нужно назначить.",
    )
    return parser.parse_args(args or sys.argv[1:])


def main(argv: Iterable[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        with _ensure_app_context():
            if args.command == "list":
                list_users()
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


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
