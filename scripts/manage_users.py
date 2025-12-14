#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""CLI-утилита для управления пользователями в проекте KP Generator."""

from __future__ import annotations

import argparse
import sqlite3
import sys
from pathlib import Path
from typing import Iterable

from werkzeug.security import generate_password_hash

# Определяем путь к базе данных относительно корня проекта
PROJECT_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_DB_PATH = PROJECT_ROOT / "kp_generator.db"


def get_connection(db_path: Path) -> sqlite3.Connection:
    """Создаёт соединение с базой данных."""
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn


def list_users(conn: sqlite3.Connection, detailed: bool = False) -> None:
    """Выводит список всех пользователей."""
    if detailed:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT 
                id, 
                username, 
                role, 
                last_name, 
                first_name, 
                contact_info, 
                created_at, 
                last_login, 
                password_hash
            FROM users 
            ORDER BY username
        """)
        rows = cursor.fetchall()

        if not rows:
            print("Пользователи не найдены.")
            return

        print(f"\n{'='*100}")
        print(f"Найдено пользователей: {len(rows)}")
        print(f"{'='*100}\n")

        for i, row in enumerate(rows, 1):
            full_name = f"{row['last_name'] or ''} {row['first_name'] or ''}".strip()
            if not full_name:
                full_name = "Не указано"

            print(f"Пользователь #{i}")
            print(f"  ID: {row['id']}")
            print(f"  Логин: {row['username']}")
            print(f"  Роль: {row['role']}")
            print(f"  ФИО: {full_name}")
            print(f"  Контактная информация: {row['contact_info'] or 'Не указана'}")
            print(f"  Дата создания: {row['created_at'] or '—'}")
            print(f"  Последний вход: {row['last_login'] or '—'}")
            print(f"  Хэш пароля: {row['password_hash'][:60]}...")
            print(f"  (Полный хэш: {row['password_hash']})")
            print(f"{'-'*100}\n")
    else:
        rows = conn.execute(
            "SELECT id, username, role, created_at, last_login FROM users ORDER BY username"
        ).fetchall()

        if not rows:
            print("Пользователи не найдены.")
            return

        print(f"Найдено {len(rows)} пользователей:")
        for row in rows:
            created = row["created_at"] or "—"
            last_login = row["last_login"] or "—"
            print(
                f"- {row['username']} | роль: {row['role']} | создан: {created} | послед. вход: {last_login}"
            )


def reset_password(conn: sqlite3.Connection, username: str, new_password: str) -> bool:
    """Сбрасывает пароль для указанного пользователя."""
    if not new_password or len(new_password.strip()) < 3:
        print("Ошибка: пароль должен содержать минимум 3 символа.")
        return False

    cursor = conn.cursor()

    # Проверяем, существует ли пользователь
    cursor.execute("SELECT id, username, role FROM users WHERE username = ?", (username,))
    user = cursor.fetchone()

    if not user:
        print(f"Ошибка: пользователь '{username}' не найден.")
        return False

    # Генерируем новый хэш пароля
    password_hash = generate_password_hash(new_password.strip())

    # Обновляем пароль
    cursor.execute(
        "UPDATE users SET password_hash = ? WHERE username = ?",
        (password_hash, username)
    )
    conn.commit()

    print("[OK] Пароль успешно сброшен для пользователя:", user['username'])
    print("  Роль:", user['role'])
    print("  Новый пароль:", new_password)
    return True


def update_role(conn: sqlite3.Connection, username: str, role: str) -> None:
    """Изменяет роль пользователя."""
    cur = conn.execute(
        "UPDATE users SET role = ? WHERE username = ?",
        (role, username),
    )
    if cur.rowcount == 0:
        raise ValueError(f"Пользователь '{username}' не найден.")
    conn.commit()


def parse_args(args: Iterable[str] | None = None) -> argparse.Namespace:
    """Парсит аргументы командной строки."""
    parser = argparse.ArgumentParser(
        description="Управление пользователями (просмотр, сброс паролей, управление ролями).",
    )

    parser.add_argument(
        "--db",
        type=Path,
        default=DEFAULT_DB_PATH,
        help="Путь к файлу базы данных (по умолчанию kp_generator.db в корне проекта).",
    )

    subparsers = parser.add_subparsers(dest="command", required=True)

    # Команда list
    list_parser = subparsers.add_parser("list", help="Показать всех пользователей.")
    list_parser.add_argument(
        "--detailed",
        action="store_true",
        help="Показать подробную информацию, включая хэши паролей.",
    )

    # Команда reset-password
    reset_parser = subparsers.add_parser(
        "reset-password", help="Сбросить пароль пользователя."
    )
    reset_parser.add_argument("username", help="Логин пользователя (username).")
    reset_parser.add_argument("new_password", help="Новый пароль (минимум 3 символа).")

    # Команда set-role
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

    db_path: Path = args.db
    if not db_path.exists():
        print(f"Файл базы данных не найден: {db_path}", file=sys.stderr)
        return 1

    try:
        with get_connection(db_path) as conn:
            if args.command == "list":
                list_users(conn, detailed=getattr(args, "detailed", False))
            elif args.command == "reset-password":
                success = reset_password(conn, args.username, args.new_password)
                return 0 if success else 1
            elif args.command == "set-role":
                update_role(conn, args.username, args.role)
                print(f"Роль пользователя '{args.username}' изменена на '{args.role}'.")
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        return 1
    except sqlite3.Error as exc:
        print(f"Ошибка работы с базой данных: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:
        print(f"Неожиданная ошибка: {exc}", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())

