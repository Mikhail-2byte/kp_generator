#!/usr/bin/env python
"""Утилита для управления ролями пользователей в базе kp_generator.db."""

from __future__ import annotations

import argparse
import sqlite3
import sys
from pathlib import Path
from typing import Iterable


DEFAULT_DB_PATH = Path(__file__).resolve().parent / "kp_generator.db"


def get_connection(db_path: Path) -> sqlite3.Connection:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn


def list_users(conn: sqlite3.Connection) -> None:
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


def update_role(conn: sqlite3.Connection, username: str, role: str) -> None:
    cur = conn.execute(
        "UPDATE users SET role = ? WHERE username = ?",
        (role, username),
    )
    if cur.rowcount == 0:
        raise ValueError(f"Пользователь '{username}' не найден.")
    conn.commit()


def parse_args(args: Iterable[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Управление ролями пользователей (admin/user).",
    )

    parser.add_argument(
        "--db",
        type=Path,
        default=DEFAULT_DB_PATH,
        help="Путь к файлу базы данных (по умолчанию kp_generator.db в корне проекта).",
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

    return parser.parse_args(args)


def main(argv: Iterable[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])

    db_path: Path = args.db
    if not db_path.exists():
        print(f"Файл базы данных не найден: {db_path}", file=sys.stderr)
        return 1

    try:
        with get_connection(db_path) as conn:
            if args.command == "list":
                list_users(conn)
            elif args.command == "set-role":
                update_role(conn, args.username, args.role)
                print(
                    f"Роль пользователя '{args.username}' изменена на '{args.role}'."
                )
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        return 1
    except sqlite3.Error as exc:
        print(f"Ошибка работы с базой данных: {exc}", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
