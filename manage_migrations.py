#!/usr/bin/env python
"""CLI-утилита для управления миграциями Alembic в проекте KP Generator."""

from __future__ import annotations

import argparse
import sys
from typing import Iterable

from alembic import command

from app.database.database import (
    apply_migrations,
    downgrade_migrations,
    get_alembic_config,
    get_migration_status,
)


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description='Управление миграциями Alembic (статус/upgrade/downgrade/history).'
    )
    subparsers = parser.add_subparsers(dest='command', required=True)

    subparsers.add_parser('status', help='Показать текущие версии миграций и наличие отставания.')

    upgrade_parser = subparsers.add_parser(
        'upgrade', help='Применить миграции до указанной ревизии (по умолчанию head).'
    )
    upgrade_parser.add_argument(
        'revision',
        nargs='?',
        default='head',
        help='Целевая ревизия (формат Alembic). По умолчанию head.',
    )

    downgrade_parser = subparsers.add_parser(
        'downgrade', help='Откатить миграции до указанной ревизии.'
    )
    downgrade_parser.add_argument(
        'revision',
        help='Целевая ревизия (например, "-1" для отката на одну миграцию).',
    )

    history_parser = subparsers.add_parser('history', help='Показать историю миграций Alembic.')
    history_parser.add_argument(
        '--verbose',
        action='store_true',
        help='Выводить полные сведения (ревизия, автор, комментарии).',
    )

    return parser


def _cmd_status() -> int:
    try:
        status = get_migration_status()
    except Exception as exc:  # pragma: no cover - CLI вывод
        print(f'Не удалось получить статус миграций: {exc}', file=sys.stderr)
        return 1

    if not status.get('configured'):
        print('Файл alembic.ini не найден. Настройте Alembic перед запуском миграций.')
        return 1

    current = status.get('current_revision') or '—'
    head = status.get('head_revision') or '—'
    pending = status.get('pending', False)

    print(f"База данных: {status.get('database_url')}")
    print(f"Текущая ревизия: {current}")
    print(f"Последняя ревизия: {head}")
    print(f"Наличие не применённых миграций: {'да' if pending else 'нет'}")

    return 0


def _cmd_upgrade(revision: str) -> int:
    try:
        apply_migrations(revision)
    except Exception as exc:  # pragma: no cover - CLI вывод
        print(f'Ошибка применения миграций: {exc}', file=sys.stderr)
        return 1

    print(f'Миграции успешно применены до ревизии {revision}.')
    return 0


def _cmd_downgrade(revision: str) -> int:
    try:
        downgrade_migrations(revision)
    except Exception as exc:  # pragma: no cover - CLI вывод
        print(f'Ошибка отката миграций: {exc}', file=sys.stderr)
        return 1

    print(f'Миграции откатены до ревизии {revision}.')
    return 0


def _cmd_history(verbose: bool) -> int:
    try:
        config = get_alembic_config(strict=True)
    except FileNotFoundError as exc:  # pragma: no cover - CLI вывод
        print(str(exc), file=sys.stderr)
        return 1

    command.history(config, verbose=verbose)
    return 0


def main(argv: Iterable[str] | None = None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv or sys.argv[1:])

    if args.command == 'status':
        return _cmd_status()
    if args.command == 'upgrade':
        return _cmd_upgrade(args.revision)
    if args.command == 'downgrade':
        return _cmd_downgrade(args.revision)
    if args.command == 'history':
        return _cmd_history(args.verbose)

    parser.error('Неизвестная команда.')
    return 1


if __name__ == '__main__':  # pragma: no cover - CLI
    raise SystemExit(main())

