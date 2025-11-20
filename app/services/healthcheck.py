from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Dict, List

from sqlalchemy import text

from app.core.extensions import SessionLocal
from app.database import get_migration_status
from app.presentation.helpers import check_templates_exist
from app.services import datasets

PROJECT_ROOT = Path(__file__).resolve().parents[2]
REQUIRED_DIRECTORIES = (
    PROJECT_ROOT / 'logs',
    PROJECT_ROOT / 'templates_docs',
)


def _build_result(ok: bool, details: Dict[str, Any] | None = None) -> Dict[str, Any]:
    payload = {'ok': bool(ok)}
    if details:
        payload['details'] = details
    return payload


def _check_database() -> Dict[str, Any]:
    status = get_migration_status()
    result_ok = status.get('configured') and not status.get('pending', False)
    details: Dict[str, Any] = {
        'configured': status.get('configured', False),
        'pending_migrations': bool(status.get('pending')),
        'current_revision': status.get('current_revision'),
        'head_revision': status.get('head_revision'),
        'database_url': status.get('database_url'),
    }

    connection_error = None
    session = None
    try:
        session = SessionLocal()
        session.execute(text('SELECT 1'))
    except Exception as exc:  # pragma: no cover - логирование в вызывающем коде
        connection_error = str(exc)
        result_ok = False
    finally:
        if session is not None:
            session.close()

    if connection_error:
        details['connection_error'] = connection_error

    return _build_result(result_ok, details)


def _check_templates() -> Dict[str, Any]:
    errors = check_templates_exist()
    if errors:
        return _build_result(False, {'missing': errors})
    return _build_result(True)


def _check_directories() -> Dict[str, Any]:
    missing: List[str] = []
    for path in REQUIRED_DIRECTORIES:
        if not path.exists() or not path.is_dir():
            missing.append(path.as_posix())
    return _build_result(not missing, {'missing': missing} if missing else None)


def _check_dataset(label: str, provider: Callable[[], List[Dict[str, Any]]]) -> Dict[str, Any]:
    try:
        data = provider() or []
        return {
            'label': label,
            'ok': True,
            'items': len(data),
        }
    except Exception as exc:  # pragma: no cover - защитное ветвление
        return {
            'label': label,
            'ok': False,
            'error': str(exc),
        }


def _check_datasets() -> Dict[str, Any]:
    providers = {
        'duty_catalog': lambda: datasets.get_duty_catalog(),
        'gb_materials': lambda: datasets.get_gb_materials(),
        'logistics_cities': lambda: datasets.load_logistics_cities(),
        'orders': lambda: datasets.get_orders_documents(),
        'templates': lambda: datasets.get_task_templates(),
        'instructions': lambda: datasets.get_task_instructions(),
    }

    checks: List[Dict[str, Any]] = []
    overall_ok = True
    for label, provider in providers.items():
        result = _check_dataset(label, provider)
        checks.append(result)
        if not result.get('ok'):
            overall_ok = False

    return {
        'ok': overall_ok,
        'checks': checks,
    }


def run_health_checks() -> Dict[str, Any]:
    checks = {
        'database': _check_database(),
        'templates': _check_templates(),
        'directories': _check_directories(),
        'datasets': _check_datasets(),
    }

    failing = [name for name, payload in checks.items() if not payload.get('ok')]
    overall_status = 'ok' if not failing else 'error'

    return {
        'status': overall_status,
        'timestamp': datetime.now(timezone.utc).isoformat(),
        'failing_checks': failing,
        'checks': checks,
    }

