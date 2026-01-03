import sys
from datetime import datetime, timezone
from pathlib import Path

# Добавляем корень проекта в путь для возможности прямого запуска
project_root = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(project_root))

import pytest


def test_health_endpoint_ok(monkeypatch, client):
    payload = {
        'status': 'ok',
        'timestamp': datetime.now(timezone.utc).isoformat(),
        'failing_checks': [],
        'checks': {},
    }
    monkeypatch.setattr('app.routes.health.run_health_checks', lambda: payload)

    response = client.get('/health')

    assert response.status_code == 200
    assert response.get_json() == payload


def test_health_endpoint_error(monkeypatch, client):
    payload = {
        'status': 'error',
        'timestamp': datetime.now(timezone.utc).isoformat(),
        'failing_checks': ['database'],
        'checks': {'database': {'ok': False}},
    }
    monkeypatch.setattr('app.routes.health.run_health_checks', lambda: payload)

    response = client.get('/healthz')

    assert response.status_code == 503
    assert response.get_json() == payload


if __name__ == "__main__":
    # Запуск тестов при прямом выполнении файла
    pytest.main([__file__, "-v"])

