from app import create_app


def test_health_smoke():
    """
    Дымовой тест: реальный вызов /health без мока run_health_checks.
    Проверяет, что приложение поднимается и зависимости доступны.
    """
    app = create_app()
    app.config.update(TESTING=True, WTF_CSRF_ENABLED=False)

    with app.test_client() as client:
        response = client.get("/health")
        assert response.status_code == 200
        payload = response.get_json()
        assert payload.get("status") == "ok"
        assert "checks" in payload

