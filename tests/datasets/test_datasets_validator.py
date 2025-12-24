from app.services import datasets_validator


def test_run_all_validations_returns_results():
    results = datasets_validator.run_all_validations()
    assert isinstance(results, list)
    assert results, "Ожидается как минимум один результат проверки"

    names = {r.name for r in results}
    # Проверяем наличие ключевых справочников
    assert "duty_catalog_combined" in names  # Объединённый каталог (ручные + ТН ВЭД)
    assert "logistics_cities_combined" in names

    # Статус в допустимом наборе
    for r in results:
        assert r.status in {"ok", "warning", "error"}

