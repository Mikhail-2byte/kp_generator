"""
Демонстрационные тесты для проверки обновленных RF тарифов.
"""

import sys
from pathlib import Path

# Добавляем корень проекта в путь для возможности прямого запуска
project_root = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(project_root))

import pytest

from app.services.logistics_calculator import (
    get_rf_tariff_per_1000km,
    calculate_ekb_plus_rf_route,
    RF_TARIFF_TABLE
)


def test_all_rf_tariffs():
    """Проверка всех тарифов из таблицы RF_TARIFF_TABLE."""
    expected_tariffs = {
        1000: 30000,    # 1 тонна
        2000: 52000,    # 2 тонны
        3000: 69000,    # 3 тонны
        5000: 90000,    # 5 тонн
        7000: 105000,   # 7 тонн
        10000: 115000,  # 10 тонн
        15000: 118000,  # 15 тонн
        20000: 120000   # 20 тонн (полная фура)
    }
    
    for weight_kg, expected_tariff in expected_tariffs.items():
        actual_tariff = get_rf_tariff_per_1000km(weight_kg)
        assert actual_tariff == expected_tariff, \
            f"Тариф для {weight_kg} кг: ожидалось {expected_tariff}, получено {actual_tariff}"


def test_example_calculation_3_tons():
    """Пример расчета для 3 тонн на расстояние 1000 км."""
    result = calculate_ekb_plus_rf_route(
        weight_kg=3000,
        distance_from_ekb_km=1000
    )
    
    # КНР → ЕКБ: (1,100,000 / 20,000) × 3,000 = 165,000 руб
    china_to_ekb = result['route_details']['china_to_ekb']['price']
    assert china_to_ekb == 165000
    
    # ЕКБ → Город: 69,000 руб (тариф для 3 тонн за 1000 км)
    ekb_to_city = result['route_details']['ekb_to_destination']['price']
    assert ekb_to_city == 69000
    
    # Итого: 165,000 + 69,000 = 234,000 руб
    total = result['total_price']
    assert total == 234000


def test_example_calculation_7_tons():
    """Пример расчета для 7 тонн на расстояние 500 км."""
    result = calculate_ekb_plus_rf_route(
        weight_kg=7000,
        distance_from_ekb_km=500
    )
    
    # КНР → ЕКБ: (1,100,000 / 20,000) × 7,000 = 385,000 руб
    china_to_ekb = result['route_details']['china_to_ekb']['price']
    assert china_to_ekb == 385000
    
    # ЕКБ → Город: (105,000 / 1000) × 500 = 52,500 руб
    ekb_to_city = result['route_details']['ekb_to_destination']['price']
    assert ekb_to_city == 52500
    
    # Итого: 385,000 + 52,500 = 437,500 руб
    total = result['total_price']
    assert total == 437500


def test_tariff_table_integrity():
    """Проверка, что таблица тарифов содержит все необходимые значения."""
    assert len(RF_TARIFF_TABLE) == 8, "Таблица должна содержать 8 тарифов"
    
    # Проверяем что тарифы растут с увеличением веса
    sorted_weights = sorted(RF_TARIFF_TABLE.keys())
    for i in range(len(sorted_weights) - 1):
        current_weight = sorted_weights[i]
        next_weight = sorted_weights[i + 1]
        
        current_tariff = RF_TARIFF_TABLE[current_weight]
        next_tariff = RF_TARIFF_TABLE[next_weight]
        
        assert next_tariff > current_tariff, \
            f"Тариф для {next_weight} кг должен быть больше чем для {current_weight} кг"


if __name__ == '__main__':
    pytest.main([__file__, '-v'])

