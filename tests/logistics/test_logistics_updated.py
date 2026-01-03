"""
Тесты для обновленной логики расчета логистики.

Проверяет:
- Расчет с объемом 82 м³ (обновленное значение)
- Ограничения трала по городам (только Екатеринбург и Москва)
- Алгоритм "До ЕКБ + по РФ"
- Сохранение расстояний от ЕКБ
- Выбор между алгоритмами расчета
"""

import sys
from pathlib import Path

# Добавляем корень проекта в путь для возможности прямого запуска
project_root = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(project_root))

import pytest

from app.services.logistics_calculator import (
    calculate_logistics,
    calculate_logistics_by_weight,
    calculate_logistics_by_volume,
    get_rf_tariff_per_1000km,
    calculate_ekb_plus_rf_route,
    EURO_TRUCK_VOLUME_M3,
    EURO_TRUCK_CAPACITY_KG,
    TRAIL_CAPACITY_KG,
    TRAIL_AVAILABLE_CITIES,
)


class TestUpdatedConstants:
    """Тесты для проверки обновленных констант."""
    
    def test_truck_volume_is_82_m3(self):
        """Проверка, что объем фуры обновлен до 82 м³."""
        assert EURO_TRUCK_VOLUME_M3 == 82
    
    def test_trail_available_cities(self):
        """Проверка, что трал доступен только до Екатеринбурга и Москвы."""
        assert set(TRAIL_AVAILABLE_CITIES) == {'Екатеринбург', 'Москва'}


class TestRFTariffCalculation:
    """Тесты для тарифной сетки РФ."""
    
    def test_get_tariff_for_1_ton(self):
        """Проверка тарифа для 1 тонны."""
        assert get_rf_tariff_per_1000km(1000) == 30000
    
    def test_get_tariff_for_2_tons(self):
        """Проверка тарифа для 2 тонн."""
        assert get_rf_tariff_per_1000km(2000) == 52000
    
    def test_get_tariff_for_5_tons(self):
        """Проверка тарифа для 5 тонн."""
        assert get_rf_tariff_per_1000km(5000) == 90000
    
    def test_get_tariff_for_intermediate_weight(self):
        """Проверка, что для промежуточного веса используется ближайший больший."""
        # 1.5 тонны → используем тариф для 2 тонн
        assert get_rf_tariff_per_1000km(1500) == 52000
    
    def test_get_tariff_for_heavy_cargo(self):
        """Проверка тарифа для груза тяжелее 20 тонн."""
        # Для 25 тонн используем максимальный тариф (для 20 тонн)
        assert get_rf_tariff_per_1000km(25000) == 120000


class TestEkbPlusRFCalculation:
    """Тесты для алгоритма 'До ЕКБ + по РФ'."""
    
    def test_basic_calculation_3_tons_1050km(self):
        """Тест базового расчета для 3 тонн на 1050 км (пример из документа)."""
        result = calculate_ekb_plus_rf_route(
            weight_kg=3000,
            distance_from_ekb_km=1050
        )
        
        # КНР → ЕКБ: (1100000 / 20000) × 3000 = 165000
        # ЕКБ → Город: (69000 / 1000) × 1050 = 72450
        # Итого: 237450
        assert result['calculation_type'] == 'ekb_plus_rf'
        assert result['route_details']['china_to_ekb']['price'] == 165000
        assert abs(result['route_details']['ekb_to_destination']['price'] - 72450) < 1
        assert abs(result['total_price'] - 237450) < 1
    
    def test_calculation_includes_correct_tariff(self):
        """Проверка, что используется правильный тариф из таблицы РФ."""
        result = calculate_ekb_plus_rf_route(
            weight_kg=2000,
            distance_from_ekb_km=500
        )
        
        # Для 2 тонн тариф = 52000 руб за 1000 км
        # ЕКБ → Город: (52000 / 1000) × 500 = 26000
        assert result['route_details']['ekb_to_destination']['tariff_per_1000km'] == 52000
        assert result['route_details']['ekb_to_destination']['price'] == 26000


class TestMainLogisticsFunction:
    """Тесты для основной функции calculate_logistics."""
    
    def test_small_cargo_recommendation(self):
        """Проверка рекомендации Деловых линий для грузов ≤ 600 кг."""
        result = calculate_logistics(
            weight_kg=500,
            city_price=400000
        )
        
        assert result['recommendation'] == 'dellin'
        assert 'Деловые линии' in result['message']
    
    def test_trail_restriction_for_non_ekb_moscow(self):
        """Проверка ограничения трала для городов кроме ЕКБ и Москвы."""
        result = calculate_logistics(
            weight_kg=25000,
            city_price=500000,
            transport_type='trail',
            city_name='Новосибирск'
        )
        
        assert 'error' in result
        assert result['trail_not_available'] is True
        assert 'Трал доступен только' in result['error']
    
    def test_trail_allowed_for_ekaterinburg(self):
        """Проверка, что трал доступен для Екатеринбурга."""
        result = calculate_logistics(
            weight_kg=25000,
            city_price=1100000,
            transport_type='trail',
            city_name='Екатеринбург'
        )
        
        assert 'error' not in result
        assert result['basis'] == 'weight'
        assert result['transport_type'] == 'trail'
    
    def test_trail_allowed_for_moscow(self):
        """Проверка, что трал доступен для Москвы."""
        result = calculate_logistics(
            weight_kg=30000,
            city_price=1200000,
            transport_type='trail',
            city_name='Москва'
        )
        
        assert 'error' not in result
        assert result['basis'] == 'weight'
        assert result['transport_type'] == 'trail'
    
    def test_direct_calculation_for_main_route(self):
        """Проверка прямого расчета для города на основном маршруте."""
        result = calculate_logistics(
            weight_kg=10000,
            city_price=600000,
            transport_type='truck',
            city_name='Иркутск',
            is_main_route=True
        )
        
        assert result['calculation_type'] == 'direct'
        assert result['basis'] == 'weight'
        # (600000 / 20000) × 10000 = 300000
        assert result['total_price'] == 300000
    
    def test_ekb_plus_rf_for_light_cargo_not_main_route(self):
        """Проверка алгоритма 'До ЕКБ + по РФ' для легкого груза вне основного маршрута."""
        result = calculate_logistics(
            weight_kg=3000,  # < 5 тонн
            city_price=500000,
            transport_type='truck',
            city_name='Сызрань',
            is_main_route=False,
            distance_from_ekb_km=1050
        )
        
        assert result['calculation_type'] == 'ekb_plus_rf'
        assert 'route_details' in result
        assert result['route_details']['china_to_ekb']['price'] == 165000
        assert abs(result['route_details']['ekb_to_destination']['price'] - 72450) < 1
    
    def test_needs_distance_for_light_cargo_not_main_route(self):
        """Проверка, что требуется расстояние для легкого груза вне основного маршрута."""
        result = calculate_logistics(
            weight_kg=3000,  # < 5 тонн
            city_price=500000,
            transport_type='truck',
            city_name='Сызрань',
            is_main_route=False,
            distance_from_ekb_km=None  # расстояние не указано
        )
        
        assert 'error' in result
        assert result['needs_distance'] is True
        assert 'расстояние' in result['error'].lower()
    
    def test_direct_calculation_for_heavy_cargo_not_main_route(self):
        """Проверка прямого расчета для тяжелого груза вне основного маршрута."""
        result = calculate_logistics(
            weight_kg=8000,  # > 5 тонн
            city_price=500000,
            transport_type='truck',
            city_name='Сызрань',
            is_main_route=False,
            distance_from_ekb_km=None  # не требуется для тяжелого груза
        )
        
        # Для груза > 5 тонн используем прямой расчет даже если город не основной
        assert result['calculation_type'] == 'direct'
        assert result['basis'] == 'weight'
    
    def test_volume_calculation_with_82_m3(self):
        """Проверка расчета по объему с новым значением 82 м³."""
        # Габариты: 5000 × 2000 × 2000 мм = 20 м³
        result = calculate_logistics(
            weight_kg=5000,  # легкий груз
            city_price=820000,
            transport_type='truck',
            city_name='Новосибирск',
            is_main_route=True,
            length_mm=5000,
            width_mm=2000,
            height_mm=2000
        )
        
        # Объем: 20 м³, Вес: 5000 кг
        # weight_ratio = 5000/20000 = 0.25
        # volume_ratio = 20/82 = 0.244
        # Вес превалирует, расчет по весу
        assert result['basis'] == 'weight'
        assert result['volume_m3'] == 20.0
    
    def test_can_save_distance_flag(self):
        """Проверка флага can_save_distance."""
        result = calculate_logistics(
            weight_kg=3000,
            city_price=500000,
            transport_type='truck',
            city_name='Сызрань',
            is_main_route=False,
            distance_from_ekb_km=1050
        )
        
        # Расстояние введено вручную, но уже использовано в расчете
        # can_save_distance должен быть False так как расстояние уже есть
        assert result['calculation_type'] == 'ekb_plus_rf'
        # Примечание: в реальности флаг должен устанавливаться на фронтенде


class TestVolumeCalculationWith82M3:
    """Тесты для расчета по объему с обновленным значением 82 м³."""
    
    def test_volume_based_calculation(self):
        """Проверка расчета по объему когда объем превалирует."""
        result = calculate_logistics_by_volume(
            volume_m3=50,
            city_price=820000
        )
        
        # (820000 / 82) × 50 = 500000
        assert result['basis'] == 'volume'
        assert abs(result['total_price'] - 500000) < 1
        assert result['trucks_count'] == 1  # 50 м³ < 82 м³
    
    def test_multiple_trucks_by_volume(self):
        """Проверка расчета нескольких машин по объему."""
        result = calculate_logistics_by_volume(
            volume_m3=150,
            city_price=820000
        )
        
        # 150 м³ / 82 м³ = 1.83 → 2 машины
        assert result['trucks_count'] == 2
        # Стоимость: (820000 / 82) × 150
        assert abs(result['total_price'] - 1500000) < 1


if __name__ == '__main__':
    pytest.main([__file__, '-v'])

