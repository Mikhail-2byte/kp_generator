"""
Тесты для проверки правильности извлечения и обработки веса в расчете логистики.

Проверяет:
- Правильность извлечения веса из разных форматов
- Корректность передачи weight_kg в функции расчета
- Правильность отображения веса в результатах
- Разные варианты веса (3000кг, 4 500кг, 5 тонн и т.д.)
"""

import sys
from pathlib import Path

# Добавляем корень проекта в путь
BASE_DIR = Path(__file__).resolve().parent.parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

import pytest

from app.services.logistics_calculator import (
    EURO_TRUCK_CAPACITY_KG,
    calculate_logistics,
    calculate_logistics_by_weight,
)


class TestWeightExtraction:
    """Тесты для извлечения веса из сообщений."""
    
    def test_extract_weight_3000kg(self):
        """Проверка извлечения веса 3000кг."""
        from ai_agent.agent import AIAgent
        agent = AIAgent()
        message = "рассчитай логистику из КНР в ЕКБ за габаритный груз весом 3000кг."
        intent = agent._extract_intent(message)
        
        assert intent['is_logistics'] is True
        assert intent['logistics_params'].get('weight_kg') == 3000.0
        assert intent['logistics_params'].get('city_name') == 'Екатеринбург'
    
    def test_extract_weight_4500kg_with_space(self):
        """Проверка извлечения веса 4 500кг (с пробелом)."""
        from ai_agent.agent import AIAgent
        agent = AIAgent()
        message = "Посчитай логистику от КНР до ЕКб за груз весом 4 500кг."
        intent = agent._extract_intent(message)
        
        assert intent['is_logistics'] is True
        assert intent['logistics_params'].get('weight_kg') == 4500.0
    
    def test_extract_weight_5000kg(self):
        """Проверка извлечения веса 5000кг."""
        from ai_agent.agent import AIAgent
        agent = AIAgent()
        message = "Рассчитай стоимость логистики для габаритного груза от КНР до ЕКБ за груз весом 5000кг."
        intent = agent._extract_intent(message)
        
        assert intent['is_logistics'] is True
        assert intent['logistics_params'].get('weight_kg') == 5000.0
    
    def test_extract_weight_5_tons(self):
        """Проверка извлечения веса 5 тонн."""
        from ai_agent.agent import AIAgent
        agent = AIAgent()
        message = "Логистика для груза 5 тонн"
        intent = agent._extract_intent(message)
        
        assert intent['is_logistics'] is True
        assert intent['logistics_params'].get('weight_kg') == 5000.0
    
    def test_extract_weight_5t(self):
        """Проверка извлечения веса 5т."""
        from ai_agent.agent import AIAgent
        agent = AIAgent()
        message = "Логистика для груза 5т"
        intent = agent._extract_intent(message)
        
        assert intent['is_logistics'] is True
        assert intent['logistics_params'].get('weight_kg') == 5000.0
    
    def test_extract_weight_with_thousands_separator(self):
        """Проверка извлечения веса с разделителем тысяч (5,000 кг)."""
        from ai_agent.agent import AIAgent
        agent = AIAgent()
        message = "Рассчитай логистику для груза весом 5,000 кг"
        intent = agent._extract_intent(message)
        
        assert intent['is_logistics'] is True
        assert intent['logistics_params'].get('weight_kg') == 5000.0
    
    def test_extract_weight_1_5_tons(self):
        """Проверка извлечения веса 1.5 тонн."""
        from ai_agent.agent import AIAgent
        agent = AIAgent()
        message = "Рассчитай логистику для груза весом 1.5 тонн"
        intent = agent._extract_intent(message)
        
        assert intent['is_logistics'] is True
        assert intent['logistics_params'].get('weight_kg') == 1500.0


class TestWeightInCalculation:
    """Тесты для проверки правильности веса в расчетах."""
    
    def test_weight_3000kg_in_calculation(self):
        """Проверка, что вес 3000кг правильно передается в расчет."""
        city_price = 1100000  # Цена до Екатеринбурга
        weight_kg = 3000
        
        result = calculate_logistics_by_weight(
            weight_kg=weight_kg,
            city_price=city_price,
            truck_capacity_kg=EURO_TRUCK_CAPACITY_KG
        )
        
        # Проверяем, что weight_kg правильный
        assert result['weight_kg'] == 3000
        assert result['weight_kg'] != 3000000  # Не должен быть умножен на 1000
        
        # Проверяем формулу
        assert '× 3,000 кг' in result['calculation_formula']
        assert '× 3,000,000 кг' not in result['calculation_formula']
        
        # Проверяем расчет
        expected_price_per_kg = city_price / EURO_TRUCK_CAPACITY_KG  # 55.0
        expected_total = expected_price_per_kg * weight_kg  # 165000
        assert abs(result['total_price'] - expected_total) < 1
        
        # Проверяем количество машин
        expected_trucks = 1  # 3000 кг < 20000 кг
        assert result['trucks_count'] == expected_trucks
    
    def test_weight_4500kg_in_calculation(self):
        """Проверка, что вес 4500кг правильно передается в расчет."""
        city_price = 1100000
        weight_kg = 4500
        
        result = calculate_logistics_by_weight(
            weight_kg=weight_kg,
            city_price=city_price
        )
        
        assert result['weight_kg'] == 4500
        assert '× 4,500 кг' in result['calculation_formula']
        
        # Расчет: (1100000 / 20000) × 4500 = 247500
        expected_total = (city_price / EURO_TRUCK_CAPACITY_KG) * weight_kg
        assert abs(result['total_price'] - expected_total) < 1
    
    def test_weight_5000kg_in_calculation(self):
        """Проверка, что вес 5000кг правильно передается в расчет."""
        city_price = 1100000
        weight_kg = 5000
        
        result = calculate_logistics_by_weight(
            weight_kg=weight_kg,
            city_price=city_price
        )
        
        assert result['weight_kg'] == 5000
        assert '× 5,000 кг' in result['calculation_formula']
        
        # Расчет: (1100000 / 20000) × 5000 = 275000
        expected_total = (city_price / EURO_TRUCK_CAPACITY_KG) * weight_kg
        assert abs(result['total_price'] - expected_total) < 1
    
    def test_weight_preserved_in_calculate_logistics(self):
        """Проверка, что weight_kg сохраняется в calculate_logistics."""
        weight_kg = 3000
        city_price = 1100000
        
        result = calculate_logistics(
            weight_kg=weight_kg,
            city_price=city_price,
            city_name='Екатеринбург',
            is_main_route=True
        )
        
        # Проверяем, что weight_kg правильный
        assert result['weight_kg'] == 3000
        assert result['weight_kg'] != 3000000
        
        # Проверяем формулу
        if 'calculation_formula' in result:
            assert '× 3,000 кг' in result['calculation_formula']
            assert '× 3,000,000 кг' not in result['calculation_formula']


class TestSmallCargo:
    """Тесты для мелкогабаритных грузов."""
    
    def test_small_cargo_450kg(self):
        """Проверка обработки мелкогабаритного груза 450кг."""
        weight_kg = 450
        
        result = calculate_logistics(
            weight_kg=weight_kg,
            city_price=1100000,
            city_name='Екатеринбург'
        )
        
        assert result['basis'] == 'small_cargo'
        assert result['weight_kg'] == 450
        assert result['recommendation'] == 'dellin'
        assert 'Деловые линии' in result['message']
    
    def test_small_cargo_600kg(self):
        """Проверка обработки мелкогабаритного груза 600кг (граничное значение)."""
        weight_kg = 600
        
        result = calculate_logistics(
            weight_kg=weight_kg,
            city_price=1100000,
            city_name='Екатеринбург'
        )
        
        assert result['basis'] == 'small_cargo'
        assert result['weight_kg'] == 600
    
    def test_not_small_cargo_601kg(self):
        """Проверка, что груз 601кг не считается мелкогабаритным."""
        weight_kg = 601
        
        result = calculate_logistics(
            weight_kg=weight_kg,
            city_price=1100000,
            city_name='Екатеринбург',
            is_main_route=True
        )
        
        assert result['basis'] != 'small_cargo'
        assert result['weight_kg'] == 601


class TestHeavyCargo:
    """Тесты для тяжелых грузов."""
    
    def test_heavy_cargo_25000kg(self):
        """Проверка расчета для тяжелого груза 25 тонн."""
        weight_kg = 25000
        city_price = 1100000
        
        result = calculate_logistics_by_weight(
            weight_kg=weight_kg,
            city_price=city_price
        )
        
        assert result['weight_kg'] == 25000
        assert '× 25,000 кг' in result['calculation_formula']
        
        # Расчет: (1100000 / 20000) × 25000 = 1375000
        expected_total = (city_price / EURO_TRUCK_CAPACITY_KG) * weight_kg
        assert abs(result['total_price'] - expected_total) < 1
        
        # Количество машин: 25000 / 20000 = 2 машины
        assert result['trucks_count'] == 2
    
    def test_very_heavy_cargo_50000kg(self):
        """Проверка расчета для очень тяжелого груза 50 тонн."""
        weight_kg = 50000
        city_price = 1100000
        
        result = calculate_logistics_by_weight(
            weight_kg=weight_kg,
            city_price=city_price
        )
        
        assert result['weight_kg'] == 50000
        assert '× 50,000 кг' in result['calculation_formula']
        
        # Количество машин: 50000 / 20000 = 3 машины
        assert result['trucks_count'] == 3


class TestDifferentCities:
    """Тесты для разных городов."""
    
    def test_ekaterinburg_3000kg(self):
        """Проверка расчета для Екатеринбурга с весом 3000кг."""
        weight_kg = 3000
        city_price = 1100000  # Цена до Екатеринбурга
        
        result = calculate_logistics(
            weight_kg=weight_kg,
            city_price=city_price,
            city_name='Екатеринбург',
            is_main_route=True
        )
        
        assert result['weight_kg'] == 3000
        assert result['calculation_type'] == 'direct'
        assert result['basis'] == 'weight'
    
    def test_moscow_5000kg(self):
        """Проверка расчета для Москвы с весом 5000кг."""
        weight_kg = 5000
        city_price = 1200000  # Примерная цена до Москвы
        
        result = calculate_logistics(
            weight_kg=weight_kg,
            city_price=city_price,
            city_name='Москва',
            is_main_route=True
        )
        
        assert result['weight_kg'] == 5000
        assert result['calculation_type'] == 'direct'


class TestWeightFormatting:
    """Тесты для форматирования веса в результатах."""
    
    def test_weight_formatting_in_formula(self):
        """Проверка форматирования веса в формуле."""
        test_cases = [
            (3000, '× 3,000 кг'),
            (4500, '× 4,500 кг'),
            (5000, '× 5,000 кг'),
            (10000, '× 10,000 кг'),
            (25000, '× 25,000 кг'),
        ]
        
        city_price = 1100000
        
        for weight_kg, expected_in_formula in test_cases:
            result = calculate_logistics_by_weight(
                weight_kg=weight_kg,
                city_price=city_price
            )
            
            assert expected_in_formula in result['calculation_formula']
            # Проверяем, что нет неправильного формата с миллионами
            wrong_format = f'× {weight_kg * 1000:,} кг'
            assert wrong_format not in result['calculation_formula']


if __name__ == '__main__':
    pytest.main([__file__, '-v'])

