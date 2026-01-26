import math
import sys
from pathlib import Path

# Добавляем корень проекта в путь для возможности прямого запуска
project_root = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(project_root))

import pytest

from app.services.multi_position_calculator import MultiPositionCalculator


def _sample_positions():
    return [
        {
            'product': 'Деталь А',
            'quantity': '10',
            'cost_price': '1200',
            'weight': '3.5',
            'duty_percent': '5',
        },
        {
            'product': 'Деталь B',
            'quantity': '5',
            'cost_price': '800',
            'weight': '2',
            'duty_percent': '10',
        },
    ]


def test_calculate_positions_global_margin():
    calculator = MultiPositionCalculator({
        'pricing': {'mode': 'global'},
        'calculation_constants': {
            'conversion_rate': 10,
            'logistics_cnr_ratio': 0.4,
            'logistics_rf_ratio': 0.6,
            'conversion_fee_rate': 0.02,
            'credit_rate': 0.12,
        }
    })

    result = calculator.calculate_positions(_sample_positions(), 20000, 30, 20)

    assert result['positions'], 'Должны быть рассчитаны позиции'
    assert pytest.approx(result['actual_margin'], rel=1e-3) == 20
    assert result['total_revenue'] > result['total_costs']
    assert result['price_coefficient'] is not None


def test_calculate_positions_per_position_mode():
    calculator = MultiPositionCalculator({'pricing': {'mode': 'per_position'}})
    result = calculator.calculate_positions(_sample_positions(), 10000, 25, 15)

    assert result['price_coefficient'] is None
    for position in result['positions']:
        assert pytest.approx(position['margin'], rel=1e-3) == 15


def test_calculate_legacy_single_position():
    calculator = MultiPositionCalculator()
    position = {
        'product': 'Ось',
        'quantity': '4',
        'cost_price': '1000',
        'weight': '5',
        'duty_percent': '0',
    }
    result = calculator.calculate_legacy_single_position(position, 5000, 20, 25)
    assert math.isclose(result['general_price'], result['final_price'] * 4, rel_tol=1e-6)
    # Проверяем, что фактическая маржа близка к целевой (допускаем небольшую погрешность)
    assert pytest.approx(result['margin'], abs=0.5) == 25


def test_calculate_position_costs():
    """Тест расчета затрат для позиции."""
    calculator = MultiPositionCalculator({
        'calculation_constants': {
            'conversion_rate': 12,
            'logistics_cnr_ratio': 0.3,
            'logistics_rf_ratio': 0.7,
            'conversion_fee_rate': 0.032,
            'credit_rate': 0.16,
        }
    })
    
    position = {
        'quantity': 10,
        'cost_price': 1000,
        'weight': 5,
        'duty_percent': 5
    }
    
    costs = calculator.calculate_position_costs(position, 50000, 30, 50)
    
    assert 'cost_per_unit' in costs
    assert 'total_cost' in costs
    assert 'logistics_cnr_per_unit' in costs
    assert 'logistics_rf_per_unit' in costs
    assert 'duty_per_unit' in costs
    assert costs['total_cost'] == costs['cost_per_unit'] * position['quantity']


def test_calculate_multi_position_empty_list():
    """Тест расчета для пустого списка позиций."""
    calculator = MultiPositionCalculator()
    
    result = calculator.calculate_multi_position_prices([], 50000, 30, 30)
    
    assert result['positions'] == []
    assert result['total_costs'] == 0
    assert result['total_revenue'] == 0
    assert result['actual_margin'] == 0


def test_calculate_multi_position_zero_weight():
    """Тест расчета с нулевым весом."""
    calculator = MultiPositionCalculator()
    
    positions = [{
        'quantity': 10,
        'cost_price': 1000,
        'weight': 0,
        'duty_percent': 5
    }]
    
    # Должен обработать нулевой вес без ошибок
    result = calculator.calculate_multi_position_prices(positions, 50000, 30, 30)
    
    assert len(result['positions']) == 1


def test_additional_expenses_global_mode():
    """Тест расчета маржи с дополнительными расходами в режиме global."""
    calculator = MultiPositionCalculator({
        'pricing': {'mode': 'global'},
        'calculation_constants': {
            'conversion_rate': 10,
            'logistics_cnr_ratio': 0.4,
            'logistics_rf_ratio': 0.6,
            'conversion_fee_rate': 0.02,
            'credit_rate': 0.12,
        }
    })
    
    positions = _sample_positions()
    target_margin = 20
    additional_expenses_yuan = 500  # 500 юаней дополнительных расходов
    
    # Расчет без дополнительных расходов
    result_without = calculator.calculate_multi_position_prices(
        positions, 20000, 30, target_margin
    )
    
    # Расчет с дополнительными расходами
    result_with = calculator.calculate_multi_position_prices(
        positions, 20000, 30, target_margin,
        additional_expenses_total_yuan=additional_expenses_yuan
    )
    
    # Проверяем, что маржа остается целевой
    assert pytest.approx(result_with['actual_margin'], rel=1e-2) == target_margin
    
    # Проверяем, что затраты увеличились на дополнительные расходы (в юанях)
    expected_cost_increase = additional_expenses_yuan  # total_costs теперь в юанях
    assert pytest.approx(
        result_with['total_costs'] - result_without['total_costs'],
        rel=1e-2
    ) == expected_cost_increase
    
    # Проверяем, что выручка тоже увеличилась для сохранения маржи
    assert result_with['total_revenue'] > result_without['total_revenue']
    
    # Проверяем, что дополнительные расходы не добавляются дважды
    # Разница в затратах должна быть равна дополнительным расходам (в юанях)
    cost_difference = result_with['total_costs'] - result_without['total_costs']
    assert pytest.approx(cost_difference, rel=1e-2) == expected_cost_increase


def test_additional_expenses_per_position_mode():
    """Тест расчета маржи с дополнительными расходами в режиме per_position."""
    calculator = MultiPositionCalculator({
        'pricing': {'mode': 'per_position'},
        'calculation_constants': {
            'conversion_rate': 10,
            'logistics_cnr_ratio': 0.4,
            'logistics_rf_ratio': 0.6,
            'conversion_fee_rate': 0.02,
            'credit_rate': 0.12,
        }
    })
    
    positions = _sample_positions()
    target_margin = 15
    additional_expenses_yuan = 300  # 300 юаней дополнительных расходов
    
    # Расчет с дополнительными расходами
    result = calculator.calculate_multi_position_prices(
        positions, 10000, 25, target_margin,
        additional_expenses_total_yuan=additional_expenses_yuan
    )
    
    # Проверяем, что маржа каждой позиции остается целевой
    for position in result['positions']:
        assert pytest.approx(position['margin'], rel=1e-2) == target_margin
    
    # Проверяем, что итоговая маржа равна целевой
    assert pytest.approx(result['actual_margin'], rel=1e-2) == target_margin


def test_additional_expenses_no_double_counting():
    """Тест проверки отсутствия двойного добавления дополнительных расходов."""
    calculator = MultiPositionCalculator({
        'pricing': {'mode': 'global'},
        'calculation_constants': {
            'conversion_rate': 10,
            'logistics_cnr_ratio': 0.4,
            'logistics_rf_ratio': 0.6,
            'conversion_fee_rate': 0.02,
            'credit_rate': 0.12,
        }
    })
    
    positions = [{
        'quantity': '10',
        'cost_price': '1000',
        'weight': '5',
        'duty_percent': '5',
    }]
    
    target_margin = 20
    additional_expenses_yuan = 200
    
    # Расчет без дополнительных расходов
    result_without = calculator.calculate_multi_position_prices(
        positions, 10000, 30, target_margin
    )
    
    # Расчет с дополнительными расходами
    result_with = calculator.calculate_multi_position_prices(
        positions, 10000, 30, target_margin,
        additional_expenses_total_yuan=additional_expenses_yuan
    )
    
    # Проверяем, что дополнительные расходы добавляются только один раз
    expected_cost_increase = additional_expenses_yuan  # total_costs теперь в юанях
    actual_cost_increase = result_with['total_costs'] - result_without['total_costs']
    
    # Разница должна быть равна дополнительным расходам (не двойному значению)
    assert pytest.approx(actual_cost_increase, rel=1e-2) == expected_cost_increase
    
    # Проверяем, что маржа остается целевой
    assert pytest.approx(result_with['actual_margin'], rel=1e-2) == target_margin


def test_additional_expenses_margin_consistency():
    """Тест проверки корректности фактической маржи (должна быть равна целевой)."""
    calculator = MultiPositionCalculator({
        'pricing': {'mode': 'global'},
        'calculation_constants': {
            'conversion_rate': 11.5,
            'logistics_cnr_ratio': 0.3,
            'logistics_rf_ratio': 0.7,
            'conversion_fee_rate': 0.032,
            'credit_rate': 0.16,
        }
    })
    
    positions = _sample_positions()
    
    # Тестируем разные значения маржи и дополнительных расходов
    test_cases = [
        (20, 0),      # Без дополнительных расходов
        (20, 500),    # С дополнительными расходами
        (30, 1000),   # Другая маржа и расходы
        (15, 200),    # Меньшая маржа
    ]
    
    for target_margin, additional_expenses_yuan in test_cases:
        result = calculator.calculate_multi_position_prices(
            positions, 20000, 30, target_margin,
            additional_expenses_total_yuan=additional_expenses_yuan
        )
        
        # Проверяем, что фактическая маржа равна целевой (с небольшой погрешностью)
        assert pytest.approx(result['actual_margin'], abs=0.5) == target_margin, \
            f"Маржа {result['actual_margin']:.2f}% не равна целевой {target_margin}% " \
            f"при дополнительных расходах {additional_expenses_yuan} юаней"


def test_position_margins_single_position():
    """Тест с одной позицией с индивидуальной маржой."""
    calculator = MultiPositionCalculator({
        'pricing': {'mode': 'global'},
        'calculation_constants': {
            'conversion_rate': 11.5,
            'logistics_cnr_ratio': 0.3,
            'logistics_rf_ratio': 0.7,
            'conversion_fee_rate': 0.032,
            'credit_rate': 0.16,
        }
    })
    
    positions = [{
        'quantity': '10',
        'cost_price': '1000',
        'weight': '5',
        'duty_percent': '5',
    }]
    
    target_margin = 30.0
    position_margins = {0: 25.0}  # Индивидуальная маржа 25% для первой позиции
    
    result = calculator.calculate_multi_position_prices(
        positions, 50000, 30, target_margin,
        position_margins=position_margins
    )
    
    # Проверяем, что позиция имеет индивидуальную маржу
    assert len(result['positions']) == 1
    assert pytest.approx(result['positions'][0]['margin'], abs=0.5) == 25.0
    
    # Когда все позиции имеют индивидуальную маржу, общая маржа равна марже этой позиции
    # (так как только одна позиция)
    assert pytest.approx(result['actual_margin'], abs=0.5) == 25.0


def test_position_margins_multiple_positions():
    """Тест с несколькими позициями, где часть имеет индивидуальную маржу."""
    calculator = MultiPositionCalculator({
        'pricing': {'mode': 'global'},
        'calculation_constants': {
            'conversion_rate': 11.5,
            'logistics_cnr_ratio': 0.3,
            'logistics_rf_ratio': 0.7,
            'conversion_fee_rate': 0.032,
            'credit_rate': 0.16,
        }
    })
    
    positions = _sample_positions()
    target_margin = 30.0
    # Первая позиция с индивидуальной маржой 25%, вторая использует общую маржу
    position_margins = {0: 25.0}
    
    result = calculator.calculate_multi_position_prices(
        positions, 20000, 30, target_margin,
        position_margins=position_margins
    )
    
    # Проверяем, что первая позиция имеет индивидуальную маржу
    assert pytest.approx(result['positions'][0]['margin'], abs=0.5) == 25.0
    
    # Проверяем, что общая маржа соответствует целевой
    assert pytest.approx(result['actual_margin'], abs=0.5) == target_margin
    
    # Проверяем, что вторая позиция пересчитана для достижения общей маржи
    assert result['positions'][1]['margin'] != 25.0  # Должна быть другая маржа


def test_position_margins_all_fixed():
    """Тест когда все позиции имеют индивидуальную маржу."""
    calculator = MultiPositionCalculator({
        'pricing': {'mode': 'global'},
        'calculation_constants': {
            'conversion_rate': 11.5,
            'logistics_cnr_ratio': 0.3,
            'logistics_rf_ratio': 0.7,
            'conversion_fee_rate': 0.032,
            'credit_rate': 0.16,
        }
    })
    
    positions = _sample_positions()
    target_margin = 30.0  # Целевая маржа (может не совпадать с фактической)
    # Все позиции с индивидуальными маржами
    position_margins = {0: 25.0, 1: 35.0}
    
    result = calculator.calculate_multi_position_prices(
        positions, 20000, 30, target_margin,
        position_margins=position_margins
    )
    
    # Проверяем, что позиции имеют свои индивидуальные маржи
    assert pytest.approx(result['positions'][0]['margin'], abs=0.5) == 25.0
    assert pytest.approx(result['positions'][1]['margin'], abs=0.5) == 35.0
    
    # Фактическая маржа может отличаться от целевой, так как все позиции фиксированы
    assert result['actual_margin'] > 0


def test_position_margins_per_position_mode():
    """Тест индивидуальных марж в режиме per_position."""
    calculator = MultiPositionCalculator({
        'pricing': {'mode': 'per_position'},
        'calculation_constants': {
            'conversion_rate': 11.5,
            'logistics_cnr_ratio': 0.3,
            'logistics_rf_ratio': 0.7,
            'conversion_fee_rate': 0.032,
            'credit_rate': 0.16,
        }
    })
    
    positions = _sample_positions()
    target_margin = 30.0
    # Первая позиция с индивидуальной маржой 25%
    position_margins = {0: 25.0}
    
    result = calculator.calculate_multi_position_prices(
        positions, 20000, 30, target_margin,
        position_margins=position_margins
    )
    
    # В режиме per_position каждая позиция должна иметь свою маржу
    assert pytest.approx(result['positions'][0]['margin'], abs=0.5) == 25.0
    assert pytest.approx(result['positions'][1]['margin'], abs=0.5) == target_margin


def test_position_margins_with_bank_guarantee():
    """Тест индивидуальных марж с банковской гарантией."""
    calculator = MultiPositionCalculator({
        'pricing': {'mode': 'global'},
        'calculation_constants': {
            'conversion_rate': 11.5,
            'logistics_cnr_ratio': 0.3,
            'logistics_rf_ratio': 0.7,
            'conversion_fee_rate': 0.032,
            'credit_rate': 0.16,
        }
    })
    
    positions = _sample_positions()
    target_margin = 30.0
    position_margins = {0: 25.0}  # Первая позиция с индивидуальной маржой
    
    result = calculator.calculate_multi_position_prices(
        positions, 20000, 30, target_margin,
        position_margins=position_margins,
        use_bank_guarantee=True,
        payment_days=30
    )
    
    # Проверяем, что первая позиция имеет индивидуальную маржу
    assert pytest.approx(result['positions'][0]['margin'], abs=0.5) == 25.0
    
    # Проверяем, что общая маржа соответствует целевой (с учетом банковской гарантии)
    assert pytest.approx(result['actual_margin'], abs=1.0) == target_margin


if __name__ == "__main__":
    # Запуск тестов при прямом выполнении файла
    pytest.main([__file__, "-v"])
