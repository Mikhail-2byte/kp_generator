import math

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
    assert result['margin'] == 25


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
