import pytest

from app.calculate import calculate_selling_price


def test_calculate_selling_price_basic():
    result = calculate_selling_price(
        quantity=10,
        purchase_cost=100.0,
        logistics_rub=1200.0,
        duty_percent=5.0,
        weight=2.5,
        delivery_time=30,
        margin_percent=30.0,
    )

    expected = 170.950097
    assert result == pytest.approx(expected, rel=1e-6)


def test_calculate_selling_price_zero_quantity():
    with pytest.raises(ValueError):
        calculate_selling_price(
            quantity=0,
            purchase_cost=100.0,
            logistics_rub=100.0,
            duty_percent=5.0,
            weight=1.0,
            delivery_time=10,
        )


def test_calculate_selling_price_with_custom_config():
    config = {
        'calculation_constants': {
            'conversion_rate': 10,
            'logistics_cnr_ratio': 0.4,
            'logistics_rf_ratio': 0.6,
            'conversion_fee_rate': 0.05,
            'credit_rate': 0.2,
        }
    }

    result = calculate_selling_price(
        quantity=5,
        purchase_cost=200.0,
        logistics_rub=500.0,
        duty_percent=10.0,
        weight=1.2,
        delivery_time=20,
        margin_percent=25.0,
        config=config,
    )

    expected = 323.455708
    assert result == pytest.approx(expected, rel=1e-6)
