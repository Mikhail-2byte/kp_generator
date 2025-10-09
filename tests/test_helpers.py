import pytest

from app.helpers import validate_form_data


def _base_form(**overrides):
    base = {
        'company': 'Test Co',
        'product': 'Part',
        'quantity': '5',
        'cost_price': '100',
        'weight': '2',
        'logistics': '300',
        'duty_percent': '5',
        'delivery_time': '20',
        'margin_percent': '25',
    }
    base.update(overrides)
    return base


def test_validate_form_data_accepts_positive_cost_price_per_kg():
    form = _base_form(cost_price_per_kg='50')
    assert validate_form_data(form) == []


def test_validate_form_data_rejects_negative_cost_price_per_kg():
    form = _base_form(cost_price_per_kg='-10')
    errors = validate_form_data(form)
    assert any('cost_price_per_kg' in error for error in errors)


@pytest.mark.parametrize('value', ['', ' ', None])
def test_validate_form_data_allows_empty_cost_price_per_kg(value):
    form = _base_form(cost_price_per_kg=value)
    assert validate_form_data(form) == []
