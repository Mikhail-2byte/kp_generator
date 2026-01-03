import sys
from pathlib import Path

# Добавляем корень проекта в путь для возможности прямого запуска
project_root = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(project_root))

import pytest

from app.presentation.validators import FormValidationResult, validate_form_data


def _build_valid_form() -> dict[str, str]:
    return {
        'company': '  ООО Тест  ',
        'logistics': '1 500,5',
        'margin_percent': '25',
        'delivery_time': '30',
        'comment': 'Пример',
        'product': 'Позиция 1',
        'quantity': '10',
        'cost_price': '100',
        'weight': '2,5',
        'duty_percent': '5',
        'product_2': 'Позиция 2',
        'quantity_2': '5',
        'cost_price_2': '200',
        'weight_2': '1',
        'duty_percent_2': '0',
    }


def test_validate_form_data_success_multi_positions():
    result = validate_form_data(_build_valid_form())

    assert isinstance(result, FormValidationResult)
    assert not result.errors
    assert not result.invalid_fields
    assert len(result.positions) == 2

    assert result.cleaned_data['company'] == 'ООО Тест'
    assert result.cleaned_data['logistics'] == '1500.5'
    assert result.positions[0]['quantity'] == 10
    assert result.positions[1]['cost_price'] == 200.0


def test_validate_form_data_requires_company():
    form = _build_valid_form()
    form['company'] = ' '

    result = validate_form_data(form)

    assert 'company' in result.invalid_fields
    assert any('Укажите название компании.' in error for error in result.errors)


if __name__ == "__main__":
    # Запуск тестов при прямом выполнении файла
    pytest.main([__file__, "-v"])

