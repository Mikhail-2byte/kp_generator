"""Тесты для создания бюджета и получения данных для создания."""

import pytest
from unittest.mock import Mock, patch, MagicMock

from app.core.exceptions import ValidationError, CalculationError
from app.services.generation_orchestrator import GenerationOrchestrator
from app.services.multi_position_calculator import MultiPositionCalculator
from app.presentation.validators import validate_form_data
from app.presentation.helpers import extract_positions_from_form


@pytest.fixture
def app_config():
    """Конфигурация приложения для тестов."""
    return {
        'calculation_constants': {
            'conversion_rate': 12,
            'logistics_cnr_ratio': 0.3,
            'logistics_rf_ratio': 0.7,
            'conversion_fee_rate': 0.032,
            'credit_rate': 0.16,
        }
    }


@pytest.fixture
def orchestrator(app_config):
    """Создает экземпляр оркестратора для тестов."""
    return GenerationOrchestrator(app_config)


@pytest.fixture
def single_position_form_data():
    """Данные формы для одной позиции."""
    return {
        'company': 'ООО Тестовая Компания',
        'product': 'Деталь А',
        'quantity': '10',
        'cost_price': '1000',
        'weight': '5',
        'logistics': '50000',
        'margin_percent': '30',
        'delivery_time': '30',
        'duty_percent': '5',
        'drawing_number': 'Ч-001',
        'material': 'Сталь',
        'delivery_address': 'Москва, ул. Тестовая, д. 1',
    }


@pytest.fixture
def multi_position_form_data():
    """Данные формы для множественных позиций."""
    return {
        'company': 'ООО Тестовая Компания',
        'logistics': '50000',
        'margin_percent': '30',
        'delivery_time': '30',
        'delivery_address': 'Москва, ул. Тестовая, д. 1',
        'positions_payload': '[{"product": "Деталь А", "quantity": "10", "cost_price": "1000", "weight": "5", "duty_percent": "5"}, {"product": "Деталь B", "quantity": "5", "cost_price": "2000", "weight": "3", "duty_percent": "10"}]',
    }


class TestBudgetCreation:
    """Тесты для создания бюджета."""
    
    def test_budget_creation_single_position(self, orchestrator, single_position_form_data):
        """Тест создания бюджета для одной позиции."""
        # Мокаем шаблоны и генерацию документов
        with patch('app.services.generation_orchestrator.check_templates_exist') as mock_check:
            mock_check.return_value = []
            
            with patch.object(orchestrator, 'save_history') as mock_save:
                mock_save.return_value = True
                
                with patch.object(orchestrator, 'generate_documents') as mock_gen:
                    mock_zip = MagicMock()
                    mock_zip.seek = Mock()
                    mock_gen.return_value = (mock_zip, 'test_prefix')
                    
                    # Выполняем оркестрацию
                    zip_buffer, file_prefix = orchestrator.orchestrate(
                        single_position_form_data, None
                    )
                    
                    # Проверяем, что бюджет создан
                    assert file_prefix == 'test_prefix'
                    assert mock_save.called
                    assert mock_gen.called
                    
                    # Проверяем, что валидация была вызвана
                    call_args = mock_gen.call_args
                    assert call_args is not None
                    form_data_passed = call_args[0][0]
                    assert form_data_passed['company'] == 'ООО Тестовая Компания'
    
    def test_budget_creation_multi_position(self, orchestrator, multi_position_form_data):
        """Тест создания бюджета для множественных позиций."""
        with patch('app.services.generation_orchestrator.check_templates_exist') as mock_check:
            mock_check.return_value = []
            
            with patch.object(orchestrator, 'save_history') as mock_save:
                mock_save.return_value = True
                
                with patch.object(orchestrator, 'generate_documents') as mock_gen:
                    mock_zip = MagicMock()
                    mock_zip.seek = Mock()
                    mock_gen.return_value = (mock_zip, 'multi_test_prefix')
                    
                    # Выполняем оркестрацию
                    zip_buffer, file_prefix = orchestrator.orchestrate(
                        multi_position_form_data, None
                    )
                    
                    # Проверяем, что бюджет создан
                    assert file_prefix == 'multi_test_prefix'
                    assert mock_save.called
                    assert mock_gen.called
                    
                    # Проверяем, что переданы множественные позиции
                    call_args = mock_gen.call_args
                    positions_passed = call_args[0][1]
                    assert len(positions_passed) == 2
    
    def test_budget_data_validation(self, orchestrator, single_position_form_data):
        """Тест валидации данных для создания бюджета."""
        # Тест с валидными данными
        cleaned_data, positions, errors = orchestrator.validate_request(single_position_form_data)
        
        assert len(errors) == 0
        assert len(positions) > 0
        assert cleaned_data['company'] == 'ООО Тестовая Компания'
        assert cleaned_data['logistics'] == '50000'
        assert cleaned_data['margin_percent'] == '30'
        
        # Тест с невалидными данными (отсутствует компания)
        invalid_data = single_position_form_data.copy()
        invalid_data['company'] = ''
        
        cleaned_data, positions, errors = orchestrator.validate_request(invalid_data)
        
        assert len(errors) > 0
        assert any('компани' in error.lower() for error in errors)
    
    def test_budget_price_calculation_with_margin(self, orchestrator):
        """Тест расчета продажной цены на основе итоговой маржи."""
        # Тест для одной позиции
        positions = [{
            'quantity': 10,
            'cost_price': 1000,
            'weight': 5,
            'duty_percent': 5
        }]
        
        position_prices, total_general_price = orchestrator.calculate_prices(
            positions, 50000, 30, 30
        )
        
        assert len(position_prices) == 1
        assert total_general_price > 0
        
        # Проверяем, что цена рассчитана с учетом маржи
        final_price = position_prices[0]['final_price']
        assert final_price > 0
        
        # Тест для множественных позиций с единой маржой
        multi_positions = [
            {'quantity': 10, 'cost_price': 1000, 'weight': 5, 'duty_percent': 5},
            {'quantity': 5, 'cost_price': 2000, 'weight': 3, 'duty_percent': 10}
        ]
        
        position_prices, total_general_price = orchestrator.calculate_prices(
            multi_positions, 50000, 30, 30
        )
        
        assert len(position_prices) == 2
        assert total_general_price > 0
        
        # Проверяем, что итоговая маржа соответствует целевой
        calculator = orchestrator.calculator
        result = calculator.calculate_multi_position_prices(
            multi_positions, 50000, 30, 30
        )
        
        # Маржа должна быть близка к целевой (30%)
        assert abs(result['actual_margin'] - 30) < 1.0  # Допускаем небольшую погрешность
    
    def test_budget_get_data_for_creation(self, orchestrator, single_position_form_data):
        """Тест получения данных для создания бюджета."""
        # Получаем данные через валидацию
        cleaned_data, positions, errors = orchestrator.validate_request(single_position_form_data)
        
        # Проверяем, что данные извлечены корректно
        assert 'company' in cleaned_data
        assert 'logistics' in cleaned_data
        assert 'margin_percent' in cleaned_data
        assert 'delivery_time' in cleaned_data
        
        # Проверяем, что позиции извлечены
        assert len(positions) > 0
        first_position = positions[0]
        assert 'quantity' in first_position
        assert 'cost_price' in first_position
        assert 'weight' in first_position
        
        # Проверяем извлечение позиций напрямую
        extracted_positions = extract_positions_from_form(cleaned_data)
        assert len(extracted_positions) > 0
    
    def test_budget_calculation_margin_accuracy(self, orchestrator):
        """Тест точности расчета маржи при создании бюджета."""
        # Создаем позиции с известными параметрами
        positions = [
            {'quantity': 10, 'cost_price': 1000, 'weight': 5, 'duty_percent': 5},
            {'quantity': 5, 'cost_price': 2000, 'weight': 3, 'duty_percent': 10}
        ]
        
        target_margin = 25.0
        
        # Рассчитываем цены
        calculator = orchestrator.calculator
        result = calculator.calculate_multi_position_prices(
            positions, 50000, 30, target_margin
        )
        
        # Проверяем, что фактическая маржа близка к целевой
        actual_margin = result['actual_margin']
        assert abs(actual_margin - target_margin) < 0.5  # Допускаем погрешность 0.5%
        
        # Проверяем, что выручка больше затрат
        assert result['total_revenue'] > result['total_costs']
        
        # Проверяем расчет маржи: (выручка - затраты) / выручка * 100
        calculated_margin = (result['total_revenue'] - result['total_costs']) / result['total_revenue'] * 100
        assert abs(calculated_margin - actual_margin) < 0.01
    
    def test_budget_single_position_margin(self, orchestrator):
        """Тест расчета маржи для одной позиции."""
        position = {
            'quantity': 10,
            'cost_price': 1000,
            'weight': 5,
            'duty_percent': 5
        }
        
        target_margin = 30.0
        
        result = orchestrator.calculator.calculate_legacy_single_position(
            position, 50000, 30, target_margin
        )
        
        # Проверяем, что маржа соответствует целевой
        assert result['margin'] == target_margin
        
        # Проверяем, что цена рассчитана корректно
        assert result['final_price'] > 0
        assert result['general_price'] == result['final_price'] * position['quantity']
    
    def test_budget_validation_required_fields(self, orchestrator):
        """Тест валидации обязательных полей для создания бюджета."""
        # Тест с отсутствующими обязательными полями, но с позицией
        incomplete_data = {
            'company': 'ООО Тест',
            'product': 'Товар',
            'quantity': '10',
            'cost_price': '1000',
            'weight': '5',
            # Отсутствуют обязательные поля: logistics, margin_percent, delivery_time
        }
        
        cleaned_data, positions, errors = orchestrator.validate_request(incomplete_data)
        
        # Должны быть ошибки валидации
        assert len(errors) > 0
        
        # Проверяем, что ошибки указывают на отсутствующие поля
        error_text = ' '.join(errors).lower()
        assert 'логистик' in error_text or 'logistics' in error_text
        assert 'марж' in error_text or 'margin' in error_text
        assert 'поставк' in error_text or 'delivery' in error_text
    
    def test_budget_extract_positions_from_payload(self, orchestrator, multi_position_form_data):
        """Тест извлечения позиций из JSON payload."""
        cleaned_data, positions, errors = orchestrator.validate_request(multi_position_form_data)
        
        # Проверяем, что позиции извлечены из payload
        assert len(positions) == 2, f"Ожидалось 2 позиции, получено {len(positions)}. Ошибки: {errors}"
        
        # Проверяем данные первой позиции
        first_pos = positions[0]
        assert first_pos.get('product') == 'Деталь А'
        # После валидации quantity и cost_price могут быть строками или числами
        assert int(first_pos.get('quantity', 0)) == 10
        assert float(first_pos.get('cost_price', 0)) == 1000
        
        # Проверяем данные второй позиции
        second_pos = positions[1]
        assert second_pos.get('product') == 'Деталь B'
        assert int(second_pos.get('quantity', 0)) == 5
        assert float(second_pos.get('cost_price', 0)) == 2000

