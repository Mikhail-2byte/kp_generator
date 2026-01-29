"""Тесты для GenerationOrchestrator."""

import sys
from pathlib import Path
from unittest.mock import MagicMock, Mock, patch

# Добавляем корень проекта в путь для возможности прямого запуска
project_root = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(project_root))

import pytest

from app.core.exceptions import CalculationError, DocumentGenerationError, ValidationError
from app.services.generation_orchestrator import GenerationOrchestrator


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
def app_config_production():
    """Конфигурация как в config/settings.json (курс 11.5, НДС 0.22)."""
    return {
        'calculation_constants': {
            'conversion_rate': 11.5,
            'logistics_cnr_ratio': 0.3,
            'logistics_rf_ratio': 0.7,
            'conversion_fee_rate': 0.032,
            'credit_rate': 0.16,
            'vat_rate': 0.22,
        }
    }


@pytest.fixture
def orchestrator(app_config):
    """Создает экземпляр оркестратора для тестов."""
    return GenerationOrchestrator(app_config)


@pytest.fixture
def valid_form_data():
    """Валидные данные формы для тестов."""
    return {
        'company': 'Тестовая компания',
        'product': 'Товар',
        'quantity': '10',
        'cost_price': '100',
        'weight': '5',
        'logistics': '50000',
        'margin_percent': '30',
        'delivery_time': '30',
        'duty_percent': '5',
        'drawing_number': 'Ч-001',
        'material': 'Сталь',
        'delivery_address': 'Москва',
    }


class TestGenerationOrchestrator:
    """Тесты для GenerationOrchestrator."""
    
    def test_init(self, app_config):
        """Тест инициализации оркестратора."""
        orchestrator = GenerationOrchestrator(app_config)
        assert orchestrator.app_config == app_config
        assert orchestrator.calculator is not None
    
    def test_validate_request_success(self, orchestrator, valid_form_data):
        """Тест успешной валидации запроса."""
        with patch('app.services.generation_orchestrator.validate_form_data') as mock_validate:
            mock_validation = Mock()
            mock_validation.cleaned_data = valid_form_data
            mock_validation.positions = None
            mock_validation.errors = []
            mock_validation.invalid_fields = []
            mock_validate.return_value = mock_validation

            with patch('app.services.generation_orchestrator.extract_positions_from_form') as mock_extract:
                mock_extract.return_value = [{'quantity': 10, 'cost_price': 100}]
                
                cleaned_data, positions, errors, invalid_fields = orchestrator.validate_request(valid_form_data)

                assert errors == []
                assert len(positions) > 0
    
    def test_validate_request_with_errors(self, orchestrator, valid_form_data):
        """Тест валидации с ошибками."""
        with patch('app.services.generation_orchestrator.validate_form_data') as mock_validate:
            mock_validation = Mock()
            mock_validation.cleaned_data = valid_form_data
            mock_validation.positions = None
            mock_validation.errors = ['Ошибка валидации']
            mock_validation.invalid_fields = ['company']
            mock_validate.return_value = mock_validation

            cleaned_data, positions, errors, invalid_fields = orchestrator.validate_request(valid_form_data)

            assert len(errors) > 0
    
    def test_calculate_prices_single_position(self, orchestrator):
        """Тест расчета цен для одной позиции."""
        positions = [{
            'quantity': 10,
            'cost_price': 100,
            'weight': 5,
            'duty_percent': 5
        }]
        
        with patch.object(orchestrator.calculator, 'calculate_legacy_single_position') as mock_calc:
            mock_calc.return_value = {
                'position': positions[0],
                'final_price': 150.0,
                'general_price': 1500.0,
                'margin': 30.0
            }
            
            position_prices, total_general_price = orchestrator.calculate_prices(
                positions, 50000, 30, 30
            )
            
            assert len(position_prices) == 1
            assert total_general_price > 0
    
    def test_calculate_prices_multi_position(self, orchestrator):
        """Тест расчета цен для множественных позиций."""
        positions = [
            {'quantity': 10, 'cost_price': 100, 'weight': 5, 'duty_percent': 5},
            {'quantity': 5, 'cost_price': 200, 'weight': 3, 'duty_percent': 5}
        ]
        
        with patch.object(orchestrator.calculator, 'calculate_multi_position_prices') as mock_calc:
            mock_calc.return_value = {
                'positions': [
                    {
                        'position': positions[0],
                        'final_price': 150.0,
                        'general_price': 1500.0,
                        'costs': {'cost_per_unit': 115.0}
                    },
                    {
                        'position': positions[1],
                        'final_price': 300.0,
                        'general_price': 1500.0,
                        'costs': {'cost_per_unit': 230.0}
                    }
                ],
                'total_costs': 2300.0,
                'total_revenue': 3000.0,
                'target_margin': 30.0,
                'actual_margin': 30.0,
                'price_coefficient': 1.3
            }
            
            position_prices, total_general_price = orchestrator.calculate_prices(
                positions, 50000, 30, 30
            )
            
            assert len(position_prices) == 2
            assert total_general_price > 0
    
    def test_orchestrate_success(self, orchestrator, valid_form_data):
        """Тест успешной оркестрации генерации."""
        with patch.object(orchestrator, 'validate_request') as mock_validate:
            mock_validate.return_value = (valid_form_data, [{'quantity': 10}], [], [])

            with patch.object(orchestrator, 'calculate_prices') as mock_calc:
                mock_calc.return_value = ([{'final_price': 150.0}], 1500.0)
                
                with patch.object(orchestrator, 'save_history') as mock_save:
                    mock_save.return_value = True
                    
                    with patch.object(orchestrator, 'generate_documents') as mock_gen:
                        mock_zip = MagicMock()
                        mock_zip.seek = Mock()
                        mock_gen.return_value = (mock_zip, 'test_prefix')
                        
                        with patch('app.services.generation_orchestrator.check_templates_exist') as mock_check:
                            mock_check.return_value = []
                            
                            zip_buffer, file_prefix = orchestrator.orchestrate(
                                valid_form_data, None
                            )
                            
                            assert file_prefix == 'test_prefix'
    
    def test_orchestrate_validation_error(self, orchestrator, valid_form_data):
        """Тест обработки ошибок валидации."""
        with patch.object(orchestrator, 'validate_request') as mock_validate:
            mock_validate.return_value = (valid_form_data, [], ['Ошибка валидации'], [])

            with pytest.raises(ValidationError):
                orchestrator.orchestrate(valid_form_data, None)
    
    def test_orchestrate_calculation_error(self, orchestrator, valid_form_data):
        """Тест обработки ошибок расчета."""
        with patch.object(orchestrator, 'validate_request') as mock_validate:
            mock_validate.return_value = (valid_form_data, [{'quantity': 10}], [], [])

            with patch.object(orchestrator, 'calculate_prices') as mock_calc:
                mock_calc.side_effect = Exception('Ошибка расчета')
                
                with pytest.raises(CalculationError):
                    orchestrator.orchestrate(valid_form_data, None)
    
    def test_orchestrate_document_error(self, orchestrator, valid_form_data):
        """Тест обработки ошибок генерации документов."""
        with patch.object(orchestrator, 'validate_request') as mock_validate:
            mock_validate.return_value = (valid_form_data, [{'quantity': 10}], [], [])

            with patch.object(orchestrator, 'calculate_prices') as mock_calc:
                mock_calc.return_value = ([{'final_price': 150.0}], 1500.0)
                
                with patch.object(orchestrator, 'save_history') as mock_save:
                    mock_save.return_value = True
                    
                    with patch.object(orchestrator, 'generate_documents') as mock_gen:
                        mock_gen.side_effect = Exception('Ошибка генерации')
                        
                        with patch('app.services.generation_orchestrator.check_templates_exist') as mock_check:
                            mock_check.return_value = []
                            
                            with pytest.raises(DocumentGenerationError):
                                orchestrator.orchestrate(valid_form_data, None)

    def test_preview_calculation_reference_excel(self, app_config_production):
        """
        Предварительный расчёт: 1 позиция, закуп 10 ¥/кг, 10 шт, вес 1000 кг,
        пошлина 0, срок поставки 150, условия оплаты 30 дн, Кредит, маржа 30%.
        Логистика 10000 руб (как в эталонном Excel — даёт ~869.57 ¥).
        Ожидаемые значения сверяются с эталонным Excel (ТД РИНАКО).
        """
        orchestrator = GenerationOrchestrator(app_config_production)
        # Закуп 10 за кг, 10 шт, вес 1000 кг → cost_price = 10 * 1000 = 10000 ¥/шт
        positions = [{
            'quantity': 10,
            'cost_price': 10000.0,
            'weight': 1000.0,
            'duty_percent': 0,
            'product': 'тест ч.1111',
            'drawing_number': '1111',
            'material': '1111',
        }]
        logistics_rub = 10000.0
        delivery_time = 150
        margin_percent = 30.0
        use_credit = True
        payment_days = 30

        position_prices, total_general_price = orchestrator.calculate_prices(
            positions,
            logistics_rub,
            delivery_time,
            margin_percent,
            use_credit=use_credit,
            payment_days=payment_days,
        )

        assert len(position_prices) == 1
        res = position_prices[0]
        final_price = res['final_price']
        general_price = res['general_price']
        margin = res['margin']

        # Эталонный Excel: Цена без НДС за ед. ~15994 ¥, Итого по спецификации ~159940 ¥
        assert abs(final_price - 15994.28) < 1.0, f'final_price={final_price}'
        assert abs(general_price - 159942.8) < 2.0, f'general_price={general_price}'
        assert abs(total_general_price - 159942.8) < 2.0
        assert abs(margin - 30.0) < 0.5, f'margin={margin}'

        # Проверка составляющих через калькулятор (логистика КНР/РФ, конвертация, кредит)
        costs = orchestrator.calculator.calculate_position_costs(
            positions[0], logistics_rub, delivery_time, 10000.0,
            use_credit=use_credit, payment_days=payment_days
        )
        logistics_cnr_total = costs['logistics_cnr_per_unit'] * 10
        logistics_rf_total = costs['logistics_rf_per_unit'] * 10
        conversion_total = costs['conversion_fee_per_unit'] * 10
        credit_total = costs['credit_cost_per_unit'] * 10
        # Эталон: Логистика КНР ~260.90 ¥, РФ ~608.70 ¥, конвертация 3200 ¥, кредит ~7890.41 ¥
        assert abs(logistics_cnr_total - 260.87) < 1.0
        assert abs(logistics_rf_total - 608.70) < 1.0
        assert abs(conversion_total - 3200.0) < 0.1
        assert abs(credit_total - 7890.41) < 1.0


if __name__ == "__main__":
    # Запуск тестов при прямом выполнении файла
    pytest.main([__file__, "-v"])

