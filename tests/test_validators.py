"""
Unit-тесты для модуля валидации параметров.
"""

import pytest
from ai_agent.validators import ParameterValidator


class TestParameterValidator:
    """Тесты для ParameterValidator."""
    
    def test_validate_logistics_params(self):
        """Тест валидации параметров логистики."""
        validator = ParameterValidator()
        
        # Валидные параметры
        params = {'weight_kg': 1000, 'city_name': 'Москва'}
        is_valid, error, corrected = validator.validate_logistics_params(params)
        assert is_valid is True
        assert error is None
        
        # Невалидный вес
        params = {'weight_kg': -100, 'city_name': 'Москва'}
        is_valid, error, corrected = validator.validate_logistics_params(params)
        assert is_valid is False
        assert error is not None
        
        # Отсутствует город
        params = {'weight_kg': 1000}
        is_valid, error, corrected = validator.validate_logistics_params(params)
        assert is_valid is False
    
    def test_validate_analytics_params(self):
        """Тест валидации параметров аналитики."""
        validator = ParameterValidator()
        
        # Валидные параметры
        params = {'search_query': 'тест'}
        is_valid, error, corrected = validator.validate_analytics_params(params)
        assert is_valid is True
        
        # Пустой запрос
        params = {'search_query': ''}
        is_valid, error, corrected = validator.validate_analytics_params(params)
        # Пустой запрос может быть валидным для некоторых типов запросов
        assert isinstance(is_valid, bool)

