"""
Unit-тесты для модуля извлечения намерений.
"""

import pytest
from unittest.mock import Mock
from ai_agent.intent_extractor import IntentExtractor


class TestIntentExtractor:
    """Тесты для IntentExtractor."""
    
    @pytest.fixture
    def intent_extractor(self):
        """Создает экземпляр IntentExtractor для тестов."""
        logistics_helper = Mock()
        analytics_helper = Mock()
        analytics_helper.get_categories.return_value = ['Барабаны', 'Вал']
        return IntentExtractor(logistics_helper, analytics_helper)
    
    def test_extract_logistics_intent(self, intent_extractor):
        """Тест извлечения намерения логистики."""
        message = "Рассчитай логистику из КНР в Москву за груз весом 5000 кг"
        intent = intent_extractor.extract(message)
        
        assert intent['is_logistics'] is True
        assert 'logistics_params' in intent
    
    def test_extract_analytics_intent(self, intent_extractor):
        """Тест извлечения намерения аналитики."""
        message = "Найди в 1с чертеж 35.1774.000 СБ"
        intent = intent_extractor.extract(message)
        
        assert intent['is_analytics'] is True
        assert 'analytics_params' in intent
    
    def test_extract_materials_intent(self, intent_extractor):
        """Тест извлечения намерения поиска материалов."""
        message = "Найди материал 10"
        intent = intent_extractor.extract(message)
        
        assert intent['is_materials'] is True
        assert 'materials_query' in intent

