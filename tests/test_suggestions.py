"""
Unit-тесты для модуля предложений.
"""

import pytest
from ai_agent.suggestions import SuggestionsGenerator


class TestSuggestionsGenerator:
    """Тесты для SuggestionsGenerator."""
    
    def test_generate_empty_result_message(self):
        """Тест генерации сообщения при пустых результатах."""
        message = SuggestionsGenerator.generate_empty_result_message(
            'search',
            query='тест',
            category='Барабаны'
        )
        assert 'тест' in message
        assert 'Барабаны' in message
        assert len(message) > 0
    
    def test_get_popular_queries(self):
        """Тест получения популярных запросов."""
        queries = SuggestionsGenerator.get_popular_queries(limit=3)
        assert len(queries) == 3
        assert all(isinstance(q, str) for q in queries)
        
        # С историей
        history = ['Топ заказчиков', 'Топ заказчиков', 'Статистика']
        popular = SuggestionsGenerator.get_popular_queries(history, limit=2)
        assert len(popular) == 2
        assert 'Топ заказчиков' in popular
    
    def test_generate_autocomplete_suggestions(self):
        """Тест генерации автодополнения."""
        options = ['Москва', 'Московская область', 'Санкт-Петербург', 'Екатеринбург']
        
        suggestions = SuggestionsGenerator.generate_autocomplete_suggestions(
            'Моск',
            options,
            limit=3
        )
        assert len(suggestions) <= 3
        assert any('Москва' in s for s in suggestions)
    
    def test_generate_demo_queries(self):
        """Тест генерации демонстрационных запросов."""
        analytics_queries = SuggestionsGenerator.generate_demo_queries('analytics')
        assert len(analytics_queries) > 0
        assert all(isinstance(q, str) for q in analytics_queries)
        
        all_queries = SuggestionsGenerator.generate_demo_queries('all')
        assert len(all_queries) > len(analytics_queries)

