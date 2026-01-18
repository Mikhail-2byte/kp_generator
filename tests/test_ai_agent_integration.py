"""
Интеграционные тесты для AI агента.
Покрывают сценарии использования, обработку ошибок, benchmarks.
"""

import pytest
from unittest.mock import Mock, patch
from ai_agent.agent import AIAgent
from ai_agent.analytics_helper import AnalyticsHelper


class TestAIAgentIntegration:
    """Интеграционные тесты для AIAgent."""
    
    @pytest.fixture
    def agent(self):
        """Создает экземпляр AIAgent для тестов."""
        # Используем мок API ключа для тестов
        with patch('ai_agent.agent.get_api_key', return_value='test_key'):
            agent = AIAgent(api_key='test_key')
            return agent
    
    def test_logistics_calculation_flow(self, agent):
        """Тест полного потока расчета логистики."""
        message = "Рассчитай логистику из КНР в Москву за груз весом 5000 кг"
        
        # Мокаем API вызов
        with patch.object(agent, '_call_api') as mock_api:
            mock_api.return_value = {
                'choices': [{
                    'message': {
                        'content': 'Расчет логистики выполнен.'
                    }
                }]
            }
            
            response = agent.chat(message)
            assert response is not None
            assert len(response) > 0
    
    def test_analytics_search_flow(self, agent):
        """Тест полного потока поиска в аналитике."""
        message = "Найди в 1с чертеж 35.1774.000 СБ"
        
        # Мокаем API вызов
        with patch.object(agent, '_call_api') as mock_api:
            mock_api.return_value = {
                'choices': [{
                    'message': {
                        'content': 'Поиск выполнен.'
                    }
                }]
            }
            
            response = agent.chat(message)
            assert response is not None
    
    def test_error_handling(self, agent):
        """Тест обработки ошибок."""
        message = "Тестовый запрос"
        
        # Мокаем ошибку API
        with patch.object(agent, '_call_api') as mock_api:
            from ai_agent.agent import APIError
            mock_api.side_effect = APIError("Ошибка API")
            
            response = agent.chat(message)
            assert response is not None
            assert 'ошибка' in response.lower() or 'error' in response.lower()
    
    def test_cache_usage(self, agent):
        """Тест использования кеша."""
        message = "Повторный запрос"
        
        # Первый запрос
        with patch.object(agent, '_call_api') as mock_api:
            mock_api.return_value = {
                'choices': [{
                    'message': {
                        'content': 'Первый ответ'
                    }
                }]
            }
            response1 = agent.chat(message)
        
        # Второй запрос (должен использовать кеш)
        response2 = agent.chat(message)
        
        # Проверяем, что ответы одинаковые (если кеш работает)
        assert response1 == response2
    
    def test_context_processing(self, agent):
        """Тест обработки контекста из предыдущих сообщений."""
        # Первое сообщение
        message1 = "Найди в 1с чертеж 35.1774.000 СБ"
        agent.chat(message1)
        
        # Второе сообщение с уточнением
        message2 = "Покажи график"
        
        # Проверяем, что контекст учитывается
        with patch.object(agent, '_call_api') as mock_api:
            mock_api.return_value = {
                'choices': [{
                    'message': {
                        'content': 'График показан.'
                    }
                }]
            }
            response = agent.chat(message2)
            assert response is not None


class TestAnalyticsHelperIntegration:
    """Интеграционные тесты для AnalyticsHelper."""
    
    @pytest.fixture
    def analytics_helper(self):
        """Создает экземпляр AnalyticsHelper для тестов."""
        return AnalyticsHelper(lazy_load=True)
    
    def test_data_loading_and_search(self, analytics_helper):
        """Тест загрузки данных и поиска."""
        # Загружаем данные
        analytics_helper._ensure_data_loaded()
        
        # Проверяем, что данные загружены
        categories = analytics_helper.get_categories()
        assert len(categories) > 0
        
        # Пробуем поиск
        results = analytics_helper.search_by_nomenclature('тест', limit=5)
        assert isinstance(results, type(analytics_helper.data_cache.get(categories[0])))
    
    def test_chart_generation(self, analytics_helper):
        """Тест генерации графиков."""
        analytics_helper._ensure_data_loaded()
        
        categories = analytics_helper.get_categories()
        if categories:
            df = analytics_helper.data_cache.get(categories[0])
            if not df.empty:
                chart = analytics_helper.generate_chart_customers(df, limit=5)
                # График должен быть base64 строкой или None
                assert chart is None or isinstance(chart, str)
    
    def test_filtering_and_sorting(self, analytics_helper):
        """Тест фильтрации и сортировки."""
        analytics_helper._ensure_data_loaded()
        
        categories = analytics_helper.get_categories()
        if categories:
            df = analytics_helper.data_cache.get(categories[0])
            if not df.empty and 'Цена подачи' in df.columns:
                # Фильтрация
                filters = {'price_min': 1000, 'price_max': 10000}
                filtered = analytics_helper.filter_results(df, filters)
                assert len(filtered) <= len(df)
                
                # Сортировка
                sorted_df = analytics_helper.sort_results(df, 'Цена подачи', 'desc')
                assert len(sorted_df) == len(df)


class TestPerformanceBenchmarks:
    """Тесты производительности (benchmarks)."""
    
    def test_search_performance(self):
        """Тест производительности поиска."""
        import time
        helper = AnalyticsHelper(lazy_load=True)
        helper._ensure_data_loaded()
        
        start = time.time()
        for _ in range(10):
            helper.search_by_nomenclature('тест', limit=10)
        duration = time.time() - start
        
        # Поиск должен быть быстрым (< 1 секунды для 10 запросов)
        assert duration < 1.0, f"Поиск слишком медленный: {duration:.2f}s"
    
    def test_chart_generation_performance(self):
        """Тест производительности генерации графиков."""
        import time
        helper = AnalyticsHelper(lazy_load=True)
        helper._ensure_data_loaded()
        
        categories = helper.get_categories()
        if categories:
            df = helper.data_cache.get(categories[0])
            if not df.empty:
                start = time.time()
                chart = helper.generate_chart_customers(df, limit=10)
                duration = time.time() - start
                
                # Генерация графика должна быть разумно быстрой (< 5 секунд)
                assert duration < 5.0, f"Генерация графика слишком медленная: {duration:.2f}s"
                assert chart is None or isinstance(chart, str)

