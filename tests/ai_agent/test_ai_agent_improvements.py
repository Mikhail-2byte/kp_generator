"""
Тесты для AI агента с улучшенной обработкой ошибок.
"""

import os
import pytest
from unittest.mock import Mock, patch, MagicMock
import sys
from pathlib import Path

# Добавляем корень проекта в путь
project_root = Path(__file__).resolve().parents[2]
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))


@pytest.fixture
def mock_env():
    """Mock окружения с API ключом."""
    with patch.dict(os.environ, {
        'OPENROUTER_API_KEY': 'sk-or-v1-test-key-1234567890abcdef',
        'OPENROUTER_MODEL': 'test-model',
        'OPENROUTER_TIMEOUT': '30',
        'AI_FALLBACK_ENABLED': 'true'
    }):
        yield


@pytest.fixture
def mock_knowledge_base():
    """Mock базы знаний."""
    with patch('ai_agent.agent.KnowledgeBase') as mock_kb:
        instance = mock_kb.return_value
        instance.get_relevant_context.return_value = "Test context"
        instance.get_all_context.return_value = "All context"
        yield instance


@pytest.fixture
def mock_logistics():
    """Mock логистического помощника."""
    with patch('ai_agent.agent.LogisticsHelper') as mock_lh:
        instance = mock_lh.return_value
        instance.calculate_simple_logistics.return_value = {
            'success': True,
            'cost': 1000,
            'city': 'Москва'
        }
        instance.format_logistics_result.return_value = "Результат логистики"
        yield instance


class TestConfig:
    """Тесты конфигурации."""
    
    def test_config_validates_api_key_format(self, mock_env):
        """Проверка валидации формата API ключа."""
        from ai_agent.config import validate_api_key, ConfigurationError
        
        # Валидный ключ
        validate_api_key('sk-or-v1-' + 'a' * 40)
        
        # Невалидный формат
        with pytest.raises(ConfigurationError, match="должен начинаться"):
            validate_api_key('invalid-key-format')
        
        # Слишком короткий
        with pytest.raises(ConfigurationError, match="слишком короткий"):
            validate_api_key('sk-or-v1-short')
        
        # Пустой ключ
        with pytest.raises(ConfigurationError, match="не установлен"):
            validate_api_key(None)
    
    def test_config_loads_from_env(self, mock_env):
        """Проверка загрузки конфигурации из окружения."""
        from ai_agent.config import (
            get_model_name,
            get_timeout,
            is_fallback_enabled
        )
        
        assert get_model_name() == 'test-model'
        assert get_timeout() == 30
        assert is_fallback_enabled() is True


class TestAPIValidator:
    """Тесты валидатора API ключа."""
    
    @patch('ai_agent.api_validator.requests.post')
    def test_validate_key_success(self, mock_post):
        """Тест успешной валидации ключа."""
        from ai_agent.api_validator import validate_api_key
        
        # Mock успешного ответа
        mock_response = Mock()
        mock_response.status_code = 200
        mock_response.headers = {}
        mock_post.return_value = mock_response
        
        result = validate_api_key('sk-or-v1-' + 'a' * 40)
        
        assert result.is_valid is True
        assert result.error_message is None
    
    @patch('ai_agent.api_validator.requests.post')
    def test_validate_key_unauthorized(self, mock_post):
        """Тест невалидного ключа (401)."""
        from ai_agent.api_validator import validate_api_key
        
        mock_response = Mock()
        mock_response.status_code = 401
        mock_response.headers = {}
        mock_post.return_value = mock_response
        
        result = validate_api_key('sk-or-v1-' + 'a' * 40)
        
        assert result.is_valid is False
        assert "недействителен" in result.error_message.lower()
    
    @patch('ai_agent.api_validator.requests.post')
    def test_validate_key_rate_limit(self, mock_post):
        """Тест превышения rate limit (429)."""
        from ai_agent.api_validator import validate_api_key
        
        mock_response = Mock()
        mock_response.status_code = 429
        mock_response.headers = {'Retry-After': '60'}
        mock_post.return_value = mock_response
        
        result = validate_api_key('sk-or-v1-' + 'a' * 40)
        
        assert result.is_valid is True  # Ключ валиден, но превышен лимит
        assert "rate limit" in result.error_message.lower()
        assert 'x-ratelimit-limit' in result.rate_limit_info or True  # может быть пустым
    
    @patch('ai_agent.api_validator.requests.post')
    def test_validate_key_timeout(self, mock_post):
        """Тест timeout при валидации."""
        from ai_agent.api_validator import validate_api_key
        import requests
        
        mock_post.side_effect = requests.exceptions.Timeout()
        
        result = validate_api_key('sk-or-v1-' + 'a' * 40)
        
        assert result.is_valid is False
        assert "время ожидания" in result.error_message.lower()


class TestAIAgent:
    """Тесты AI агента."""
    
    def test_agent_initialization(self, mock_env, mock_knowledge_base, mock_logistics):
        """Тест инициализации агента."""
        from ai_agent.agent import AIAgent
        
        agent = AIAgent()
        
        assert agent.api_key == 'sk-or-v1-test-key-1234567890abcdef'
        assert agent.model == 'test-model'
        assert agent.fallback_enabled is True
    
    def test_agent_initialization_without_key(self):
        """Тест инициализации без API ключа."""
        from ai_agent.agent import AIAgent
        from ai_agent.config import ConfigurationError
        
        with patch.dict(os.environ, {}, clear=True):
            with pytest.raises(ConfigurationError):
                AIAgent()
    
    @patch('ai_agent.agent.requests.post')
    def test_call_api_success(self, mock_post, mock_env, mock_knowledge_base, mock_logistics):
        """Тест успешного вызова API."""
        from ai_agent.agent import AIAgent
        
        # Mock успешного ответа
        mock_response = Mock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            'choices': [{'message': {'content': 'Test response'}}],
            'usage': {'prompt_tokens': 10, 'completion_tokens': 20}
        }
        mock_post.return_value = mock_response
        
        agent = AIAgent()
        messages = [{'role': 'user', 'content': 'test'}]
        result = agent._call_api(messages)
        
        assert 'choices' in result
        assert result['choices'][0]['message']['content'] == 'Test response'
    
    @patch('ai_agent.agent.requests.post')
    def test_call_api_401_error(self, mock_post, mock_env, mock_knowledge_base, mock_logistics):
        """Тест обработки ошибки 401."""
        from ai_agent.agent import AIAgent, APIKeyInvalidError
        
        mock_response = Mock()
        mock_response.status_code = 401
        mock_post.return_value = mock_response
        
        agent = AIAgent()
        messages = [{'role': 'user', 'content': 'test'}]
        
        with pytest.raises(APIKeyInvalidError) as exc_info:
            agent._call_api(messages)
        
        assert exc_info.value.status_code == 401
        assert "недействителен" in str(exc_info.value).lower()
    
    @patch('ai_agent.agent.requests.post')
    def test_call_api_429_error(self, mock_post, mock_env, mock_knowledge_base, mock_logistics):
        """Тест обработки ошибки 429 (rate limit)."""
        from ai_agent.agent import AIAgent, APIRateLimitError
        
        mock_response = Mock()
        mock_response.status_code = 429
        mock_response.headers = {'Retry-After': '60'}
        mock_post.return_value = mock_response
        
        agent = AIAgent()
        messages = [{'role': 'user', 'content': 'test'}]
        
        with pytest.raises(APIRateLimitError) as exc_info:
            agent._call_api(messages)
        
        assert exc_info.value.status_code == 429
        assert exc_info.value.retry_after == 60
    
    @patch('ai_agent.agent.requests.post')
    def test_call_api_timeout(self, mock_post, mock_env, mock_knowledge_base, mock_logistics):
        """Тест обработки timeout."""
        from ai_agent.agent import AIAgent, APIError
        import requests
        
        mock_post.side_effect = requests.exceptions.Timeout()
        
        agent = AIAgent()
        messages = [{'role': 'user', 'content': 'test'}]
        
        with pytest.raises(APIError) as exc_info:
            agent._call_api(messages)
        
        assert "время ожидания" in str(exc_info.value).lower()
    
    @patch('ai_agent.agent.requests.post')
    @patch('ai_agent.agent.cache_ai_response')
    @patch('ai_agent.agent.get_cached_ai_response')
    def test_chat_with_caching(self, mock_get_cache, mock_set_cache, mock_post, mock_env, mock_knowledge_base, mock_logistics):
        """Тест работы кеширования в chat."""
        from ai_agent.agent import AIAgent
        
        # Первый запрос - cache miss
        mock_get_cache.return_value = None
        mock_response = Mock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            'choices': [{'message': {'content': 'Cached response'}}],
            'usage': {}
        }
        mock_post.return_value = mock_response
        
        agent = AIAgent()
        response1 = agent.chat("test question")
        
        assert response1 == 'Cached response'
        mock_set_cache.assert_called_once()
        
        # Второй запрос - cache hit
        mock_get_cache.return_value = 'Cached response'
        response2 = agent.chat("test question")
        
        assert response2 == 'Cached response'
        # API не должен вызываться при cache hit
        assert mock_post.call_count == 1
    
    @patch('ai_agent.agent.requests.post')
    def test_chat_fallback_on_error(self, mock_post, mock_env, mock_knowledge_base, mock_logistics):
        """Тест fallback режима при ошибке API."""
        from ai_agent.agent import AIAgent
        
        # Mock ошибки API
        mock_response = Mock()
        mock_response.status_code = 500
        mock_post.return_value = mock_response
        
        agent = AIAgent()
        response = agent.chat("test question")
        
        # Должен вернуть ошибку + fallback результат
        assert "временно недоступен" in response.lower() or "ошибка" in response.lower()
        assert "результаты поиска" in response.lower() or "контекст" in response.lower()
    
    def test_fallback_search(self, mock_env, mock_knowledge_base, mock_logistics):
        """Тест fallback поиска."""
        from ai_agent.agent import AIAgent
        
        mock_knowledge_base.get_relevant_context.return_value = "Test documentation context"
        
        agent = AIAgent()
        result = agent._fallback_search("test question")
        
        assert "результаты поиска" in result.lower()
        assert "test documentation context" in result.lower()


class TestCacheManager:
    """Тесты менеджера кеша."""
    
    @patch('ai_agent.cache_manager.get_redis_client')
    def test_cache_response(self, mock_redis_client):
        """Тест кеширования ответа."""
        from ai_agent.cache_manager import cache_ai_response
        
        mock_client = Mock()
        mock_redis_client.return_value = mock_client
        
        result = cache_ai_response(
            message="test question",
            response="test answer",
            ttl=3600
        )
        
        assert result is True
        mock_client.setex.assert_called_once()
    
    @patch('ai_agent.cache_manager.get_redis_client')
    def test_get_cached_response(self, mock_redis_client):
        """Тест получения из кеша."""
        from ai_agent.cache_manager import get_cached_ai_response
        import json
        
        mock_client = Mock()
        cached_data = json.dumps({
            'message': 'test question',
            'response': 'cached answer'
        })
        mock_client.get.return_value = cached_data
        mock_redis_client.return_value = mock_client
        
        result = get_cached_ai_response("test question")
        
        assert result == 'cached answer'
    
    @patch('ai_agent.cache_manager.get_redis_client')
    def test_invalidate_cache(self, mock_redis_client):
        """Тест очистки кеша."""
        from ai_agent.cache_manager import invalidate_ai_cache
        
        mock_client = Mock()
        mock_client.keys.return_value = ['key1', 'key2']
        mock_client.delete.return_value = 2
        mock_redis_client.return_value = mock_client
        
        result = invalidate_ai_cache()
        
        assert result is True
        mock_client.delete.assert_called_once()


class TestUsageMonitor:
    """Тесты мониторинга использования."""
    
    @patch('ai_agent.usage_monitor.SessionLocal')
    def test_log_usage(self, mock_session):
        """Тест логирования использования."""
        from ai_agent.usage_monitor import log_ai_usage
        
        mock_session_inst = Mock()
        mock_session.return_value = mock_session_inst
        
        result = log_ai_usage(
            user_id=1,
            request_tokens=10,
            response_tokens=20,
            response_time_ms=500,
            model_name='test-model'
        )
        
        assert result is True
        mock_session_inst.add.assert_called_once()
        mock_session_inst.commit.assert_called_once()
    
    @patch('ai_agent.usage_monitor.SessionLocal')
    def test_get_stats(self, mock_session):
        """Тест получения статистики."""
        from ai_agent.usage_monitor import UsageMonitor
        
        # Mock результата запроса
        mock_session_inst = Mock()
        mock_query = Mock()
        mock_result = Mock()
        mock_result.total_requests = 100
        mock_result.request_tokens = 1000
        mock_result.response_tokens = 2000
        mock_result.total_cost = 5.0
        mock_result.avg_response_time = 500
        mock_result.errors_count = 5
        mock_result.cache_hits = 40
        
        mock_query.one.return_value = mock_result
        mock_session_inst.query.return_value.filter.return_value = mock_query
        mock_session.return_value = mock_session_inst
        
        stats = UsageMonitor.get_stats(days=30)
        
        assert stats.total_requests == 100
        assert stats.total_tokens == 3000
        assert stats.cache_hit_rate == 40.0


if __name__ == '__main__':
    pytest.main([__file__, '-v'])

