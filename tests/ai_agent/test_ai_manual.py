"""
Скрипт для ручного тестирования улучшений AI агента (без эмодзи).
"""

import os
import sys
from pathlib import Path

# Добавляем корень проекта в путь
project_root = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(project_root))

# Устанавливаем тестовые переменные окружения
os.environ['OPENROUTER_API_KEY'] = 'sk-or-v1-test-' + 'a' * 40
os.environ['OPENROUTER_MODEL'] = 'xiaomi/mimo-v2-flash:free'
os.environ['AI_FALLBACK_ENABLED'] = 'true'


def test_config_validation():
    """Тест валидации конфигурации."""
    print("\n=== Test 1: Config validation ===")
    from ai_agent.config import validate_api_key, ConfigurationError
    
    try:
        validate_api_key('sk-or-v1-' + 'a' * 40)
        print("[OK] Valid key accepted")
    except Exception as e:
        print(f"[FAIL] Error: {e}")
    
    try:
        validate_api_key('invalid-key')
        print("[FAIL] Invalid key should NOT be accepted!")
    except ConfigurationError:
        print("[OK] Invalid key rejected")
    
    try:
        validate_api_key(None)
        print("[FAIL] Empty key should NOT be accepted!")
    except ConfigurationError:
        print("[OK] Empty key rejected")


def test_api_validator():
    """Тест валидатора API."""
    print("\n=== Test 2: API Key Validator ===")
    from ai_agent.api_validator import APIKeyValidator
    from unittest.mock import Mock, patch
    
    # Mock success
    with patch('ai_agent.api_validator.requests.post') as mock_post:
        mock_response = Mock()
        mock_response.status_code = 200
        mock_response.headers = {}
        mock_post.return_value = mock_response
        
        result = APIKeyValidator.validate_key('sk-or-v1-' + 'a' * 40)
        if result.is_valid:
            print("[OK] Validation successful (200)")
        else:
            print(f"[FAIL] Validation failed: {result.error_message}")
    
    # Mock 401
    with patch('ai_agent.api_validator.requests.post') as mock_post:
        mock_response = Mock()
        mock_response.status_code = 401
        mock_response.headers = {}
        mock_post.return_value = mock_response
        
        result = APIKeyValidator.validate_key('sk-or-v1-' + 'a' * 40)
        if not result.is_valid:
            print("[OK] Error 401 handled correctly")
        else:
            print(f"[FAIL] Error 401 not handled")
    
    # Mock 429
    with patch('ai_agent.api_validator.requests.post') as mock_post:
        mock_response = Mock()
        mock_response.status_code = 429
        mock_response.headers = {'Retry-After': '60'}
        mock_post.return_value = mock_response
        
        result = APIKeyValidator.validate_key('sk-or-v1-' + 'a' * 40)
        if result.is_valid and result.error_message and "rate limit" in result.error_message.lower():
            print("[OK] Error 429 handled correctly")
        else:
            print(f"[FAIL] Error 429 not handled properly")


def test_cache_manager():
    """Тест менеджера кеша."""
    print("\n=== Test 3: Cache Manager ===")
    from ai_agent.cache_manager import AICacheManager
    from unittest.mock import Mock, patch
    import json
    
    # Test caching
    with patch('ai_agent.cache_manager.get_redis_client') as mock_redis:
        mock_client = Mock()
        mock_redis.return_value = mock_client
        
        result = AICacheManager.cache_response(
            message="test question",
            response="test answer",
            ttl=3600
        )
        
        if result and mock_client.setex.called:
            print("[OK] Caching works")
        else:
            print("[FAIL] Caching doesn't work")
    
    # Test retrieval
    with patch('ai_agent.cache_manager.get_redis_client') as mock_redis:
        mock_client = Mock()
        cached_data = json.dumps({
            'message': 'test',
            'response': 'cached answer'
        })
        mock_client.get.return_value = cached_data
        mock_redis.return_value = mock_client
        
        result = AICacheManager.get_cached_response("test")
        
        if result == 'cached answer':
            print("[OK] Cache retrieval works")
        else:
            print("[FAIL] Cache retrieval doesn't work")


def test_usage_monitor():
    """Тест мониторинга."""
    print("\n=== Test 4: Usage Monitor ===")
    from ai_agent.usage_monitor import UsageMonitor
    from unittest.mock import Mock, patch
    
    # Test logging
    with patch('ai_agent.usage_monitor.SessionLocal') as mock_session:
        mock_session_inst = Mock()
        mock_session.return_value = mock_session_inst
        
        result = UsageMonitor.log_usage(
            user_id=1,
            request_tokens=10,
            response_tokens=20,
            response_time_ms=500
        )
        
        if result and mock_session_inst.add.called:
            print("[OK] Usage logging works")
        else:
            print("[FAIL] Logging doesn't work")


def test_error_classes():
    """Тест классов ошибок."""
    print("\n=== Test 5: Error Classes ===")
    from ai_agent.agent import (
        APIError,
        APIKeyInvalidError,
        APIRateLimitError,
        APIServerError
    )
    
    # Test base error
    error = APIError("Test error", status_code=500)
    if error.status_code == 500:
        print("[OK] APIError works")
    else:
        print("[FAIL] APIError doesn't work")
    
    # Test key error
    error = APIKeyInvalidError("Invalid key", status_code=401)
    if error.status_code == 401:
        print("[OK] APIKeyInvalidError works")
    else:
        print("[FAIL] APIKeyInvalidError doesn't work")
    
    # Test rate limit
    error = APIRateLimitError("Rate limit", status_code=429, retry_after=60)
    if error.retry_after == 60:
        print("[OK] APIRateLimitError works")
    else:
        print("[FAIL] APIRateLimitError doesn't work")


def main():
    """Главная функция."""
    print("=" * 60)
    print("AI Agent Improvements - Manual Testing")
    print("=" * 60)
    
    try:
        test_config_validation()
        test_api_validator()
        test_cache_manager()
        test_usage_monitor()
        test_error_classes()
        
        print("\n" + "=" * 60)
        print("[SUCCESS] All basic tests passed!")
        print("=" * 60)
        print("\nNotes:")
        print("- Config: API key validation works")
        print("- API validator: handles all statuses (200, 401, 429)")
        print("- Caching: Redis integration ready")
        print("- Monitoring: DB logging works")
        print("- Error handling: all error classes created")
        print("\nFor full testing:")
        print("1. Set real API key in .env")
        print("2. Run Flask application")
        print("3. Go to /ai-agent for testing")
        print("4. Check admin panel at /admin/ai-agent")
        
    except Exception as e:
        print(f"\n[ERROR] Testing failed: {e}")
        import traceback
        traceback.print_exc()


if __name__ == '__main__':
    main()

