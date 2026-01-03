"""
AI Agent для консультаций по проекту KP Generator.

Модуль предоставляет AI агента на базе Xiaomi Mimo V2 через OpenRouter API
для ответов на вопросы по документации и расчета логистики.
"""

__version__ = "0.1.0"

# Импортируем подмодули напрямую при загрузке пакета
# Это необходимо для работы patch('ai_agent.module') с unittest.mock.patch
# который использует pkgutil.resolve_name, требующий реальных атрибутов модуля

# Базовые модули импортируем напрямую через относительные импорты
# Это стандартный способ Python для пакетов
from . import config
from . import api_validator
from . import agent
from . import knowledge_base
from . import logistics_helper
from . import materials_helper
from . import duty_helper

# Модули, которые требуют Flask (загружаются лениво через __getattr__)
import importlib
import sys
_FLASK_DEPENDENT_MODULES = ['cache_manager', 'usage_monitor']
_current_module = sys.modules[__name__]

# Делаем классы и функции доступными через from ai_agent import ...
# Используем __getattr__ для ленивой загрузки классов, функций и Flask-зависимых модулей
def __getattr__(name: str):
    """Lazy импорт классов и функций для удобного доступа через from ai_agent import ..."""
    
    # Ленивая загрузка модулей, зависящих от Flask
    if name in _FLASK_DEPENDENT_MODULES:
        try:
            if name == 'cache_manager':
                from . import cache_manager
                setattr(_current_module, name, cache_manager)
                return cache_manager
            elif name == 'usage_monitor':
                from . import usage_monitor
                setattr(_current_module, name, usage_monitor)
                return usage_monitor
        except ImportError as e:
            import warnings
            warnings.warn(f"Не удалось загрузить модуль ai_agent.{name}: {e}", ImportWarning)
            raise AttributeError(f"module '{__name__}' has no attribute '{name}' (Flask not initialized?)")
    
    # Импортируем классы и функции для удобного доступа через from ai_agent import ...
    if name == 'AIAgent':
        from ai_agent.agent import AIAgent
        return AIAgent
    if name == 'APIError':
        from ai_agent.agent import APIError
        return APIError
    if name == 'APIKeyInvalidError':
        from ai_agent.agent import APIKeyInvalidError
        return APIKeyInvalidError
    if name == 'APIRateLimitError':
        from ai_agent.agent import APIRateLimitError
        return APIRateLimitError
    if name == 'ConfigurationError':
        from ai_agent.config import ConfigurationError
        return ConfigurationError
    if name == 'validate_api_key':
        from ai_agent.config import validate_api_key
        return validate_api_key
    if name == 'get_model_name':
        from ai_agent.config import get_model_name
        return get_model_name
    if name == 'get_timeout':
        from ai_agent.config import get_timeout
        return get_timeout
    if name == 'is_fallback_enabled':
        from ai_agent.config import is_fallback_enabled
        return is_fallback_enabled
    if name == 'APIKeyValidator':
        from ai_agent.api_validator import APIKeyValidator
        return APIKeyValidator
    if name == 'AICacheManager':
        from ai_agent.cache_manager import AICacheManager
        return AICacheManager
    if name == 'cache_ai_response':
        from ai_agent.cache_manager import cache_ai_response
        return cache_ai_response
    if name == 'get_cached_ai_response':
        from ai_agent.cache_manager import get_cached_ai_response
        return cache_ai_response
    if name == 'invalidate_ai_cache':
        from ai_agent.cache_manager import invalidate_ai_cache
        return invalidate_ai_cache
    if name == 'UsageMonitor':
        from ai_agent.usage_monitor import UsageMonitor
        return UsageMonitor
    if name == 'log_ai_usage':
        from ai_agent.usage_monitor import log_ai_usage
        return log_ai_usage
    
    raise AttributeError(f"module '{__name__}' has no attribute '{name}'")

# Для совместимости с pkgutil.resolve_name, который может использовать dir()
# добавляем подмодули в __dir__ (опционально, но помогает)
def __dir__():
    """Возвращает список доступных атрибутов."""
    return sorted([
        'agent', 'config', 'api_validator', 'cache_manager',
        'usage_monitor', 'knowledge_base', 'logistics_helper',
        'materials_helper', 'duty_helper',
        'AIAgent', 'APIError', 'APIKeyInvalidError', 'APIRateLimitError',
        'ConfigurationError', 'validate_api_key', 'get_model_name', 
        'get_timeout', 'is_fallback_enabled',
        'APIKeyValidator',
        'AICacheManager', 'cache_ai_response', 'get_cached_ai_response', 
        'invalidate_ai_cache',
        'UsageMonitor', 'log_ai_usage',
    ])

__all__ = [
    # Agent
    'AIAgent',
    'APIError',
    'APIKeyInvalidError',
    'APIRateLimitError',
    # Config
    'ConfigurationError',
    'get_model_name',
    'get_timeout',
    'is_fallback_enabled',
    'validate_api_key',
    # API Validator
    'APIKeyValidator',
    # Cache Manager
    'AICacheManager',
    'cache_ai_response',
    'get_cached_ai_response',
    'invalidate_ai_cache',
    # Usage Monitor
    'UsageMonitor',
    'log_ai_usage',
]
