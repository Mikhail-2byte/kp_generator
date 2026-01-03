"""Модуль кэширования для справочников и часто используемых данных."""

import json
from functools import lru_cache, wraps
from typing import Any, Callable, Dict, Optional, TypeVar

from flask import current_app

T = TypeVar('T')

# Попытка импорта Redis
try:
    import redis
    REDIS_AVAILABLE = True
except ImportError:
    REDIS_AVAILABLE = False
    redis = None


class CacheManager:
    """Менеджер кэша для справочников."""
    
    def __init__(self):
        self._cache: Dict[str, Any] = {}
    
    def get(self, key: str) -> Optional[Any]:
        """Получает значение из кэша."""
        return self._cache.get(key)
    
    def set(self, key: str, value: Any) -> None:
        """Устанавливает значение в кэш."""
        self._cache[key] = value
    
    def invalidate(self, key: Optional[str] = None) -> None:
        """
        Инвалидирует кэш.
        
        Args:
            key: Ключ для инвалидации. Если None, очищает весь кэш.
        """
        if key:
            self._cache.pop(key, None)
        else:
            self._cache.clear()
    
    def clear_all(self) -> None:
        """Очищает весь кэш."""
        self._cache.clear()


# Глобальный экземпляр менеджера кэша
_cache_manager = CacheManager()


def get_cache_manager() -> CacheManager:
    """Возвращает глобальный экземпляр менеджера кэша."""
    return _cache_manager


def cached_dataset(maxsize: int = 128):
    """
    Декоратор для кэширования функций загрузки справочников.
    
    Args:
        maxsize: Максимальный размер кэша (по умолчанию 128)
    """
    def decorator(func: Callable[..., T]) -> Callable[..., T]:
        @lru_cache(maxsize=maxsize)
        @wraps(func)
        def wrapper(*args: Any, **kwargs: Any) -> T:
            return func(*args, **kwargs)
        
        # Добавляем метод для очистки кэша
        wrapper.cache_clear = wrapper.cache_clear  # type: ignore
        
        return wrapper
    return decorator


def invalidate_dataset_cache(func_name: Optional[str] = None) -> None:
    """
    Инвалидирует кэш справочников.
    
    Args:
        func_name: Имя функции для инвалидации. Если None, очищает весь кэш.
    """
    if func_name:
        # Находим функцию и очищаем её кэш
        try:
            func = globals().get(func_name)
            if func and hasattr(func, 'cache_clear'):
                func.cache_clear()  # type: ignore
        except Exception:
            pass
    else:
        # Очищаем весь кэш менеджера
        _cache_manager.clear_all()
    
    # Также очищаем Redis кэш если доступен
    if REDIS_AVAILABLE:
        try:
            redis_client = get_redis_client()
            if redis_client:
                if func_name:
                    redis_client.delete(f"dataset:{func_name}")
                else:
                    # Удаляем все ключи с префиксом dataset:
                    keys = redis_client.keys("dataset:*")
                    if keys:
                        redis_client.delete(*keys)
        except Exception:
            pass


def get_redis_client() -> Optional[Any]:
    """
    Получает клиент Redis если он настроен и доступен.
    
    Returns:
        Redis клиент или None если Redis недоступен
    """
    if not REDIS_AVAILABLE:
        return None
    
    try:
        app = current_app._get_current_object() if hasattr(current_app, '_get_current_object') else current_app
        if not app:
            return None
        
        # Проверяем, есть ли уже клиент в контексте приложения
        if hasattr(app, 'redis_client'):
            return app.redis_client
        
        # Пытаемся создать клиент из конфигурации
        config = app.config.get('APP_SETTINGS', {})
        redis_config = config.get('redis', {})
        
        if not redis_config.get('enabled', False):
            return None
        
        host = redis_config.get('host', 'localhost')
        port = redis_config.get('port', 6379)
        db = redis_config.get('db', 0)
        password = redis_config.get('password')
        
        client = redis.Redis(
            host=host,
            port=port,
            db=db,
            password=password,
            decode_responses=True,
            socket_connect_timeout=2,
            socket_timeout=2
        )
        
        # Проверяем подключение
        client.ping()
        
        # Сохраняем клиент в контексте приложения
        app.redis_client = client
        return client
    except Exception:
        return None


def cached_redis(key_prefix: str, ttl: int = 3600):
    """
    Декоратор для кэширования результатов функций в Redis.
    
    Args:
        key_prefix: Префикс для ключа кэша
        ttl: Время жизни кэша в секундах (по умолчанию 1 час)
    
    Returns:
        Декоратор функции
    """
    def decorator(func: Callable[..., T]) -> Callable[..., T]:
        @wraps(func)
        def wrapper(*args: Any, **kwargs: Any) -> T:
            # Пытаемся получить из Redis
            redis_client = get_redis_client()
            if redis_client:
                try:
                    # Создаем ключ кэша
                    cache_key = f"{key_prefix}:{func.__name__}:{hash(str(args) + str(kwargs))}"
                    cached_value = redis_client.get(cache_key)
                    
                    if cached_value:
                        return json.loads(cached_value)
                except Exception:
                    pass
            
            # Если не нашли в Redis, выполняем функцию
            result = func(*args, **kwargs)
            
            # Сохраняем в Redis
            if redis_client:
                try:
                    cache_key = f"{key_prefix}:{func.__name__}:{hash(str(args) + str(kwargs))}"
                    redis_client.setex(
                        cache_key,
                        ttl,
                        json.dumps(result, default=str, ensure_ascii=False)
                    )
                except Exception:
                    pass
            
            return result
        
        return wrapper
    return decorator


__all__ = [
    'CacheManager',
    'get_cache_manager',
    'cached_dataset',
    'invalidate_dataset_cache',
    'get_redis_client',
    'cached_redis',
    'REDIS_AVAILABLE'
]

