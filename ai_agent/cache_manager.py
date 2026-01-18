"""
Менеджер кеширования для AI агента с использованием Redis.
"""

import hashlib
import json
import logging
import os
import re
from typing import Optional

from app.core.cache import REDIS_AVAILABLE, get_redis_client

logger = logging.getLogger(__name__)


class AICacheManager:
    """Менеджер кеширования ответов AI агента."""
    
    # TTL для кеша (24 часа по умолчанию)
    DEFAULT_TTL = int(os.getenv('AI_CACHE_TTL', '86400'))
    
    # Префикс для ключей кеша
    CACHE_PREFIX = 'ai_agent'
    
    # Префикс для кеша запросов к данным 1С
    ANALYTICS_CACHE_PREFIX = 'ai_agent:analytics'
    
    # Префикс для кеша графиков
    CHART_CACHE_PREFIX = 'ai_agent:charts'
    
    @classmethod
    def _normalize_message(cls, message: str) -> str:
        """
        Нормализует сообщение для кеширования.
        
        Args:
            message: Исходное сообщение
            
        Returns:
            Нормализованное сообщение
        """
        # Приводим к нижнему регистру
        normalized = message.lower().strip()
        
        # Удаляем множественные пробелы
        normalized = re.sub(r'\s+', ' ', normalized)
        
        # Удаляем знаки препинания в конце
        normalized = normalized.rstrip('.,!?;:')
        
        return normalized
    
    @classmethod
    def _generate_cache_key(cls, message: str, context: Optional[str] = None) -> str:
        """
        Генерирует ключ кеша на основе сообщения и контекста.
        
        Args:
            message: Сообщение пользователя
            context: Дополнительный контекст
            
        Returns:
            Ключ кеша
        """
        # Нормализуем сообщение
        normalized_message = cls._normalize_message(message)
        
        # Создаем строку для хеширования
        cache_string = normalized_message
        if context:
            cache_string += f"|{context}"
        
        # Генерируем хеш
        hash_object = hashlib.sha256(cache_string.encode())
        hash_hex = hash_object.hexdigest()[:16]  # Используем первые 16 символов
        
        return f"{cls.CACHE_PREFIX}:response:{hash_hex}"
    
    @classmethod
    def get_cached_response(cls, message: str, context: Optional[str] = None) -> Optional[str]:
        """
        Получает закешированный ответ для сообщения.
        
        Args:
            message: Сообщение пользователя
            context: Дополнительный контекст
            
        Returns:
            Закешированный ответ или None
        """
        if not REDIS_AVAILABLE:
            return None
        
        try:
            redis_client = get_redis_client()
            if not redis_client:
                return None
            
            cache_key = cls._generate_cache_key(message, context)
            cached_value = redis_client.get(cache_key)
            
            if cached_value:
                logger.debug(f"Cache HIT для сообщения: {message[:50]}...")
                # Десериализуем JSON
                cached_data = json.loads(cached_value)
                return cached_data.get('response')
            else:
                logger.debug(f"Cache MISS для сообщения: {message[:50]}...")
                return None
        
        except Exception as e:
            logger.warning(f"Ошибка при получении из кеша: {e}")
            return None
    
    @classmethod
    def cache_response(
        cls,
        message: str,
        response: str,
        context: Optional[str] = None,
        ttl: Optional[int] = None
    ) -> bool:
        """
        Кеширует ответ для сообщения.
        
        Args:
            message: Сообщение пользователя
            response: Ответ для кеширования
            context: Дополнительный контекст
            ttl: Время жизни кеша в секундах (по умолчанию DEFAULT_TTL)
            
        Returns:
            True если кеширование успешно
        """
        if not REDIS_AVAILABLE:
            return False
        
        try:
            redis_client = get_redis_client()
            if not redis_client:
                return False
            
            cache_key = cls._generate_cache_key(message, context)
            ttl = ttl or cls.DEFAULT_TTL
            
            # Сериализуем данные
            cache_data = {
                'message': message,
                'response': response,
                'context': context
            }
            cache_value = json.dumps(cache_data, ensure_ascii=False)
            
            # Сохраняем в Redis с TTL
            redis_client.setex(cache_key, ttl, cache_value)
            logger.debug(f"Закеширован ответ для: {message[:50]}... (TTL={ttl}s)")
            
            return True
        
        except Exception as e:
            logger.warning(f"Ошибка при кешировании: {e}")
            return False
    
    @classmethod
    def invalidate_cache(cls, message: Optional[str] = None) -> bool:
        """
        Инвалидирует кеш.
        
        Args:
            message: Конкретное сообщение для инвалидации. Если None - очищается весь кеш AI
            
        Returns:
            True если инвалидация успешна
        """
        if not REDIS_AVAILABLE:
            return False
        
        try:
            redis_client = get_redis_client()
            if not redis_client:
                return False
            
            if message:
                # Инвалидируем конкретное сообщение
                cache_key = cls._generate_cache_key(message)
                deleted = redis_client.delete(cache_key)
                logger.info(f"Инвалидирован кеш для сообщения: {message[:50]}...")
                return deleted > 0
            else:
                # Инвалидируем весь кеш AI агента
                pattern = f"{cls.CACHE_PREFIX}:response:*"
                keys = redis_client.keys(pattern)
                if keys:
                    deleted = redis_client.delete(*keys)
                    logger.info(f"Инвалидировано {deleted} записей кеша AI агента")
                    return True
                return False
        
        except Exception as e:
            logger.warning(f"Ошибка при инвалидации кеша: {e}")
            return False
    
    @classmethod
    def get_cache_stats(cls) -> dict:
        """
        Получает статистику кеша.
        
        Returns:
            Словарь со статистикой
        """
        if not REDIS_AVAILABLE:
            return {'enabled': False}
        
        try:
            redis_client = get_redis_client()
            if not redis_client:
                return {'enabled': False}
            
            pattern = f"{cls.CACHE_PREFIX}:response:*"
            keys = redis_client.keys(pattern)
            
            total_size = 0
            if keys:
                for key in keys:
                    try:
                        # Получаем размер значения
                        value = redis_client.get(key)
                        if value:
                            total_size += len(value)
                    except Exception:
                        pass
            
            return {
                'enabled': True,
                'total_keys': len(keys) if keys else 0,
                'total_size_bytes': total_size,
                'total_size_kb': round(total_size / 1024, 2)
            }
        
        except Exception as e:
            logger.warning(f"Ошибка при получении статистики кеша: {e}")
            return {'enabled': False, 'error': str(e)}


# Удобные функции для использования в коде

def get_cached_ai_response(message: str, context: Optional[str] = None) -> Optional[str]:
    """
    Удобная функция для получения закешированного ответа.
    
    Args:
        message: Сообщение пользователя
        context: Дополнительный контекст
        
    Returns:
        Закешированный ответ или None
    """
    return AICacheManager.get_cached_response(message, context)


def cache_ai_response(
    message: str,
    response: str,
    context: Optional[str] = None,
    ttl: Optional[int] = None
) -> bool:
    """
    Удобная функция для кеширования ответа.
    
    Args:
        message: Сообщение пользователя
        response: Ответ для кеширования
        context: Дополнительный контекст
        ttl: Время жизни кеша в секундах
        
    Returns:
        True если кеширование успешно
    """
    return AICacheManager.cache_response(message, response, context, ttl)


def invalidate_ai_cache(message: Optional[str] = None) -> bool:
    """
    Удобная функция для инвалидации кеша.
    
    Args:
        message: Конкретное сообщение или None для очистки всего кеша
        
    Returns:
        True если инвалидация успешна
    """
    return AICacheManager.invalidate_cache(message)


def cache_analytics_query(
    query_type: str,
    params: dict,
    result: str,
    ttl: Optional[int] = None
) -> bool:
    """
    Кеширует результат запроса к данным 1С.
    
    Args:
        query_type: Тип запроса (search, statistics, top, etc.)
        params: Параметры запроса
        result: Результат запроса
        ttl: Время жизни кеша в секундах
        
    Returns:
        True если кеширование успешно
    """
    if not REDIS_AVAILABLE:
        return False
    
    try:
        redis_client = get_redis_client()
        if not redis_client:
            return False
        
        # Создаем ключ кеша на основе типа запроса и параметров
        import hashlib
        import json
        cache_string = f"{query_type}:{json.dumps(params, sort_keys=True)}"
        hash_hex = hashlib.sha256(cache_string.encode()).hexdigest()[:16]
        cache_key = f"{AICacheManager.ANALYTICS_CACHE_PREFIX}:{hash_hex}"
        
        ttl = ttl or AICacheManager.DEFAULT_TTL
        cache_data = {
            'query_type': query_type,
            'params': params,
            'result': result
        }
        cache_value = json.dumps(cache_data, ensure_ascii=False)
        
        redis_client.setex(cache_key, ttl, cache_value)
        logger.debug(f"Закеширован результат аналитики: {query_type} (TTL={ttl}s)")
        return True
    except Exception as e:
        logger.warning(f"Ошибка при кешировании результата аналитики: {e}")
        return False


def get_cached_analytics_query(
    query_type: str,
    params: dict
) -> Optional[str]:
    """
    Получает закешированный результат запроса к данным 1С.
    
    Args:
        query_type: Тип запроса
        params: Параметры запроса
        
    Returns:
        Закешированный результат или None
    """
    if not REDIS_AVAILABLE:
        return None
    
    try:
        redis_client = get_redis_client()
        if not redis_client:
            return None
        
        import hashlib
        import json
        cache_string = f"{query_type}:{json.dumps(params, sort_keys=True)}"
        hash_hex = hashlib.sha256(cache_string.encode()).hexdigest()[:16]
        cache_key = f"{AICacheManager.ANALYTICS_CACHE_PREFIX}:{hash_hex}"
        
        cached_value = redis_client.get(cache_key)
        if cached_value:
            cached_data = json.loads(cached_value)
            logger.debug(f"Cache HIT для аналитики: {query_type}")
            return cached_data.get('result')
        else:
            logger.debug(f"Cache MISS для аналитики: {query_type}")
            return None
    except Exception as e:
        logger.warning(f"Ошибка при получении из кеша аналитики: {e}")
        return None


def cache_chart(
    chart_type: str,
    params: dict,
    chart_base64: str,
    ttl: Optional[int] = None
) -> bool:
    """
    Кеширует сгенерированный график.
    
    Args:
        chart_type: Тип графика (customers, materials, etc.)
        params: Параметры графика
        chart_base64: Base64 строка изображения
        ttl: Время жизни кеша в секундах (по умолчанию 7 дней для графиков)
        
    Returns:
        True если кеширование успешно
    """
    if not REDIS_AVAILABLE:
        return False
    
    try:
        redis_client = get_redis_client()
        if not redis_client:
            return False
        
        import hashlib
        import json
        cache_string = f"{chart_type}:{json.dumps(params, sort_keys=True)}"
        hash_hex = hashlib.sha256(cache_string.encode()).hexdigest()[:16]
        cache_key = f"{AICacheManager.CHART_CACHE_PREFIX}:{hash_hex}"
        
        # Для графиков используем более длинный TTL (7 дней)
        ttl = ttl or (7 * 24 * 60 * 60)
        cache_data = {
            'chart_type': chart_type,
            'params': params,
            'chart_base64': chart_base64
        }
        cache_value = json.dumps(cache_data, ensure_ascii=False)
        
        redis_client.setex(cache_key, ttl, cache_value)
        logger.debug(f"Закеширован график: {chart_type} (TTL={ttl}s)")
        return True
    except Exception as e:
        logger.warning(f"Ошибка при кешировании графика: {e}")
        return False


def get_cached_chart(
    chart_type: str,
    params: dict
) -> Optional[str]:
    """
    Получает закешированный график.
    
    Args:
        chart_type: Тип графика
        params: Параметры графика
        
    Returns:
        Base64 строка изображения или None
    """
    if not REDIS_AVAILABLE:
        return None
    
    try:
        redis_client = get_redis_client()
        if not redis_client:
            return None
        
        import hashlib
        import json
        cache_string = f"{chart_type}:{json.dumps(params, sort_keys=True)}"
        hash_hex = hashlib.sha256(cache_string.encode()).hexdigest()[:16]
        cache_key = f"{AICacheManager.CHART_CACHE_PREFIX}:{hash_hex}"
        
        cached_value = redis_client.get(cache_key)
        if cached_value:
            cached_data = json.loads(cached_value)
            logger.debug(f"Cache HIT для графика: {chart_type}")
            return cached_data.get('chart_base64')
        else:
            logger.debug(f"Cache MISS для графика: {chart_type}")
            return None
    except Exception as e:
        logger.warning(f"Ошибка при получении графика из кеша: {e}")
        return None


def invalidate_analytics_cache() -> bool:
    """
    Инвалидирует весь кеш запросов к данным 1С.
    
    Returns:
        True если инвалидация успешна
    """
    if not REDIS_AVAILABLE:
        return False
    
    try:
        redis_client = get_redis_client()
        if not redis_client:
            return False
        
        pattern = f"{AICacheManager.ANALYTICS_CACHE_PREFIX}:*"
        keys = redis_client.keys(pattern)
        if keys:
            deleted = redis_client.delete(*keys)
            logger.info(f"Инвалидировано {deleted} записей кеша аналитики")
            return True
        return False
    except Exception as e:
        logger.warning(f"Ошибка при инвалидации кеша аналитики: {e}")
        return False

