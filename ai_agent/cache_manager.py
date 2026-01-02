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
                    except:
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

