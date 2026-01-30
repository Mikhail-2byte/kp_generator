"""Модуль для работы с Redis сессиями Flask."""

import logging
from typing import Any, Dict, Optional

from flask import Flask

logger = logging.getLogger(__name__)


def _wrap_session_interface_for_werkzeug3(app: Flask) -> None:
    """
    Оборачивает session_interface.save_session так, чтобы значение cookie
    всегда передавалось в set_cookie как str (Werkzeug 3.x не принимает bytes).
    Flask-Session с use_signer=True передаёт signer.sign() -> bytes.
    """
    original_interface = app.session_interface
    original_save = original_interface.save_session

    def save_session(app: Flask, session: Any, response: Any) -> None:
        original_set_cookie = response.set_cookie

        def set_cookie_str(name: str, value: Any, *args: Any, **kwargs: Any) -> None:
            if isinstance(value, bytes):
                value = value.decode("utf-8")
            original_set_cookie(name, value, *args, **kwargs)

        response.set_cookie = set_cookie_str
        try:
            original_save(app, session, response)
        finally:
            response.set_cookie = original_set_cookie

    original_interface.save_session = save_session

# Пытаемся импортировать redis, если не установлен - будет использован fallback
try:
    import redis
    REDIS_AVAILABLE = True
except ImportError:
    REDIS_AVAILABLE = False
    logger.warning('Модуль redis не установлен. Redis сессии недоступны.')


def init_redis_connection(config: Dict[str, Any]) -> Optional[Any]:
    """
    Инициализирует подключение к Redis.
    
    Args:
        config: Словарь конфигурации с настройками Redis
        
    Returns:
        Экземпляр Redis клиента или None при ошибке подключения
    """
    if not REDIS_AVAILABLE:
        return None
        
    redis_config = config.get('redis', {})
    
    # Получаем параметры подключения из конфига или переменных окружения
    import os
    host = os.environ.get('REDIS_HOST') or redis_config.get('host', 'localhost')
    port = int(os.environ.get('REDIS_PORT') or redis_config.get('port', 6379))
    password = os.environ.get('REDIS_PASSWORD') or redis_config.get('password')
    db = int(os.environ.get('REDIS_DB') or redis_config.get('db', 0))
    
    socket_timeout = redis_config.get('socket_timeout', 5)
    socket_connect_timeout = redis_config.get('socket_connect_timeout', 5)
    retry_on_timeout = redis_config.get('retry_on_timeout', True)
    
    try:
        redis_client = redis.Redis(
            host=host,
            port=port,
            password=password,
            db=db,
            socket_timeout=socket_timeout,
            socket_connect_timeout=socket_connect_timeout,
            retry_on_timeout=retry_on_timeout,
            decode_responses=False  # Данные сессии — pickle (binary); cookie приводим к str в _wrap_session_interface_for_werkzeug3
        )
        
        # Проверяем подключение
        redis_client.ping()
        logger.info(f'Redis подключен успешно: {host}:{port}/{db}')
        return redis_client
        
    except (redis.ConnectionError, redis.TimeoutError, Exception) as exc:
        logger.warning(f'Не удалось подключиться к Redis: {exc}. Используется fallback на cookie-based сессии.')
        return None


def setup_redis_session(app: Flask, config: Dict[str, Any]) -> bool:
    """
    Настраивает Flask-Session с Redis бэкендом.
    
    Args:
        app: Экземпляр Flask приложения
        config: Словарь конфигурации
        
    Returns:
        True если Redis сессии настроены успешно, False иначе (будет использован fallback)
    """
    if not REDIS_AVAILABLE:
        logger.warning('Redis недоступен (модуль не установлен). Используется fallback на cookie-based сессии.')
        return False
        
    try:
        from flask_session import Session
        
        redis_client = init_redis_connection(config)
        
        if redis_client is None:
            return False
        
        # Настройка Flask-Session
        session_config = config.get('session', {})
        app.config['SESSION_TYPE'] = 'redis'
        app.config['SESSION_REDIS'] = redis_client
        app.config['SESSION_PERMANENT'] = session_config.get('permanent', True)
        app.config['SESSION_USE_SIGNER'] = True  # Подписываем сессии для безопасности
        app.config['SESSION_KEY_PREFIX'] = 'session:'
        
        # Время жизни сессии в секундах (по умолчанию 24 часа)
        lifetime = session_config.get('lifetime', 86400)
        app.config['PERMANENT_SESSION_LIFETIME'] = lifetime
        
        # Инициализируем Flask-Session
        Session(app)
        
        # Обёртка: Flask-Session с use_signer=True передаёт в set_cookie bytes
        # (signer.sign возвращает bytes), а Werkzeug 3.x требует str
        _wrap_session_interface_for_werkzeug3(app)
        
        logger.info('Flask-Session с Redis бэкендом настроен успешно')
        return True
        
    except ImportError:
        logger.error('Flask-Session не установлен. Установите через: pip install Flask-Session')
        return False
    except Exception as exc:
        logger.error(f'Ошибка при настройке Redis сессий: {exc}')
        return False

