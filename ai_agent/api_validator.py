"""
Валидатор API ключа OpenRouter с кешированием статуса.
"""

import logging
import time
from typing import Dict, Optional, Tuple

import requests

from ai_agent.config import get_api_url

logger = logging.getLogger(__name__)


class APIValidationResult:
    """Результат валидации API ключа."""
    
    def __init__(
        self,
        is_valid: bool,
        error_message: Optional[str] = None,
        rate_limit_info: Optional[Dict[str, any]] = None
    ):
        self.is_valid = is_valid
        self.error_message = error_message
        self.rate_limit_info = rate_limit_info or {}
    
    def __bool__(self):
        return self.is_valid
    
    def __repr__(self):
        return f"APIValidationResult(valid={self.is_valid}, error={self.error_message})"


class APIKeyValidator:
    """Валидатор API ключа OpenRouter с кешированием."""
    
    # Кеш результатов валидации: {api_key: (timestamp, result)}
    _validation_cache: Dict[str, Tuple[float, APIValidationResult]] = {}
    
    # Время жизни кеша в секундах (5 минут)
    CACHE_TTL = 300
    
    @classmethod
    def validate_key(cls, api_key: str, force_check: bool = False) -> APIValidationResult:
        """
        Валидирует API ключ через тестовый запрос к OpenRouter.
        
        Args:
            api_key: API ключ для проверки
            force_check: Принудительная проверка (игнорировать кеш)
            
        Returns:
            APIValidationResult с результатом валидации
        """
        # Проверяем кеш, если не force_check
        if not force_check:
            cached_result = cls._get_from_cache(api_key)
            if cached_result:
                logger.debug("Используем закешированный результат валидации ключа")
                return cached_result
        
        logger.info("Выполняем тестовый запрос для валидации API ключа...")
        
        # Выполняем минимальный запрос к API
        headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        }
        
        # Минимальный payload для проверки ключа
        payload = {
            "model": "xiaomi/mimo-v2-flash:free",
            "messages": [
                {"role": "user", "content": "test"}
            ],
            "max_tokens": 1  # Минимальный запрос
        }
        
        try:
            response = requests.post(
                url=get_api_url(),
                headers=headers,
                json=payload,
                timeout=10
            )
            
            # Извлекаем информацию о rate limits из headers
            rate_limit_info = cls._extract_rate_limit_info(response.headers)
            
            if response.status_code == 200:
                result = APIValidationResult(
                    is_valid=True,
                    rate_limit_info=rate_limit_info
                )
                logger.info("API ключ валиден")
            elif response.status_code == 401:
                result = APIValidationResult(
                    is_valid=False,
                    error_message="API ключ недействителен или отозван"
                )
                logger.error("API ключ недействителен (401)")
            elif response.status_code == 403:
                result = APIValidationResult(
                    is_valid=False,
                    error_message="Недостаточно прав или баланса на аккаунте"
                )
                logger.error("Недостаточно прав или баланса (403)")
            elif response.status_code == 429:
                result = APIValidationResult(
                    is_valid=True,  # Ключ валиден, но превышен лимит
                    error_message="Превышен лимит запросов (rate limit)",
                    rate_limit_info=rate_limit_info
                )
                logger.warning("Превышен rate limit (429)")
            else:
                result = APIValidationResult(
                    is_valid=False,
                    error_message=f"Неожиданный статус: {response.status_code}"
                )
                logger.error(f"Неожиданный статус при валидации: {response.status_code}")
            
            # Кешируем результат
            cls._save_to_cache(api_key, result)
            return result
            
        except requests.exceptions.Timeout:
            result = APIValidationResult(
                is_valid=False,
                error_message="Превышено время ожидания при проверке ключа"
            )
            logger.error("Timeout при валидации API ключа")
            return result
        
        except requests.exceptions.ConnectionError:
            result = APIValidationResult(
                is_valid=False,
                error_message="Не удалось подключиться к OpenRouter API"
            )
            logger.error("Connection error при валидации API ключа")
            return result
        
        except Exception as e:
            result = APIValidationResult(
                is_valid=False,
                error_message=f"Ошибка при валидации: {str(e)}"
            )
            logger.exception("Неожиданная ошибка при валидации API ключа")
            return result
    
    @classmethod
    def _extract_rate_limit_info(cls, headers: Dict[str, str]) -> Dict[str, any]:
        """
        Извлекает информацию о rate limits из headers ответа.
        
        Args:
            headers: HTTP headers из ответа
            
        Returns:
            Словарь с информацией о лимитах
        """
        rate_limit_info = {}
        
        # OpenRouter может возвращать различные headers для rate limits
        rate_limit_headers = [
            'x-ratelimit-limit',
            'x-ratelimit-remaining',
            'x-ratelimit-reset',
            'x-ratelimit-limit-requests',
            'x-ratelimit-remaining-requests',
            'x-ratelimit-limit-tokens',
            'x-ratelimit-remaining-tokens',
        ]
        
        for header in rate_limit_headers:
            value = headers.get(header)
            if value:
                try:
                    # Пытаемся преобразовать в число
                    rate_limit_info[header] = int(value)
                except ValueError:
                    rate_limit_info[header] = value
        
        return rate_limit_info
    
    @classmethod
    def _get_from_cache(cls, api_key: str) -> Optional[APIValidationResult]:
        """
        Получает результат валидации из кеша.
        
        Args:
            api_key: API ключ
            
        Returns:
            Закешированный результат или None
        """
        if api_key not in cls._validation_cache:
            return None
        
        timestamp, result = cls._validation_cache[api_key]
        
        # Проверяем, не истек ли кеш
        if time.time() - timestamp > cls.CACHE_TTL:
            del cls._validation_cache[api_key]
            return None
        
        return result
    
    @classmethod
    def _save_to_cache(cls, api_key: str, result: APIValidationResult) -> None:
        """
        Сохраняет результат валидации в кеш.
        
        Args:
            api_key: API ключ
            result: Результат валидации
        """
        cls._validation_cache[api_key] = (time.time(), result)
    
    @classmethod
    def clear_cache(cls, api_key: Optional[str] = None) -> None:
        """
        Очищает кеш валидации.
        
        Args:
            api_key: Конкретный ключ для очистки. Если None - очищается весь кеш
        """
        if api_key:
            cls._validation_cache.pop(api_key, None)
            logger.debug("Очищен кеш валидации для ключа")
        else:
            cls._validation_cache.clear()
            logger.debug("Очищен весь кеш валидации")


def validate_api_key(api_key: str, force_check: bool = False) -> APIValidationResult:
    """
    Удобная функция для валидации API ключа.
    
    Args:
        api_key: API ключ для проверки
        force_check: Принудительная проверка без кеша
        
    Returns:
        APIValidationResult с результатом валидации
    """
    return APIKeyValidator.validate_key(api_key, force_check=force_check)

