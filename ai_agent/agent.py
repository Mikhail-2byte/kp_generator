"""
Основной модуль AI агента с интеграцией OpenRouter API.
"""

import json
import logging
import re
import sys
import time
from typing import Any, Dict, List, Optional

import requests

from ai_agent.cache_manager import cache_ai_response, get_cached_ai_response
from ai_agent.config import (
    ConfigurationError,
    get_api_key,
    get_api_url,
    get_model_name,
    get_timeout,
    is_fallback_enabled,
    is_reasoning_enabled,
    is_usage_monitoring_enabled,
)
from ai_agent.knowledge_base import KnowledgeBase
from ai_agent.logistics_helper import LogisticsHelper
from ai_agent.materials_helper import MaterialsHelper
from ai_agent.duty_helper import DutyHelper

logger = logging.getLogger(__name__)


class APIError(Exception):
    """Базовая ошибка API."""
    def __init__(self, message: str, status_code: Optional[int] = None, retry_after: Optional[int] = None):
        super().__init__(message)
        self.status_code = status_code
        self.retry_after = retry_after


class APIKeyInvalidError(APIError):
    """Ошибка невалидного API ключа."""
    pass


class APIRateLimitError(APIError):
    """Ошибка превышения rate limit."""
    pass


class APIServerError(APIError):
    """Ошибка на стороне сервера API."""
    pass


class AIAgent:
    """AI агент для консультаций по проекту."""
    
    def __init__(self, api_key: Optional[str] = None, model: Optional[str] = None):
        """
        Инициализирует AI агента.
        
        Args:
            api_key: API ключ OpenRouter (если None, берется из config)
            model: Название модели (если None, берется из config)
            
        Raises:
            ConfigurationError: Если API ключ не установлен или невалиден
        """
        try:
            self.api_key = api_key or get_api_key()
        except ConfigurationError as e:
            logger.error(f"Ошибка конфигурации AI агента: {e}")
            raise
        
        self.api_url = get_api_url()
        self.model = model or get_model_name()
        self.reasoning_enabled = is_reasoning_enabled()
        self.timeout = get_timeout()
        self.fallback_enabled = is_fallback_enabled()
        self.usage_monitoring = is_usage_monitoring_enabled()
        
        # Инициализируем помощников
        self.knowledge_base = KnowledgeBase()
        self.logistics_helper = LogisticsHelper()
        self.materials_helper = MaterialsHelper()
        self.duty_helper = DutyHelper()
        
        # История диалога
        self.conversation_history: List[Dict[str, Any]] = []
        
        # Системный промпт
        self.system_prompt = self._build_system_prompt()
        
        logger.info(f"AI агент инициализирован с моделью {self.model}")
    
    def _build_system_prompt(self) -> str:
        """
        Формирует системный промпт для агента.
        
        Returns:
            Системный промпт
        """
        # Получаем краткий контекст документации
        docs_context = self.knowledge_base.get_all_context()
        
        prompt = """Ты - AI консультант и помощник для проекта KP Generator (генератор коммерческих предложений).

Твоя роль:
- Отвечать на вопросы по документации проекта (инструкции, руководства)
- Помогать с расчетом логистики
- Искать материалы в справочнике (gb_materials)
- Искать пошлины в справочниках (duty_rates, TNVED)
- Консультировать по использованию системы

Доступные функции:
1. Ответы на вопросы по документации - используй контекст из базы знаний
2. Расчет логистики - можешь вызывать функцию расчета для грузов
3. Поиск материалов - можешь искать по русскому названию, GB стандарту, ГОСТ
4. Поиск пошлин - можешь искать по товару, категории, коду ТН ВЭД

Правила работы с логистикой:
- Если пользователь просит рассчитать логистику, извлекай параметры:
  * Вес груза (в кг)
  * Город назначения
  * Тип транспорта (фура/трал, по умолчанию фура)
  * Габариты (опционально: длина, ширина, высота в мм)
- Если параметров недостаточно, уточни у пользователя
- После расчета логистики, представь результат в понятном формате

Правила работы с материалами:
- Если пользователь спрашивает о материалах (например: "найди материал 10", "какой GB у 15Г", "цена на сталь 20"), используй поиск по справочнику
- Ищи по русскому названию, GB стандарту или ГОСТ
- Показывай все варианты с разными типами заготовок (Лист, Прокат, Труба, Поковка)
- Если материал не рекомендуется к использованию, обязательно предупреди пользователя

Правила работы с пошлинами:
- Если пользователь спрашивает о пошлинах (например: "какая пошлина на сталь", "пошлина по коду 7214"), используй поиск по справочникам
- Ищи по названию товара, категории, коду ТН ВЭД или ключевым словам
- Показывай процент пошлины и источник (ручной справочник или ТН ВЭД)

Правила ответов:
- Отвечай на русском языке
- Будь кратким, но информативным
- ВСЕГДА используй контекст из документации для точных ответов
- Если в контексте есть конкретная информация (имена, фамилии, процедуры), обязательно используй её
- НЕ говори "не указано" или "не найдено", если информация есть в предоставленном контексте
- Если контекст содержит конкретные данные (например, кому отправлять документы), обязательно укажи их
- Если не знаешь ответа и его нет в контексте, честно скажи об этом

Контекст документации:
"""
        prompt += docs_context[:5000]  # Ограничиваем размер
        
        return prompt
    
    def _extract_intent(self, message: str) -> Dict[str, Any]:
        """
        Определяет намерение пользователя из сообщения.
        
        Args:
            message: Сообщение пользователя
            
        Returns:
            Словарь с информацией о намерении
        """
        message_lower = message.lower()
        
        # Ключевые слова для расчета логистики
        logistics_keywords = [
            'логистик', 'рассчитай', 'расчет', 'доставк', 'перевозк',
            'стоимость доставки', 'цена доставки', 'логистика для',
            'сколько стоит', 'стоимость перевозки'
        ]
        
        is_logistics = any(keyword in message_lower for keyword in logistics_keywords)
        
        # Попытка извлечь параметры логистики
        logistics_params = {}
        if is_logistics:
            # Вес - ищем числа перед словами "кг", "kg", "тонн", "т"
            # Используем более гибкий паттерн, который работает с кириллицей и латиницей
            weight = None
            is_tons = False
            
            # Паттерн для поиска веса: число (с точкой/запятой) + пробел (опционально) + единица измерения
            # Используем более широкий поиск
            weight_patterns = [
                (r'(\d+(?:[.,]\d+)?)\s*(?:кг|kg)', False),  # килограммы
                (r'(\d+(?:[.,]\d+)?)\s*(?:тонн|т)(?!\w)', True),  # тонны (не часть слова)
            ]
            
            for pattern, tons_flag in weight_patterns:
                try:
                    weight_match = re.search(pattern, message_lower, re.UNICODE)
                    if weight_match:
                        weight_str = weight_match.group(1).replace(',', '.')
                        try:
                            weight = float(weight_str)
                            is_tons = tons_flag
                            break
                        except ValueError:
                            continue
                except Exception:
                    continue
            
            # Если не нашли через регулярки, попробуем найти число рядом со словами "вес", "весом"
            if weight is None:
                # Ищем паттерн "вес[ом] ... число"
                weight_context_match = re.search(r'вес[ома]*\s+(\d+(?:[.,]\d+)?)', message_lower, re.UNICODE)
                if weight_context_match:
                    weight_str = weight_context_match.group(1).replace(',', '.')
                    try:
                        weight = float(weight_str)
                        # По умолчанию считаем килограммами, если не указано иное
                        if 'тонн' in message_lower or ('т' in message_lower and 'тонн' not in message_lower):
                            is_tons = True
                    except ValueError:
                        pass
            
            if weight is not None:
                # Если указано в тоннах, переводим в кг
                if is_tons:
                    weight *= 1000
                logistics_params['weight_kg'] = weight
            
            # Город (поиск по словам, включая сокращения)
            cities = [c.get('name') for c in self.logistics_helper._load_cities()]
            # Словарь сокращений
            city_aliases = {
                'екб': 'Екатеринбург',
                'мск': 'Москва',
                'спб': 'Санкт-Петербург',
                'питер': 'Санкт-Петербург',
            }
            
            # Сначала проверяем сокращения
            for alias, city_name in city_aliases.items():
                if alias in message_lower and city_name in cities:
                    logistics_params['city_name'] = city_name
                    break
            
            # Если не нашли по сокращению, ищем по полному названию
            if 'city_name' not in logistics_params:
                for city in cities:
                    city_lower = city.lower()
                    # Проверяем точное вхождение или начало слова
                    if city_lower in message_lower:
                        logistics_params['city_name'] = city
                        break
            
            # Тип транспорта
            if 'трал' in message_lower:
                logistics_params['transport_type'] = 'trail'
            else:
                logistics_params['transport_type'] = 'truck'
            
            # Габариты (опционально)
            dim_match = re.search(r'(\d+)\s*[xх×]\s*(\d+)\s*[xх×]\s*(\d+)\s*(?:мм|mm)', message_lower)
            if dim_match:
                logistics_params['length_mm'] = float(dim_match.group(1))
                logistics_params['width_mm'] = float(dim_match.group(2))
                logistics_params['height_mm'] = float(dim_match.group(3))
        
        # Ключевые слова для поиска материалов
        materials_keywords = [
            'материал', 'gb', 'гб', 'аналог', 'сталь', 'металл',
            'найди материал', 'какой gb', 'цена на', 'стоимость материала',
            'gost', 'гост', 'заготовк'
        ]
        
        is_materials = any(keyword in message_lower for keyword in materials_keywords)
        
        # Ключевые слова для поиска пошлин
        duty_keywords = [
            'пошлин', 'таможенн', 'код тн вэд', 'тнвэд', 'tnved',
            'какая пошлина', 'ставка пошлины', 'пошлина на', 'пошлина для'
        ]
        
        is_duty = any(keyword in message_lower for keyword in duty_keywords)
        
        # Извлекаем запрос для поиска материалов
        materials_query = None
        if is_materials:
            # Пытаемся извлечь название материала или GB стандарт
            # Ищем числа, которые могут быть названиями материалов (10, 15, 20 и т.д.)
            material_match = re.search(r'\b(\d+[А-Яа-я]*)\b', message)
            if material_match:
                materials_query = material_match.group(1)
            # Или ищем после ключевых слов
            for keyword in ['материал', 'gb', 'гб', 'сталь']:
                if keyword in message_lower:
                    # Берем следующее слово после ключевого
                    parts = message_lower.split(keyword)
                    if len(parts) > 1:
                        next_part = parts[1].strip().split()[0] if parts[1].strip() else None
                        if next_part:
                            materials_query = next_part
                            break
        
        # Извлекаем запрос для поиска пошлин
        duty_query = None
        if is_duty:
            # Ищем код ТН ВЭД (только цифры, 4-10 символов)
            code_match = re.search(r'\b(\d{4,10})\b', message)
            if code_match:
                duty_query = code_match.group(1)
            else:
                # Ищем название товара после ключевых слов
                for keyword in ['пошлина на', 'пошлина для', 'таможенн']:
                    if keyword in message_lower:
                        parts = message_lower.split(keyword)
                        if len(parts) > 1:
                            next_part = parts[1].strip().split()[0:3]  # Берем первые 3 слова
                            if next_part:
                                duty_query = ' '.join(next_part)
                                break
        
        return {
            'is_logistics': is_logistics,
            'logistics_params': logistics_params,
            'is_materials': is_materials,
            'materials_query': materials_query,
            'is_duty': is_duty,
            'duty_query': duty_query,
            'is_question': '?' in message or any(word in message_lower for word in ['как', 'что', 'почему', 'где', 'когда', 'зачем']),
        }
    
    def _call_api(self, messages: List[Dict[str, Any]], user_id: Optional[int] = None) -> Dict[str, Any]:
        """
        Вызывает API OpenRouter с улучшенной обработкой ошибок.
        
        Args:
            messages: Список сообщений для отправки
            user_id: ID пользователя для логирования
            
        Returns:
            Ответ от API
            
        Raises:
            APIKeyInvalidError: Невалидный API ключ
            APIRateLimitError: Превышен rate limit
            APIServerError: Ошибка на стороне сервера
            APIError: Другая ошибка API
        """
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }
        
        payload = {
            "model": self.model,
            "messages": messages,
        }
        
        if self.reasoning_enabled:
            payload["reasoning"] = {"enabled": True}
        
        start_time = time.time()
        error_type = None
        
        try:
            response = requests.post(
                url=self.api_url,
                headers=headers,
                data=json.dumps(payload),
                timeout=self.timeout
            )
            
            response_time_ms = int((time.time() - start_time) * 1000)
            
            # Обработка различных статус-кодов
            if response.status_code == 200:
                response_data = response.json()
                
                # Извлекаем информацию об использовании
                if self.usage_monitoring:
                    self._log_usage(
                        user_id=user_id,
                        response_data=response_data,
                        response_time_ms=response_time_ms,
                        error_type=None
                    )
                
                return response_data
            
            elif response.status_code == 401:
                error_type = "auth_error"
                error_msg = "API ключ недействителен или отозван. Обратитесь к администратору для обновления ключа."
                logger.error(f"API Error 401: {error_msg}")
                
                if self.usage_monitoring:
                    self._log_usage(user_id=user_id, response_time_ms=response_time_ms, error_type=error_type)
                
                raise APIKeyInvalidError(error_msg, status_code=401)
            
            elif response.status_code == 403:
                error_type = "forbidden"
                error_msg = "Недостаточно прав или баланса на аккаунте OpenRouter. Пополните баланс или проверьте права доступа."
                logger.error(f"API Error 403: {error_msg}")
                
                if self.usage_monitoring:
                    self._log_usage(user_id=user_id, response_time_ms=response_time_ms, error_type=error_type)
                
                raise APIKeyInvalidError(error_msg, status_code=403)
            
            elif response.status_code == 429:
                error_type = "rate_limit"
                # Пытаемся извлечь информацию о retry-after
                retry_after = response.headers.get('Retry-After')
                retry_seconds = int(retry_after) if retry_after and retry_after.isdigit() else 60
                
                error_msg = f"Превышен лимит запросов (rate limit). Попробуйте через {retry_seconds} секунд."
                logger.warning(f"API Error 429: {error_msg}")
                
                if self.usage_monitoring:
                    self._log_usage(user_id=user_id, response_time_ms=response_time_ms, error_type=error_type)
                
                raise APIRateLimitError(error_msg, status_code=429, retry_after=retry_seconds)
            
            elif response.status_code in [500, 502, 503, 504]:
                error_type = "server_error"
                error_msg = f"Ошибка на стороне сервера OpenRouter (статус {response.status_code}). Сервис временно недоступен."
                logger.error(f"API Error {response.status_code}: {error_msg}")
                
                if self.usage_monitoring:
                    self._log_usage(user_id=user_id, response_time_ms=response_time_ms, error_type=error_type)
                
                raise APIServerError(error_msg, status_code=response.status_code)
            
            else:
                error_type = "unknown_error"
                error_msg = f"Неожиданный статус ответа: {response.status_code}"
                logger.error(f"API Error {response.status_code}: {error_msg}")
                
                if self.usage_monitoring:
                    self._log_usage(user_id=user_id, response_time_ms=response_time_ms, error_type=error_type)
                
                raise APIError(error_msg, status_code=response.status_code)
        
        except requests.exceptions.Timeout:
            error_type = "timeout"
            response_time_ms = int((time.time() - start_time) * 1000)
            error_msg = f"Превышено время ожидания ({self.timeout} сек). Попробуйте упростить запрос."
            logger.error(f"API Timeout: {error_msg}")
            
            if self.usage_monitoring:
                self._log_usage(user_id=user_id, response_time_ms=response_time_ms, error_type=error_type)
            
            raise APIError(error_msg)
        
        except requests.exceptions.ConnectionError as e:
            error_type = "connection_error"
            response_time_ms = int((time.time() - start_time) * 1000)
            error_msg = "Не удалось подключиться к OpenRouter API. Проверьте подключение к интернету."
            logger.error(f"API Connection Error: {error_msg} - {str(e)}")
            
            if self.usage_monitoring:
                self._log_usage(user_id=user_id, response_time_ms=response_time_ms, error_type=error_type)
            
            raise APIError(error_msg)
        
        except (APIKeyInvalidError, APIRateLimitError, APIServerError, APIError):
            # Перебрасываем наши кастомные ошибки как есть
            raise
        
        except Exception as e:
            error_type = "unexpected_error"
            response_time_ms = int((time.time() - start_time) * 1000)
            error_msg = f"Неожиданная ошибка при обращении к API: {str(e)}"
            logger.exception(f"Unexpected API Error: {error_msg}")
            
            if self.usage_monitoring:
                self._log_usage(user_id=user_id, response_time_ms=response_time_ms, error_type=error_type)
            
            raise APIError(error_msg)
    
    def _log_usage(
        self,
        user_id: Optional[int] = None,
        response_data: Optional[Dict] = None,
        response_time_ms: Optional[int] = None,
        error_type: Optional[str] = None,
        user_message: Optional[str] = None
    ) -> None:
        """
        Логирует использование API.
        
        Args:
            user_id: ID пользователя
            response_data: Данные ответа от API
            response_time_ms: Время ответа в миллисекундах
            error_type: Тип ошибки (если была)
            user_message: Сообщение пользователя
        """
        try:
            from ai_agent.usage_monitor import log_ai_usage
            
            # Извлекаем токены и стоимость из ответа
            request_tokens = None
            response_tokens = None
            total_cost = None
            
            if response_data and 'usage' in response_data:
                usage = response_data['usage']
                request_tokens = usage.get('prompt_tokens')
                response_tokens = usage.get('completion_tokens')
                total_cost = usage.get('total_cost')  # Некоторые провайдеры возвращают стоимость
            
            log_ai_usage(
                user_id=user_id,
                request_tokens=request_tokens,
                response_tokens=response_tokens,
                total_cost=total_cost,
                response_time_ms=response_time_ms,
                model_name=self.model,
                error_type=error_type,
                user_message=user_message,
                cache_hit=False
            )
        except Exception as e:
            # Не прерываем работу если логирование не удалось
            logger.warning(f"Не удалось залогировать использование API: {e}")
    
    def _fallback_search(self, message: str) -> str:
        """
        Простой поиск по документации без AI (fallback режим).
        
        Args:
            message: Сообщение пользователя
            
        Returns:
            Результат поиска или сообщение о недоступности
        """
        logger.info("Используется fallback режим (поиск без AI)")
        
        # Получаем релевантный контекст из базы знаний
        try:
            relevant_context = self.knowledge_base.get_relevant_context(message)
            
            if relevant_context:
                return f"""🔍 **Результаты поиска по документации** (AI временно недоступен):

{relevant_context}

---
💡 Если вам нужна более точная информация, попробуйте переформулировать вопрос или обратитесь к администратору."""
            else:
                return """К сожалению, AI агент временно недоступен, и по вашему запросу не найдено информации в документации.

Вы можете:
• Переформулировать вопрос более конкретно
• Обратиться к руководству пользователя
• Связаться с администратором системы

Приносим извинения за неудобства."""
        
        except Exception as e:
            logger.error(f"Ошибка в fallback режиме: {e}")
            return """AI агент временно недоступен.

Пожалуйста, обратитесь к администратору системы или попробуйте позже.

Приносим извинения за неудобства."""
    
    def chat(self, message: str, context: Optional[Dict[str, Any]] = None, user_id: Optional[int] = None) -> str:
        """
        Основной метод для общения с агентом с поддержкой кеширования.
        
        Args:
            message: Сообщение пользователя
            context: Дополнительный контекст
            user_id: ID пользователя (для логирования)
            
        Returns:
            Ответ агента
        """
        # Проверяем кеш для обычных вопросов (не логистика)
        intent = self._extract_intent(message)
        # Логистика считается запросом, если есть вес ИЛИ есть ключевые слова логистики
        is_logistics_query = intent['is_logistics'] and (
            intent['logistics_params'].get('weight_kg') is not None or
            intent['logistics_params'].get('city_name')
        )
        
        if not is_logistics_query:
            # Пытаемся получить из кеша
            cached_response = get_cached_ai_response(message)
            if cached_response:
                logger.info("Используется закешированный ответ")
                
                # Добавляем в историю
                self.conversation_history.append({"role": "user", "content": message})
                self.conversation_history.append({"role": "assistant", "content": cached_response})
                
                # Логируем использование с пометкой cache_hit
                if self.usage_monitoring:
                    try:
                        from ai_agent.usage_monitor import log_ai_usage
                        log_ai_usage(
                            user_id=user_id,
                            user_message=message,
                            cache_hit=True,
                            model_name=self.model
                        )
                    except Exception as e:
                        logger.warning(f"Не удалось залогировать cache hit: {e}")
                
                return cached_response
        
        # Если это запрос на поиск материалов
        if intent['is_materials']:
            query = intent.get('materials_query') or message
            # Убираем ключевые слова из запроса для более точного поиска
            query_clean = re.sub(r'\b(материал|gb|гб|аналог|сталь|металл|найди|какой|цена|стоимость)\b', '', query.lower()).strip()
            if not query_clean:
                query_clean = query
            
            # Определяем тип поиска
            search_field = None
            message_lower = message.lower()
            if 'gb' in message_lower or 'гб' in message_lower:
                search_field = 'gb'
            elif 'gost' in message_lower or 'гост' in message_lower:
                search_field = 'gost'
            elif any(char.isdigit() for char in query_clean):
                # Если есть цифры, ищем по русскому названию
                search_field = 'russian'
            
            materials = self.materials_helper.find_material(query_clean, search_field)
            formatted_result = self.materials_helper.format_material_result(materials)
            
            # Добавляем в историю
            self.conversation_history.append({"role": "user", "content": message})
            self.conversation_history.append({"role": "assistant", "content": formatted_result})
            
            return formatted_result
        
        # Если это запрос на поиск пошлин
        if intent['is_duty']:
            query = intent.get('duty_query') or message
            # Убираем ключевые слова из запроса
            query_clean = re.sub(r'\b(пошлин|таможенн|код|тн вэд|тнвэд|tnved|какая|ставка|на|для)\b', '', query.lower()).strip()
            if not query_clean:
                query_clean = query
            
            # Если запрос - только цифры, это код ТН ВЭД
            if query_clean.strip().isdigit():
                duties = self.duty_helper.search_by_code(query_clean.strip())
            else:
                duties = self.duty_helper.find_duty(query_clean)
            
            formatted_result = self.duty_helper.format_duty_result(duties)
            
            # Добавляем в историю
            self.conversation_history.append({"role": "user", "content": message})
            self.conversation_history.append({"role": "assistant", "content": formatted_result})
            
            return formatted_result
        
        # Если это запрос на расчет логистики
        if is_logistics_query:
            params = intent['logistics_params']
            weight_kg = params.get('weight_kg')
            city_name = params.get('city_name', '')
            
            # Проверяем, что вес указан
            if weight_kg is None or weight_kg <= 0:
                response = "Для расчета логистики необходимо указать вес груза.\n\n"
                response += "Пожалуйста, укажите вес груза, например:\n"
                response += "• \"450 кг\"\n"
                response += "• \"1.5 тонн\"\n"
                response += "• \"500кг\""
                
                self.conversation_history.append({"role": "user", "content": message})
                self.conversation_history.append({"role": "assistant", "content": response})
                return response
            
            # Проверяем, что город указан
            if not city_name:
                response = "Для расчета логистики необходимо указать город назначения.\n\n"
                response += "Пожалуйста, укажите город, например:\n"
                response += "• \"Москва\"\n"
                response += "• \"Екатеринбург\"\n"
                response += "• \"Санкт-Петербург\""
                
                self.conversation_history.append({"role": "user", "content": message})
                self.conversation_history.append({"role": "assistant", "content": response})
                return response
            
            result = self.logistics_helper.calculate_simple_logistics(
                weight_kg=weight_kg,
                city_name=city_name,
                transport_type=params.get('transport_type', 'truck'),
                length_mm=params.get('length_mm'),
                width_mm=params.get('width_mm'),
                height_mm=params.get('height_mm'),
            )
            
            # Форматируем результат
            formatted_result = self.logistics_helper.format_logistics_result(result)
            
            # Если есть ошибка, возвращаем сразу
            if 'error' in result:
                # Добавляем в историю для контекста
                self.conversation_history.append({
                    "role": "user",
                    "content": message
                })
                self.conversation_history.append({
                    "role": "assistant",
                    "content": formatted_result
                })
                return formatted_result
            
            # Для успешного расчета возвращаем отформатированный результат напрямую
            self.conversation_history.append({
                "role": "user",
                "content": message
            })
            self.conversation_history.append({
                "role": "assistant",
                "content": formatted_result
            })
            
            return formatted_result
        
        # Обычный вопрос - используем AI с контекстом
        # Добавляем в историю
        self.conversation_history.append({
            "role": "user",
            "content": message
        })
        
        # Получаем релевантный контекст
        relevant_context = self.knowledge_base.get_relevant_context(message)
        
        # Формируем сообщения для API
        api_messages = [
            {"role": "system", "content": self.system_prompt},
        ]
        
        # Добавляем релевантный контекст, если есть
        if relevant_context:
            user_content = f"""ВАЖНО: Используй ТОЧНУЮ информацию из предоставленного контекста ниже. Если в контексте указаны конкретные имена, фамилии, процедуры или данные - обязательно используй их в ответе.

Контекст из документации:
{relevant_context}

Вопрос пользователя: {message}

Инструкция: Ответь на вопрос, используя информацию из контекста выше. Если в контексте есть конкретные данные (имена, фамилии, процедуры), обязательно укажи их точно."""
        else:
            user_content = message
        
        api_messages.append({
            "role": "user",
            "content": user_content
        })
        
        # Добавляем историю (последние 5 сообщений)
        for hist_msg in self.conversation_history[-5:]:
            msg_for_api = {
                "role": hist_msg["role"],
                "content": hist_msg["content"]
            }
            # Добавляем reasoning_details только для assistant сообщений
            if hist_msg["role"] == "assistant" and "reasoning_details" in hist_msg:
                msg_for_api["reasoning_details"] = hist_msg["reasoning_details"]
            api_messages.append(msg_for_api)
        
        try:
            # Вызываем API с обработкой ошибок
            response_data = self._call_api(api_messages, user_id=user_id)
            assistant_message = response_data['choices'][0]['message']
            content = assistant_message.get('content', '')
            
            # Сохраняем ответ в историю
            assistant_msg = {
                "role": "assistant",
                "content": content
            }
            
            # Сохраняем reasoning_details если есть
            if 'reasoning_details' in assistant_message:
                assistant_msg["reasoning_details"] = assistant_message.get('reasoning_details')
            
            self.conversation_history.append(assistant_msg)
            
            # Кешируем ответ для будущих запросов
            try:
                cache_ai_response(message, content)
                logger.debug("Ответ успешно закеширован")
            except Exception as e:
                logger.warning(f"Не удалось закешировать ответ: {e}")
            
            return content
        
        except APIKeyInvalidError as e:
            logger.error(f"API ключ невалиден: {e}")
            error_response = f"❌ **Ошибка аутентификации**: {str(e)}"
            
            # Используем fallback если включен
            if self.fallback_enabled:
                error_response += "\n\n" + self._fallback_search(message)
            
            self.conversation_history.append({
                "role": "assistant",
                "content": error_response
            })
            return error_response
        
        except APIRateLimitError as e:
            logger.warning(f"Rate limit превышен: {e}")
            error_response = f"⏱️ **Превышен лимит запросов**: {str(e)}"
            
            # Используем fallback если включен
            if self.fallback_enabled:
                error_response += "\n\n" + self._fallback_search(message)
            
            self.conversation_history.append({
                "role": "assistant",
                "content": error_response
            })
            return error_response
        
        except APIServerError as e:
            logger.error(f"Ошибка сервера API: {e}")
            error_response = f"🔧 **Сервис временно недоступен**: {str(e)}"
            
            # Используем fallback если включен
            if self.fallback_enabled:
                error_response += "\n\n" + self._fallback_search(message)
            
            self.conversation_history.append({
                "role": "assistant",
                "content": error_response
            })
            return error_response
        
        except APIError as e:
            logger.error(f"Ошибка API: {e}")
            error_response = f"⚠️ **Ошибка при обращении к AI**: {str(e)}"
            
            # Используем fallback если включен
            if self.fallback_enabled:
                error_response += "\n\n" + self._fallback_search(message)
            
            self.conversation_history.append({
                "role": "assistant",
                "content": error_response
            })
            return error_response
        
        except Exception as e:
            logger.exception(f"Неожиданная ошибка в chat(): {e}")
            error_response = f"❌ Произошла неожиданная ошибка. Пожалуйста, попробуйте позже."
            
            # Используем fallback если включен
            if self.fallback_enabled:
                error_response += "\n\n" + self._fallback_search(message)
            
            self.conversation_history.append({
                "role": "assistant",
                "content": error_response
            })
            return error_response
    
    def clear_history(self) -> None:
        """Очищает историю диалога."""
        self.conversation_history = []
        logger.debug("История диалога очищена")

