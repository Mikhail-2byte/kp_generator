"""
Основной модуль AI агента с интеграцией OpenRouter API.
"""

import json
import logging
import re
import time
from typing import Any, Dict, List, Optional

import pandas as pd
import requests

from ai_agent.cache_manager import (
    cache_ai_response, 
    get_cached_ai_response,
    cache_analytics_query,
    get_cached_analytics_query,
    cache_chart,
    get_cached_chart,
    invalidate_analytics_cache
)
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
from ai_agent.analytics_helper import AnalyticsHelper
from ai_agent.duty_helper import DutyHelper
from ai_agent.exporters import DataExporter
from ai_agent.feedback import get_feedback_collector
from ai_agent.formatters import ResponseFormatter
from ai_agent.metrics import get_metrics_collector
from ai_agent.suggestions import SuggestionsGenerator
from ai_agent.intent_extractor import IntentExtractor
from ai_agent.knowledge_base import KnowledgeBase
from ai_agent.logistics_helper import LogisticsHelper
from ai_agent.materials_helper import MaterialsHelper
from ai_agent.validators import ParameterValidator

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
        self.analytics_helper = AnalyticsHelper()
        self.exporter = DataExporter()
        
        # Инициализируем извлекатель намерений
        self.intent_extractor = IntentExtractor(
            self.logistics_helper,
            self.analytics_helper
        )
        
        # Инициализируем валидатор параметров
        self.validator = ParameterValidator()
        
        # Инициализируем сборщик метрик
        self.metrics_collector = get_metrics_collector()
        
        # Инициализируем сборщик обратной связи
        self.feedback_collector = get_feedback_collector()
        
        # История диалога
        self.conversation_history: List[Dict[str, Any]] = []
        
        # Системный промпт
        self.system_prompt = self._build_system_prompt()
        
        logger.info(f"AI агент инициализирован с моделью {self.model}")
    
    def _build_system_prompt(self, query_context: Optional[str] = None) -> str:
        """
        Формирует системный промпт для агента.
        
        Args:
            query_context: Контекст запроса для динамической загрузки релевантной документации
        
        Returns:
            Системный промпт
        """
        # Базовый промпт без документации (всегда одинаковый)
        base_prompt = """Ты - AI консультант и помощник для проекта KP Generator (генератор коммерческих предложений).

Твоя роль:
- Отвечать на вопросы по документации проекта (инструкции, руководства)
- Помогать с расчетом логистики
- Искать материалы в справочнике (gb_materials)
- Искать пошлины в справочниках (duty_rates, TNVED)
- Поиск по данным из 1С (номенклатура, чертежи, заказчики, материалы)
- Консультировать по использованию системы

Доступные функции:
1. Ответы на вопросы по документации - используй контекст из базы знаний
2. Расчет логистики - можешь вызывать функцию расчета для грузов
3. Поиск материалов - можешь искать по русскому названию, GB стандарту, ГОСТ
4. Поиск пошлин - можешь искать по товару, категории, коду ТН ВЭД
5. Поиск по данным 1С - можешь искать по номенклатуре, чертежу, заказчику, материалу

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

Правила работы с данными 1С:
- Если пользователь просит найти данные из 1С (например: "найди в 1с", "покажи данные по", "найди чертеж X"), используй поиск по данным 1С
- Доступны категории: Барабаны, Вал (и другие, если есть)
- Можно искать по:
  * Номенклатуре (например: "найди барабан самоцентрирующийся")
  * Чертежу (например: "найди чертеж 35.1774.000 СБ")
  * Заказчику (например: "покажи все записи по ВАЙЛДБЕРРИЗ")
  * Материалу (например: "найди все позиции с материалом СБ")
  * Статусу тендера (например: "покажи выигранные тендеры")
- Топ-листы и рейтинги:
  * Топ заказчиков, материалов, менеджеров
  * Топ менеджеров по выигранным тендерам (по количеству или по сумме)
  * Самый большой выигранный тендер
- Статистика:
  * Общая статистика по записям
  * Статистика по тендерам (выиграно, проиграно, отменено и т.д.)
- При большом результате (>5 записей) или явном запросе график генерируется автоматически
- Для запросов статистики ("сколько записей", "топ заказчиков") показывай сводную информацию
- Доступны различные типы графиков:
  * Обычные графики: топ заказчиков, материалы, менеджеры, цены
  * Временные ряды: "динамика тендеров по месяцам", "график по времени"
  * Scatter plots (корреляции): "корреляция вес vs цена", "зависимость веса и цены"
  * Heatmaps: "heatmap заказчик менеджер", "тепловая карта"

Правила ответов:
- Отвечай на русском языке
- Будь кратким, но информативным
- ВСЕГДА используй контекст из документации для точных ответов
- Если в контексте есть конкретная информация (имена, фамилии, процедуры), обязательно используй её
- НЕ говори "не указано" или "не найдено", если информация есть в предоставленном контексте
- Если контекст содержит конкретные данные (например, кому отправлять документы), обязательно укажи их
- Если не знаешь ответа и его нет в контексте, честно скажи об этом

"""
        
        # Динамически загружаем релевантный контекст документации
        if query_context:
            # Используем релевантный контекст вместо всей документации
            docs_context = self.knowledge_base.get_relevant_context(query_context, max_results=3)
            if docs_context:
                base_prompt += "\nКонтекст документации (релевантный для текущего запроса):\n"
                base_prompt += docs_context
            else:
                # Fallback: берем первые 2000 символов из всей документации
                all_docs = self.knowledge_base.get_all_context()
                base_prompt += "\nКонтекст документации:\n"
                base_prompt += all_docs[:2000]
        else:
            # Если контекста нет, берем краткую версию всей документации
            docs_context = self.knowledge_base.get_all_context()
            base_prompt += "\nКонтекст документации:\n"
            base_prompt += docs_context[:3000]  # Уменьшили с 5000 до 3000
        
        return base_prompt
    
    def _extract_intent(self, message: str) -> Dict[str, Any]:
        """
        Определяет намерение пользователя из сообщения.
        Использует улучшенный многоуровневый извлекатель намерений.
        
        Args:
            message: Сообщение пользователя
            
        Returns:
            Словарь с информацией о намерении
        """
        # Используем новый извлекатель намерений
        intent_result = self.intent_extractor.extract(message)
        
        # Сохраняем обратную совместимость со старым форматом
        return intent_result
    
    def _call_api(self, messages: List[Dict[str, Any]], user_id: Optional[int] = None) -> Dict[str, Any]:
        is_logistics = any(keyword in message_lower for keyword in logistics_keywords)
        
        # Попытка извлечь параметры логистики
        logistics_params = {}
        if is_logistics:
            # Вес - ищем числа перед словами "кг", "kg", "тонн", "т"
            # Используем более гибкий паттерн, который работает с кириллицей и латиницей
            weight = None
            is_tons = False
            
            def parse_weight_number(weight_str: str) -> float:
                """
                Парсит строку с весом, правильно обрабатывая разделители тысяч.
                
                Args:
                    weight_str: Строка с числом (может содержать пробелы, запятые, точки)
                    
                Returns:
                    Число в виде float
                """
                # Убираем пробелы по краям
                weight_str = weight_str.strip()
                
                # Убираем пробелы внутри числа (разделители тысяч)
                # Пробелы между цифрами считаем разделителями тысяч
                weight_str = re.sub(r'(\d)\s+(\d)', r'\1\2', weight_str)
                
                # Если есть запятая и точка, определяем что есть что
                # Обычно: запятая = разделитель тысяч, точка = десятичный разделитель
                # Или наоборот: точка = разделитель тысяч, запятая = десятичный разделитель
                
                # Сначала проверяем, есть ли и запятая, и точка
                has_comma = ',' in weight_str
                has_dot = '.' in weight_str
                
                if has_comma and has_dot:
                    # Определяем по позиции: что идет последним, то и есть десятичный разделитель
                    comma_pos = weight_str.rfind(',')
                    dot_pos = weight_str.rfind('.')
                    if comma_pos > dot_pos:
                        # Запятая - десятичный разделитель, точка - разделитель тысяч
                        weight_str = weight_str.replace('.', '').replace(',', '.')
                    else:
                        # Точка - десятичный разделитель, запятая - разделитель тысяч
                        weight_str = weight_str.replace(',', '')
                elif has_comma:
                    # Только запятая - может быть и разделитель тысяч, и десятичный разделитель
                    # Если после запятой 3 цифры - это разделитель тысяч
                    parts = weight_str.split(',')
                    if len(parts) == 2 and len(parts[1]) == 3 and len(parts[0]) > 0:
                        # Вероятно разделитель тысяч (например, "5,000")
                        weight_str = weight_str.replace(',', '')
                    else:
                        # Вероятно десятичный разделитель
                        weight_str = weight_str.replace(',', '.')
                elif has_dot:
                    # Только точка - может быть и разделитель тысяч, и десятичный разделитель
                    # Если после точки 3 цифры - это разделитель тысяч
                    parts = weight_str.split('.')
                    if len(parts) == 2 and len(parts[1]) == 3 and len(parts[0]) > 0:
                        # Вероятно разделитель тысяч (например, "5.000")
                        weight_str = weight_str.replace('.', '')
                    # Иначе точка - десятичный разделитель, оставляем как есть
                
                return float(weight_str)
            
            # Паттерн для поиска веса: число (с пробелами, точкой/запятой) + пробел (опционально) + единица измерения
            # Используем более широкий поиск, включая случаи без пробела и с пробелами внутри числа
            # Для кириллицы используем отрицательный просмотр вперед вместо \b
            # Поддерживаем пробелы внутри числа как разделители тысяч (например, "4 500кг")
            weight_patterns = [
                (r'(\d+(?:\s+\d+)*(?:[.,]\d+)?)\s*(?:кг|kg)(?![а-яёa-z])', False),  # килограммы (с пробелами в числе)
                (r'(\d+(?:\s+\d+)*(?:[.,]\d+)?)\s*(?:тонн|т)(?![а-яёa-z])', True),  # тонны (с пробелами в числе)
            ]
            
            for pattern, tons_flag in weight_patterns:
                try:
                    # Ищем все совпадения и выбираем самое длинное (вероятно, полное число)
                    weight_matches = re.findall(pattern, message_lower, re.UNICODE)
                    if weight_matches:
                        # Выбираем самое длинное совпадение (вероятно, полное число с разделителями)
                        weight_str = max(weight_matches, key=len)
                        try:
                            weight = parse_weight_number(weight_str)
                            is_tons = tons_flag
                            logger.debug(f"Извлечен вес: {weight} ({'тонн' if is_tons else 'кг'}) из строки '{weight_str}'")
                            break
                        except ValueError as e:
                            logger.warning(f"Не удалось распарсить вес из '{weight_str}': {e}")
                            continue
                except Exception as e:
                    logger.warning(f"Ошибка при поиске веса по паттерну '{pattern}': {e}")
                    continue
            
            # Если не нашли через регулярки, попробуем найти число рядом со словами "вес", "весом"
            if weight is None:
                # Ищем паттерн "вес[ом] ... число"
                weight_context_match = re.search(r'вес[ома]*\s+(\d+(?:[.,]\d+)?)', message_lower, re.UNICODE)
                if weight_context_match:
                    weight_str = weight_context_match.group(1)
                    try:
                        weight = parse_weight_number(weight_str)
                        # По умолчанию считаем килограммами, если не указано иное
                        if 'тонн' in message_lower or (re.search(r'\bт\b', message_lower) and 'тонн' not in message_lower):
                            is_tons = True
                        logger.debug(f"Извлечен вес из контекста: {weight} ({'тонн' if is_tons else 'кг'}) из строки '{weight_str}'")
                    except ValueError as e:
                        logger.warning(f"Не удалось распарсить вес из контекста '{weight_str}': {e}")
            
            if weight is not None:
                # Если указано в тоннах, переводим в кг
                if is_tons:
                    logger.debug(f"Вес указан в тоннах: {weight} т, переводим в кг")
                    weight *= 1000
                
                # Проверка на разумность значения (максимум 1000 тонн = 1,000,000 кг)
                # Если вес больше, вероятно ошибка парсинга
                if weight > 1_000_000:
                    logger.warning(f"Обнаружен подозрительно большой вес: {weight} кг. Возможно ошибка парсинга.")
                    # Пытаемся исправить: если вес очень большой, возможно это было в граммах или ошибка
                    # Но для безопасности оставляем как есть и просто логируем
                
                logistics_params['weight_kg'] = weight
                logger.info(f"Установлен вес для логистики: {logistics_params['weight_kg']} кг (извлечено из сообщения: '{message[:100]}', is_tons={is_tons})")
            
            # Город (улучшенный поиск с учетом контекста)
            cities = [c.get('name') for c in self.logistics_helper._load_cities()]
            # Словарь сокращений
            city_aliases = {
                'екб': 'Екатеринбург',
                'мск': 'Москва',
                'спб': 'Санкт-Петербург',
                'питер': 'Санкт-Петербург',
            }
            
            # Шаг 1: Проверяем сокращения
            for alias, city_name in city_aliases.items():
                if alias in message_lower and city_name in cities:
                    logistics_params['city_name'] = city_name
                    break
            
            # Шаг 2: Ищем город после предлогов "в", "до", "на" (приоритетный поиск)
            if 'city_name' not in logistics_params:
                # Ищем паттерны: "в [Город]", "до [Город]", "на [Город]", "из КНР в [Город]"
                preposition_patterns = [
                    r'(?:из\s+кнр\s+)?в\s+(\w+(?:\s+\w+)*)',
                    r'(?:из\s+кнр\s+)?до\s+(\w+(?:\s+\w+)*)',
                    r'на\s+(\w+(?:\s+\w+)*)',
                ]
                
                for pattern in preposition_patterns:
                    match = re.search(pattern, message_lower, re.UNICODE)
                    if match:
                        potential_city = match.group(1).strip()
                        # Ищем точное совпадение с городом из списка
                        for city in cities:
                            city_lower = city.lower()
                            # Точное совпадение или совпадение начала
                            if city_lower == potential_city or potential_city.startswith(city_lower):
                                logistics_params['city_name'] = city
                                logger.debug(f"Найден город '{city}' после предлога из фразы '{potential_city}'")
                                break
                        if 'city_name' in logistics_params:
                            break
            
            # Шаг 3: Если не нашли по предлогам, ищем по полному названию с границами слов
            if 'city_name' not in logistics_params:
                for city in cities:
                    city_lower = city.lower()
                    # Используем регулярку с границами слов для точного поиска
                    # Ищем слово целиком, не как часть другого слова
                    city_pattern = r'\b' + re.escape(city_lower) + r'\b'
                    if re.search(city_pattern, message_lower, re.UNICODE):
                        logistics_params['city_name'] = city
                        logger.debug(f"Найден город '{city}' по границам слов")
                        break
            
            # Шаг 4: Если все еще не нашли, используем частичное совпадение (fallback)
            if 'city_name' not in logistics_params:
                for city in cities:
                    city_lower = city.lower()
                    # Проверяем вхождение города в сообщение
                    if city_lower in message_lower:
                        logistics_params['city_name'] = city
                        logger.debug(f"Найден город '{city}' по частичному совпадению (fallback)")
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
        
        # Ключевые слова для работы с данными 1С
        analytics_keywords = [
            '1с', '1c', 'барабаны', 'вал', 'бандаж', 'блок', 'болт', 'бочка',
            'найди в 1с', 'покажи данные по', 'сколько записей',
            'топ заказчиков', 'топ материалов', 'статистика по',
            'график', 'визуализируй', 'покажи график',
            'менеджер', 'топ менеджеров', 'лучший менеджер',
            'тендер', 'выигран', 'выиграли', 'проиграли', 'отменен',
            'самый большой тендер', 'крупный тендер', 'большой заказ'
        ]
        
        is_analytics = any(keyword in message_lower for keyword in analytics_keywords)
        
        # Извлекаем параметры для поиска в 1С
        analytics_params = {}
        if is_analytics:
            # Определяем категорию
            categories = self.analytics_helper.get_categories()
            for cat in categories:
                cat_lower = cat.lower()
                if cat_lower in message_lower:
                    analytics_params['category'] = cat
                    break
            
            # Проверяем, нужны ли графики
            chart_keywords = ['график', 'визуализируй', 'покажи график', 'диаграмм']
            analytics_params['need_chart'] = any(keyword in message_lower for keyword in chart_keywords)
            
            # Извлекаем поисковый запрос
            search_patterns = [
                r'найди\s+(?:в\s+1с\s+)?(?:по\s+)?(?:чертежу|чертеж)\s+([^\s]+)',
                r'найди\s+(?:в\s+1с\s+)?(?:барабан|вал|бандаж|номенклатур[ауе])\s+([^\s]+)',
                r'найди\s+чертеж\s+([^\s]+)',
                r'покажи\s+(?:данные\s+)?(?:по\s+)?(?:заказчику|заказчик)\s+([^\s]+)',
            ]
            
            for pattern in search_patterns:
                match = re.search(pattern, message_lower, re.UNICODE)
                if match:
                    analytics_params['search_query'] = match.group(1).strip()
                    break
            
            # Если не нашли паттерном, пытаемся извлечь из контекста
            if 'search_query' not in analytics_params:
                # Ищем после ключевых слов
                for keyword in ['найди', 'покажи', 'дай', 'найти']:
                    if keyword in message_lower:
                        parts = message_lower.split(keyword)
                        if len(parts) > 1:
                            next_part = parts[1].strip()
                            # Убираем служебные слова
                            next_part = re.sub(r'\b(в\s+1с|по|чертежу|заказчику|материалу|номенклатуре)\b', '', next_part).strip()
                            if next_part:
                                analytics_params['search_query'] = next_part.split()[0]  # Берем первое слово
                                break
        
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
            'is_analytics': is_analytics,
            'analytics_params': analytics_params,
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
                
                # Проверяем наличие choices в ответе
                if 'choices' not in response_data:
                    logger.error(f"API ответ не содержит 'choices': {response_data}")
                    raise APIError(f"Некорректный ответ API: отсутствует 'choices'. Ответ: {response_data}")
                
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
                # Логируем тело ответа для диагностики
                try:
                    response_body = response.text[:1000]  # Ограничиваем размер лога
                except Exception:
                    response_body = "Не удалось получить тело ответа"
                
                error_msg = f"Неожиданный статус ответа: {response.status_code}"
                logger.error(f"API Error {response.status_code}: {error_msg}")
                logger.error(f"Response body: {response_body}")
                logger.error(f"Request URL: {self.api_url}")
                logger.error(f"Request model: {self.model}")
                
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
    
    def _enrich_intent_with_context(self, intent: Dict[str, Any], message: str) -> Dict[str, Any]:
        """
        Обогащает намерение контекстом из предыдущих сообщений.
        Обрабатывает уточняющие вопросы и использует умные дефолты.
        
        Args:
            intent: Извлеченное намерение
            message: Текущее сообщение пользователя
            
        Returns:
            Обогащенное намерение
        """
        message_lower = message.lower()
        
        # Проверяем, является ли это уточняющим вопросом
        clarification_keywords = ['а', 'а по', 'а по сумме', 'а по количеству', 'покажи', 'добавь', 'еще']
        is_clarification = any(keyword in message_lower for keyword in clarification_keywords)
        
        if is_clarification and len(self.conversation_history) >= 2:
            # Ищем последний запрос пользователя в истории
            last_user_message = None
            for msg in reversed(self.conversation_history):
                if msg.get('role') == 'user':
                    last_user_message = msg.get('content', '')
                    break
            
            if last_user_message:
                # Извлекаем намерение из предыдущего сообщения
                last_intent = self._extract_intent(last_user_message)
                
                # Если предыдущий запрос был аналитикой, применяем контекст
                if last_intent.get('is_analytics'):
                    # Обогащаем текущее намерение параметрами из предыдущего
                    if not intent.get('is_analytics'):
                        intent['is_analytics'] = True
                        intent['analytics_params'] = last_intent.get('analytics_params', {}).copy()
                    
                    # Обрабатываем уточнения
                    if 'по сумме' in message_lower or 'сумм' in message_lower:
                        intent['analytics_params']['by_sum'] = True
                    elif 'по количеству' in message_lower:
                        intent['analytics_params']['by_sum'] = False
                    
                    if 'график' in message_lower or 'покажи график' in message_lower:
                        intent['analytics_params']['need_chart'] = True
                    
                    # Сохраняем категорию из предыдущего запроса, если не указана новая
                    if 'category' not in intent['analytics_params'] and 'category' in last_intent.get('analytics_params', {}):
                        intent['analytics_params']['category'] = last_intent['analytics_params']['category']
                
                # Если предыдущий запрос был логистикой, применяем контекст
                elif last_intent.get('is_logistics'):
                    if not intent.get('is_logistics'):
                        intent['is_logistics'] = True
                        intent['logistics_params'] = last_intent.get('logistics_params', {}).copy()
        
        return intent
    
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
        # Определяем тип запроса для метрик
        intent = self._extract_intent(message)
        
        # Определяем тип запроса для метрик
        request_type = 'question'
        if intent.get('is_logistics'):
            request_type = 'logistics'
        elif intent.get('is_materials'):
            request_type = 'materials'
        elif intent.get('is_duty'):
            request_type = 'duty'
        elif intent.get('is_analytics'):
            request_type = 'analytics'
        
        # Начинаем отслеживание метрик
        metrics = self.metrics_collector.start_request(request_type)
        cache_hit = False
        records_count = 0
        error_type = None
        
        # Обогащаем намерение контекстом из предыдущих сообщений
        intent = self._enrich_intent_with_context(intent, message)
        
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
                cache_hit = True
                
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
                
                # Завершаем отслеживание метрик
                self.metrics_collector.end_request(metrics, success=True, cache_hit=True)
                return cached_response
            
        # Если это запрос на поиск материалов
        if intent['is_materials']:
            query = intent.get('materials_query') or message
            
            # Валидируем запрос материалов
            is_valid, error_msg, corrected_query = self.validator.validate_materials_query(query)
            if not is_valid:
                examples = [
                    "Найди материал 10",
                    "Какой GB у стали 20",
                    "Цена на сталь 15Г"
                ]
                error_response = self.validator.generate_validation_error_message(
                    'materials', [error_msg], examples
                )
                self.conversation_history.append({"role": "user", "content": message})
                self.conversation_history.append({"role": "assistant", "content": error_response})
                return error_response
            
            if corrected_query:
                query = corrected_query
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
            
            # Валидируем запрос пошлин
            is_valid, error_msg, corrected_query = self.validator.validate_duty_query(query)
            if not is_valid:
                examples = [
                    "Какая пошлина на сталь",
                    "Пошлина по коду 7214",
                    "Таможенная пошлина для металлопроката"
                ]
                error_response = self.validator.generate_validation_error_message(
                    'duty', [error_msg], examples
                )
                self.conversation_history.append({"role": "user", "content": message})
                self.conversation_history.append({"role": "assistant", "content": error_response})
                return error_response
            
            if corrected_query:
                query = corrected_query
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
        
        # Если это запрос на работу с данными 1С
        if intent['is_analytics']:
            params = intent['analytics_params']
            
            # Валидируем параметры аналитики
            is_valid, error_msg, corrected_params = self.validator.validate_analytics_params(params)
            if not is_valid:
                # Используем исправленные параметры
                if corrected_params != params:
                    params = corrected_params
                    intent['analytics_params'] = corrected_params
                    logger.info(f"Параметры аналитики исправлены: {error_msg}")
                else:
                    # Если исправление невозможно, возвращаем ошибку
                    examples = [
                        "Найди в 1с чертеж 35.1774.000 СБ",
                        "Топ менеджеров по сумме",
                        "Статистика по тендерам"
                    ]
                    error_response = self.validator.generate_validation_error_message(
                        'analytics', [error_msg], examples
                    )
                    self.conversation_history.append({"role": "user", "content": message})
                    self.conversation_history.append({"role": "assistant", "content": error_response})
                    return error_response
            
            category = params.get('category')
            search_query = params.get('search_query')
            need_chart = params.get('need_chart', False)
            
            message_lower = message.lower()
            
            # Определяем тип запроса: статистика, поиск или топ
            is_statistics = any(keyword in message_lower for keyword in ['сколько', 'статистика', 'стат', 'сколько записей'])
            is_top = any(keyword in message_lower for keyword in ['топ', 'top', 'лучшие', 'популярные'])
            is_manager = any(keyword in message_lower for keyword in ['менеджер', 'менеджеров'])
            is_tender = any(keyword in message_lower for keyword in ['тендер', 'выигран', 'выиграли', 'проиграли', 'отменен', 'крупный', 'большой'])
            is_largest_tender = any(keyword in message_lower for keyword in ['самый большой', 'крупный тендер', 'большой заказ', 'самый крупный'])
            
            # Определяем запросы на новые типы графиков
            is_timeline = any(keyword in message_lower for keyword in ['динамика', 'временной ряд', 'по месяцам', 'по кварталам', 'по годам', 'график по времени'])
            is_scatter = any(keyword in message_lower for keyword in ['корреляция', 'зависимость', 'вес vs цена', 'вес и цена', 'scatter'])
            is_heatmap = any(keyword in message_lower for keyword in ['heatmap', 'тепловая карта', 'матрица', 'заказчик менеджер'])
            
            # Определяем запросы на экспорт
            is_export = any(keyword in message_lower for keyword in ['экспорт', 'export', 'скачать', 'выгрузить', 'сохранить в', 'экспортировать'])
            export_format = None
            if is_export:
                if 'excel' in message_lower or 'xlsx' in message_lower:
                    export_format = 'excel'
                elif 'csv' in message_lower:
                    export_format = 'csv'
                elif 'png' in message_lower or 'график' in message_lower:
                    export_format = 'png'
                else:
                    export_format = 'csv'  # По умолчанию CSV
            
            results_df = pd.DataFrame()
            
            if is_statistics:
                # Запрос статистики
                # Проверяем кеш
                cache_params = {'type': 'statistics', 'category': category}
                cached_result = get_cached_analytics_query('statistics', cache_params)
                if cached_result:
                    logger.info("Используется закешированный результат статистики")
                    response = cached_result
                else:
                    stats = self.analytics_helper.get_statistics(category)
                    response_parts = [
                        "📊 Статистика по данным 1С"
                    ]
                    
                    if category:
                        response_parts.append(f"Категория: {category}")
                    else:
                        response_parts.append("Категории: " + ", ".join(self.analytics_helper.get_categories()))
                    
                    response_parts.append("")
                    
                    if stats:
                        response_parts.append(f"Всего записей: {ResponseFormatter.format_count(stats.get('total_records', 0))}")
                        
                        if 'total_weight_kg' in stats:
                            weight = ResponseFormatter.format_weight(stats['total_weight_kg'])
                            response_parts.append(f"Общий вес: {weight} кг")
                        
                        if 'total_price' in stats:
                            price = ResponseFormatter.format_price(stats['total_price'])
                            response_parts.append(f"Общая стоимость: {price} руб")
                        
                        if 'unique_customers' in stats:
                            response_parts.append(f"Уникальных заказчиков: {ResponseFormatter.format_count(stats['unique_customers'])}")
                        
                        if 'unique_materials' in stats:
                            response_parts.append(f"Уникальных материалов: {ResponseFormatter.format_count(stats['unique_materials'])}")
                    
                    response = "\n".join(response_parts)
                    
                    # Кешируем результат
                    cache_analytics_query('statistics', cache_params, response)
                    
                    # Генерируем графики если нужно
                    if need_chart or stats.get('total_records', 0) > 5:
                        # Топ заказчиков - используем метод поиска для получения данных
                        if category and category in self.analytics_helper.data_cache:
                            customer_df = self.analytics_helper.data_cache[category]
                        else:
                            # Объединяем все категории
                            dfs = [self.analytics_helper.data_cache[cat] for cat in self.analytics_helper.categories if cat in self.analytics_helper.data_cache]
                            customer_df = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()
                        
                        if not customer_df.empty:
                            # Проверяем кеш графика
                            chart_params = {'type': 'customers', 'category': category}
                            chart_img = get_cached_chart('customers', chart_params)
                            if not chart_img:
                                chart_img = self.analytics_helper.generate_chart_customers(customer_df)
                                if chart_img:
                                    cache_chart('customers', chart_params, chart_img)
                            if chart_img:
                                response += f"\n\n![Топ заказчиков](data:image/png;base64,{chart_img})"
                
            elif is_top:
                # Запрос топа
                if 'заказчик' in message_lower:
                    # Проверяем кеш
                    cache_params = {'type': 'top_customers', 'category': category, 'limit': 10}
                    cached_result = get_cached_analytics_query('top_customers', cache_params)
                    if cached_result:
                        logger.info("Используется закешированный результат топ заказчиков")
                        response = cached_result
                    else:
                        top_df = self.analytics_helper.get_top_customers(category, limit=10)
                        if not top_df.empty:
                            response_parts = ["📊 Топ-10 заказчиков по количеству заказов:\n"]
                            for idx, row in top_df.iterrows():
                                count = ResponseFormatter.format_count(row['Количество'])
                                response_parts.append(f"{idx + 1}. {row['Заказчик']}: {count} заказов")
                            response = "\n".join(response_parts)
                            
                            # Кешируем результат
                            cache_analytics_query('top_customers', cache_params, response)
                            
                            # График
                            if need_chart or len(top_df) > 5:
                                if category and category in self.analytics_helper.data_cache:
                                    customer_df = self.analytics_helper.data_cache[category]
                                else:
                                    dfs = [self.analytics_helper.data_cache[cat] for cat in self.analytics_helper.categories if cat in self.analytics_helper.data_cache]
                                    customer_df = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()
                                
                                if not customer_df.empty:
                                    # Проверяем кеш графика
                                    chart_params = {'type': 'customers', 'category': category}
                                    chart_img = get_cached_chart('customers', chart_params)
                                    if not chart_img:
                                        chart_img = self.analytics_helper.generate_chart_customers(customer_df)
                                        if chart_img:
                                            cache_chart('customers', chart_params, chart_img)
                                    if chart_img:
                                        response += f"\n\n![Топ заказчиков](data:image/png;base64,{chart_img})"
                        else:
                            response = SuggestionsGenerator.generate_empty_result_message(
                                'top',
                                category=category
                            )
                            # Добавляем подсказки
                            suggestions = SuggestionsGenerator.suggest_alternatives(
                                message,
                                self.analytics_helper.get_categories(),
                                ['номенклатура', 'чертеж', 'заказчик', 'материал']
                            )
                            response += "\n\n" + "\n".join(suggestions)
                        
                elif 'материал' in message_lower:
                    # Проверяем кеш
                    cache_params = {'type': 'top_materials', 'category': category, 'limit': 10}
                    cached_result = get_cached_analytics_query('top_materials', cache_params)
                    if cached_result:
                        logger.info("Используется закешированный результат топ материалов")
                        response = cached_result
                    else:
                        top_df = self.analytics_helper.get_top_materials(category, limit=10)
                        if not top_df.empty:
                            response_parts = ["📊 Топ-10 материалов по количеству использований:\n"]
                            for idx, row in top_df.iterrows():
                                count = ResponseFormatter.format_count(row['Количество'])
                                response_parts.append(f"{idx + 1}. {row['Материал']}: {count} использований")
                            response = "\n".join(response_parts)
                            
                            # Кешируем результат
                            cache_analytics_query('top_materials', cache_params, response)
                            
                            # График
                            if need_chart or len(top_df) > 5:
                                if category and category in self.analytics_helper.data_cache:
                                    material_df = self.analytics_helper.data_cache[category]
                                else:
                                    dfs = [self.analytics_helper.data_cache[cat] for cat in self.analytics_helper.categories if cat in self.analytics_helper.data_cache]
                                    material_df = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()
                                
                                if not material_df.empty:
                                    # Проверяем кеш графика
                                    chart_params = {'type': 'materials', 'category': category}
                                    chart_img = get_cached_chart('materials', chart_params)
                                    if not chart_img:
                                        chart_img = self.analytics_helper.generate_chart_materials(material_df)
                                        if chart_img:
                                            cache_chart('materials', chart_params, chart_img)
                                    if chart_img:
                                        response += f"\n\n![Популярные материалы](data:image/png;base64,{chart_img})"
                        else:
                            response = SuggestionsGenerator.generate_empty_result_message(
                                'top',
                                category=category
                            )
                            suggestions = SuggestionsGenerator.suggest_alternatives(
                                message,
                                self.analytics_helper.get_categories(),
                                ['номенклатура', 'чертеж', 'заказчик', 'материал']
                            )
                            response += "\n\n" + "\n".join(suggestions)
                elif 'менеджер' in message_lower:
                    # Топ менеджеров по выигранным тендерам
                    by_sum = 'сумм' in message_lower or 'сумме' in message_lower or 'деньг' in message_lower
                    # Проверяем кеш
                    cache_params = {'type': 'top_managers', 'category': category, 'limit': 10, 'by_sum': by_sum}
                    cached_result = get_cached_analytics_query('top_managers', cache_params)
                    if cached_result:
                        logger.info("Используется закешированный результат топ менеджеров")
                        response = cached_result
                    else:
                        top_df = self.analytics_helper.get_top_managers_by_won_tenders(category, limit=10, by_sum=by_sum)
                        if not top_df.empty:
                            if by_sum and 'Сумма_руб' in top_df.columns:
                                response_parts = ["📊 Топ-10 менеджеров по сумме выигранных тендеров:\n"]
                                for idx, row in top_df.iterrows():
                                    sum_rub = ResponseFormatter.format_price(row['Сумма_руб'])
                                    count = ResponseFormatter.format_count(row['Количество_тендеров'])
                                    response_parts.append(f"{idx + 1}. {row['Менеджер']}: {sum_rub} руб ({count} тендеров)")
                            else:
                                response_parts = ["📊 Топ-10 менеджеров по количеству выигранных тендеров:\n"]
                                for idx, row in top_df.iterrows():
                                    count = ResponseFormatter.format_count(row['Количество_тендеров'])
                                    response_parts.append(f"{idx + 1}. {row['Менеджер']}: {count} тендеров")
                            response = "\n".join(response_parts)
                            
                            # Кешируем результат
                            cache_analytics_query('top_managers', cache_params, response)
                            
                            # График
                            if need_chart or len(top_df) > 5:
                                if category and category in self.analytics_helper.data_cache:
                                    manager_df = self.analytics_helper.data_cache[category]
                                else:
                                    dfs = [self.analytics_helper.data_cache[cat] for cat in self.analytics_helper.categories if cat in self.analytics_helper.data_cache]
                                    manager_df = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()
                                
                                if not manager_df.empty:
                                    # Проверяем кеш графика
                                    chart_params = {'type': 'managers', 'category': category, 'by_sum': by_sum}
                                    chart_img = get_cached_chart('managers', chart_params)
                                    if not chart_img:
                                        chart_img = self.analytics_helper.generate_chart_managers(manager_df, by_sum=by_sum)
                                        if chart_img:
                                            cache_chart('managers', chart_params, chart_img)
                                    if chart_img:
                                        response += f"\n\n![Топ менеджеров](data:image/png;base64,{chart_img})"
                        else:
                            response = SuggestionsGenerator.generate_empty_result_message(
                                'top',
                                category=category
                            )
                            suggestions = SuggestionsGenerator.suggest_alternatives(
                                message,
                                self.analytics_helper.get_categories(),
                                ['номенклатура', 'чертеж', 'заказчик', 'материал']
                            )
                            response += "\n\n" + "\n".join(suggestions)
                else:
                    response = "Уточните, что показать: топ заказчиков, топ материалов или топ менеджеров?"
            
            elif is_largest_tender:
                # Самый большой выигранный тендер
                # Проверяем кеш
                cache_params = {'type': 'largest_tender', 'category': category}
                cached_result = get_cached_analytics_query('largest_tender', cache_params)
                if cached_result:
                    logger.info("Используется закешированный результат самого большого тендера")
                    response = cached_result
                else:
                    largest = self.analytics_helper.get_largest_won_tender(category)
                    if largest:
                        response_parts = ["🏆 Самый большой выигранный тендер:\n"]
                        response_parts.append(f"**{largest.get('Номенклатура', 'Без названия')}**")
                        response_parts.append(f"Чертеж: {largest.get('Чертеж', '')}")
                        response_parts.append(f"Заказчик: {largest.get('Заказчик', '')}")
                        response_parts.append(f"Менеджер: {largest.get('Менеджер', '')}")
                        count = ResponseFormatter.format_count(largest.get('Количество', 0))
                        response_parts.append(f"Количество: {count} шт")
                        price = ResponseFormatter.format_price(largest.get('Цена_подачи', 0))
                        response_parts.append(f"Цена за единицу: {price} руб")
                        total_sum = ResponseFormatter.format_price(largest.get('Общая_сумма', 0))
                        response_parts.append(f"**Общая сумма: {total_sum} руб**")
                        response_parts.append(f"Статус: {largest.get('Статус_тендера', '')}")
                        response_parts.append(f"Дата тендера: {largest.get('Дата_тендера', '')}")
                        if largest.get('Категория'):
                            response_parts.append(f"Категория: {largest.get('Категория', '')}")
                        response = "\n".join(response_parts)
                        
                        # Кешируем результат
                        cache_analytics_query('largest_tender', cache_params, response)
                    else:
                        response = SuggestionsGenerator.generate_empty_result_message(
                            'search',
                            query=None,
                            category=category
                        )
                        response += "\n\n" + SuggestionsGenerator.generate_help_message('search')
            
            elif is_tender and is_statistics:
                # Статистика по тендерам
                # Проверяем кеш
                cache_params = {'type': 'tender_statistics', 'category': category}
                cached_result = get_cached_analytics_query('tender_statistics', cache_params)
                if cached_result:
                    logger.info("Используется закешированный результат статистики по тендерам")
                    response = cached_result
                else:
                    tender_stats = self.analytics_helper.get_tender_statistics(category)
                    if tender_stats:
                        response_parts = ["📊 Статистика по тендерам"]
                        if category:
                            response_parts.append(f"Категория: {category}")
                        else:
                            response_parts.append("Категории: " + ", ".join(self.analytics_helper.get_categories()))
                        response_parts.append("")
                        response_parts.append(f"Всего тендеров: {ResponseFormatter.format_count(tender_stats.get('total_tenders', 0))}")
                        response_parts.append(f"✅ Выиграно: {ResponseFormatter.format_count(tender_stats.get('won', 0))}")
                        response_parts.append(f"✅ Выиграно частично: {ResponseFormatter.format_count(tender_stats.get('won_partially', 0))}")
                        response_parts.append(f"❌ Проиграно: {ResponseFormatter.format_count(tender_stats.get('lost', 0))}")
                        response_parts.append(f"🚫 Отменено: {ResponseFormatter.format_count(tender_stats.get('cancelled', 0))}")
                        response_parts.append(f"📤 Подано: {ResponseFormatter.format_count(tender_stats.get('submitted', 0))}")
                        response_parts.append(f"⚪ Не участвовали: {ResponseFormatter.format_count(tender_stats.get('not_participated', 0))}")
                        if tender_stats.get('won_sum_rub', 0) > 0:
                            won_sum = ResponseFormatter.format_price(tender_stats.get('won_sum_rub', 0))
                            response_parts.append(f"\n💰 Сумма выигранных тендеров: {won_sum} руб")
                        response = "\n".join(response_parts)
                        
                        # Кешируем результат
                        cache_analytics_query('tender_statistics', cache_params, response)
                    else:
                        response = SuggestionsGenerator.generate_empty_result_message(
                            'statistics',
                            category=category
                        )
                        response += "\n\n" + SuggestionsGenerator.generate_help_message('statistics')
            
            elif search_query:
                # Поиск по запросу
                # Определяем тип поиска
                search_type = 'drawing' if ('чертеж' in message_lower or 'чертежу' in message_lower) else \
                              'customer' if ('заказчик' in message_lower or 'заказчику' in message_lower) else \
                              'material' if ('материал' in message_lower or 'материалу' in message_lower) else \
                              'combined'
                
                # Проверяем кеш
                cache_params = {
                    'type': 'search',
                    'search_type': search_type,
                    'query': search_query,
                    'category': category
                }
                cached_result = get_cached_analytics_query('search', cache_params)
                if cached_result:
                    logger.info("Используется закешированный результат поиска")
                    response = cached_result
                else:
                    # Определяем тип поиска по контексту
                    if 'чертеж' in message_lower or 'чертежу' in message_lower:
                        results_df = self.analytics_helper.search_by_drawing(search_query, category)
                    elif 'заказчик' in message_lower or 'заказчику' in message_lower:
                        results_df = self.analytics_helper.search_by_customer(search_query, category)
                    elif 'материал' in message_lower or 'материалу' in message_lower:
                        results_df = self.analytics_helper.search_by_material(search_query, category)
                    else:
                        # Пробуем разные типы поиска
                        results_df = self.analytics_helper.search_by_drawing(search_query, category)
                        if results_df.empty:
                            results_df = self.analytics_helper.search_by_nomenclature(search_query, category)
                        if results_df.empty:
                            results_df = self.analytics_helper.search_by_customer(search_query, category)
                
                    # Форматируем результат с пагинацией
                    include_charts = need_chart or len(results_df) > 5
                    # Используем дефолтный лимит 20 записей для отображения
                    formatted = self.analytics_helper.format_search_result(
                        results_df, 
                        include_charts=include_charts, 
                        max_records=20,
                        offset=0
                    )
                    
                    response_parts = [formatted['text']]
                    
                    # Добавляем записи
                    if formatted['records']:
                        response_parts.append("\n**Найденные записи:**\n")
                        for i, record in enumerate(formatted['records'], 1):
                            record_num = formatted.get('offset', 0) + i
                            record_parts = [f"**{record_num}. {record.get('Номенклатура', 'Без названия')}**"]
                            if record.get('Чертеж'):
                                record_parts.append(f"Чертеж: {record['Чертеж']}")
                            if record.get('Заказчик, сокр'):
                                record_parts.append(f"Заказчик: {record['Заказчик, сокр']}")
                            if record.get('Материал'):
                                record_parts.append(f"Материал: {record['Материал']}")
                            if record.get('Цена подачи'):
                                record_parts.append(f"Цена: {record['Цена подачи']} руб")
                            response_parts.append(" | ".join(record_parts))
                    
                    # Добавляем информацию о пагинации
                    if formatted.get('has_more', False):
                        remaining = formatted['total_count'] - formatted['shown_count']
                        response_parts.append(f"\n... и еще {remaining} записей (всего найдено {formatted['total_count']})")
                    
                    response = "\n".join(response_parts)
                    
                    # Добавляем графики
                    for chart in formatted.get('charts', []):
                        # Используем markdown формат для изображений
                        # Frontend должен поддерживать отображение base64 изображений
                        if 'image' in chart:
                            # chart['image'] уже содержит полный data URL из format_search_result
                            if chart['image'].startswith('data:image'):
                                response += f"\n\n![{chart.get('title', 'График')}]({chart['image']})"
                            elif chart['image']:  # Если это только base64 без префикса
                                response += f"\n\n![{chart.get('title', 'График')}](data:image/png;base64,{chart['image']})"
                    
                    # Кешируем результат поиска
                    cache_analytics_query('search', cache_params, response)
                    
                    # Обработка экспорта результатов
                    if is_export and not results_df.empty:
                        try:
                            if export_format == 'excel':
                                export_result = self.exporter.export_to_excel(results_df)
                                response += f"\n\n📥 **Экспорт в Excel выполнен:**\n"
                                response += f"Файл: `{export_result['filename']}`\n"
                                response += f"Записей: {export_result['records_count']}\n"
                                response += f"Размер: {export_result['size'] / 1024:.2f} КБ\n"
                                response += f"Путь: `{export_result['path']}`"
                            elif export_format == 'png':
                                # Экспортируем графики, если есть
                                if formatted.get('charts'):
                                    export_result = self.exporter.export_multiple_charts(formatted['charts'])
                                    response += f"\n\n📥 **Экспорт графиков в PNG выполнен:**\n"
                                    response += f"Файлов: {export_result['count']}\n"
                                    response += f"Общий размер: {export_result['total_size'] / 1024:.2f} КБ\n"
                                    for file_info in export_result['files']:
                                        response += f"• `{file_info['filename']}` ({file_info['size'] / 1024:.2f} КБ)\n"
                                else:
                                    response += "\n\n⚠️ Графики для экспорта не найдены."
                            else:  # CSV по умолчанию
                                export_result = self.exporter.export_to_csv(results_df)
                                response += f"\n\n📥 **Экспорт в CSV выполнен:**\n"
                                response += f"Файл: `{export_result['filename']}`\n"
                                response += f"Записей: {export_result['records_count']}\n"
                                response += f"Размер: {export_result['size'] / 1024:.2f} КБ\n"
                                response += f"Путь: `{export_result['path']}`"
                        except Exception as e:
                            logger.error(f"Ошибка при экспорте: {e}", exc_info=True)
                            response += f"\n\n❌ Ошибка при экспорте: {str(e)}"
            
            elif is_timeline:
                # Запрос на временной ряд
                # Получаем данные
                if category and category in self.analytics_helper.data_cache:
                    timeline_df = self.analytics_helper.data_cache[category]
                else:
                    dfs = [self.analytics_helper.data_cache[cat] for cat in self.analytics_helper.categories if cat in self.analytics_helper.data_cache]
                    timeline_df = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()
                
                if not timeline_df.empty:
                    # Определяем группировку
                    group_by = 'month'
                    if 'квартал' in message_lower or 'quarter' in message_lower:
                        group_by = 'quarter'
                    elif 'год' in message_lower or 'year' in message_lower:
                        group_by = 'year'
                    
                    # Проверяем кеш
                    chart_params = {'type': 'timeline', 'category': category, 'group_by': group_by}
                    chart_img = get_cached_chart('timeline', chart_params)
                    if not chart_img:
                        chart_img = self.analytics_helper.generate_chart_timeline(
                            timeline_df,
                            group_by=group_by
                        )
                        if chart_img:
                            cache_chart('timeline', chart_params, chart_img)
                    
                    if chart_img:
                        response = f"📈 Динамика тендеров по {group_by}:\n\n![Динамика тендеров](data:image/png;base64,{chart_img})"
                    else:
                        response = "Не удалось сгенерировать график динамики. Проверьте наличие данных с датами."
                else:
                    response = SuggestionsGenerator.generate_empty_result_message('statistics', category=category)
            
            elif is_scatter:
                # Запрос на scatter plot (корреляция)
                # Получаем данные
                if category and category in self.analytics_helper.data_cache:
                    scatter_df = self.analytics_helper.data_cache[category]
                else:
                    dfs = [self.analytics_helper.data_cache[cat] for cat in self.analytics_helper.categories if cat in self.analytics_helper.data_cache]
                    scatter_df = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()
                
                if not scatter_df.empty:
                    # Определяем оси (по умолчанию вес vs цена)
                    x_col = 'ГП вес, кг'
                    y_col = 'Цена подачи'
                    
                    # Проверяем кеш
                    chart_params = {'type': 'scatter', 'category': category, 'x': x_col, 'y': y_col}
                    chart_img = get_cached_chart('scatter', chart_params)
                    if not chart_img:
                        chart_img = self.analytics_helper.generate_chart_scatter(
                            scatter_df,
                            x_column=x_col,
                            y_column=y_col
                        )
                        if chart_img:
                            cache_chart('scatter', chart_params, chart_img)
                    
                    if chart_img:
                        response = f"📊 Корреляция: {x_col} vs {y_col}\n\n![Корреляция](data:image/png;base64,{chart_img})"
                    else:
                        response = "Не удалось сгенерировать график корреляции. Проверьте наличие необходимых данных."
                else:
                    response = SuggestionsGenerator.generate_empty_result_message('statistics', category=category)
            
            elif is_heatmap:
                # Запрос на heatmap
                # Получаем данные
                if category and category in self.analytics_helper.data_cache:
                    heatmap_df = self.analytics_helper.data_cache[category]
                else:
                    dfs = [self.analytics_helper.data_cache[cat] for cat in self.analytics_helper.categories if cat in self.analytics_helper.data_cache]
                    heatmap_df = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()
                
                if not heatmap_df.empty:
                    # По умолчанию: заказчик vs менеджер
                    x_col = 'Заказчик, сокр' if 'Заказчик, сокр' in heatmap_df.columns else 'Заказчик'
                    y_col = 'Менеджер'
                    value_col = 'Цена подачи'
                    
                    # Проверяем кеш
                    chart_params = {'type': 'heatmap', 'category': category, 'x': x_col, 'y': y_col, 'value': value_col}
                    chart_img = get_cached_chart('heatmap', chart_params)
                    if not chart_img:
                        chart_img = self.analytics_helper.generate_chart_heatmap(
                            heatmap_df,
                            x_column=x_col,
                            y_column=y_col,
                            value_column=value_col
                        )
                        if chart_img:
                            cache_chart('heatmap', chart_params, chart_img)
                    
                    if chart_img:
                        response = f"🔥 Heatmap: {y_col} vs {x_col}\n\n![Heatmap](data:image/png;base64,{chart_img})"
                    else:
                        response = "Не удалось сгенерировать heatmap. Проверьте наличие необходимых данных."
                else:
                    response = SuggestionsGenerator.generate_empty_result_message('statistics', category=category)
            
            # Проверяем запросы на комбинированные расчеты (логистика + данные 1С)
            is_logistics_for_1c = any(keyword in message_lower for keyword in [
                'рассчитай логистику для', 'логистика для записей', 'логистика по заказчику',
                'рассчитай доставку для', 'стоимость доставки для'
            ])
            
            if is_logistics_for_1c and search_query:
                # Комбинированный запрос: поиск в 1С + расчет логистики
                # Сначала находим записи
                if 'заказчик' in message_lower or 'заказчику' in message_lower:
                    results_df = self.analytics_helper.search_by_customer(search_query, category)
                else:
                    results_df = self.analytics_helper.search_by_nomenclature(search_query, category)
                
                if not results_df.empty:
                    # Извлекаем город из запроса
                    city_name = None
                    cities = [c.get('name') for c in self.logistics_helper._load_cities()]
                    for city in cities:
                        if city.lower() in message_lower:
                            city_name = city
                            break
                    
                    if not city_name:
                        response = "Для расчета логистики необходимо указать город назначения.\n\n"
                        response += f"Найдено {len(results_df)} записей по запросу '{search_query}'.\n"
                        response += "Укажите город, например: \"рассчитай логистику для заказчика X в Москву\""
                    else:
                        # Рассчитываем логистику для каждой записи
                        logistics_results = []
                        for idx, row in results_df.iterrows():
                            weight = row.get('ГП вес, кг')
                            if pd.notna(weight) and weight > 0:
                                logistics_result = self.logistics_helper.calculate_simple_logistics(
                                    weight_kg=float(weight),
                                    city_name=city_name
                                )
                                if 'error' not in logistics_result:
                                    logistics_results.append({
                                        'record': row.to_dict(),
                                        'logistics': logistics_result
                                    })
                        
                        if logistics_results:
                            response_parts = [
                                f"📦 Расчет логистики для {len(logistics_results)} записей",
                                f"Город назначения: {city_name}",
                                f"Запрос: {search_query}",
                                ""
                            ]
                            
                            total_cost = 0
                            for i, item in enumerate(logistics_results[:10], 1):  # Показываем первые 10
                                record = item['record']
                                logistics = item['logistics']
                                cost = logistics.get('total_cost', 0)
                                total_cost += cost
                                
                                response_parts.append(f"**{i}. {record.get('Номенклатура', 'Без названия')}**")
                                if record.get('Чертеж'):
                                    response_parts.append(f"Чертеж: {record['Чертеж']}")
                                weight = ResponseFormatter.format_weight(record.get('ГП вес, кг', 0))
                                response_parts.append(f"Вес: {weight} кг")
                                logistics_cost = ResponseFormatter.format_price(cost)
                                response_parts.append(f"Логистика: {logistics_cost} руб")
                                response_parts.append("")
                            
                            if len(logistics_results) > 10:
                                response_parts.append(f"... и еще {len(logistics_results) - 10} записей")
                            
                            # Общая сумма
                            total_cost_formatted = ResponseFormatter.format_price(total_cost)
                            response_parts.append(f"\n**Общая стоимость логистики: {total_cost_formatted} руб**")
                            
                            response = "\n".join(response_parts)
                        else:
                            response = f"Не удалось рассчитать логистику. Проверьте наличие веса в записях."
                else:
                    response = f"По запросу '{search_query}' ничего не найдено в данных 1С."
            
            else:
                # Неопределенный запрос - просим уточнить
                response = "Для поиска в данных 1С уточните:\n"
                response += "• Что искать (номенклатура, чертеж, заказчик, материал)\n"
                response += "• В какой категории (Барабаны, Вал) - или во всех\n"
                response += "• Нужны ли графики (укажите 'покажи график')\n"
                response += "• Можно рассчитать логистику для найденных записей\n\n"
                response += "Примеры запросов:\n"
                response += "• \"Найди в 1с чертеж 35.1774.000 СБ\"\n"
                response += "• \"Покажи все записи по заказчику ВАЙЛДБЕРРИЗ\"\n"
                response += "• \"Сколько записей в Барабанах\"\n"
                response += "• \"Топ заказчиков с графиком\"\n"
                response += "• \"Топ менеджеров по выигранным тендерам\"\n"
                response += "• \"Самый большой выигранный тендер\"\n"
                response += "• \"Статистика по тендерам\"\n"
                response += "• \"Рассчитай логистику для заказчика X в Москву\""
            
            # Добавляем в историю
            self.conversation_history.append({"role": "user", "content": message})
            self.conversation_history.append({"role": "assistant", "content": response})
            
            # Завершаем отслеживание метрик
            self.metrics_collector.end_request(metrics, success=True, cache_hit=cache_hit, records_count=records_count)
            
            return response
        
        # Если это запрос на расчет логистики
        if is_logistics_query:
            params = intent['logistics_params']
            
            # Валидируем параметры логистики
            is_valid, error_msg, corrected_params = self.validator.validate_logistics_params(params)
            if not is_valid:
                # Используем исправленные параметры, если они есть
                if corrected_params != params:
                    params = corrected_params
                    intent['logistics_params'] = corrected_params
                    logger.info(f"Параметры логистики исправлены: {error_msg}")
                else:
                    # Если исправление невозможно, возвращаем ошибку
                    examples = [
                        "Рассчитай логистику из КНР в Екатеринбург за груз весом 5000 кг",
                        "Логистика в Москву, вес 2 тонны",
                        "Расчет доставки в Санкт-Петербург, 1000 кг"
                    ]
                    error_response = self.validator.generate_validation_error_message(
                        'logistics', [error_msg], examples
                    )
                    self.conversation_history.append({"role": "user", "content": message})
                    self.conversation_history.append({"role": "assistant", "content": error_response})
                    return error_response
            
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
            
            logger.debug(f"Вызываем calculate_simple_logistics с weight_kg={weight_kg}, city_name={city_name}")
            
            result = self.logistics_helper.calculate_simple_logistics(
                weight_kg=weight_kg,
                city_name=city_name,
                transport_type=params.get('transport_type', 'truck'),
                length_mm=params.get('length_mm'),
                width_mm=params.get('width_mm'),
                height_mm=params.get('height_mm'),
            )
            
            logger.debug(f"Результат calculate_simple_logistics: weight_kg={result.get('weight_kg')}, formula={result.get('calculation_formula')}")
            
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
        
        # Формируем динамический системный промпт с учетом текущего запроса
        # Это оптимизирует размер промпта и делает его более релевантным
        dynamic_system_prompt = self._build_system_prompt(query_context=message)
        
        # Формируем сообщения для API
        api_messages = [
            {"role": "system", "content": dynamic_system_prompt},
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
            
            # Проверяем наличие choices
            if 'choices' not in response_data or not response_data['choices']:
                logger.error(f"Ответ API не содержит choices: {response_data}")
                return "⚠️ Ошибка: API вернул некорректный ответ. Попробуйте повторить запрос."
            
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
            error_type = type(e).__name__
            error_response = f"⚠️ **Ошибка при обращении к AI**: {str(e)}"
            
            # Используем fallback если включен
            if self.fallback_enabled:
                error_response += "\n\n" + self._fallback_search(message)
            
            self.conversation_history.append({
                "role": "assistant",
                "content": error_response
            })
            
            # Завершаем отслеживание метрик с ошибкой
            self.metrics_collector.end_request(metrics, success=False, error_type=error_type)
            return error_response
        
        except Exception as e:
            logger.exception(f"Неожиданная ошибка в chat(): {e}")
            error_type = type(e).__name__
            error_response = "❌ Произошла неожиданная ошибка. Пожалуйста, попробуйте позже."
            
            # Используем fallback если включен
            if self.fallback_enabled:
                error_response += "\n\n" + self._fallback_search(message)
            
            self.conversation_history.append({
                "role": "assistant",
                "content": error_response
            })
            
            # Завершаем отслеживание метрик с ошибкой
            self.metrics_collector.end_request(metrics, success=False, error_type=error_type)
            return error_response
    
    def clear_history(self) -> None:
        """Очищает историю диалога."""
        self.conversation_history = []
        logger.debug("История диалога очищена")

