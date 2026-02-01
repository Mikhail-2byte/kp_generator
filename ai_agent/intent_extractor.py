"""
Многоуровневая система извлечения намерений из пользовательских запросов.
Поддерживает комбинированные запросы и приоритетные паттерны.
"""

import logging
import re
from typing import Any, Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)


class IntentExtractor:
    """Класс для извлечения намерений из пользовательских запросов."""
    
    def __init__(self, logistics_helper, analytics_helper):
        """
        Инициализирует извлекатель намерений.
        
        Args:
            logistics_helper: Экземпляр LogisticsHelper для доступа к городам
            analytics_helper: Экземпляр AnalyticsHelper для доступа к категориям
        """
        self.logistics_helper = logistics_helper
        self.analytics_helper = analytics_helper
        
        # Приоритетные паттерны для разных типов запросов
        # Порядок важен - более специфичные паттерны должны быть выше
        self.intent_patterns = self._build_intent_patterns()
    
    def _build_intent_patterns(self) -> List[Tuple[str, Dict[str, Any]]]:
        """
        Строит список паттернов для распознавания намерений с приоритетами.
        
        Returns:
            Список кортежей (pattern, metadata) отсортированных по приоритету
        """
        patterns = []
        
        # Высокий приоритет: комбинированные запросы
        patterns.extend([
            # Комбинированные запросы аналитики
            (r'топ\s+менеджеров?\s+(?:по\s+)?(?:сумм[еа]|деньг[аам]|выручк[еа])\s+(?:в\s+)?([^\s]+)?', {
                'type': 'analytics',
                'subtype': 'top_managers',
                'by_sum': True,
                'priority': 10
            }),
            (r'топ\s+менеджеров?\s+(?:по\s+)?(?:количеству|тендерам)\s+(?:в\s+)?([^\s]+)?', {
                'type': 'analytics',
                'subtype': 'top_managers',
                'by_sum': False,
                'priority': 10
            }),
            (r'топ\s+(?:заказчиков?|материалов?|менеджеров?)\s+(?:в\s+)?([^\s]+)?', {
                'type': 'analytics',
                'subtype': 'top',
                'priority': 9
            }),
            (r'самый\s+(?:большой|крупный)\s+(?:выигранный\s+)?тендер\s+(?:в\s+)?([^\s]+)?', {
                'type': 'analytics',
                'subtype': 'largest_tender',
                'priority': 10
            }),
            (r'статистика\s+(?:по\s+)?тендерам\s+(?:в\s+)?([^\s]+)?', {
                'type': 'analytics',
                'subtype': 'tender_statistics',
                'priority': 9
            }),
        ])
        
        # Средний приоритет: специфичные запросы
        patterns.extend([
            # Логистика с полными параметрами
            (r'(?:рассчитай|посчитай|расчет)\s+логистик[уи]\s+(?:из\s+кнр\s+)?(?:в|до)\s+(\w+)\s+(?:для|за|весом|вес)\s+(\d+(?:\s+\d+)*(?:[.,]\d+)?)\s*(?:кг|тонн|т)', {
                'type': 'logistics',
                'priority': 8
            }),
            # Поиск по чертежу в 1С
            (r'найди\s+(?:в\s+1с\s+)?(?:по\s+)?чертеж[у]?\s+([^\s]+(?:\s+[^\s]+)*)', {
                'type': 'analytics',
                'subtype': 'search',
                'search_type': 'drawing',
                'priority': 8
            }),
            # Поиск по заказчику в 1С
            (r'покажи\s+(?:все\s+)?(?:записи|данные)\s+(?:по\s+)?заказчику\s+([^\s]+)', {
                'type': 'analytics',
                'subtype': 'search',
                'search_type': 'customer',
                'priority': 8
            }),
        ])
        
        # Низкий приоритет: общие паттерны
        patterns.extend([
            (r'логистик[аи]|рассчитай|расчет\s+(?:стоимости|цены)', {
                'type': 'logistics',
                'priority': 5
            }),
            (r'материал|gb|гб|gost|гост', {
                'type': 'materials',
                'priority': 5
            }),
            (r'пошлин[аи]|тн\s*вэд|tnved', {
                'type': 'duty',
                'priority': 5
            }),
            (r'1с|1c', {
                'type': 'analytics',
                'priority': 5
            }),
        ])
        
        # Сортируем по приоритету (высокий приоритет = выше в списке)
        patterns.sort(key=lambda x: x[1].get('priority', 0), reverse=True)
        
        return patterns
    
    def extract(self, message: str) -> Dict[str, Any]:
        """
        Извлекает намерение из сообщения пользователя.
        
        Args:
            message: Сообщение пользователя
            
        Returns:
            Словарь с информацией о намерении и параметрах
        """
        message_lower = message.lower()
        result = {
            'is_logistics': False,
            'is_materials': False,
            'is_duty': False,
            'is_analytics': False,
            'is_question': False,
            'logistics_params': {},
            'materials_query': None,
            'duty_query': None,
            'analytics_params': {},
            'confidence': 0.0,  # Уверенность в определении намерения
        }
        
        # Определяем, является ли это вопросом
        question_words = ['как', 'что', 'почему', 'где', 'когда', 'зачем', 'кто']
        result['is_question'] = '?' in message or any(
            word in message_lower for word in question_words
        )
        
        # Применяем паттерны в порядке приоритета
        matched_pattern = None
        for pattern, metadata in self.intent_patterns:
            match = re.search(pattern, message_lower, re.UNICODE | re.IGNORECASE)
            if match:
                matched_pattern = (pattern, metadata, match)
                result['confidence'] = metadata.get('priority', 0) / 10.0
                break
        
        # Обрабатываем найденный паттерн
        if matched_pattern:
            pattern, metadata, match = matched_pattern
            intent_type = metadata.get('type')
            
            if intent_type == 'logistics':
                result['is_logistics'] = True
                result['logistics_params'] = self._extract_logistics_params(message, match)
            elif intent_type == 'materials':
                result['is_materials'] = True
                result['materials_query'] = self._extract_materials_query(message, match)
            elif intent_type == 'duty':
                result['is_duty'] = True
                result['duty_query'] = self._extract_duty_query(message, match)
            elif intent_type == 'analytics':
                result['is_analytics'] = True
                result['analytics_params'] = self._extract_analytics_params(
                    message, match, metadata
                )
        
        # Если не нашли по паттернам, используем fallback (старый метод)
        if not any([result['is_logistics'], result['is_materials'], 
                   result['is_duty'], result['is_analytics']]):
            result = self._fallback_extraction(message, result)
        
        return result
    
    def _extract_logistics_params(self, message: str, match: re.Match) -> Dict[str, Any]:
        """Извлекает параметры логистики из сообщения."""
        message_lower = message.lower()
        params = {}
        
        # Извлекаем город из паттерна, если есть
        if match.groups():
            city_candidate = match.group(1) if len(match.groups()) >= 1 else None
            if city_candidate:
                cities = [c.get('name') for c in self.logistics_helper._load_cities()]
                city_candidate_lower = city_candidate.lower()
                for city in cities:
                    if city.lower() == city_candidate_lower or city_candidate_lower.startswith(city.lower()):
                        params['city_name'] = city
                        break
        
        # Извлекаем вес из паттерна, если есть
        if len(match.groups()) >= 2:
            weight_str = match.group(2)
            weight = self._parse_weight(weight_str)
            if weight:
                params['weight_kg'] = weight
        
        # Дополнительное извлечение параметров (если не нашли в паттерне)
        if 'weight_kg' not in params:
            weight = self._extract_weight(message_lower)
            if weight:
                params['weight_kg'] = weight
        
        if 'city_name' not in params:
            city = self._extract_city(message_lower)
            if city:
                params['city_name'] = city
        
        # Тип транспорта
        if 'трал' in message_lower:
            params['transport_type'] = 'trail'
        else:
            params['transport_type'] = 'truck'
        
        # Габариты
        dim_match = re.search(r'(\d+)\s*[xх×]\s*(\d+)\s*[xх×]\s*(\d+)\s*(?:мм|mm)', message_lower)
        if dim_match:
            params['length_mm'] = float(dim_match.group(1))
            params['width_mm'] = float(dim_match.group(2))
            params['height_mm'] = float(dim_match.group(3))
        
        return params
    
    def _extract_materials_query(self, message: str, match: re.Match) -> Optional[str]:
        """Извлекает запрос для поиска материалов."""
        # Ищем числа, которые могут быть названиями материалов
        material_match = re.search(r'\b(\d+[А-Яа-я]*)\b', message)
        if material_match:
            return material_match.group(1)
        
        # Ищем после ключевых слов
        message_lower = message.lower()
        for keyword in ['материал', 'gb', 'гб', 'сталь']:
            if keyword in message_lower:
                parts = message_lower.split(keyword)
                if len(parts) > 1:
                    next_part = parts[1].strip().split()[0] if parts[1].strip() else None
                    if next_part:
                        return next_part
        
        return None
    
    def _extract_duty_query(self, message: str, match: re.Match) -> Optional[str]:
        """Извлекает запрос для поиска пошлин."""
        # Ищем код ТН ВЭД
        code_match = re.search(r'\b(\d{4,10})\b', message)
        if code_match:
            return code_match.group(1)
        
        # Ищем название товара
        message_lower = message.lower()
        for keyword in ['пошлина на', 'пошлина для', 'таможенн']:
            if keyword in message_lower:
                parts = message_lower.split(keyword)
                if len(parts) > 1:
                    next_part = parts[1].strip().split()[0:3]
                    if next_part:
                        return ' '.join(next_part)
        
        return None
    
    def _extract_analytics_params(
        self, message: str, match: re.Match, metadata: Dict[str, Any]
    ) -> Dict[str, Any]:
        """Извлекает параметры для аналитики."""
        message_lower = message.lower()
        params = {}
        
        # Определяем подтип запроса
        subtype = metadata.get('subtype')
        if subtype:
            params['subtype'] = subtype
        
        # Извлекаем категорию из паттерна, если есть
        if match.groups():
            category_candidate = match.group(1) if len(match.groups()) >= 1 else None
            if category_candidate:
                categories = self.analytics_helper.get_categories()
                category_candidate_lower = category_candidate.lower()
                for cat in categories:
                    if cat.lower() == category_candidate_lower:
                        params['category'] = cat
                        break
        
        # Если не нашли в паттерне, ищем в сообщении
        if 'category' not in params:
            categories = self.analytics_helper.get_categories()
            for cat in categories:
                cat_lower = cat.lower()
                # Прямое совпадение
                if cat_lower in message_lower:
                    params['category'] = cat
                    break
                # Проверяем начало слова (для склонений: Барабаны -> барабан)
                cat_base = cat_lower.rstrip('ы').rstrip('и').rstrip('а').rstrip('я')
                if len(cat_base) >= 3:  # Минимум 3 символа для базы
                    # Ищем слова, начинающиеся с базы категории
                    words = message_lower.split()
                    for word in words:
                        # Убираем знаки препинания
                        word_clean = word.strip('.,!?;:()[]{}"\'')
                        if word_clean.startswith(cat_base) and len(word_clean) >= len(cat_base):
                            params['category'] = cat
                            break
                    if 'category' in params:
                        break
        
        # Специфичные параметры для разных подтипов
        if subtype == 'top_managers':
            params['by_sum'] = metadata.get('by_sum', False)
        elif subtype == 'search':
            search_type = metadata.get('search_type')
            if search_type:
                params['search_type'] = search_type
                # Извлекаем поисковый запрос
                if match.groups():
                    params['search_query'] = match.group(1).strip()
        
        # Проверяем, нужны ли графики
        chart_keywords = ['график', 'визуализируй', 'покажи график', 'диаграмм']
        params['need_chart'] = any(keyword in message_lower for keyword in chart_keywords)
        
        # Дополнительное извлечение поискового запроса, если не нашли
        if 'search_query' not in params and subtype == 'search':
            params['search_query'] = self._extract_search_query(message_lower)
        
        return params
    
    def _extract_weight(self, message_lower: str) -> Optional[float]:
        """Извлекает вес из сообщения."""
        weight_patterns = [
            (r'(\d+(?:\s+\d+)*(?:[.,]\d+)?)\s*(?:кг|kg)(?![а-яёa-z])', False),
            (r'(\d+(?:\s+\d+)*(?:[.,]\d+)?)\s*(?:тонн|т)(?![а-яёa-z])', True),
        ]
        
        for pattern, is_tons in weight_patterns:
            matches = re.findall(pattern, message_lower, re.UNICODE)
            if matches:
                weight_str = max(matches, key=len)
                weight = self._parse_weight(weight_str)
                if weight:
                    if is_tons:
                        weight *= 1000
                    return weight
        
        return None
    
    def _parse_weight(self, weight_str: str) -> Optional[float]:
        """Парсит строку с весом."""
        try:
            weight_str = weight_str.strip()
            weight_str = re.sub(r'(\d)\s+(\d)', r'\1\2', weight_str)
            
            has_comma = ',' in weight_str
            has_dot = '.' in weight_str
            
            if has_comma and has_dot:
                comma_pos = weight_str.rfind(',')
                dot_pos = weight_str.rfind('.')
                if comma_pos > dot_pos:
                    weight_str = weight_str.replace('.', '').replace(',', '.')
                else:
                    weight_str = weight_str.replace(',', '')
            elif has_comma:
                parts = weight_str.split(',')
                if len(parts) == 2 and len(parts[1]) == 3 and len(parts[0]) > 0:
                    weight_str = weight_str.replace(',', '')
                else:
                    weight_str = weight_str.replace(',', '.')
            elif has_dot:
                parts = weight_str.split('.')
                if len(parts) == 2 and len(parts[1]) == 3 and len(parts[0]) > 0:
                    weight_str = weight_str.replace('.', '')
            
            return float(weight_str)
        except (ValueError, AttributeError):
            return None
    
    def _extract_city(self, message_lower: str) -> Optional[str]:
        """Извлекает город из сообщения."""
        cities = [c.get('name') for c in self.logistics_helper._load_cities()]
        city_aliases = {
            'екб': 'Екатеринбург',
            'мск': 'Москва',
            'спб': 'Санкт-Петербург',
            'питер': 'Санкт-Петербург',
        }
        
        # Проверяем сокращения
        for alias, city_name in city_aliases.items():
            if alias in message_lower and city_name in cities:
                return city_name
        
        # Ищем после предлогов
        preposition_patterns = [
            r'(?:из\s+кнр\s+)?в\s+(\w+(?:\s+\w+)*)',
            r'(?:из\s+кнр\s+)?до\s+(\w+(?:\s+\w+)*)',
            r'на\s+(\w+(?:\s+\w+)*)',
        ]
        
        for pattern in preposition_patterns:
            match = re.search(pattern, message_lower, re.UNICODE)
            if match:
                potential_city = match.group(1).strip()
                for city in cities:
                    city_lower = city.lower()
                    if city_lower == potential_city or potential_city.startswith(city_lower):
                        return city
        
        # Поиск по границам слов
        for city in cities:
            city_lower = city.lower()
            city_pattern = r'\b' + re.escape(city_lower) + r'\b'
            if re.search(city_pattern, message_lower, re.UNICODE):
                return city
        
        # Частичное совпадение
        for city in cities:
            city_lower = city.lower()
            if city_lower in message_lower:
                return city
        
        return None
    
    def _extract_search_query(self, message_lower: str) -> Optional[str]:
        """Извлекает поисковый запрос из сообщения."""
        search_patterns = [
            r'найди\s+(?:в\s+1с\s+)?(?:по\s+)?(?:чертежу|чертеж)\s+([^\s]+(?:\s+[^\s]+)*)',
            r'найди\s+(?:в\s+1с\s+)?(?:барабан|вал|бандаж|номенклатур[ауе])\s+([^\s]+(?:\s+[^\s]+)*)',
            r'найди\s+чертеж\s+([^\s]+(?:\s+[^\s]+)*)',
            r'покажи\s+(?:данные\s+)?(?:по\s+)?(?:заказчику|заказчик)\s+([^\s]+)',
        ]
        
        for pattern in search_patterns:
            match = re.search(pattern, message_lower, re.UNICODE)
            if match:
                return match.group(1).strip()
        
        # Fallback: ищем после ключевых слов
        for keyword in ['найди', 'покажи', 'дай', 'найти']:
            if keyword in message_lower:
                parts = message_lower.split(keyword)
                if len(parts) > 1:
                    next_part = parts[1].strip()
                    next_part = re.sub(r'\b(в\s+1с|по|чертежу|заказчику|материалу|номенклатуре)\b', '', next_part).strip()
                    if next_part:
                        return next_part.split()[0]
        
        return None
    
    def _extract_materials_query_fallback(self, message_lower: str) -> Optional[str]:
        """Fallback метод для извлечения запроса материалов."""
        # Ищем числа, которые могут быть названиями материалов
        material_match = re.search(r'\b(\d+[а-яё]*)\b', message_lower, re.UNICODE)
        if material_match:
            return material_match.group(1)
        
        # Ищем после ключевых слов
        for keyword in ['материал', 'gb', 'гб', 'сталь']:
            if keyword in message_lower:
                parts = message_lower.split(keyword)
                if len(parts) > 1:
                    next_part = parts[1].strip().split()[0] if parts[1].strip() else None
                    if next_part:
                        return next_part
        
        return None
    
    def _extract_duty_query_fallback(self, message_lower: str) -> Optional[str]:
        """Fallback метод для извлечения запроса пошлин."""
        # Ищем код ТН ВЭД
        code_match = re.search(r'\b(\d{4,10})\b', message_lower, re.UNICODE)
        if code_match:
            return code_match.group(1)
        
        # Ищем название товара
        for keyword in ['пошлина на', 'пошлина для', 'таможенн']:
            if keyword in message_lower:
                parts = message_lower.split(keyword)
                if len(parts) > 1:
                    next_part = parts[1].strip().split()[0:3]
                    if next_part:
                        return ' '.join(next_part)
        
        return None
    
    def _extract_analytics_params_fallback(self, message_lower: str) -> Dict[str, Any]:
        """Fallback метод для извлечения параметров аналитики."""
        params = {}
        
        # Определяем категорию
        categories = self.analytics_helper.get_categories()
        for cat in categories:
            cat_lower = cat.lower()
            # Прямое совпадение
            if cat_lower in message_lower:
                params['category'] = cat
                break
            # Проверяем начало слова (для склонений: Барабаны -> барабан)
            cat_base = cat_lower.rstrip('ы').rstrip('и').rstrip('а').rstrip('я')
            if len(cat_base) >= 3:  # Минимум 3 символа для базы
                # Ищем слова, начинающиеся с базы категории
                words = message_lower.split()
                for word in words:
                    # Убираем знаки препинания
                    word_clean = word.strip('.,!?;:()[]{}"\'')
                    if word_clean.startswith(cat_base) and len(word_clean) >= len(cat_base):
                        params['category'] = cat
                        break
                if 'category' in params:
                    break
        
        # Проверяем, нужны ли графики
        chart_keywords = ['график', 'визуализируй', 'покажи график', 'диаграмм']
        params['need_chart'] = any(keyword in message_lower for keyword in chart_keywords)
        
        # Извлекаем поисковый запрос
        search_query = self._extract_search_query(message_lower)
        if search_query:
            params['search_query'] = search_query
        
        return params
    
    def _fallback_extraction(self, message: str, result: Dict[str, Any]) -> Dict[str, Any]:
        """
        Fallback метод для извлечения намерений (старая логика).
        Используется, если паттерны не сработали.
        """
        message_lower = message.lower()
        
        # Простые ключевые слова
        logistics_keywords = [
            'логистик', 'рассчитай', 'расчет', 'доставк', 'перевозк',
            'стоимость доставки', 'цена доставки', 'логистика для',
            'сколько стоит', 'стоимость перевозки'
        ]
        result['is_logistics'] = any(keyword in message_lower for keyword in logistics_keywords)
        
        materials_keywords = [
            'материал', 'gb', 'гб', 'аналог', 'сталь', 'металл',
            'найди материал', 'какой gb', 'цена на', 'стоимость материала',
            'gost', 'гост', 'заготовк'
        ]
        result['is_materials'] = any(keyword in message_lower for keyword in materials_keywords)
        
        duty_keywords = [
            'пошлин', 'таможенн', 'код тн вэд', 'тнвэд', 'tnved',
            'какая пошлина', 'ставка пошлины', 'пошлина на', 'пошлина для'
        ]
        result['is_duty'] = any(keyword in message_lower for keyword in duty_keywords)
        
        analytics_keywords = [
            '1с', '1c', 'найди в 1с', 'покажи данные по', 'сколько записей',
            'топ заказчиков', 'топ материалов', 'статистика по',
            'график', 'визуализируй', 'покажи график',
            'менеджер', 'топ менеджеров', 'лучший менеджер',
            'тендер', 'выигран', 'выиграли', 'проиграли', 'отменен',
            'самый большой тендер', 'крупный тендер', 'большой заказ'
        ]
        result['is_analytics'] = any(keyword in message_lower for keyword in analytics_keywords)
        
        # Извлекаем параметры для каждого типа
        if result['is_logistics']:
            result['logistics_params'] = {
                'weight_kg': self._extract_weight(message_lower),
                'city_name': self._extract_city(message_lower),
                'transport_type': 'trail' if 'трал' in message_lower else 'truck'
            }
        
        if result['is_materials']:
            result['materials_query'] = self._extract_materials_query_fallback(message_lower)
        
        if result['is_duty']:
            result['duty_query'] = self._extract_duty_query_fallback(message_lower)
        
        if result['is_analytics']:
            result['analytics_params'] = self._extract_analytics_params_fallback(message_lower)
        
        return result

