"""
Модуль для валидации параметров, извлеченных из пользовательских запросов.
"""

import logging
import re
from typing import Any, Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)


class ParameterValidator:
    """Класс для валидации параметров запросов."""
    
    # Диапазоны допустимых значений
    WEIGHT_MIN_KG = 0.1  # Минимальный вес 100 грамм
    WEIGHT_MAX_KG = 1_000_000  # Максимальный вес 1000 тонн
    
    # Допустимые форматы
    DRAWING_PATTERN = re.compile(r'^[\w\.\-\s]+$', re.UNICODE)
    CITY_NAME_PATTERN = re.compile(r'^[А-Яа-яЁё\s\-]+$', re.UNICODE)
    
    def validate_logistics_params(self, params: Dict[str, Any]) -> Tuple[bool, Optional[str], Dict[str, Any]]:
        """
        Валидирует параметры логистики.
        
        Args:
            params: Словарь с параметрами логистики
            
        Returns:
            Кортеж (is_valid, error_message, corrected_params)
        """
        corrected_params = params.copy()
        errors = []
        
        # Валидация веса
        weight = params.get('weight_kg')
        if weight is not None:
            if not isinstance(weight, (int, float)):
                errors.append("Вес должен быть числом")
            elif weight < self.WEIGHT_MIN_KG:
                errors.append(f"Вес слишком мал (минимум {self.WEIGHT_MIN_KG} кг)")
            elif weight > self.WEIGHT_MAX_KG:
                errors.append(f"Вес слишком велик (максимум {self.WEIGHT_MAX_KG} кг)")
                # Попытка исправления: возможно, вес был указан в граммах
                if weight > self.WEIGHT_MAX_KG * 1000:
                    corrected_weight = weight / 1000
                    if self.WEIGHT_MIN_KG <= corrected_weight <= self.WEIGHT_MAX_KG:
                        corrected_params['weight_kg'] = corrected_weight
                        errors.append(f"Исправлено: вес был указан в граммах, переведен в {corrected_weight:.2f} кг")
        
        # Валидация города
        city_name = params.get('city_name')
        if city_name:
            if not isinstance(city_name, str):
                errors.append("Название города должно быть строкой")
            elif not self.CITY_NAME_PATTERN.match(city_name):
                errors.append("Название города содержит недопустимые символы")
        
        # Валидация типа транспорта
        transport_type = params.get('transport_type')
        if transport_type and transport_type not in ['truck', 'trail']:
            errors.append(f"Неизвестный тип транспорта: {transport_type}")
            corrected_params['transport_type'] = 'truck'  # Дефолт
        
        # Валидация габаритов
        for dim_name in ['length_mm', 'width_mm', 'height_mm']:
            dim_value = params.get(dim_name)
            if dim_value is not None:
                if not isinstance(dim_value, (int, float)):
                    errors.append(f"{dim_name} должен быть числом")
                elif dim_value <= 0:
                    errors.append(f"{dim_name} должен быть положительным числом")
                elif dim_value > 20_000:  # Максимум 20 метров
                    errors.append(f"{dim_name} слишком велик (максимум 20000 мм)")
        
        is_valid = len(errors) == 0
        error_message = "; ".join(errors) if errors else None
        
        return is_valid, error_message, corrected_params
    
    def validate_analytics_params(self, params: Dict[str, Any]) -> Tuple[bool, Optional[str], Dict[str, Any]]:
        """
        Валидирует параметры аналитики.
        
        Args:
            params: Словарь с параметрами аналитики
            
        Returns:
            Кортеж (is_valid, error_message, corrected_params)
        """
        corrected_params = params.copy()
        errors = []
        
        # Валидация категории
        category = params.get('category')
        if category:
            if not isinstance(category, str):
                errors.append("Категория должна быть строкой")
        
        # Валидация поискового запроса
        search_query = params.get('search_query')
        if search_query:
            if not isinstance(search_query, str):
                errors.append("Поисковый запрос должен быть строкой")
            elif len(search_query.strip()) == 0:
                errors.append("Поисковый запрос не может быть пустым")
            else:
                # Очищаем запрос от лишних пробелов
                corrected_params['search_query'] = search_query.strip()
        
        # Валидация чертежа (если это поиск по чертежу)
        if params.get('search_type') == 'drawing' and search_query:
            if not self.DRAWING_PATTERN.match(search_query):
                errors.append("Номер чертежа содержит недопустимые символы")
        
        # Валидация флагов
        if 'need_chart' in params and not isinstance(params['need_chart'], bool):
            corrected_params['need_chart'] = bool(params['need_chart'])
        
        if 'by_sum' in params and not isinstance(params['by_sum'], bool):
            corrected_params['by_sum'] = bool(params['by_sum'])
        
        is_valid = len(errors) == 0
        error_message = "; ".join(errors) if errors else None
        
        return is_valid, error_message, corrected_params
    
    def validate_materials_query(self, query: Optional[str]) -> Tuple[bool, Optional[str], Optional[str]]:
        """
        Валидирует запрос для поиска материалов.
        
        Args:
            query: Поисковый запрос
            
        Returns:
            Кортеж (is_valid, error_message, corrected_query)
        """
        if query is None:
            return True, None, None
        
        if not isinstance(query, str):
            return False, "Запрос должен быть строкой", None
        
        query = query.strip()
        if len(query) == 0:
            return False, "Запрос не может быть пустым", None
        
        if len(query) > 100:
            return False, "Запрос слишком длинный (максимум 100 символов)", query[:100]
        
        return True, None, query
    
    def validate_duty_query(self, query: Optional[str]) -> Tuple[bool, Optional[str], Optional[str]]:
        """
        Валидирует запрос для поиска пошлин.
        
        Args:
            query: Поисковый запрос
            
        Returns:
            Кортеж (is_valid, error_message, corrected_query)
        """
        if query is None:
            return True, None, None
        
        if not isinstance(query, str):
            return False, "Запрос должен быть строкой", None
        
        query = query.strip()
        if len(query) == 0:
            return False, "Запрос не может быть пустым", None
        
        # Если это код ТН ВЭД (только цифры)
        if query.isdigit():
            if len(query) < 4 or len(query) > 10:
                return False, "Код ТН ВЭД должен содержать от 4 до 10 цифр", None
        
        return True, None, query
    
    def generate_validation_error_message(
        self, 
        intent_type: str, 
        errors: List[str],
        examples: Optional[List[str]] = None
    ) -> str:
        """
        Генерирует понятное сообщение об ошибке валидации.
        
        Args:
            intent_type: Тип намерения (logistics, analytics, etc.)
            errors: Список ошибок
            examples: Примеры корректных запросов
            
        Returns:
            Отформатированное сообщение об ошибке
        """
        message_parts = [
            "❌ Обнаружены ошибки в параметрах запроса:",
            ""
        ]
        
        for i, error in enumerate(errors, 1):
            message_parts.append(f"{i}. {error}")
        
        if examples:
            message_parts.append("")
            message_parts.append("Примеры корректных запросов:")
            for example in examples:
                message_parts.append(f"• {example}")
        
        return "\n".join(message_parts)


