"""
Модуль-обертка для поиска по справочникам пошлин.
Упрощает использование справочников duty_rates и TNVED для AI агента.
"""

import sys
from pathlib import Path
from typing import Any, Dict, List, Optional

# Добавляем корень проекта в путь для импорта
BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR))

from app.services import datasets


class DutyHelper:
    """Класс-помощник для поиска пошлин в справочниках."""
    
    def __init__(self):
        """Инициализирует помощник пошлин."""
        self.duty_catalog_cache: Optional[List[Dict[str, Any]]] = None
    
    def _load_catalog(self) -> List[Dict[str, Any]]:
        """Загружает объединенный каталог пошлин (с кэшированием)."""
        if self.duty_catalog_cache is None:
            try:
                self.duty_catalog_cache = datasets.get_duty_catalog()
            except Exception as e:
                print(f"Ошибка при загрузке каталога пошлин: {e}")
                self.duty_catalog_cache = []
        return self.duty_catalog_cache
    
    def search_by_product(self, query: str, limit: int = 10) -> List[Dict[str, Any]]:
        """
        Ищет пошлины по названию товара/продукта.
        
        Args:
            query: Поисковый запрос
            limit: Максимальное количество результатов
            
        Returns:
            Список найденных записей
        """
        catalog = self._load_catalog()
        query_lower = query.lower().strip()
        results = []
        
        for item in catalog:
            product = str(item.get('product', '') or item.get('keywords_display', '')).lower()
            if query_lower in product or product in query_lower:
                results.append(item)
                if len(results) >= limit:
                    break
        
        return results
    
    def search_by_category(self, query: str, limit: int = 10) -> List[Dict[str, Any]]:
        """
        Ищет пошлины по категории.
        
        Args:
            query: Поисковый запрос
            limit: Максимальное количество результатов
            
        Returns:
            Список найденных записей
        """
        catalog = self._load_catalog()
        query_lower = query.lower().strip()
        results = []
        
        for item in catalog:
            category = str(item.get('category', '') or item.get('description', '')).lower()
            if query_lower in category:
                results.append(item)
                if len(results) >= limit:
                    break
        
        return results
    
    def search_by_code(self, code: str, limit: int = 10) -> List[Dict[str, Any]]:
        """
        Ищет пошлины по коду ТН ВЭД.
        
        Args:
            code: Код ТН ВЭД
            limit: Максимальное количество результатов
            
        Returns:
            Список найденных записей
        """
        catalog = self._load_catalog()
        code_upper = code.strip().upper()
        results = []
        
        for item in catalog:
            item_code = str(item.get('code', '')).strip().upper()
            if code_upper in item_code or item_code in code_upper:
                results.append(item)
                if len(results) >= limit:
                    break
        
        return results
    
    def search_by_keywords(self, query: str, limit: int = 10) -> List[Dict[str, Any]]:
        """
        Ищет пошлины по ключевым словам.
        
        Args:
            query: Поисковый запрос
            limit: Максимальное количество результатов
            
        Returns:
            Список найденных записей
        """
        catalog = self._load_catalog()
        query_lower = query.lower().strip()
        results = []
        
        for item in catalog:
            # Проверяем ключевые слова
            keywords = item.get('keywords_list', [])
            keywords_text = str(item.get('keywords_display', '')).lower()
            search_blob = str(item.get('search_blob', '')).lower()
            
            if (query_lower in keywords_text or 
                query_lower in search_blob or
                any(query_lower in str(kw).lower() for kw in keywords)):
                results.append(item)
                if len(results) >= limit:
                    break
        
        return results
    
    def search_all_fields(self, query: str, limit: int = 10) -> List[Dict[str, Any]]:
        """
        Ищет пошлины по всем полям.
        
        Args:
            query: Поисковый запрос
            limit: Максимальное количество результатов
            
        Returns:
            Список найденных записей
        """
        catalog = self._load_catalog()
        query_lower = query.lower().strip()
        results = []
        
        for item in catalog:
            # Проверяем все поисковые поля
            search_blob = str(item.get('search_blob', '')).lower()
            if query_lower in search_blob:
                results.append(item)
                if len(results) >= limit:
                    break
        
        return results
    
    def find_duty(self, query: str, search_type: Optional[str] = None) -> List[Dict[str, Any]]:
        """
        Универсальный поиск пошлин.
        
        Args:
            query: Поисковый запрос
            search_type: Тип поиска ('product', 'category', 'code', 'keywords', 'all' или None)
            
        Returns:
            Список найденных записей
        """
        # Если запрос похож на код ТН ВЭД (только цифры)
        if query.strip().isdigit():
            return self.search_by_code(query)
        
        if search_type == 'product':
            return self.search_by_product(query)
        elif search_type == 'category':
            return self.search_by_category(query)
        elif search_type == 'code':
            return self.search_by_code(query)
        elif search_type == 'keywords':
            return self.search_by_keywords(query)
        else:
            # По умолчанию ищем по всем полям
            return self.search_all_fields(query)
    
    def format_duty_result(self, items: List[Dict[str, Any]]) -> str:
        """
        Форматирует результаты поиска пошлин для вывода.
        
        Args:
            items: Список найденных записей
            
        Returns:
            Отформатированная строка с результатами
        """
        if not items:
            return "❌ Пошлины не найдены. Попробуйте изменить поисковый запрос или уточните название товара/категории."
        
        if len(items) == 1:
            item = items[0]
            result = "✅ Найдена пошлина:\n\n"
            
            # Код ТН ВЭД (если есть)
            code = item.get('code', '')
            if code:
                result += f"**Код ТН ВЭД:** {code}\n"
            
            # Название/ключевые слова
            product = item.get('product') or item.get('keywords_display', '')
            if product:
                result += f"**Товар:** {product}\n"
            
            # Категория/описание
            category = item.get('category') or item.get('description', '')
            if category:
                result += f"**Категория:** {category}\n"
            
            # Пошлина
            duty_percent = item.get('duty_percent')
            duty_text = item.get('duty_text', '')
            if duty_percent is not None:
                result += f"**Пошлина:** {duty_percent}%\n"
            elif duty_text:
                result += f"**Пошлина:** {duty_text}\n"
            else:
                result += "**Пошлина:** не указана\n"
            
            # Примеры (если есть)
            examples = item.get('examples', '')
            if examples:
                result += f"**Примеры:** {examples}\n"
            
            # Источник
            source = item.get('source', '')
            if source == 'manual':
                result += "\n📝 Источник: Ручной справочник"
            elif source == 'tnved':
                result += "\n📋 Источник: Каталог ТН ВЭД"
            
            return result
        
        # Несколько результатов
        result = f"✅ Найдено записей: {len(items)}\n\n"
        for i, item in enumerate(items[:10], 1):  # Показываем максимум 10
            code = item.get('code', '')
            product = item.get('product') or item.get('keywords_display', '') or item.get('title', '')
            category = item.get('category') or item.get('description', '')
            duty_percent = item.get('duty_percent')
            duty_text = item.get('duty_text', '')
            
            result += f"{i}. "
            if code:
                result += f"**{code}** - "
            result += f"{product or category or 'Без названия'}"
            
            if duty_percent is not None:
                result += f" - {duty_percent}%"
            elif duty_text:
                result += f" - {duty_text}"
            
            result += "\n"
        
        if len(items) > 10:
            result += f"\n... и еще {len(items) - 10} результатов. Уточните запрос для более точного поиска."
        
        return result

