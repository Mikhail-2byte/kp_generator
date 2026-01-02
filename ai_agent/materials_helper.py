"""
Модуль-обертка для поиска по справочнику материалов.
Упрощает использование справочника gb_materials для AI агента.
"""

import sys
from pathlib import Path
from typing import Any, Dict, List, Optional

# Добавляем корень проекта в путь для импорта
BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR))

from app.services import datasets


class MaterialsHelper:
    """Класс-помощник для поиска материалов в справочнике."""
    
    def __init__(self):
        """Инициализирует помощник материалов."""
        self.materials_cache: Optional[List[Dict[str, Any]]] = None
    
    def _load_materials(self) -> List[Dict[str, Any]]:
        """Загружает список материалов (с кэшированием)."""
        if self.materials_cache is None:
            try:
                self.materials_cache = datasets.load_gb_materials()
            except Exception as e:
                print(f"Ошибка при загрузке материалов: {e}")
                self.materials_cache = []
        return self.materials_cache
    
    def search_by_russian(self, query: str, limit: int = 10) -> List[Dict[str, Any]]:
        """
        Ищет материалы по русскому названию.
        
        Args:
            query: Поисковый запрос
            limit: Максимальное количество результатов
            
        Returns:
            Список найденных материалов
        """
        materials = self._load_materials()
        query_lower = query.lower().strip()
        results = []
        
        for material in materials:
            russian = str(material.get('russian', '')).lower()
            if query_lower in russian or russian in query_lower:
                results.append(material)
                if len(results) >= limit:
                    break
        
        return results
    
    def search_by_gb(self, query: str, limit: int = 10) -> List[Dict[str, Any]]:
        """
        Ищет материалы по GB стандарту.
        
        Args:
            query: Поисковый запрос
            limit: Максимальное количество результатов
            
        Returns:
            Список найденных материалов
        """
        materials = self._load_materials()
        query_lower = query.lower().strip()
        results = []
        
        for material in materials:
            gb = str(material.get('gb', '')).lower()
            # Игнорируем предупреждения в поле GB
            if 'не использовать' in gb:
                continue
            if query_lower in gb or gb in query_lower:
                results.append(material)
                if len(results) >= limit:
                    break
        
        return results
    
    def search_by_gost(self, query: str, limit: int = 10) -> List[Dict[str, Any]]:
        """
        Ищет материалы по ГОСТ.
        
        Args:
            query: Поисковый запрос
            limit: Максимальное количество результатов
            
        Returns:
            Список найденных материалов
        """
        materials = self._load_materials()
        query_lower = query.lower().strip()
        results = []
        
        for material in materials:
            gost = str(material.get('gost', '')).lower()
            if query_lower in gost:
                results.append(material)
                if len(results) >= limit:
                    break
        
        return results
    
    def search_all_fields(self, query: str, limit: int = 10) -> List[Dict[str, Any]]:
        """
        Ищет материалы по всем полям.
        
        Args:
            query: Поисковый запрос
            limit: Максимальное количество результатов
            
        Returns:
            Список найденных материалов
        """
        materials = self._load_materials()
        query_lower = query.lower().strip()
        results = []
        
        for material in materials:
            # Проверяем все поля
            russian = str(material.get('russian', '')).lower()
            gb = str(material.get('gb', '')).lower()
            gost = str(material.get('gost', '')).lower()
            price = str(material.get('price', '')).lower()
            workpiece_type = str(material.get('workpiece_type', '')).lower()
            
            if (query_lower in russian or 
                query_lower in gb or 
                query_lower in gost or 
                query_lower in price or 
                query_lower in workpiece_type):
                results.append(material)
                if len(results) >= limit:
                    break
        
        return results
    
    def find_material(self, query: str, search_field: Optional[str] = None) -> List[Dict[str, Any]]:
        """
        Универсальный поиск материалов.
        
        Args:
            query: Поисковый запрос
            search_field: Поле для поиска ('russian', 'gb', 'gost', 'all' или None)
            
        Returns:
            Список найденных материалов
        """
        if search_field == 'russian':
            return self.search_by_russian(query)
        elif search_field == 'gb':
            return self.search_by_gb(query)
        elif search_field == 'gost':
            return self.search_by_gost(query)
        else:
            # По умолчанию ищем по всем полям
            return self.search_all_fields(query)
    
    def format_material_result(self, materials: List[Dict[str, Any]]) -> str:
        """
        Форматирует результаты поиска материалов для вывода.
        
        Args:
            materials: Список найденных материалов
            
        Returns:
            Отформатированная строка с результатами
        """
        if not materials:
            return "❌ Материалы не найдены. Попробуйте изменить поисковый запрос."
        
        if len(materials) == 1:
            material = materials[0]
            result = f"✅ Найден материал:\n\n"
            result += f"**Русское название:** {material.get('russian', '—')}\n"
            result += f"**GB стандарт:** {material.get('gb', '—')}\n"
            result += f"**ГОСТ:** {material.get('gost', '—')}\n"
            
            price = material.get('price', '')
            if price:
                result += f"**Цена:** {price} юань/кг\n"
            else:
                result += f"**Цена:** не указана\n"
            
            workpiece_type = material.get('workpiece_type', '')
            if workpiece_type:
                result += f"**Тип заготовки:** {workpiece_type}\n"
            
            notes = material.get('notes', '')
            if notes:
                result += f"**Примечания:** {notes}\n"
            
            # Проверяем предупреждения
            gb = str(material.get('gb', '')).lower()
            if 'не использовать' in gb:
                result += f"\n⚠️ **ВНИМАНИЕ:** Этот материал не рекомендуется к использованию. "
                if 'замена' in gb:
                    result += "Рекомендуется замена.\n"
            
            return result
        
        # Несколько результатов
        result = f"✅ Найдено материалов: {len(materials)}\n\n"
        for i, material in enumerate(materials[:10], 1):  # Показываем максимум 10
            result += f"{i}. **{material.get('russian', '—')}** "
            result += f"(GB: {material.get('gb', '—')}, "
            result += f"ГОСТ: {material.get('gost', '—')})"
            
            price = material.get('price', '')
            if price:
                result += f" - {price} юань/кг"
            
            workpiece_type = material.get('workpiece_type', '')
            if workpiece_type:
                result += f" ({workpiece_type})"
            
            result += "\n"
        
        if len(materials) > 10:
            result += f"\n... и еще {len(materials) - 10} результатов. Уточните запрос для более точного поиска."
        
        return result

