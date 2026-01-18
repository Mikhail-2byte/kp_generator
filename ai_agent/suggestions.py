"""
Модуль для генерации умных предложений и альтернатив при пустых результатах.
Поддерживает автодополнение, популярные запросы, интерактивные подсказки.
"""

import logging
from collections import Counter
from typing import Dict, List, Optional

logger = logging.getLogger(__name__)


class SuggestionsGenerator:
    """Класс для генерации предложений и альтернатив."""
    
    @staticmethod
    def generate_empty_result_message(
        query_type: str,
        query: Optional[str] = None,
        category: Optional[str] = None
    ) -> str:
        """
        Генерирует дружелюбное сообщение при отсутствии результатов.
        
        Args:
            query_type: Тип запроса (search, statistics, top, etc.)
            query: Поисковый запрос
            category: Категория поиска
            
        Returns:
            Сообщение с предложениями
        """
        messages = []
        
        if query_type == 'search':
            messages.append("По вашему запросу ничего не найдено.")
            
            if query:
                messages.append(f"\n**Запрос**: {query}")
            
            if category:
                messages.append(f"**Категория**: {category}")
            
            messages.append("\n**Возможные причины:**")
            messages.append("• Запрос содержит опечатки")
            messages.append("• Данные отсутствуют в указанной категории")
            messages.append("• Используется другое название или формат")
            
            messages.append("\n**Что можно попробовать:**")
            messages.append("• Проверьте правильность написания")
            messages.append("• Попробуйте более общий запрос")
            messages.append("• Убедитесь, что выбрана правильная категория")
            messages.append("• Попробуйте поиск по другому полю (чертеж, заказчик, материал)")
        
        elif query_type == 'statistics':
            messages.append("Статистика не найдена.")
            
            if category:
                messages.append(f"\n**Категория**: {category}")
                messages.append("Возможно, в этой категории нет данных.")
            else:
                messages.append("Возможно, данные еще не загружены.")
            
            messages.append("\n**Что можно попробовать:**")
            messages.append("• Проверьте доступные категории")
            messages.append("• Убедитесь, что данные загружены")
        
        elif query_type == 'top':
            messages.append("Топ-лист не найден.")
            
            if category:
                messages.append(f"\n**Категория**: {category}")
            
            messages.append("\n**Что можно попробовать:**")
            messages.append("• Проверьте доступные категории")
            messages.append("• Убедитесь, что в данных есть соответствующие записи")
        
        else:
            messages.append("Результаты не найдены.")
            messages.append("\nПопробуйте уточнить запрос или обратитесь к администратору.")
        
        return "\n".join(messages)
    
    @staticmethod
    def suggest_alternatives(
        query: str,
        available_categories: List[str],
        search_fields: List[str] = None
    ) -> List[str]:
        """
        Предлагает альтернативные варианты запроса.
        
        Args:
            query: Исходный запрос
            available_categories: Список доступных категорий
            search_fields: Список доступных полей для поиска
            
        Returns:
            Список предложений
        """
        suggestions = []
        
        # Предложения по категориям
        if available_categories:
            suggestions.append(f"**Доступные категории**: {', '.join(available_categories)}")
        
        # Предложения по полям поиска
        if search_fields:
            suggestions.append(f"**Можно искать по**: {', '.join(search_fields)}")
        
        # Предложения по формату запроса
        suggestions.append("**Примеры запросов:**")
        suggestions.append("• \"Найди в 1с чертеж [номер]\"")
        suggestions.append("• \"Покажи все записи по заказчику [название]\"")
        suggestions.append("• \"Найди все позиции с материалом [название]\"")
        suggestions.append("• \"Топ заказчиков\"")
        suggestions.append("• \"Статистика по тендерам\"")
        
        return suggestions
    
    @staticmethod
    def generate_help_message(query_type: str) -> str:
        """
        Генерирует справочное сообщение для типа запроса.
        
        Args:
            query_type: Тип запроса
            
        Returns:
            Справочное сообщение
        """
        help_messages = {
            'search': """
**Как искать в данных 1С:**

1. **По номенклатуре**: "Найди в 1с барабан самоцентрирующийся"
2. **По чертежу**: "Найди чертеж 35.1774.000 СБ"
3. **По заказчику**: "Покажи все записи по заказчику ВАЙЛДБЕРРИЗ"
4. **По материалу**: "Найди все позиции с материалом СБ"
5. **По статусу тендера**: "Покажи выигранные тендеры"

Можно указать категорию: "Найди в Барабанах чертеж X"
""",
            'statistics': """
**Статистика по данным 1С:**

• "Сколько записей в Барабанах"
• "Статистика по всем категориям"
• "Статистика по тендерам"

Показывает общее количество записей, вес, стоимость, уникальных заказчиков и материалов.
""",
            'top': """
**Топ-листы:**

• "Топ заказчиков"
• "Топ материалов"
• "Топ менеджеров по выигранным тендерам"
• "Топ менеджеров по сумме выигранных тендеров"

Можно указать категорию и запросить график: "Топ заказчиков с графиком"
"""
        }
        
        return help_messages.get(query_type, "Попробуйте уточнить запрос.")
    
    @staticmethod
    def suggest_similar_queries(
        query: str,
        query_type: str
    ) -> List[str]:
        """
        Предлагает похожие запросы на основе исходного.
        
        Args:
            query: Исходный запрос
            query_type: Тип запроса
            
        Returns:
            Список похожих запросов
        """
        suggestions = []
        
        if query_type == 'search':
            # Предлагаем варианты поиска
            if len(query) > 3:
                # Предлагаем более короткий запрос
                suggestions.append(f"Попробуйте более короткий запрос: \"{query[:3]}...\"")
            
            # Предлагаем поиск по другим полям
            suggestions.append("Попробуйте поиск по другому полю:")
            suggestions.append("• По чертежу вместо номенклатуры")
            suggestions.append("• По заказчику вместо материала")
        
        return suggestions
    
    @staticmethod
    def get_popular_queries(
        query_history: Optional[List[str]] = None,
        limit: int = 5
    ) -> List[str]:
        """
        Получает популярные запросы на основе истории.
        
        Args:
            query_history: История запросов (если None, возвращает дефолтные)
            limit: Количество популярных запросов
            
        Returns:
            Список популярных запросов
        """
        if query_history:
            # Подсчитываем частоту запросов
            query_counter = Counter(query_history)
            popular = [query for query, count in query_counter.most_common(limit)]
            return popular
        
        # Дефолтные популярные запросы
        return [
            "Топ заказчиков",
            "Статистика по тендерам",
            "Топ менеджеров",
            "Самый большой тендер",
            "Динамика тендеров"
        ]
    
    @staticmethod
    def generate_autocomplete_suggestions(
        partial_query: str,
        available_options: List[str],
        limit: int = 5
    ) -> List[str]:
        """
        Генерирует предложения автодополнения на основе частичного запроса.
        
        Args:
            partial_query: Частичный запрос пользователя
            available_options: Доступные варианты для автодополнения
            limit: Максимальное количество предложений
            
        Returns:
            Список предложений автодополнения
        """
        if not partial_query:
            return available_options[:limit]
        
        partial_lower = partial_query.lower().strip()
        suggestions = []
        
        # Точные совпадения (приоритет)
        exact_matches = [opt for opt in available_options if opt.lower().startswith(partial_lower)]
        suggestions.extend(exact_matches[:limit])
        
        # Частичные совпадения
        if len(suggestions) < limit:
            partial_matches = [
                opt for opt in available_options
                if partial_lower in opt.lower() and opt not in suggestions
            ]
            suggestions.extend(partial_matches[:limit - len(suggestions)])
        
        return suggestions[:limit]
    
    @staticmethod
    def generate_demo_queries(query_type: str = 'all') -> List[str]:
        """
        Генерирует демонстрационные запросы для примера.
        
        Args:
            query_type: Тип запросов ('all', 'analytics', 'logistics', 'materials')
            
        Returns:
            Список демонстрационных запросов
        """
        demo_queries = {
            'analytics': [
                "Найди в 1с чертеж 35.1774.000 СБ",
                "Топ-10 заказчиков по количеству заказов",
                "Статистика по тендерам в категории Барабаны",
                "Самый большой выигранный тендер",
                "Динамика тендеров по месяцам",
                "Корреляция вес vs цена",
                "Heatmap заказчик менеджер"
            ],
            'logistics': [
                "Рассчитай логистику из КНР в Екатеринбург за груз весом 5000 кг",
                "Логистика в Москву, вес 2 тонны",
                "Расчет доставки в Санкт-Петербург, 1000 кг"
            ],
            'materials': [
                "Найди материал 10",
                "Какой GB у стали 20",
                "Свойства материала 09Г2С"
            ],
            'all': []
        }
        
        if query_type == 'all':
            all_queries = []
            for queries in demo_queries.values():
                all_queries.extend(queries)
            return all_queries
        
        return demo_queries.get(query_type, [])
    
    @staticmethod
    def format_autocomplete_response(
        suggestions: List[str],
        query_type: str = 'search'
    ) -> str:
        """
        Форматирует ответ с предложениями автодополнения.
        
        Args:
            suggestions: Список предложений
            query_type: Тип запроса
            
        Returns:
            Отформатированная строка
        """
        if not suggestions:
            return ""
        
        lines = ["**Возможные варианты:**"]
        for i, suggestion in enumerate(suggestions, 1):
            lines.append(f"{i}. {suggestion}")
        
        return "\n".join(lines)
    
    @staticmethod
    def generate_interactive_hints(
        context: Optional[Dict] = None
    ) -> List[str]:
        """
        Генерирует интерактивные подсказки на основе контекста.
        
        Args:
            context: Контекст (категория, тип запроса, etc.)
            
        Returns:
            Список подсказок
        """
        hints = []
        
        if context:
            category = context.get('category')
            query_type = context.get('query_type', 'general')
            
            if category:
                hints.append(f"💡 Попробуйте поиск в категории '{category}'")
            
            if query_type == 'analytics':
                hints.append("💡 Можно добавить 'с графиком' для визуализации")
                hints.append("💡 Используйте 'экспорт в CSV' для сохранения результатов")
            elif query_type == 'logistics':
                hints.append("💡 Укажите вес и город для точного расчета")
                hints.append("💡 Можно рассчитать логистику для записей из 1С")
        
        # Общие подсказки
        if not hints:
            hints.extend([
                "💡 Используйте конкретные запросы для лучших результатов",
                "💡 Можно комбинировать несколько условий",
                "💡 Добавьте 'покажи график' для визуализации данных"
            ])
        
        return hints

