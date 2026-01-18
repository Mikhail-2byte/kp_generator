"""
Модуль для создания интерактивных элементов в ответах агента.
Поддерживает кнопки-действия, структурированные ответы, ссылки.
"""

import logging
from typing import Any, Dict, List, Optional

logger = logging.getLogger(__name__)


class InteractiveElements:
    """Класс для создания интерактивных элементов в ответах."""
    
    @staticmethod
    def create_action_button(
        text: str,
        action: str,
        params: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """
        Создает кнопку-действие.
        
        Args:
            text: Текст кнопки
            action: Действие (например, 'export_csv', 'show_chart', 'filter')
            params: Параметры действия
            
        Returns:
            Словарь с данными кнопки
        """
        return {
            'type': 'button',
            'text': text,
            'action': action,
            'params': params or {}
        }
    
    @staticmethod
    def create_link(
        text: str,
        url: str,
        description: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Создает ссылку.
        
        Args:
            text: Текст ссылки
            url: URL
            description: Описание (опционально)
            
        Returns:
            Словарь с данными ссылки
        """
        return {
            'type': 'link',
            'text': text,
            'url': url,
            'description': description
        }
    
    @staticmethod
    def create_structured_response(
        title: str,
        sections: List[Dict[str, Any]],
        actions: Optional[List[Dict[str, Any]]] = None
    ) -> Dict[str, Any]:
        """
        Создает структурированный ответ с разделами.
        
        Args:
            title: Заголовок ответа
            sections: Список разделов (каждый раздел - словарь с 'title' и 'content')
            actions: Список действий (кнопки, ссылки)
            
        Returns:
            Словарь со структурированным ответом
        """
        return {
            'type': 'structured',
            'title': title,
            'sections': sections,
            'actions': actions or []
        }
    
    @staticmethod
    def format_buttons_markdown(buttons: List[Dict[str, Any]]) -> str:
        """
        Форматирует кнопки в markdown для отображения.
        
        Args:
            buttons: Список кнопок
            
        Returns:
            Markdown строка с кнопками
        """
        if not buttons:
            return ""
        
        lines = ["\n**Действия:**"]
        for i, button in enumerate(buttons, 1):
            action_text = button.get('text', f'Действие {i}')
            lines.append(f"{i}. [{action_text}](#action:{button.get('action', '')})")
        
        return "\n".join(lines)
    
    @staticmethod
    def format_links_markdown(links: List[Dict[str, Any]]) -> str:
        """
        Форматирует ссылки в markdown.
        
        Args:
            links: Список ссылок
            
        Returns:
            Markdown строка со ссылками
        """
        if not links:
            return ""
        
        lines = []
        for link in links:
            text = link.get('text', 'Ссылка')
            url = link.get('url', '#')
            description = link.get('description')
            
            if description:
                lines.append(f"- [{text}]({url}) - {description}")
            else:
                lines.append(f"- [{text}]({url})")
        
        return "\n".join(lines)
    
    @staticmethod
    def format_structured_response_markdown(
        structured_response: Dict[str, Any]
    ) -> str:
        """
        Форматирует структурированный ответ в markdown.
        
        Args:
            structured_response: Структурированный ответ
            
        Returns:
            Markdown строка
        """
        lines = []
        
        title = structured_response.get('title', '')
        if title:
            lines.append(f"## {title}\n")
        
        sections = structured_response.get('sections', [])
        for section in sections:
            section_title = section.get('title', '')
            section_content = section.get('content', '')
            
            if section_title:
                lines.append(f"### {section_title}")
            if section_content:
                lines.append(section_content)
            lines.append("")
        
        actions = structured_response.get('actions', [])
        if actions:
            buttons = [a for a in actions if a.get('type') == 'button']
            links = [a for a in actions if a.get('type') == 'link']
            
            if buttons:
                lines.append(InteractiveElements.format_buttons_markdown(buttons))
            if links:
                lines.append(InteractiveElements.format_links_markdown(links))
        
        return "\n".join(lines)
    
    @staticmethod
    def create_export_actions(
        export_type: str = 'csv',
        filename: Optional[str] = None
    ) -> List[Dict[str, Any]]:
        """
        Создает кнопки для экспорта результатов.
        
        Args:
            export_type: Тип экспорта ('csv', 'excel', 'png')
            filename: Имя файла (опционально)
            
        Returns:
            Список кнопок экспорта
        """
        actions = []
        
        if export_type in ['csv', 'all']:
            actions.append(InteractiveElements.create_action_button(
                '📥 Экспорт в CSV',
                'export_csv',
                {'filename': filename} if filename else {}
            ))
        
        if export_type in ['excel', 'all']:
            actions.append(InteractiveElements.create_action_button(
                '📊 Экспорт в Excel',
                'export_excel',
                {'filename': filename} if filename else {}
            ))
        
        if export_type in ['png', 'all']:
            actions.append(InteractiveElements.create_action_button(
                '🖼️ Экспорт графиков',
                'export_png',
                {'filename': filename} if filename else {}
            ))
        
        return actions
    
    @staticmethod
    def create_chart_actions(
        chart_types: List[str],
        category: Optional[str] = None
    ) -> List[Dict[str, Any]]:
        """
        Создает кнопки для генерации различных типов графиков.
        
        Args:
            chart_types: Список типов графиков ('customers', 'materials', 'timeline', etc.)
            category: Категория (опционально)
            
        Returns:
            Список кнопок графиков
        """
        actions = []
        
        chart_labels = {
            'customers': '📊 График заказчиков',
            'materials': '📊 График материалов',
            'managers': '📊 График менеджеров',
            'timeline': '📈 Временной ряд',
            'scatter': '📈 Корреляция',
            'heatmap': '🔥 Heatmap'
        }
        
        for chart_type in chart_types:
            label = chart_labels.get(chart_type, f'График {chart_type}')
            actions.append(InteractiveElements.create_action_button(
                label,
                'show_chart',
                {'chart_type': chart_type, 'category': category} if category else {'chart_type': chart_type}
            ))
        
        return actions
    
    @staticmethod
    def create_filter_actions(
        available_filters: List[str]
    ) -> List[Dict[str, Any]]:
        """
        Создает кнопки для применения фильтров.
        
        Args:
            available_filters: Список доступных фильтров
            
        Returns:
            Список кнопок фильтров
        """
        actions = []
        
        filter_labels = {
            'date': '📅 Фильтр по датам',
            'price': '💰 Фильтр по цене',
            'weight': '⚖️ Фильтр по весу',
            'customer': '👤 Фильтр по заказчику',
            'material': '🔩 Фильтр по материалу'
        }
        
        for filter_name in available_filters:
            label = filter_labels.get(filter_name, f'Фильтр {filter_name}')
            actions.append(InteractiveElements.create_action_button(
                label,
                'apply_filter',
                {'filter_type': filter_name}
            ))
        
        return actions

