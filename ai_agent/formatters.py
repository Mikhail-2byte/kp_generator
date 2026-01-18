"""
Модуль для единообразного форматирования ответов агента.
Предоставляет функции для форматирования чисел, таблиц, списков и других элементов.
"""

import logging
from typing import Any, Dict, List, Optional, Union

logger = logging.getLogger(__name__)


class ResponseFormatter:
    """Класс для форматирования ответов агента."""
    
    # Настройки форматирования
    THOUSANDS_SEPARATOR = ' '  # Разделитель тысяч
    DECIMAL_SEPARATOR = '.'  # Десятичный разделитель
    DECIMAL_PLACES_PRICE = 2  # Количество знаков после запятой для цен
    DECIMAL_PLACES_WEIGHT = 2  # Количество знаков после запятой для веса
    DECIMAL_PLACES_COUNT = 0  # Количество знаков после запятой для количества
    
    @staticmethod
    def format_number(
        value: Union[int, float],
        decimal_places: int = 2,
        thousands_separator: Optional[str] = None,
        decimal_separator: Optional[str] = None
    ) -> str:
        """
        Форматирует число с разделителями тысяч и округлением.
        
        Args:
            value: Число для форматирования
            decimal_places: Количество знаков после запятой
            thousands_separator: Разделитель тысяч (по умолчанию из настроек)
            decimal_separator: Десятичный разделитель (по умолчанию из настроек)
            
        Returns:
            Отформатированная строка
        """
        if value is None:
            return ''
        
        if not isinstance(value, (int, float)):
            try:
                value = float(value)
            except (ValueError, TypeError):
                return str(value)
        
        thousands_sep = thousands_separator or ResponseFormatter.THOUSANDS_SEPARATOR
        decimal_sep = decimal_separator or ResponseFormatter.DECIMAL_SEPARATOR
        
        # Округляем
        rounded = round(value, decimal_places)
        
        # Форматируем с разделителями
        if decimal_places == 0:
            # Целое число
            formatted = f"{int(rounded):,}".replace(',', thousands_sep)
        else:
            # С дробной частью
            formatted = f"{rounded:,.{decimal_places}f}".replace(',', thousands_sep)
            if decimal_sep != '.':
                formatted = formatted.replace('.', decimal_sep)
        
        return formatted
    
    @staticmethod
    def format_price(value: Union[int, float]) -> str:
        """
        Форматирует цену в рублях.
        
        Args:
            value: Цена для форматирования
            
        Returns:
            Отформатированная строка с ценой
        """
        return ResponseFormatter.format_number(
            value,
            decimal_places=ResponseFormatter.DECIMAL_PLACES_PRICE
        )
    
    @staticmethod
    def format_weight(value: Union[int, float]) -> str:
        """
        Форматирует вес в килограммах.
        
        Args:
            value: Вес для форматирования
            
        Returns:
            Отформатированная строка с весом
        """
        return ResponseFormatter.format_number(
            value,
            decimal_places=ResponseFormatter.DECIMAL_PLACES_WEIGHT
        )
    
    @staticmethod
    def format_count(value: Union[int, float]) -> str:
        """
        Форматирует количество (целое число).
        
        Args:
            value: Количество для форматирования
            
        Returns:
            Отформатированная строка с количеством
        """
        return ResponseFormatter.format_number(
            value,
            decimal_places=ResponseFormatter.DECIMAL_PLACES_COUNT
        )
    
    @staticmethod
    def format_percentage(value: Union[int, float], decimal_places: int = 1) -> str:
        """
        Форматирует процентное значение.
        
        Args:
            value: Процент для форматирования
            decimal_places: Количество знаков после запятой
            
        Returns:
            Отформатированная строка с процентом
        """
        formatted = ResponseFormatter.format_number(value, decimal_places)
        return f"{formatted}%"
    
    @staticmethod
    def format_table(
        headers: List[str],
        rows: List[List[Any]],
        align: Optional[List[str]] = None,
        separator: str = " | "
    ) -> str:
        """
        Форматирует таблицу в текстовом виде.
        
        Args:
            headers: Заголовки столбцов
            rows: Строки данных
            align: Выравнивание для каждого столбца ('left', 'right', 'center')
            separator: Разделитель между столбцами
            
        Returns:
            Отформатированная таблица
        """
        if not headers or not rows:
            return ""
        
        # Определяем ширину столбцов
        col_widths = [len(str(header)) for header in headers]
        for row in rows:
            for i, cell in enumerate(row):
                if i < len(col_widths):
                    col_widths[i] = max(col_widths[i], len(str(cell)))
        
        # Формируем строку заголовка
        header_parts = []
        for i, header in enumerate(headers):
            width = col_widths[i]
            align_type = (align[i] if align and i < len(align) else 'left').lower()
            if align_type == 'right':
                header_parts.append(str(header).rjust(width))
            elif align_type == 'center':
                header_parts.append(str(header).center(width))
            else:
                header_parts.append(str(header).ljust(width))
        
        header_line = separator.join(header_parts)
        
        # Формируем разделитель
        separator_line = separator.join(['-' * width for width in col_widths])
        
        # Формируем строки данных
        data_lines = []
        for row in rows:
            row_parts = []
            for i, cell in enumerate(row):
                if i < len(col_widths):
                    width = col_widths[i]
                    align_type = (align[i] if align and i < len(align) else 'left').lower()
                    cell_str = str(cell) if cell is not None else ''
                    if align_type == 'right':
                        row_parts.append(cell_str.rjust(width))
                    elif align_type == 'center':
                        row_parts.append(cell_str.center(width))
                    else:
                        row_parts.append(cell_str.ljust(width))
            data_lines.append(separator.join(row_parts))
        
        # Объединяем все части
        result = [header_line, separator_line] + data_lines
        return "\n".join(result)
    
    @staticmethod
    def format_list(
        items: List[Any],
        numbered: bool = True,
        prefix: str = "",
        start_number: int = 1
    ) -> str:
        """
        Форматирует список элементов.
        
        Args:
            items: Список элементов
            numbered: Использовать нумерацию (True) или буллеты (False)
            prefix: Префикс для каждого элемента
            start_number: Начальный номер для нумерации
            
        Returns:
            Отформатированный список
        """
        if not items:
            return ""
        
        lines = []
        for i, item in enumerate(items, start=start_number):
            if numbered:
                marker = f"{i}."
            else:
                marker = "•"
            
            item_str = str(item) if item is not None else ''
            lines.append(f"{prefix}{marker} {item_str}")
        
        return "\n".join(lines)
    
    @staticmethod
    def format_section(title: str, content: str, level: int = 1) -> str:
        """
        Форматирует секцию с заголовком.
        
        Args:
            title: Заголовок секции
            content: Содержимое секции
            level: Уровень заголовка (1-6, соответствует markdown)
            
        Returns:
            Отформатированная секция
        """
        if level < 1:
            level = 1
        elif level > 6:
            level = 6
        
        header_prefix = "#" * level
        return f"{header_prefix} {title}\n\n{content}"
    
    @staticmethod
    def format_separator(char: str = "-", length: int = 50) -> str:
        """
        Форматирует разделитель.
        
        Args:
            char: Символ для разделителя
            length: Длина разделителя
            
        Returns:
            Строка-разделитель
        """
        return char * length
    
    @staticmethod
    def format_progress_bar(
        current: Union[int, float],
        total: Union[int, float],
        length: int = 20,
        filled_char: str = "█",
        empty_char: str = "░"
    ) -> str:
        """
        Форматирует текстовый прогресс-бар.
        
        Args:
            current: Текущее значение
            total: Максимальное значение
            length: Длина прогресс-бара в символах
            filled_char: Символ для заполненной части
            empty_char: Символ для пустой части
            
        Returns:
            Отформатированный прогресс-бар
        """
        if total == 0:
            return empty_char * length
        
        percentage = min(100, max(0, (current / total) * 100))
        filled_length = int((percentage / 100) * length)
        empty_length = length - filled_length
        
        bar = filled_char * filled_length + empty_char * empty_length
        return f"[{bar}] {percentage:.1f}%"
    
    @staticmethod
    def format_record(
        record: Dict[str, Any],
        fields: Optional[List[str]] = None,
        field_labels: Optional[Dict[str, str]] = None
    ) -> str:
        """
        Форматирует запись (словарь) в читаемый вид.
        
        Args:
            record: Словарь с данными записи
            fields: Список полей для отображения (если None, все поля)
            field_labels: Словарь переименования полей (старое_имя -> новое_имя)
            
        Returns:
            Отформатированная запись
        """
        if not record:
            return ""
        
        fields_to_show = fields or list(record.keys())
        field_labels = field_labels or {}
        
        lines = []
        for field in fields_to_show:
            if field not in record:
                continue
            
            label = field_labels.get(field, field)
            value = record[field]
            
            # Форматируем значение в зависимости от типа
            if isinstance(value, (int, float)):
                # Определяем тип форматирования по названию поля
                if 'цена' in field.lower() or 'price' in field.lower() or 'стоимость' in field.lower():
                    formatted_value = ResponseFormatter.format_price(value)
                elif 'вес' in field.lower() or 'weight' in field.lower():
                    formatted_value = ResponseFormatter.format_weight(value)
                elif 'количество' in field.lower() or 'count' in field.lower():
                    formatted_value = ResponseFormatter.format_count(value)
                else:
                    formatted_value = ResponseFormatter.format_number(value)
            else:
                formatted_value = str(value) if value is not None else ''
            
            lines.append(f"**{label}**: {formatted_value}")
        
        return "\n".join(lines)
    
    @staticmethod
    def add_separator(text: str, separator: str = "\n---\n") -> str:
        """
        Добавляет разделитель к тексту.
        
        Args:
            text: Исходный текст
            separator: Разделитель
            
        Returns:
            Текст с разделителем
        """
        if not text:
            return ""
        return f"{text}{separator}"
    
    @staticmethod
    def format_summary(
        total: int,
        shown: int,
        label: str = "записей"
    ) -> str:
        """
        Форматирует сводку о количестве результатов.
        
        Args:
            total: Общее количество
            shown: Показано количество
            label: Название единицы измерения
            
        Returns:
            Отформатированная сводка
        """
        if total == shown:
            return f"Показаны все {total} {label}"
        else:
            return f"Найдено {total} {label}, показано {shown}"

