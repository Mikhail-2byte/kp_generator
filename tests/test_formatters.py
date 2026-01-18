"""
Unit-тесты для модуля форматирования.
"""

import pytest
from ai_agent.formatters import ResponseFormatter


class TestResponseFormatter:
    """Тесты для ResponseFormatter."""
    
    def test_format_price(self):
        """Тест форматирования цены."""
        assert ResponseFormatter.format_price(1234567.89) == "1 234 567.89"
        assert ResponseFormatter.format_price(1000) == "1 000.00"
        assert ResponseFormatter.format_price(0) == "0.00"
    
    def test_format_weight(self):
        """Тест форматирования веса."""
        assert ResponseFormatter.format_weight(1234.56) == "1 234.56"
        assert ResponseFormatter.format_weight(100) == "100.00"
        assert ResponseFormatter.format_weight(0) == "0.00"
    
    def test_format_count(self):
        """Тест форматирования количества."""
        assert ResponseFormatter.format_count(12345) == "12 345"
        assert ResponseFormatter.format_count(0) == "0"
        assert ResponseFormatter.format_count(1) == "1"
    
    def test_format_table(self):
        """Тест форматирования таблицы."""
        headers = ['Название', 'Количество', 'Цена']
        rows = [
            ['Товар 1', 10, 1000.50],
            ['Товар 2', 25, 2500.75]
        ]
        result = ResponseFormatter.format_table(headers, rows)
        assert 'Название' in result
        assert 'Товар 1' in result
        assert 'Товар 2' in result
    
    def test_format_list(self):
        """Тест форматирования списка."""
        items = ['Первый', 'Второй', 'Третий']
        
        # Нумерованный список
        numbered = ResponseFormatter.format_list(items, numbered=True)
        assert '1.' in numbered
        assert 'Первый' in numbered
        
        # Маркированный список
        bulleted = ResponseFormatter.format_list(items, numbered=False)
        assert 'Первый' in bulleted

